import os
import streamlit as st
import PyPDF2
import io
import pandas as pd
import openai
from docx import Document
import json
import openpyxl
import base64

# Initialise OpenAI
openai.api_key = os.getenv('OPENAI_API_KEY')

#Setting the Page Title
st.set_page_config(page_title="Extract Anything", layout='wide')

#Code to hide the made with streamlit text
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

def read_file(file):
    if file.type == 'application/pdf':
        pdfReader = PyPDF2.PdfReader(file)
        text = " ".join([page.extract_text() for page in pdfReader.pages])
    elif file.type == 'application/msword':
        doc = Document(io.BytesIO(file.getvalue()))
        text = " ".join([para.text for para in doc.paragraphs])
    else:
        st.write("Unsupported format")
        text = ''
    return text


def extract_document(cv, inputs):
    print("THE INPUTS ARE", inputs)
    prompt = f"This is the document {cv}. Can you extract the relevant variables {', '.join(inputs)}"
    
    functions = []
    properties = {}

    for input in inputs:
        properties[input] = {
            "type": "string",
            "description" : f"Extracted {input} value",
        }

    functions.append(
        {
            "name": "extract_anything",
            "description": "Extracts any variable from text",
            "parameters": {
                "type": "object",
                "properties": properties
            }
        }
    )
    
    completion = openai.ChatCompletion.create(
    model="gpt-3.5-turbo-0613",
    messages=[{"role": "user", "content": prompt}],
    functions=functions,
    function_call= "auto",
    )
    
    arguments = json.loads(completion.choices[0].message.function_call["arguments"])
    print("THE ARGUMENTS ARE", arguments)
    return arguments

def main():
    st.title("Extract anything")
    st.text("Enter the names of the fields you would like to extract (separated by comma)")
    
    input_str = st.text_input("Enter fields to extract")
    
    inputs = [i.strip() for i in input_str.split(",")]
    
    multiple_files = st.file_uploader("Upload Files", type=['pdf', 'doc', 'docx'], accept_multiple_files=True)

    if st.button("Extract"):
        if multiple_files is not None:
            for file in multiple_files:
                file_text = read_file(file)
                arguments = extract_document(file_text, inputs)
                
                temp_dict = {input: arguments.get(input, 'N/A') for input in inputs}

                # Create a dataframe here
                df = pd.DataFrame(temp_dict, index=[0])
                st.table(df)

                # Convert DataFrame to Excel
                towrite = io.BytesIO()
                df.to_excel(towrite, index=False, engine='openpyxl')
                towrite.seek(0)  # reset pointer
                b64 = base64.b64encode(towrite.read()).decode()  # some strings
                linko= f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="myfilename.xlsx">Download excel file</a>'
                st.markdown(linko, unsafe_allow_html=True)

        else:
            st.write("Please upload the files")

if __name__ == '__main__':
    main()
