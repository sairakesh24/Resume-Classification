#!/usr/bin/env python
# coding: utf-8

# In[2]:


import os
import pandas as pd
import pickle
from pypdf import PdfReader
import re
import streamlit as st
from docx import Document
from io import BytesIO  # Import BytesIO to handle file streams
import win32com.client  # For handling .doc files on Windows
import tempfile  # To create temporary files for the uploaded .doc files
import pythoncom  # For initializing COM

# Load models
word_vector = pickle.load(open("tfidf.pkl", "rb"))
model = pickle.load(open("model.pkl", "rb"))

def cleanResume(txt):
    cleanText = re.sub('http\S+\s', ' ', txt)
    cleanText = re.sub('RT|cc', ' ', cleanText)
    cleanText = re.sub('#\S+\s', ' ', cleanText)
    cleanText = re.sub('@\S+', '  ', cleanText)  
    cleanText = re.sub('[%s]' % re.escape("""!"#$%&'()*+,-./:;<=>?@[\]^_`{|}~"""), ' ', cleanText)
    cleanText = re.sub(r'[^\x00-\x7f]', ' ', cleanText) 
    cleanText = re.sub('\s+', ' ', cleanText)
    return cleanText

category_mapping = {
    1: "Peoplesoft Resume",
    2: "React Developer",
    3: "SQL Developer",
    4: "Workday",
}

def extract_text_from_docx(uploaded_file):
    # Use BytesIO to handle the uploaded file stream correctly
    docx_file = BytesIO(uploaded_file.read())
    doc = Document(docx_file)
    text = "\n".join([para.text for para in doc.paragraphs])
    return text

def extract_text_from_doc(uploaded_file):
    # Initialize COM
    pythoncom.CoInitialize()
    
    # Save the uploaded .doc file to a temporary directory
    with tempfile.NamedTemporaryFile(delete=False, suffix=".doc") as temp_file:
        temp_file.write(uploaded_file.getbuffer())  # Save the uploaded file content
        temp_file_path = temp_file.name

    # Using win32com to extract text from .doc files on Windows
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(temp_file_path)  # Open the saved .doc file
    text = doc.Content.Text
    doc.Close()
    word.Quit()
    
    # Remove the temporary file after processing
    os.remove(temp_file_path)
    
    # Uninitialize COM to clean up
    pythoncom.CoUninitialize()
    
    return text

def categorize_resumes(uploaded_files, output_directory):
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    
    results = []
    
    for uploaded_file in uploaded_files:
        text = ""

        if uploaded_file.name.endswith('.pdf'):  # Handle PDF files
            reader = PdfReader(uploaded_file)
            page = reader.pages[0]
            text = page.extract_text()

        elif uploaded_file.name.endswith('.docx'):  # Handle DOCX files
            text = extract_text_from_docx(uploaded_file)
        
        elif uploaded_file.name.endswith('.doc'):  # Handle DOC files
            text = extract_text_from_doc(uploaded_file)

        # If the file is neither PDF, DOC, nor DOCX, skip processing
        if not text:
            st.warning(f"Skipping {uploaded_file.name}. Unsupported file format.")
            continue

        cleaned_resume = cleanResume(text)

        input_features = word_vector.transform([cleaned_resume])
        prediction_id = model.predict(input_features)[0]
        category_name = category_mapping.get(prediction_id, "Unknown")
        
        category_folder = os.path.join(output_directory, category_name)
        
        if not os.path.exists(category_folder):
            os.makedirs(category_folder)
        
        target_path = os.path.join(category_folder, uploaded_file.name)
        with open(target_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        results.append({'filename': uploaded_file.name, 'category': category_name})
    
    results_df = pd.DataFrame(results)
    return results_df

st.title("Resume Categorizer Application")
st.subheader("With Python & Machine Learning")

uploaded_files = st.file_uploader("Choose PDF, DOC, or DOCX files", type=["pdf", "doc", "docx"], accept_multiple_files=True)
output_directory = st.text_input("Output Directory", "categorized_resumes")

if st.button("Categorize Resumes"):
    if uploaded_files and output_directory:
        results_df = categorize_resumes(uploaded_files, output_directory)
        st.write(results_df)
        results_csv = results_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Download results as CSV",
            data=results_csv,
            file_name='categorized_resumes.csv',
            mime='text/csv',
        )
        st.success("Resumes categorization and processing completed.")
    else:
        st.error("Please upload files and specify the output directory.")


# In[ ]:




