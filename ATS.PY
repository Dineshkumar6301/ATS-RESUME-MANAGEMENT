import streamlit as st
import mysql.connector
import pandas as pd
import re
import os
from docx import Document
import PyPDF2
import docx2txt
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import requests

# Google Drive setup
SERVICE_ACCOUNT_FILE = 'C:\\Users\\mariy\\PycharmProjects\\ATS\\client_secret.json'
SCOPES = ['https://www.googleapis.com/auth/drive.file']

# Authenticate and create the Google Drive API client
credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=credentials)

# MySQL database connection
def get_db_connection():
    return mysql.connector.connect(
        host='127.0.0.1',
        user='root',
        password='630186',  # Update with your DB password
        database='resume_db'
    )

# User authentication functions
def signup(username, password):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO users (username, password) VALUES (%s, %s)", (username, password))
        conn.commit()
        st.success('Account created successfully!')
    except Exception as e:
        st.error(f"Error creating account: {e}")
    finally:
        cursor.close()
        conn.close()

def login(username, password):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE username = %s AND password = %s", (username, password))
    user = cursor.fetchone()
    cursor.close()
    conn.close()
    return user

# Functions for resume processing and extraction
def input_pdf_text(uploaded_file):
    reader = PyPDF2.PdfReader(uploaded_file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() or ""
    return text.strip()

def input_docx_text(uploaded_file):
    return docx2txt.process(uploaded_file)

# Function to extract details from resume text
def extract_details_from_resume(text):
    # Match name (assumes the name is at the start of the document)
    name = re.search(r'^[A-Z][a-zA-Z\s]+', text)
    
    # Match phone numbers in various formats
    phone = re.search(r'\b(?:\+\d{1,2}\s?)?\d{10}\b', text)
    
    # Match email addresses
    email = re.search(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
    
    # Match job titles (extendable list of titles)
    job_title = re.search(r'(Software Engineer|Data Scientist|Project Manager|Software Developer|Systems Analyst|Engineer)', text)
    
    # Extract relevant skills (add more skills based on the resume content)
    skills = re.findall(r'\b(Python|NumPy|Pandas|Matplotlib|PowerBI|SQL|MySQL|Excel|Machine Learning|Data Analysis|Automation)\b', text)
    
    # Match location using a pattern with city, optional postal code
    location = re.search(r'Location:\s*([A-Za-z\s]+),?\s*(\d{6})?', text)

    # Clean and return extracted data
    return {
        'Name': name.group(0).strip() if name else 'Not Found',
        'Phone Number': phone.group(0).strip() if phone else 'Not Found',
        'Email ID': email.group(0).strip() if email else 'Not Found',
        'Job Title': job_title.group(0).strip() if job_title else 'Not Found',
        'Skills': ', '.join(set(skills)) if skills else 'Not Found',
        'Location': location.group(1).strip() if location else 'Not Found',
    }

# Function to upload file to Google Drive
def upload_file_to_google_drive(file_path, file_name):
    file_metadata = {
        'name': file_name,
        'mimeType': 'application/pdf' if file_name.endswith('.pdf') else 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    }
    media = MediaFileUpload(file_path, mimetype=file_metadata['mimeType'])
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    return f"https://drive.google.com/file/d/{file.get('id')}/view"

# Function to save data to Excel
def save_full_data_to_excel(data, file_name):
    df = pd.DataFrame(data)
    df.to_excel(file_name, index=False)

# Function to search resumes in the database
def search_resumes(search_query, filter_by):
    conn = get_db_connection()
    cursor = conn.cursor()
    if filter_by == 'Name':
        cursor.execute("SELECT * FROM resumes WHERE name LIKE %s", (f'%{search_query}%',))
    elif filter_by == 'Phone Number':
        cursor.execute("SELECT * FROM resumes WHERE phone_number LIKE %s", (f'%{search_query}%',))
    elif filter_by == 'Email ID':
        cursor.execute("SELECT * FROM resumes WHERE email_id LIKE %s", (f'%{search_query}%',))
    elif filter_by == 'Job Title':
        cursor.execute("SELECT * FROM resumes WHERE job_title LIKE %s", (f'%{search_query}%',))
    elif filter_by == 'Skills':
        skills = search_query.split(',')
        query = "SELECT * FROM resumes WHERE " + " OR ".join([f"skills LIKE %s" for _ in skills])
        cursor.execute(query, tuple([f'%{skill.strip()}%' for skill in skills]))

    results = cursor.fetchall()
    cursor.close()
    conn.close()
    return results

# Streamlit app layout
st.title('Resume Management System')

# User authentication state
if 'user' not in st.session_state:
    st.session_state.user = None

# Signup and Login forms
if st.session_state.user is None:
    option = st.selectbox('Choose an action', ['Sign Up', 'Log In'])

    if option == 'Sign Up':
        st.subheader('Create an Account')
        username = st.text_input('Username')
        password = st.text_input('Password', type='password')
        if st.button('Sign Up'):
            signup(username, password)

    elif option == 'Log In':
        st.subheader('Log In')
        username = st.text_input('Username')
        password = st.text_input('Password', type='password')
        if st.button('Log In'):
            user = login(username, password)
            if user:
                st.session_state.user = username
                st.success(f'Welcome {username}!')
            else:
                st.error('Invalid credentials')

# Resume upload section
if st.session_state.user is not None:
    st.subheader('Upload Resumes')
    upload_option = st.radio("Select upload option:", ["Single Resume Upload", "Bulk Resume Upload"])

    # Handle single resume upload
    if upload_option == "Single Resume Upload":
        uploaded_file = st.file_uploader("Choose a file (PDF or DOCX)", type=['pdf', 'docx'])
        if uploaded_file is not None:
            if uploaded_file.size > 5 * 1024 * 1024:  # 5 MB limit
                st.error("File size exceeds the 5MB limit!", icon="🚨")
            else:
                try:
                    text = ""
                    if uploaded_file.type == 'application/pdf':
                        text = input_pdf_text(uploaded_file)
                    elif uploaded_file.type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
                        text = input_docx_text(uploaded_file)

                    if text:
                        extracted_data = extract_details_from_resume(text)

                        # Save the uploaded file temporarily
                        temp_file_path = f'temp/{uploaded_file.name}'
                        os.makedirs(os.path.dirname(temp_file_path), exist_ok=True)
                        with open(temp_file_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())

                        # Upload to Google Drive
                        drive_url = upload_file_to_google_drive(temp_file_path, uploaded_file.name)
                        extracted_data['Resume URL'] = drive_url

                        # Display extracted data
                        st.dataframe(pd.DataFrame([extracted_data]))
                        st.success("Data extracted successfully.")

                        # Save extracted data to Excel
                        if st.button("Save Data to Excel"):
                            save_full_data_to_excel([extracted_data], "extracted_data.xlsx")
                            st.success("Data saved to Excel.")
                except Exception as e:
                    st.error(f"Error processing file: {e}")

    # Handle bulk resume upload
    if upload_option == "Bulk Resume Upload":
        st.subheader("Bulk Resume Upload from Folder Path")

        folder_path = st.text_input("Enter the folder path containing resumes")

        if st.button("Upload Resumes from Folder"):
            if os.path.exists(folder_path):
                extracted_data_list = []

                for file_name in os.listdir(folder_path):
                    file_path = os.path.join(folder_path, file_name)

                    # Skip temporary and hidden files
                    if file_name.startswith('~$') or file_name.startswith('.'):
                        st.warning(f"Skipping temporary/hidden file: {file_name}")
                        continue

                    # Process PDF files
                    if file_name.endswith('.pdf'):
                        with open(file_path, "rb") as f:
                            try:
                                text = input_pdf_text(f)
                                extracted_data = extract_details_from_resume(text)
                                extracted_data['Resume URL'] = upload_file_to_google_drive(file_path, file_name)
                                extracted_data_list.append(extracted_data)
                            except Exception as e:
                                st.error(f"Error processing PDF file {file_name}: {e}")

                    # Process DOCX files
                    elif file_name.endswith('.docx'):
                        try:
                            text = input_docx_text(file_path)
                            extracted_data = extract_details_from_resume(text)
                            extracted_data['Resume URL'] = upload_file_to_google_drive(file_path, file_name)
                            extracted_data_list.append(extracted_data)
                        except Exception as e:
                            st.error(f"Error processing DOCX file {file_name}: {e}")

                # Display all extracted data
                if extracted_data_list:
                    st.dataframe(pd.DataFrame(extracted_data_list))
                    st.success("Bulk data extracted successfully.")

                # Save to Excel
                if st.button("Save All Data to Excel"):
                    save_full_data_to_excel(extracted_data_list, "bulk_extracted_data.xlsx")
                    st.success("All data saved to Excel.")
            else:
                st.error("The specified folder path does not exist. Please check the path and try again.")

    # Search resumes
    st.subheader('Search Resumes')
    search_query = st.text_input('Enter your search query (e.g., skills, name, etc.)')
    filter_by = st.selectbox("Filter by:", ['Name', 'Phone Number', 'Email ID', 'Job Title', 'Skills'])

    if st.button('Search'):
        results = search_resumes(search_query, filter_by)
        if results:
            st.dataframe(pd.DataFrame(results, columns=['ID', 'Name', 'Phone Number', 'Email ID', 'Job Title', 'Skills', 'Location', 'Resume URL']))
        else:
            st.info("No results found.")
