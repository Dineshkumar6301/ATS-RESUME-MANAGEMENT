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

# MySQL database connection function
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
    name = re.search(r'^[A-Z][a-zA-Z\s]+', text)
    phone = re.search(r'\b(?:\+\d{1,2}\s?)?\d{10}\b', text)
    email = re.search(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
    job_title = re.search(r'(Software Engineer|Data Scientist|Project Manager|Software Developer|Systems Analyst|Engineer)', text)
    skills = re.findall(r'\b(Python|NumPy|Pandas|Matplotlib|PowerBI|SQL|MySQL|Excel|Machine Learning|Data Analysis|Automation)\b', text)
    location = re.search(r'Location:\s*([A-Za-z\s]+),?\s*(\d{6})?', text)

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

# Function to insert resume data into the database
def insert_resume_data(name, phone_number, email_id, job_title, skills, location, resume_url):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        query = """
            INSERT INTO resumes (name, phone_number, email_id, job_title, skills, location, resume_url)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        """
        values = (name, phone_number, email_id, job_title, skills, location, resume_url)
        cursor.execute(query, values)
        conn.commit()
        st.success("Data saved to the database successfully.")
    except Exception as e:
        st.error(f"Error saving data to database: {e}")
    finally:
        cursor.close()
        conn.close()

# Function to search resumes in the database
def search_resumes(search_query, filter_by):
    conn = get_db_connection()
    cursor = conn.cursor()

    # Construct the search query based on the selected filter
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
    
    # Convert results to a DataFrame for display
    columns = ['ID', 'Name', 'Phone Number', 'Email ID', 'Job Title', 'Skills', 'Location', 'Resume URL']
    return pd.DataFrame(results, columns=columns) if results else pd.DataFrame(columns=columns)

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
                        temp_file_path = f'temp/{uploaded_file.name}'
                        os.makedirs(os.path.dirname(temp_file_path), exist_ok=True)
                        with open(temp_file_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())

                        drive_url = upload_file_to_google_drive(temp_file_path, uploaded_file.name)
                        extracted_data['Resume URL'] = drive_url
                        
                        # Insert data into the database
                        insert_resume_data(
                            extracted_data['Name'],
                            extracted_data['Phone Number'],
                            extracted_data['Email ID'],
                            extracted_data['Job Title'],
                            extracted_data['Skills'],
                            extracted_data['Location'],
                            extracted_data['Resume URL']
                        )
                        
                        st.dataframe(pd.DataFrame([extracted_data]))
                        st.success("Data extracted successfully.")

                        if st.button("Save Data to Excel"):
                            save_full_data_to_excel([extracted_data], "extracted_data.xlsx")
                            st.success("Data saved to Excel.")
                except Exception as e:
                    st.error(f"Error processing file: {e}")

    elif upload_option == "Bulk Resume Upload":
        uploaded_files = st.file_uploader("Choose files (PDF or DOCX)", type=['pdf', 'docx'], accept_multiple_files=True)
        if uploaded_files:
            extracted_data_list = []
            for uploaded_file in uploaded_files:
                if uploaded_file.size > 5 * 1024 * 1024:  # 5 MB limit
                    st.error(f"{uploaded_file.name} exceeds the 5MB limit!", icon="🚨")
                else:
                    try:
                        text = ""
                        if uploaded_file.type == 'application/pdf':
                            text = input_pdf_text(uploaded_file)
                        elif uploaded_file.type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
                            text = input_docx_text(uploaded_file)

                        if text:
                            extracted_data = extract_details_from_resume(text)
                            temp_file_path = f'temp/{uploaded_file.name}'
                            os.makedirs(os.path.dirname(temp_file_path), exist_ok=True)
                            with open(temp_file_path, "wb") as f:
                                f.write(uploaded_file.getbuffer())

                            drive_url = upload_file_to_google_drive(temp_file_path, uploaded_file.name)
                            extracted_data['Resume URL'] = drive_url
                            
                            # Insert data into the database
                            insert_resume_data(
                                extracted_data['Name'],
                                extracted_data['Phone Number'],
                                extracted_data['Email ID'],
                                extracted_data['Job Title'],
                                extracted_data['Skills'],
                                extracted_data['Location'],
                                extracted_data['Resume URL']
                            )
                            
                            extracted_data_list.append(extracted_data)
                    except Exception as e:
                        st.error(f"Error processing {uploaded_file.name}: {e}")

            if extracted_data_list:
                st.dataframe(pd.DataFrame(extracted_data_list))
                if st.button("Save All Data to Excel"):
                    save_full_data_to_excel(extracted_data_list, "extracted_data.xlsx")
                    st.success("All data saved to Excel.")

    # Search section
    st.subheader('Search Resumes')
    search_query = st.text_input("Enter search query:")
    filter_by = st.selectbox("Filter by:", ['Name', 'Phone Number', 'Email ID', 'Job Title', 'Skills'])
    if st.button("Search"):
        if search_query:
            results = search_resumes(search_query, filter_by)
            if not results.empty:
                st.dataframe(results)
            else:
                st.warning("No results found.")
        else:
            st.warning("Please enter a search query.")
