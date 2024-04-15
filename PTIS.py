import streamlit as st
import pandas as pd
import os
from PyPDF2 import PdfReader, PdfWriter # type: ignore
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import fitz  # type: ignore # PyMuPDF library for rendering PDF pages
from PIL import Image
from io import BytesIO

# Function to extract employee code from PDF
def extract_employee_code(page_text):
    # Extract the first number encountered in the page text as the employee code
    match = re.search(r'\b(\d+)\b', page_text)
    if match:
        return match.group(1)
    else:
        return None

# Function to split PDF into individual files based on employee codes
def split_pdf(pdf_path):
    with open(pdf_path, "rb") as pdf_file:
        pdf_reader = PdfReader(pdf_file)
        for page_num, page in enumerate(pdf_reader.pages):
            page_text = page.extract_text()
            employee_code = extract_employee_code(page_text)

            if employee_code:
                os.makedirs('employees_invoice', exist_ok=True)
                output_file_name = f"employees_invoice/{employee_code}.pdf"
                pdf_writer = PdfWriter()
                pdf_writer.add_page(page)
                with open(output_file_name, 'wb') as output_file:
                    pdf_writer.write(output_file)

# Function to read email addresses and employee information from CSV
def read_file_names(directory):
    file_names = os.listdir(directory)
    all_files = {}

    for file_name in file_names:
        path = os.path.join(directory, file_name)
        employee_id = file_name.split('.')[-2]
        all_files[employee_id] = path

    return all_files

# Function to send email with attached PDF payslip
def sent_email(receiver, app_password, sender_email, pdf_file_path, employee_name, salary):
    message = MIMEMultipart() # initiate class
    message['From'] = sender_email # sender email
    message['To'] = receiver # receiver email
    message['Subject'] = 'Monthly Salary Payslip' # E-mail subject

    # Generate email body with the name of the person
    body = f"Dear {employee_name},\n\nPlease find attached your monthly salary payslip. Your net salary for this month is PKR: {salary}.\n\nIf you have any question or concerns regarding your salary or deductions, please don't hesitate to reach out our HR department.\n\nThank you for your hard work and dedication\n\nBest regards,\nPTIS\nHR Department"

    message.attach(MIMEText(body, 'plain')) # attach body

    # Attach PDF file
    with open(pdf_file_path, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())

    encoders.encode_base64(part)
    part.add_header(
        'Content-Disposition',
        f'attachment; filename= {pdf_file_path.split("/")[-1]}',
    )
    message.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587) # start server
    server.starttls()
    server.login(sender_email, app_password)
    server.sendmail(sender_email, receiver, message.as_string()) # send email
    server.quit()

# Define Streamlit UI
def main():
    st.title("PTIS Invoice Sending Application")
    st.sidebar.title("Enter Credentials")

    # Get sender email and app password from user input
    sender_email = st.sidebar.text_input("Enter Sender Email")
    app_password = st.sidebar.text_input("Enter App Password", type="password")

    # File uploader for CSV and PDF files
    csv_file = st.sidebar.file_uploader("Upload CSV file", type=["csv"])
    pdf_file = st.sidebar.file_uploader("Upload PDF file", type=["pdf"])

    if csv_file is not None:
        df = pd.read_csv(csv_file)

    if pdf_file is not None and st.sidebar.button("Send"):
        pdf_images = read_pdf(pdf_file)
        if sender_email and app_password:
            st.subheader("CSV File Contents")
            st.write(df)

            # Save uploaded files to a temporary location
            with open("temp_pdf.pdf", "wb") as f:
                f.write(pdf_file.getvalue())

            # Perform PDF splitting and email sending
            sent_pdf(df, "temp_pdf.pdf", sender_email, app_password)

            # Remove temporary file
            os.remove("temp_pdf.pdf")

        # Display PDF file contents
            st.subheader("PDF File Contents")
            for page_num, image_bytes in enumerate(pdf_images, start=1):
                st.image(Image.open(BytesIO(image_bytes)), caption=f"Page {page_num}", use_column_width=True)
    

# Function to read CSV file
def read_csv(file):
    df = pd.read_csv(file)
    return df

def read_pdf(file):
    images = []
    file_path = os.path.join('E:\\Projects\\Payslip splitting', file.name)  # Adjust the file path accordingly
    with fitz.open(file_path) as pdf_file:  # Open the file using the adjusted path
        for page_num in range(len(pdf_file)):
            page = pdf_file.load_page(page_num)
            # Render page as an image
            image_bytes = page.get_pixmap().tobytes()
            images.append(image_bytes)
    return images

# Function to open and display the contents of a PDF file
def display_pdf_content(pdf_file):
    st.title("Uploaded PDF Content")
    if pdf_file is not None:
        pdf_display = open_pdf(pdf_file)
        st.write(pdf_display)

# Function to open and read the contents of a PDF file
def open_pdf(pdf_file):
    reader = PdfReader(pdf_file)
    pdf_text = ""
    for page in reader.pages:
        pdf_text += page.extract_text()
    return pdf_text

def extract_salary(pdf_path):
    salaries = []
    with fitz.open(pdf_path) as pdf_file:
        for page_num in range(len(pdf_file)):
            page = pdf_file.load_page(page_num)
            text = page.get_text()

            # Define a regex pattern to match salary information
            pattern = r'NET AMOUNT PAYABLE\s*:\s*(\d[\d,]*)'  # Match "NET AMOUNT PAYABLE :" followed by a numeric value

            # Search for the pattern in the text
            match = re.search(pattern, text)

            if match:
                # Extract the salary amount from the matched group
                salary_str = match.group(1).replace(',', '')  # Remove commas from the salary amount
                salary = int(salary_str)
                salaries.append(salary)

    return salary

# Function to send emails to employees
def sent_pdf(df, pdf_file, sender_email, app_password):
    split_pdf(pdf_file)
    employee_data = read_file_names('employees_invoice')
    
    for employee_id, pdf_path in employee_data.items():
        receiver_email_row = df[df['employee_id'] == int(employee_id)]
        
        if receiver_email_row.empty:
            st.write(f"Email for employee with ID {employee_id} not found.")
            continue
        
        receiver_email = receiver_email_row['email'].iloc[0]
        employee_name = receiver_email_row['Name'].iloc[0]
        
        if pd.isna(receiver_email):
            st.write(f"Email address not found for employee with ID {employee_id}. Skipping.")
            continue
        
        salary = extract_salary(pdf_path)  # Extract salary information
        sent_email(receiver_email, app_password, sender_email, pdf_path, employee_name, salary)
        
        st.write(f'Sent {pdf_path} to {receiver_email}')

if __name__ == "__main__":
    main()

