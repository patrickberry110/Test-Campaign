import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
import time
import threading
import subprocess
import sys

# Function to install missing packages
def install_package(package):
    try:
        __import__(package)
    except ImportError:
        st.warning(f"Installing missing package: {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Ensure openpyxl is installed for Excel file handling
install_package("openpyxl")

# Enforce custom branding
st.set_page_config(
    page_title="EdgePoint Campaign Manager",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Inject custom CSS for branding
st.markdown("""
    <style>
    /* Body styling */
    body {
        font-family: 'Open Sans', sans-serif;
        background-color: #ffffff; /* White background */
        color: #003366; /* EdgePoint Blue for text */
    }
    /* Header and Titles */
    h1, h2, h3, h4, h5, h6 {
        color: #003366; /* EdgePoint Blue */
    }
    /* Buttons */
    .stButton>button {
        background-color: #FFC72C; /* Gold buttons */
        color: #003366; /* EdgePoint Blue text */
        border-radius: 12px;
        border: none;
        padding: 8px 16px;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #ffdf80; /* Lighter gold for hover effect */
        color: #003366;
    }
    /* Sidebar styling */
    .css-18e3th9 {
        background-color: #E6EFF5; /* Light blue background for sidebar */
    }
    /* File uploader */
    .stFileUploader {
        border-radius: 8px;
        border: 1px solid #cccccc; /* Light gray border */
        padding: 10px;
    }
    </style>
""", unsafe_allow_html=True)

# Add EdgePoint logo
st.image("https://www.edgepointcm.com/wp-content/uploads/2020/09/logo-blue.png", width=200)

# Application title
st.title("EdgePoint Automated Campaign Manager")
st.write("Experience. Empathy. Results. Automate your email campaigns with ease.")

# Step 1: Upload Contact List
st.header("1. Upload Contact List")
uploaded_file = st.file_uploader("Upload a CSV or Excel file with email contacts:", type=["csv", "xlsx"])

if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            contacts = pd.read_csv(uploaded_file)
        elif uploaded_file.name.endswith(".xlsx"):
            contacts = pd.read_excel(uploaded_file, engine="openpyxl")
        else:
            st.error("Unsupported file type! Please upload a .csv or .xlsx file.")
            st.stop()

        st.write("Preview of your contact list:")
        st.dataframe(contacts.head())

        # Ensure email column is case-insensitive
        email_column = None
        name_column = None
        for column in contacts.columns:
            if column.strip().lower() == "email":
                email_column = column
            if column.strip().lower() == "name":
                name_column = column

        if email_column is None:
            st.error("The file must include an 'email' column (case-insensitive)!")
            st.stop()
        else:
            contacts.rename(columns={email_column: "email"}, inplace=True)

        if name_column is not None:
            contacts.rename(columns={name_column: "name"}, inplace=True)

    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")

# Step 2: Upload Campaign Materials
st.header("2. Upload Campaign Materials")
uploaded_pdf = st.file_uploader("Upload Campaign Materials Here:", type=["pdf"])

if uploaded_pdf:
    st.write("Campaign Materials uploaded successfully!")

# Step 3: Enter Mailgun API Details
st.header("3. Enter Mailgun API Details")
mailgun_domain = st.text_input("Mailgun Domain", "Your Mailgun domain here")
mailgun_api_key = st.text_input("Mailgun API Key", type="password")

def send_test_email():
    test_email = "your_verified_email@example.com"  # Replace with your verified email
    response = requests.post(
        f"https://api.mailgun.net/v3/{mailgun_domain}/messages",
        auth=("api", mailgun_api_key),
        data={
            "from": f"Test User <mailgun@{mailgun_domain}>",
            "to": test_email,
            "subject": "Mailgun Test Email",
            "text": "This is a test email to verify Mailgun credentials."
        }
    )
    return response

if st.button("Verify Mailgun Credentials"):
    if not mailgun_domain or not mailgun_api_key:
        st.error("Please provide both the Mailgun domain and API key.")
    else:
        try:
            response = send_test_email()
            if response.status_code == 200:
                st.success("Mailgun credentials verified successfully! Test email sent.")
            else:
                st.error(f"Failed to verify credentials: {response.json().get('message', response.text)}")
        except Exception as e:
            st.error(f"Error verifying credentials: {str(e)}")

# Step 4: Create Email Template
st.header("4. Create Email Template")
subject = st.text_input("Email Subject:", "Your Subject Here")
email_body = st.text_area("Email Body:", "Hi {name},\n\nThis is a personalized email template.")

st.write("You can use placeholders like `{name}`, `{company}` for personalization.")

# Step 5: Schedule Email Campaign
st.header("5. Schedule Email Campaign")
schedule_time = st.time_input("Select the time to send the campaign:", value=(datetime.now() + timedelta(minutes=5)).time())
schedule_date = st.date_input("Select the date to send the campaign:", value=datetime.now().date())

# Combine date and time
schedule_datetime = datetime.combine(schedule_date, schedule_time)

if st.button("Schedule Campaign"):
    if not uploaded_file or contacts.empty:
        st.error("Please upload a contact list before scheduling.")
    elif not mailgun_domain or not mailgun_api_key:
        st.error("Please provide Mailgun API details before scheduling.")
    else:
        st.success(f"Campaign scheduled for {schedule_datetime}.")

        def send_emails():
            st.write("Sending emails...")
            try:
                for index, contact in contacts.iterrows():
                    recipient_email = contact['email']

                    # Replace placeholders in the template
                    personalized_body = email_body
                    for column in contacts.columns:
                        if f"{{{column}}}" in personalized_body:
                            personalized_body = personalized_body.replace(f"{{{column}}}", str(contact[column]))

                    # Send email via Mailgun
                    response = requests.post(
                        f"https://api.mailgun.net/v3/{mailgun_domain}/messages",
                        auth=("api", mailgun_api_key),
                        data={
                            "from": f"Your Name <mailgun@{mailgun_domain}>",
                            "to": recipient_email,
                            "subject": subject,
                            "text": personalized_body
                        }
                    )

                    if response.status_code == 200:
                        st.write(f"Email sent to {recipient_email}.")
                    else:
                        st.error(f"Failed to send email to {recipient_email}: {response.text}")

                    time.sleep(1)  # Avoid sending emails too quickly

                st.success("All emails sent successfully!")
            except Exception as e:
                st.error(f"Failed to send emails: {str(e)}")

        # Use threading to avoid blocking UI
        threading.Thread(target=send_emails).start()