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

# Initialize Streamlit App
st.title("Automated Email Outreach Campaign")
st.write("Manage your email outreach campaigns easily.")

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

        if 'email' not in contacts.columns:
            st.error("The file must include an 'email' column!")
            st.stop()

    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")

# Step 2: Enter Mailgun API Details
st.header("2. Enter Mailgun API Details")
mailgun_domain = st.text_input("Mailgun Domain", "Your Mailgun domain here")
mailgun_api_key = st.text_input("Mailgun API Key", type="password")

if st.button("Verify Mailgun Credentials"):
    try:
        # Send a test email to verify credentials
        test_email = "your_verified_email@example.com"  # Replace with a verified recipient email
        test_response = requests.post(
            f"https://api.mailgun.net/v3/{mailgun_domain}/messages",
            auth=("api", mailgun_api_key),
            data={
                "from": f"Test User <mailgun@{mailgun_domain}>",
                "to": test_email,
                "subject": "Test Email",
                "text": "This is a test email to verify Mailgun credentials."
            }
        )
        if test_response.status_code == 200:
            st.success("Mailgun credentials verified successfully! Test email sent.")
        else:
            st.error(f"Failed to verify credentials. API response: {test_response.text}")
    except Exception as e:
        st.error(f"Error verifying credentials: {str(e)}")

# Step 3: Create Email Template
st.header("3. Create Email Template")
subject = st.text_input("Email Subject:", "Your Subject Here")
email_body = st.text_area("Email Body:", "Hi {name},\n\nThis is a personalized email template.")

st.write("You can use placeholders like `{name}`, `{company}` for personalization.")

# Step 4: Schedule Email Campaign
st.header("4. Schedule Email Campaign")
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