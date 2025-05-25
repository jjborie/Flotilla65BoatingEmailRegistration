import streamlit as st
import pandas as pd
import mammoth
import docx
import base64
import os
import tempfile
import re
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail

# Configuration
OUTPUT_HTML_PATH = "confirmation.html"

# [Existing functions: extract_images_fallback, convert_docx_to_html, personalize_html, send_email_sendgrid remain unchanged]

def main():
    st.title("Mass Email Sender for Boating Course")

    # Sidebar for SendGrid configuration
    st.sidebar.header("SendGrid Configuration")
    sender_email = st.sidebar.text_input("Sender Email", value="jjborie@hotmail.com")
    # Remove sendgrid_api_key input and use secrets
    try:
        sendgrid_api_key = st.secrets["general"]["SENDGRID_API_KEY"]
    except KeyError:
        st.error("SendGrid API Key not found in Streamlit secrets. Please add it to secrets.toml or Streamlit Cloud secrets.")
        return
    st.sidebar.markdown("""
        **Note**: Ensure the sender email is verified in SendGrid at https://app.sendgrid.com/settings/sender_auth.
        The SendGrid API key is loaded from Streamlit secrets.
    """)

    # Main UI
    st.header("Upload Files")
    excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    docx_file = st.file_uploader(
        "Upload Word Document",
        type=["docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"]
    )

    email_subject = st.text_input("Email Subject", value="Boating Course Enrollment Confirmation - June 14, 2025")

    if st.button("Send Emails"):
        if not excel_file or not docx_file or not sender_email or not sendgrid_api_key:
            st.error("Please upload both files and provide a sender email.")
            return

        # [Rest of the main function remains unchanged]
        # Save uploaded files temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
            tmp_excel.write(excel_file.read())
            excel_path = tmp_excel.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
            tmp_docx.write(docx_file.read())
            docx_path = tmp_docx.name

        try:
            # Convert DOCX to HTML
            html_content, conversion_messages = convert_docx_to_html(docx_path, OUTPUT_HTML_PATH)
            if conversion_messages:
                st.warning("DOCX conversion warnings: " + str(conversion_messages))
            if html_content is None:
                st.error("Failed to generate HTML content.")
                return

            # Read Excel file
            try:
                df = pd.read_excel(excel_path, sheet_name="Sheet1")
            except Exception as e:
                st.error(f"Failed to read Excel file: {str(e)}")
                return

            # Send emails
            st.header("Email Sending Status")
            for index, row in df.iterrows():
                first_name = str(row["First Name"]).strip()
                last_name = str(row["Last Name"]).strip()
                email = str(row["Primary Student E-mail"]).strip()

                if not email or "@" not in email:
                    st.write(f"Skipping invalid email for {first_name} {last_name}: {email}")
                    continue

                personalized_html = personalize_html(html_content, first_name, last_name)
                status = send_email_sendgrid(
                    email, email_subject, personalized_html, sender_email, sendgrid_api_key
                )
                st.write(status)

        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
        finally:
            # Clean up temporary files
            try:
                os.unlink(excel_path)
                os.unlink(docx_path)
            except Exception as e:
                st.warning(f"Failed to clean up temporary files: {str(e)}")

if __name__ == "__main__":
    main()
    