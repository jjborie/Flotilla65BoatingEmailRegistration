import streamlit as st
import pandas as pd
import mammoth
import docx
import base64
import os
import tempfile
import re
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Configuration
OUTPUT_HTML_PATH = "confirmation.html"

def extract_images_fallback(docx_path):
    """Fallback image extraction using python-docx, prioritizing heading images."""
    try:
        doc = docx.Document(docx_path)
        image_data_urls = []
        image_positions = []
        paragraph_index = 0

        # Count paragraphs and check inline shapes for images
        for para in doc.paragraphs:
            paragraph_index += 1
        for shape in doc.inline_shapes:
            try:
                if shape._inline.graphic.graphicData.uri == "http://schemas.openxmlformats.org/drawingml/2006/picture":
                    blip = shape._inline.graphic.graphicData.pic.blipFill.blip
                    rel_id = blip.embed or blip.link
                    if rel_id is None:
                        st.warning(f"Fallback: Skipping image: No embed or link attribute found.")
                        continue
                    
                    image_rel = doc.part.rels.get(rel_id)
                    if image_rel is None:
                        st.warning(f"Fallback: Skipping image: No relationship found for rel_id {rel_id}.")
                        continue
                    
                    image_part = image_rel.target_part
                    image_data = image_part.blob
                    content_type = image_part.content_type
                    ext = "png" if content_type == "image/png" else "jpg" if content_type == "image/jpeg" else "png"
                    base64_data = base64.b64encode(image_data).decode("utf-8")
                    data_url = f"data:image/{ext};base64,{base64_data}"
                    # Assume first image is at start (heading)
                    image_data_urls.append((data_url, ext))
                    image_positions.append(0)  # Force heading position
                    st.info(f"Fallback: Embedded image at paragraph index 0 (forced heading), content type: {content_type}")
            except Exception as e:
                st.warning(f"Fallback: Failed to process image: {str(e)}")
                continue

        return image_data_urls, image_positions
    except Exception as e:
        st.warning(f"Fallback: Failed to process DOCX: {str(e)}")
        return [], []

def convert_docx_to_html(docx_path, output_path):
    """Convert DOCX to HTML, embedding images as Base64 at the correct position."""
    image_counter = 0
    image_data_urls = []
    
    def handle_image(image):
        nonlocal image_counter, image_data_urls
        try:
            image_counter += 1
            ext = "png" if image.content_type == "image/png" else "jpg" if image.content_type == "image/jpeg" else "png"
            # Use get_stream() for newer mammoth versions
            try:
                image_data = image.get_stream().read()
            except AttributeError:
                # Fallback to get_reader() for older mammoth versions
                image_data = image.get_reader().read()
            base64_data = base64.b64encode(image_data).decode("utf-8")
            data_url = f"data:image/{ext};base64,{base64_data}"
            image_data_urls.append((data_url, ext))
            st.info(f"Mammoth: Embedded image {image_counter}, content type: {image.content_type}")
            return {"src": data_url, "alt": f"Course Image {image_counter}", "style": "max-width:100%;height:auto;"}
        except Exception as e:
            st.warning(f"Mammoth: Failed to process image {image_counter}: {str(e)}")
            return {"src": ""}

    try:
        with open(docx_path, "rb") as f:
            result = mammoth.convert_to_html(f, convert_image=mammoth.images.img_element(handle_image))
            html_content = result.value
            messages = result.messages
    except Exception as e:
        st.error(f"Failed to convert DOCX to HTML: {str(e)}")
        return None, []

    # Fallback to python-docx if no images extracted
    if not image_data_urls:
        st.warning("No images extracted by mammoth. Trying python-docx fallback.")
        image_data_urls, image_positions = extract_images_fallback(docx_path)
        if image_data_urls:
            # Prepend first image as heading
            image_tag = f'<img src="{image_data_urls[0][0]}" alt="Course Image 1" style="max-width:100%;height:auto;">'
            html_content = f"{image_tag}\n{html_content}"
            st.info("Added image at the start as heading")
            # Append any additional images
            for i, (data_url, ext) in enumerate(image_data_urls[1:], 2):
                image_tag = f'<img src="{data_url}" alt="Course Image {i}" style="max-width:100%;height:auto;">'
                html_content += f"\n{image_tag}"
                st.info(f"Added image tag: {image_tag}")

    # Wrap HTML in a proper structure, escaping curly braces in CSS
    final_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Boating Course Enrollment Confirmation</title>
    <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; }}
        .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
        h1 {{ color: #004080; }}
        .section {{ margin-bottom: 20px; }}
        img {{ max-width: 100%; height: auto; display: block; margin: 0 auto; }}
    </style>
</head>
<body>
    <div class="container">
        {html_content}
    </div>
</body>
</html>"""

    try:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(final_html)
    except Exception as e:
        st.warning(f"Failed to save HTML file: {str(e)}")
        return None, []

    st.info(f"HTML saved to {output_path}")
    if not image_data_urls:
        st.warning("No images extracted from the DOCX file.")
    return final_html, messages

def personalize_html(html_content, first_name, last_name):
    """Personalize HTML content by replacing placeholders."""
    try:
        # Handle both {FirstName} and {{FirstName}} (and similarly for LastName)
        personalized = html_content
        personalized = re.sub(r'\{\{FirstName\}\}|\{FirstName\}', first_name, personalized)
        personalized = re.sub(r'\{\{LastName\}\}|\{LastName\}', last_name, personalized)
        if '{{FirstName}}' in personalized or '{FirstName}' in personalized:
            st.warning(f"Placeholder '{{FirstName}}' or '{FirstName}' not replaced for {first_name}.")
        if '{{LastName}}' in personalized or '{LastName}' in personalized:
            st.warning(f"Placeholder '{{LastName}}' or '{LastName}' not replaced for {last_name}.")
        return personalized
    except Exception as e:
        st.error(f"Failed to personalize HTML: {str(e)}")
        return html_content

def send_email_smtp(to_email, subject, html_content, sender_email, smtp_password):
    """Send email using Gmail's SMTP server."""
    try:
        # Create MIME message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(html_content, 'html'))

        # Connect to Gmail's SMTP server
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()  # Enable TLS
            server.login(sender_email, smtp_password)  # Login with app-specific password
            server.send_message(msg)
        return f"Email sent to {to_email} successfully"
    except Exception as e:
        return f"Failed to send email to {to_email}: {str(e)}"

def main():
    st.title("Mass Email Sender for Boating Course")

    # Sidebar for email configuration
    st.sidebar.header("Email Configuration")
    sender_email = st.sidebar.text_input("Sender Email", value="yourname@gmail.com")
    smtp_password = os.getenv("smtp_password")
    if not smtp_password:
        st.error("SMTP password not found in .env file. Please add 'smtp_password' to your .env file.")
        return
    st.sidebar.markdown("""
        **Note**: Ensure your Gmail app-specific password is set in the .env file.
        Generate one at https://myaccount.google.com/security with two-factor authentication enabled.
    """)

    # Main UI
    st.header("Upload Files")
    excel_file = st.file_uploader("Upload Excel File", type=["xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"])
    docx_file = st.file_uploader(
        "Upload Word Document",
        type=["docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"]
    )

    email_subject = st.text_input("Email Subject", value="Boating Course Enrollment Confirmation - September 6, 2025")

    if st.button("Send Emails"):
        if not excel_file or not docx_file or not sender_email or not smtp_password:
            st.error("Please upload both files and provide sender email.")
            return

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
                status = send_email_smtp(
                    email, email_subject, personalized_html, sender_email, smtp_password
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