# Flotilla65BoatingEmailRegistration

A Streamlit application designed to streamline public education registration for the Boating Course of the USCG Auxiliary Flotilla 65. This tool automates the conversion of a Word document email template to HTML and sends personalized confirmation emails to registered students using data from an Excel spreadsheet, leveraging SendGrid for email delivery.

## Features

- **Streamlit Web Interface**: User-friendly UI for uploading Excel and Word files, configuring SendGrid settings, and sending mass emails.
- **DOCX to HTML Conversion**: Converts a Word document template to HTML with placeholders (`{FirstName}`, `{LastName}`) for personalization after "Student:".
- **Image Handling**: Extracts images from the Word document and embeds them as Base64 data URLs in the HTML email, ensuring compatibility across email clients.
- **Mass Email Sending**: Sends personalized emails via SendGrid, using recipient data (first name, last name, email) from an Excel file.
- **Secure Credential Management**: SendGrid API key is securely loaded from Streamlit secrets, avoiding hardcoded credentials.
- **Dependency Management**: Uses the UV package manager for reproducible Python environments.
- **Error Handling**: Validates email addresses, skips invalid entries, and provides detailed status messages for each email sent.
- **Temporary File Management**: Safely handles temporary files for uploaded Excel and Word documents, ensuring cleanup after processing.

## Prerequisites

- **Python**: Version 3.12 or higher.
- **UV Package Manager**: For dependency management (see [UV Documentation](https://docs.astral.sh/uv/)).
- **SendGrid Account**: A verified SendGrid account with an API key and a verified sender email.
- **Streamlit Secrets**: A configured `secrets.toml` file or Streamlit Cloud secrets with the SendGrid API key.
- **Dependencies**: Specified in `pyproject.toml`:
  - `streamlit>=1.45.1`
  - `pandas>=2.2.3`
  - `openpyxl>=3.1.5`
  - `mammoth>=1.9.0`
  - `python-docx>=1.1.2`
  - `pillow>=11.2.1`
  - `sendgrid>=6.12.2`
  - `lxml` (implicit dependency via `python-docx`)

## Installation

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/jjborie/Flotilla65BoatingEmailRegistration.git
   cd Flotilla65BoatingEmailRegistration
   ```

2. **Set Up UV Environment**:
   ```bash
   uv sync
   ```
   This installs dependencies from `pyproject.toml` into a virtual environment.

3. **Verify Dependencies**:
   ```bash
   uv pip list
   ```
   Confirm that `streamlit`, `pandas`, `openpyxl`, `mammoth`, `python-docx`, `pillow`, `sendgrid`, and `lxml` are installed.

4. **Configure Streamlit Secrets**:
   - Create a `.streamlit/secrets.toml` file in the project root or configure secrets in Streamlit Cloud:
     ```toml
     [general]
     SENDGRID_API_KEY = "your-sendgrid-api-key"
     ```
   - Replace `your-sendgrid-api-key` with your actual SendGrid API key.

## Usage

### Running the Streamlit App

1. **Prepare Input Files**:
   - **Excel File** (`.xlsx`): Must include columns `First Name`, `Last Name`, and `Primary Student E-mail`. Example: `Fake_KBYC_Source_2025.xlsx`.
   - **Word Document** (`.docx`): Email template with "Student:" as a placeholder for personalization. Example: `1 Enrollment Confirmation Boating Course June 14, 2025 CRYC.docx`.

2. **Launch the App**:
   ```bash
   uv run streamlit run app.py
   ```
   This opens the Streamlit UI in your browser (e.g., `http://localhost:8501`).

3. **Configure the UI**:
   - **SendGrid Settings** (Sidebar):
     - **Sender Email**: Your verified SendGrid sender email (e.g., `jjborie@hotmail.com`).
     - **Note**: The SendGrid API key is loaded from Streamlit secrets. Ensure the sender email is verified in SendGrid at [https://app.sendgrid.com/settings/sender_auth](https://app.sendgrid.com/settings/sender_auth).
   - **Upload Files**: Select the Excel and Word files.
   - **Email Subject**: Enter a subject (e.g., "Boating Course Enrollment Confirmation - June 14, 2025").
   - Click **Send Emails** to process and send.

4. **Output**:
   - The Word document is converted to `confirmation.html` with embedded images and "Student: {FirstName} {LastName}".
   - Emails are sent via SendGrid with personalized HTML content.
   - The UI displays status messages (success or failure) for each email.
   - Conversion warnings from `mammoth` (if any) are shown.
   - Temporary files are cleaned up automatically.

### Standalone HTML Conversion

To convert a Word document to HTML independently (e.g., for testing):
```bash
uv run python convert_docx_to_html.py
```
- **Input**: Word document at `DOCX_PATH` (default: `1 Enrollment Confirmation Boating Course June 14, 2025 CRYC.docx`).
- **Output**: `confirmation.html` with embedded Base64 images and "Student: {FirstName} {LastName}".
- **Note**: The generated HTML is optimized for email use with inline images.

## Project Structure

```plaintext
Flotilla65BoatingEmailRegistration/
├── .streamlit/
│   └── secrets.toml       # Streamlit secrets (excluded from git)
├── app.py                # Streamlit app for mass email sending
├── convert_docx_to_html.py # Script for DOCX-to-HTML conversion
├── pyproject.toml        # Project metadata and dependencies
├── uv.lock               # UV lock file for reproducible dependencies
├── README.md             # Project documentation
├── LICENSE               # MIT License
├── confirmation.html     # Generated HTML email template (created at runtime)
└── .gitignore            # Excludes sensitive files
```

## Security Notes

- **SendGrid API Key**: Never hardcode the API key in `app.py`. Use Streamlit secrets for secure storage.
- **Sensitive Files**: Use a `.gitignore` file to exclude sensitive files:
  ```plaintext
  *.xlsx
  *.docx
  confirmation.html
  .streamlit/secrets.toml
  __pycache__/
  ```
- **Sender Email Verification**: Ensure the sender email is verified in SendGrid to avoid email delivery issues.
- **Testing**: Send a test email to yourself before processing all recipients to verify SendGrid settings and email rendering.

## Troubleshooting

- **SendGrid Errors**:
  - Verify the API key and sender email in SendGrid.
  - Check SendGrid dashboard for delivery issues or blocks.
  - Ensure `secrets.toml` is correctly configured.
- **DOCX Conversion Issues**:
  - If `confirmation.html` lacks formatting or images, confirm the Word document contains inline images (not floating or linked).
  - Review `mammoth` warnings in the Streamlit UI or console output.
  - Ensure the Word document is a valid `.docx` file.
- **Invalid Emails**: The app skips rows with missing or invalid email addresses, logging them in the UI.
- **UV Issues**:
  - If dependencies fail to install, run `uv sync --refresh` to update `uv.lock`.
  - Ensure Python 3.12+ is installed and accessible to UV.
- **Streamlit Secrets Errors**:
  - If the API key is not found, verify `.streamlit/secrets.toml` exists or Streamlit Cloud secrets are configured.

## Known Issues

- **Image Display in Emails**: Some email clients may not display Base64-encoded images. Test with major clients (e.g., Gmail, Outlook) to ensure compatibility.
- **Excel File Sensitivity**: Ensure the Excel file has exact column names (`First Name`, `Last Name`, `Primary Student E-mail`) to avoid parsing errors.
- **Mammoth Conversion Limitations**: Complex Word document formatting may not translate perfectly to HTML. Simplify the document if issues arise.

## Contributing

Contributions are welcome! To contribute:
1. Fork the repository.
2. Create a feature branch:
   ```bash
   git checkout -b feature/your-feature
   ```
3. Commit changes:
   ```bash
   git commit -m 'Add your feature'
   ```
4. Push to the branch:
   ```bash
   git push origin feature/your-feature
   ```
5. Open a pull request on GitHub.

Please include tests and update documentation for new features.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contact

For questions or support, contact the USCG Auxiliary Flotilla 65 or open an issue on [GitHub](https://github.com/jjborie/Flotilla65BoatingEmailRegistration/issues).

---
Developed to support the public education mission of the USCG Auxiliary Flotilla 65.