# Flotilla65BoatingEmailRegistration

A Streamlit application designed to streamline public education registration for the Boating Course of the USCG Auxiliary Flotilla 65. This tool automates the conversion of a Word document email template to HTML and sends personalized confirmation emails to registered students using data from an Excel spreadsheet, leveraging Gmail’s SMTP server for email delivery.

## Features

- **Streamlit Web Interface**: User-friendly UI for uploading Excel and Word files, configuring Gmail sender email, and sending mass emails.
- **DOCX to HTML Conversion**: Converts a Word document template to HTML with placeholders (`{FirstName}`, `{{FirstName}}`, `{LastName}`, `{{LastName}}`) for personalization after "Student:".
- **Image Handling**: Extracts images from the Word document and embeds them as Base64 data URLs in the HTML email, ensuring compatibility across email clients.
- **Mass Email Sending**: Sends personalized emails via Gmail’s SMTP server, using recipient data (first name, last name, email) from an Excel file.
- **Secure Credential Management**: Gmail app-specific password is securely loaded from a `.env` file, avoiding hardcoded credentials.
- **Dependency Management**: Uses the UV package manager for reproducible Python environments.
- **Error Handling**: Validates email addresses, skips invalid entries, and provides detailed status messages for each email sent.
- **Temporary File Management**: Safely handles temporary files for uploaded Excel and Word documents, ensuring cleanup after processing.
- **Flexible File Upload**: Supports both `.xlsx`/`.docx` extensions and their MIME types for robust file handling.

## Prerequisites

- **Python**: Version 3.12 or higher.
- **UV Package Manager**: For dependency management (see [UV Documentation](https://docs.astral.sh/uv/)).
- **Gmail Account**: A Gmail account with two-factor authentication (2FA) enabled and an app-specific password generated for SMTP access.
- **Dependencies**: Specified in `pyproject.toml`:
  - `streamlit>=1.45.1`
  - `pandas>=2.2.3`
  - `openpyxl>=3.1.5`
  - `mammoth>=1.9.0`
  - `python-docx>=1.1.2`
  - `python-dotenv>=1.0.1`
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
   Confirm that `streamlit`, `pandas`, `openpyxl`, `mammoth`, `python-docx`, `python-dotenv`, and `lxml` are installed.

4. **Configure Gmail App-Specific Password**:
   - Enable 2FA on your Gmail account at [https://myaccount.google.com/security](https://myaccount.google.com/security).
   - Generate an app-specific password under “App passwords” (e.g., for “Mail”).
   - Create a `.env` file in the project root:
     ```plaintext
     smtp_password=your-gmail-app-specific-password
     ```
   - Replace `your-gmail-app-specific-password` with the 16-character password from Google.

## Usage

### Running the Streamlit App

1. **Prepare Input Files**:
   - **Excel File** (`.xlsx`): Must include columns `First Name`, `Last Name`, and `Primary Student E-mail`. Example: `Fake_KBYC_Source_2025.xlsx`.
   - **Word Document** (`.docx`): Email template with placeholders (e.g., `Student: {FirstName} {LastName}` or `Student: {{FirstName}} {{LastName}}`). Example: `1 Enrollment Confirmation Boating Course June 14, 2025 CRYC.docx`.

2. **Launch the App**:
   ```bash
   uv run streamlit run app.py
   ```
   This opens the Streamlit UI in your browser (e.g., `http://localhost:8501`).

3. **Configure the UI**:
   - **Email Settings** (Sidebar):
     - **Sender Email**: Your Gmail address (e.g., `yourname@gmail.com`).
     - **Note**: The Gmail app-specific password is loaded from the `.env` file. Ensure 2FA is enabled and the password is correctly set.
   - **Upload Files**: Select the Excel and Word files (supports `.xlsx`/`.docx` extensions or MIME types).
   - **Email Subject**: Enter a subject (e.g., “Boating Course Enrollment Confirmation - September 6, 2025”).
   - Click **Send Emails** to process and send.

4. **Output**:
   - The Word document is converted to `confirmation.html` with embedded images and personalized placeholders.
   - Emails are sent via Gmail’s SMTP server with personalized HTML content.
   - The UI displays status messages (success or failure) for each email.
   - Conversion warnings from `mammoth` (if any) are shown.
   - Temporary files are cleaned up automatically.

## Project Structure

```plaintext
Flotilla65BoatingEmailRegistration/
├── .env                 # Gmail app-specific password (excluded from git)
├── .streamlit/          # Optional Streamlit config (excluded from git)
├── app.py               # Streamlit app for mass email sending
├── pyproject.toml       # Project metadata and dependencies
├── uv.lock              # UV lock file for reproducible dependencies
├── README.md            # Project documentation
├── LICENSE              # MIT License
├── confirmation.html    # Generated HTML email template (created at runtime)
└── .gitignore           # Excludes sensitive files
```

## Security Notes

- **Gmail App-Specific Password**: Never hardcode the password in `app.py`. Use the `.env` file for secure storage.
- **Sensitive Files**: Use a `.gitignore` file to exclude sensitive files:
  ```plaintext
  *.xlsx
  *.docx
  confirmation.html
  .env
  .streamlit/
  __pycache__/
  ```
- **Sender Email Verification**: Ensure 2FA is enabled on your Gmail account and the app-specific password is valid.
- **Testing**: Send a test email to yourself (e.g., `jjborie@yahoo.fr`) before processing all recipients to verify Gmail settings and email rendering.

## Troubleshooting

- **Placeholder Replacement Issues**:
  - If emails show `Student: {FirstName}, {LastName}`, verify the `.docx` template uses `{FirstName}` or `{{FirstName}}` (case-sensitive).
  - Check the `confirmation.html` file for the exact placeholder text.
  - Debug with Streamlit warnings or by inspecting the HTML content in the UI.
- **Gmail SMTP Errors**:
  - Verify the app-specific password in the `.env` file (16 characters, no spaces).
  - Ensure 2FA is enabled at [https://myaccount.google.com/security](https://myaccount.google.com/security).
  - If authentication fails, regenerate the app-specific password.
- **File Upload Issues**:
  - Ensure uploaded files have `.xlsx` or `.docx` extensions or correct MIME types (`application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`, `application/vnd.openxmlformats-officedocument.wordprocessingml.document`).
  - Test with a valid `.xlsx` and `.docx` file to rule out corruption.
- **UV Issues**:
  - If dependencies fail to install, run `uv sync --refresh` to update `uv.lock`.
  - Ensure Python 3.12+ is installed and accessible to UV.
- **Gmail Sending Limits**: Gmail allows ~500 emails/day. For a few emails every three months, this should be sufficient, but monitor for limit warnings.

## Known Issues

- **Image Display in Emails**: Some email clients may not display Base64-encoded images. Test with major clients (e.g., Gmail, Outlook, Yahoo) to ensure compatibility.
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