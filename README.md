# 🧾 Invoice Automation Script
This Python script automates the creation and emailing of tutoring invoices. It pulls session and client data from a Google Sheet, generates a custom LaTeX invoice for each client based on session dates, compiles the invoice to PDF using MiKTeX’s `pdflatex`, and emails the PDF directly to the client.

This helps simplify and speed up your billing workflow while maintaining professional invoice formatting.

## ⚙️ Setup & Configuration

Before running the script, you’ll need to set up several external services and configurations:

1. **Google Cloud API Setup**
   - Create a [Google Cloud Project](https://console.cloud.google.com).
   - Enable the Google Sheets API and Google Drive API.
   - Create a service account and download its JSON credentials file (`credentials.json`).
   - Share your Google Sheet with the service account email to grant access.

2. **Google Sheet Preparation**
   - Organize your session, client, and cutoff data according to the ranges specified in your `.env` file.
   - Ensure invoice number cell and other config values are set up properly.
   - My sample setup is shown in the picture below.
  
     ![image](https://github.com/user-attachments/assets/eee97739-49a4-4169-80a1-44f978a5a96b)

3. **MiKTeX Installation (for LaTeX compilation)**
   - Install MiKTeX on your system.
   - Add MiKTeX’s `pdflatex` to your system PATH so it can be invoked from the command line.
   - [MiKTeX Download](https://miktex.org/download)

4. **Email Setup (Gmail example)**
   - Enable 2-Step Verification on your [Google Account](https://myaccount.google.com/security).
   - Generate an **App Password** specifically for the script to send emails securely.
   - Replace the email and password placeholders in your `.env` file.

5. **Environment Variables**
   - Create a `.env` file and fill it with important data.
   - See [.env.example](.env.example) for a sample setup.

## 🛠️ Tech Stack

- **Python**: Main scripting language
- **Google Sheets API**: For reading and updating spreadsheet data
- **LaTeX**: Invoice formatting
- **MiKTeX**: LaTeX distribution for PDF generation
- **subprocess**: To run `pdflatex` for compiling LaTeX to PDF
- **gspread & oauth2client**: Google Sheets API client & authentication
- **smtplib & email.message**: Sending emails via SMTP
- **dotenv**: Managing environment variables
