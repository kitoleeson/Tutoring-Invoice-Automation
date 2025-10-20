# Main Python program for automating invoice creation from Google Sheets to email

import sys # for command line arguments
import gspread # read/write data from google sheets
from oauth2client.service_account import ServiceAccountCredentials # authentification for google
import os # interacting with the operating system
import subprocess # to run external commands
from datetime import datetime, timedelta # date formatting
from dotenv import load_dotenv # for use of .env file
import glob # for globbing patterns
import smtplib # sends email via smtp
from email.message import EmailMessage # for sending emails


# ----------- SETUP -----------
load_dotenv()

# ----------- AUTHENTICATION -----------
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(os.getenv("SPREADSHEET_KEY")).sheet1

# ----------- HELPERS -----------
def shorten_date(dt, full=True):
    if full:
        return dt.strftime("%B %#d")
    return dt.strftime("%b %#d")

def extract_initials(name):
    return ''.join(word[0] for word in name.split()).upper()

def print2D(name, arr):
    print(name.upper())
    for row in arr:
        print(row)
    print()

# custom parsing functions
def parse_date(serial_number):
    return datetime(1899, 12, 30) + timedelta(days=float(serial_number))

def parse_session(row):
    return [ row[0], parse_date(row[1]), float(row[2]), float(row[3]) ]

def parse_client(row):
    return [ row[0], int(row[1]), row[2], float(row[3]), row[4], float(row[5]), float(row[6]), row[7], row[8] ]

# ----------- CORE FUNCTION -----------
def create_and_send_invoice(name):
    print(name.upper())

    # grab info from sheet
    print("Pulling info".ljust(20, '.'), end="")
    all_sessions = sheet.get(os.getenv("SESSION_RANGE"), value_render_option='UNFORMATTED_VALUE')
    all_clients = [parse_client(row) for row in sheet.get(os.getenv("CLIENT_RANGE"))]
    cutoff_dates = [parse_date(row[0]) for row in sheet.get(os.getenv("CUTOFF_RANGE"), value_render_option='UNFORMATTED_VALUE')]
    invoice_number = sheet.acell(os.getenv("INVOICE_NUMBER_RANGE")).value.zfill(4)
    print("done.")

    # custom variables
    client_sessions = [parse_session(row) for row in all_sessions if row[0] == name] # parseFloat number values
    current_sessions = [s for s in client_sessions if cutoff_dates[0] <= s[1] < cutoff_dates[1]]
    client_data = next(row for row in all_clients if row[0] == name)
   
    # create .tex string
    print("Writing tex".ljust(20, '.'), end="")
    initials = extract_initials(name)
    filename = f"INV-{invoice_number}_{initials}.tex"
    latex_content = get_invoice_template(client_data, invoice_number, current_sessions)

    # create .tex file
    invoice_folder = f"invoices/{sheet.title}/"
    os.makedirs(invoice_folder, exist_ok=True)
    tex_path = os.path.join(invoice_folder, filename)
    with open(tex_path, "w") as f:
        f.write(latex_content)
    print("done.")

    # compile .tex into .pdf file
    print("Compiling pdf".ljust(20, '.'), end="")
    subprocess.run(
        ["pdflatex", "-output-directory", invoice_folder, tex_path],
        check=True,
        stdout=subprocess.DEVNULL  # suppress stdout
    )
    print("done.")

    # send email
    pdf_path = tex_path.replace(".tex", ".pdf")
    send_invoice_email(name, client_data[7], client_data[8], pdf_path, cutoff_dates)

    # ADD FUNCTION TO UPLOAD PDF INVOICE TO DRIVE FOR STORAGE (or NAS)
    # upload pdf to google drive
    # drive_folder = os.getenv("INVOICE_FOLDER_KEY")
    # upload_to_drive(pdf_path, drive_folder)

    # remove '!' from payer name in sheet (if necessary)
    if client_data[7][-1] == '!':
        print("Updating payer".ljust(20, '.'), end="")
        for i, row in enumerate(all_clients):
            if row[7] == client_data[7]:
                sheet.update_acell(f"{chr(ord(os.getenv("CLIENT_RANGE")[0]) + 7)}{i + int(os.getenv("CLIENT_RANGE")[1])}", client_data[7][:-1])
                print("done.")
                break

    # update sheet invoice number
    sheet.update_acell(os.getenv("INVOICE_NUMBER_RANGE"), str(int(invoice_number) + 1))
    print("Invoice Complete!", end="\n\n")


# ----------- EMAIL AND DRIVE -----------
def send_invoice_email(name, payer, email, pdf_path, cutoff_dates):
    print("Sending email".ljust(20, '.'), end="")

    def fill_content(name, payer, cutoff_dates):
        # lines = [
        #     f"Good evening {payer.split(' ')[0]},",
        #     f"I'd like to welcome you to a new semester of schooling and a new semester of tutoring for {name.split(' ')[0]}.",
        #     "Throughout last semester, I built a new invoice system to help me keep my billing simple and consistent. Here's what to expect going forward:\n\t-  Invoices will now be sent biweekly directly to your email.\n\t-  Payment is due within 10 days from the day you receive the invoice.\n\t-  All fees can be paid via eTransfer using the email and phone number listed on each invoice.",
        #     "I'd also like to remind you that my sessions are billed in increments of 15 mins, rounded up or down to the nearest 0.25 hours; and that your hourly rate will never change from the rate originally set when we began working together -- even if my rates go up for new clients, yours will remain the same.",
        #     f"Please find attached your first tutoring invoice of the semester, for {shorten_date(cutoff_dates[0])} (inclusive) to {shorten_date(cutoff_dates[1])} (exclusive).",
        #     "Please feel free to reach out if you have any questions regarding invoices, payments, or scheduling.\nI appreciate your trust and support, and I'm excited to see the progress this semester will bring!",
        #     os.getenv("MY_NAME")
        # ]
        if payer[-1] == '!':
            lines = [
                f"Good evening {payer.split(' ')[0]},",
                f"I'd like to welcome you to a new semester of tutoring for {name.split(' ')[0]}.",
                "Throughout last semester, I built a new invoice system to help me keep my billing simple and consistent. Here's what to expect going forward:\n\t-  Invoices will be sent biweekly directly to your email.\n\t-  Payment is due within 10 days from the day you receive the invoice.\n\t-  All fees can be paid via eTransfer using the email and phone number listed on each invoice.",
                "I'd also like to inform/remind you that my sessions are billed in increments of 15 mins, rounded up or down to the nearest 0.25 hours; and that your hourly rate will never change from the rate originally set when we began working together -- even if my rates go up for new clients, yours will remain the same.",
                f"Please find attached your first tutoring invoice of the semester, for {shorten_date(cutoff_dates[0])} (inclusive) to {shorten_date(cutoff_dates[1])} (exclusive).",
                "Please feel free to reach out if you have any questions regarding invoices, payments, or scheduling.\nI appreciate your trust and support, and I'm excited to see the progress this semester will bring!",
                os.getenv("MY_NAME")
            ]
        else:
            lines = [
                f"Good day {payer.split(' ')[0]},",
                f"Please find attached your tutoring invoice for {shorten_date(cutoff_dates[0])} (inclusive) to {shorten_date(cutoff_dates[1])} (exclusive).",
                os.getenv("MY_NAME")
            ]
        return "\n\n".join(lines)

    msg = EmailMessage()
    msg['Subject'] = f"{name} Tutoring Invoice"
    msg['From'] = os.getenv("MY_EMAIL")
    msg['To'] = email
    msg.set_content(fill_content(name, payer, cutoff_dates))

    with open(pdf_path, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=os.path.basename(pdf_path))

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(os.getenv("MY_EMAIL"), os.getenv("MY_EMAIL_APP_PASSWORD"))
        smtp.send_message(msg)
        print("done.")

def get_invoice_template(client_data, invoice_number, current_sessions):
    # define variables
    session_rows = ""

    for s in current_sessions:
        session_rows += f"{shorten_date(s[1], False)} & {s[2]} & {s[3]} \\\\"

    session_total = sum(s[3] for s in current_sessions)
    current_tab = client_data[6]
    total_due = session_total + current_tab

    # create file
    return rf"""
    \documentclass[12pt]{{article}}
    \usepackage[margin=1in]{{geometry}}
    \usepackage{{booktabs}}
    \usepackage{{datetime}}
    \usepackage{{array}}
    \usepackage[dvipsnames]{{xcolor}}

    \newcommand{{\invoiceNumber}}{{{invoice_number}}}
    \newcommand{{\clientName}}{{{client_data[0]}}}
    \newcommand{{\subjectsList}}{{{client_data[4]}}}
    \newcommand{{\hourlyRate}}{{{client_data[5]:.2f}}}
    \newcommand{{\sessionCount}}{{{len(current_sessions)}}}
    \newcommand{{\totalHours}}{{{sum(s[2] for s in current_sessions):.2f}}}
    \newcommand{{\sessionTotal}}{{{session_total:.2f}}}
    \newcommand{{\currentTab}}{{{current_tab:.2f}}}
    \newcommand{{\totalAmount}}{{{total_due:.2f}}}
    \newcommand{{\sessionRows}}{{{session_rows}}}

    \newcommand{{\invoiceHeader}}{{
        \begin{{flushright}}
            \textbf{{{os.getenv("MY_NAME")}}}\\
            \textit{{{os.getenv("MY_CITY")}}}\\
            \textit{{e: {os.getenv("MY_EMAIL")}}}\\
            \textit{{p: {os.getenv("MY_NUMBER")}}}
        \end{{flushright}}
        
        \vspace{{2em}}
        
        \begin{{flushleft}}
            \textbf{{Invoice INV-\invoiceNumber}}\\
            \textbf{{Date: \today}}\\
            \vspace{{1em}}
            \textbf{{Client Name:}}\\
            \clientName\\
            \vspace{{1em}}
            \textbf{{Subjects:}} \subjectsList\\
            \textbf{{Hourly Rate:}} \$\hourlyRate\\
        \end{{flushleft}}
        
        \vspace{{3em}}
    }}

    % Footer with payment information
    \newcommand{{\invoiceFooter}}{{
        \vspace{{2em}}
        \begin{{flushleft}}
            \textbf{{Payment Terms:}}\\
            Payment is due within 10 days of invoice date.\\
            Please send an e-transfer to the email or phone number found at the top of this invoice.\\
            Late fee of 1.5\% per month applies to unpaid balances.
        \end{{flushleft}}
    }}

    % Main document
    \begin{{document}}
    \pagestyle{{empty}}
    \invoiceHeader

    % Session table
    \noindent
    \begin{{minipage}}[t]{{0.45\textwidth}}
        \begin{{tabular}}{{ p{{2cm}} >{{\raggedleft\arraybackslash}}p{{1cm}} >{{\raggedleft\arraybackslash}}p{{3cm}} }}
            \toprule
            \multicolumn{{2}}{{r}}{{\textbf{{Session Summary}}}} & \\
            \midrule
            \textbf{{Date}} & \textbf{{Hours}} & \textbf{{Fee (\$)}} \\
            \midrule
            \sessionRows
            \midrule
            \multicolumn{{2}}{{r}}{{\textbf{{Session Total:}}}} & \textbf{{\$\sessionTotal}} \\
            \bottomrule
        \end{{tabular}}
    \end{{minipage}}
    \hfill
    \begin{{minipage}}[t]{{0.45\textwidth}}
        \begin{{tabular}}{{@{{}} l r @{{}}}}
            \toprule
            \textbf{{Invoice Summary}} & \\
            \midrule
            \textbf{{Sessions}} & \sessionCount \\
            \textbf{{Total Hours}} & \totalHours \\
            \textbf{{Hourly Rate}} & \$\hourlyRate \\
            \textbf{{Session Total}} & \$\sessionTotal \\
            \textbf{{Current Tab}} & \$\currentTab \\
            \midrule
            \textbf{{Total Due}} & \\
            \textbf{{(Tab + Session Total)}} & \textbf{{\$\totalAmount}} \\
            \bottomrule
        \end{{tabular}}
    \end{{minipage}}

    \invoiceFooter

    \end{{document}}
    """

if __name__ == '__main__':
    # if names are provided
    if len(sys.argv) > 1:
        names = sys.argv[1:]
    else:
        # find all clients who have had a session within the time frame
        cutoff_dates = [parse_date(row[0]) for row in sheet.get(os.getenv("CUTOFF_RANGE"), value_render_option='UNFORMATTED_VALUE')]
        all_sessions = sheet.get(os.getenv("SESSION_RANGE"), value_render_option='UNFORMATTED_VALUE')
        names = set()
        for row in all_sessions:
            if cutoff_dates[0] <= parse_date(row[1]) < cutoff_dates[1]:
                names.add(row[0])
        names = list(names)

    for name in names:
        create_and_send_invoice(name)

    # remove .aux and .log files
    for file in glob.glob(f"invoices/{sheet.title}/*.aux") + glob.glob(f"invoices/{sheet.title}/*.log"):
        os.remove(file)