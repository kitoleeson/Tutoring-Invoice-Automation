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

def shorten_semester(semester):
    return semester[0] + semester[-4:]

# custom parsing functions
def parse_date(serial_number):
    return datetime(1899, 12, 30) + timedelta(days=float(serial_number))

def parse_session(row):
    return [ row[0], parse_date(row[1]), float(row[2]), float(row[3]) ]

def parse_payment(row):
    return [ row[0], parse_date(row[1]), float(row[2]) ]

def parse_client(row):
    return [ row[0], int(row[1]), row[2], float(row[3]), row[4], float(row[5]), float(row[6]), row[7], row[8] ]

# ----------- CORE FUNCTION -----------
def create_and_send_summary(name):
    print(name.upper())

    # grab info from sheet
    print("Pulling info".ljust(20, '.'), end="")
    all_sessions = sheet.get(os.getenv("SESSION_RANGE"), value_render_option='UNFORMATTED_VALUE')
    all_clients = [parse_client(row) for row in sheet.get(os.getenv("CLIENT_RANGE"))]
    semester = sheet.title
    all_payments = sheet.get(os.getenv("PAYMENT_RANGE"), value_render_option='UNFORMATTED_VALUE')
    print("done.")

    # custom variables
    client_sessions = [parse_session(row) for row in all_sessions if row[0] == name] # parseFloat number values
    client_payments = [parse_payment(row) for row in all_payments if row[0] == name]
    client_data = next(row for row in all_clients if row[0] == name)
   
    # create .tex string
    print("Writing tex".ljust(20, '.'), end="")
    initials = extract_initials(name)
    filename = f"SUM-{shorten_semester(semester)}_{initials}.tex"
    latex_content = get_summary_template(client_data, semester, client_sessions, client_payments)

    # create .tex file
    os.makedirs("invoices/", exist_ok=True)
    tex_path = os.path.join("invoices/", filename)
    with open(tex_path, "w") as f:
        f.write(latex_content)
    print("done.")

    # compile .tex into .pdf file
    print("Compiling pdf".ljust(20, '.'), end="")
    subprocess.run(
        ["pdflatex", "-output-directory", "invoices/", tex_path],
        check=True,
        stdout=subprocess.DEVNULL  # suppress stdout
    )
    print("done.")

    # send email
    pdf_path = tex_path.replace(".tex", ".pdf")
    send_summary_email(name, client_data[7], client_data[8], pdf_path, semester)

    # update sheet invoice number
    print("Summary Complete!", end="\n\n")

# ----------- EMAIL -----------
def send_summary_email(name, payer, email, pdf_path, semester):
    print("Sending email".ljust(20, '.'), end="")

    def fill_content(payer, semester):
        lines = [
            f"Good day {payer.split(' ')[0]},",
            f"Please find attached your tutoring session summary for {semester}.\nPlease let me know if you have any questions, I hope to see you again next semester!",
            os.getenv("MY_NAME")
        ]
        return "\n\n".join(lines)

    msg = EmailMessage()
    msg['Subject'] = f"{name} Tutoring Summary {semester}"
    msg['From'] = os.getenv("MY_EMAIL")
    msg['To'] = email
    msg.set_content(fill_content(payer, semester))

    with open(pdf_path, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=os.path.basename(pdf_path))

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(os.getenv("MY_EMAIL"), os.getenv("MY_EMAIL_APP_PASSWORD"))
        smtp.send_message(msg)
        print("done.")

def get_summary_template(client_data, semester, client_sessions, client_payments):
    # define variables
    session_rows = ""
    payment_rows = ""

    for s in client_sessions:
        session_rows += f"{shorten_date(s[1], False)} & {s[2]} & {s[3]} \\\\"

    for p in client_payments:
        payment_rows += f"{shorten_date(p[1], False)} & {p[2]} \\\\"

    # create file
    return rf"""
    \documentclass[12pt]{{article}}
    \usepackage[margin=1in]{{geometry}}
    \usepackage{{booktabs}}
    \usepackage{{datetime}}
    \usepackage{{array}}
    \usepackage{{tabularx}}
    \usepackage[dvipsnames]{{xcolor}}

    \newcommand{{\clientName}}{{{client_data[0]}}}
    \newcommand{{\subjectsList}}{{{client_data[4]}}}
    \newcommand{{\hourlyRate}}{{{client_data[5]:.2f}}}
    \newcommand{{\semester}}{{{semester}}}

    \newcommand{{\sessionCount}}{{{len(client_sessions)}}}
    \newcommand{{\totalHours}}{{{sum(s[2] for s in client_sessions):.2f}}}
    \newcommand{{\sessionTotal}}{{{sum(s[3] for s in client_sessions):.2f}}}
    \newcommand{{\sessionRows}}{{{session_rows}}}

    \newcommand{{\paymentCount}}{{{len(client_payments)}}}
    \newcommand{{\paymentTotal}}{{{sum(p[2] for p in client_payments):.2f}}}
    \newcommand{{\paymentRows}}{{{payment_rows}}}

    \newcommand{{\invoiceHeader}}{{
        \begin{{minipage}}[t]{{0.48\textwidth}}        
        \begin{{flushleft}}
            \textbf{{Client Name:}}\\
            \clientName\\
            \vspace{{1em}}
            \textbf{{Subjects:}} \subjectsList\\
            \textbf{{Hourly Rate:}} \$\hourlyRate\\
            \vspace{{1em}}
            \textbf{{Semester: \semester}}\\
        \end{{flushleft}}
        \end{{minipage}}
        \hfill
        \begin{{minipage}}[t]{{0.48\textwidth}}
        \begin{{flushright}}
            \textbf{{{os.getenv("MY_NAME")}}}\\
            \textit{{{os.getenv("MY_CITY")}}}\\
            \textit{{e: {os.getenv("MY_EMAIL")}}}\\
            \textit{{p: {os.getenv("MY_NUMBER")}}}
        \end{{flushright}}
        \end{{minipage}}
        \vspace{{2em}}
    }}

    % Main document
    \begin{{document}}
    \pagestyle{{empty}}
    \invoiceHeader

    % Side-by-side tables
    % Sessions and payments
    \noindent
    \begin{{minipage}}[t]{{0.45\textwidth}}
        \vspace{{0em}}
        \begin{{tabular}}{{ p{{2cm}} >{{\raggedleft\arraybackslash}}p{{1cm}} >{{\raggedleft\arraybackslash}}p{{3cm}} }}
            \toprule
            \multicolumn{{3}}{{l}}{{\textbf{{All Sessions}}}} \\
            \midrule
            \textbf{{Date}} & \textbf{{Hours}} & \textbf{{Fee (\$)}} \\
            \midrule
            \sessionRows
            \midrule
            \multicolumn{{2}}{{c}}{{\textbf{{Session Total:}}}} & \textbf{{\$\sessionTotal}} \\
            \bottomrule
        \end{{tabular}}
    \end{{minipage}}
    \hfill
    \begin{{minipage}}[t]{{0.45\textwidth}}
        \vspace{{0em}}
        \begin{{tabular}}{{ p{{3.5cm}} >{{\raggedleft\arraybackslash}}p{{2.5cm}} }}
            \toprule
            \multicolumn{{2}}{{l}}{{\textbf{{All Payments}}}} \\
            \midrule
            \textbf{{Date}} & \textbf{{Amount (\$)}} \\
            \midrule
            \paymentRows
            \midrule
            \textbf{{Payment Total:}} & \textbf{{\$\paymentTotal}} \\
            \bottomrule
        \end{{tabular}}
    \end{{minipage}}
    \vspace{{3em}}

    \noindent
    \begin{{tabularx}}{{\textwidth}}{{l >{{\raggedleft\arraybackslash}}X}}
        \toprule
        \textbf{{Semester Summary}} & \\
        \midrule
        \textbf{{Number of Sessions}} & \sessionCount \\
        \textbf{{Total Hours}} & \totalHours \\
        \textbf{{Hourly Rate}} & \$\hourlyRate \\
        \textbf{{Session Total}} & \$\sessionTotal \\
        \midrule
        \textbf{{Number of Payments}} & \paymentCount \\
        \textbf{{Total Paid}} & \paymentTotal \\
        \bottomrule
    \end{{tabularx}}

    \end{{document}}
    """

if __name__ == '__main__':
    # if names are provided
    if len(sys.argv) > 1:
        names = sys.argv[1:]
    else:
        names = [row[0] for row in sheet.get(os.getenv("CLIENT_RANGE"))]

    for name in names:
        create_and_send_summary(name)

    # remove .aux and .log files
    for file in glob.glob("invoices/*.aux") + glob.glob("invoices/*.log"):
        os.remove(file)