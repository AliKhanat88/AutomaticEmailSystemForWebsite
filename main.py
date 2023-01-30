import gspread
from oauth2client.service_account import ServiceAccountCredentials

from email.mime.audio import MIMEAudio
from mimetypes import MimeTypes
import smtplib, ssl
from email import message
from os.path import basename
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


def send_mail(send_from, send_to, subject, text, filename, password):
    from_addr = send_from
    to_addr = send_to

    subject = subject

    body = text

    msg = MIMEMultipart()

    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Subject'] = subject
    body = MIMEText(body, 'plain')
    msg.attach(body)

    with open(filename, "r", encoding="utf-8") as f:
        part = MIMEApplication(f.read(), Name=basename(filename))
        part['Content-Disposition'] = 'attachment; filename="{}"'.format(basename(filename))
        msg.attach(part)

    context=ssl.create_default_context()

    with smtplib.SMTP("smtp.gmail.com", port=587) as smtp:
        smtp.starttls(context=context)
        smtp.login(from_addr, password)
        smtp.send_message(msg)
    smtp.close()


def main():
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('taml-366419-c5b41289e628.json', scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open("TAML")
    sheet = spreadsheet.get_worksheet(1)

    file = open("categories.csv", encoding="utf-8")

    data = []
    csv_data = []
    file.readline()
    for line in file:
        csv_data.append(line)
        data.append(line.strip().split(","))

    send_from = input("Enter you username: ")
    password = input("Enter your 16-degit smtp app password: ")
    for record in sheet.get_all_records():
        var = record["Category"]
        f = open(f'data_{record["Email"]}.csv', "w", encoding="utf-8")
        for i in range(len(data)):
            sub = data[i][2].split("|")
            for cate in sub:
                if var == cate.strip():
                    f.write(csv_data[i])

        f.close()

        send_mail(send_from, record["Email"],"Category file", "This is a csv file. You can open it with google sheet or excel.", f'data_{record["Email"]}.csv',
              password)

    
main()


