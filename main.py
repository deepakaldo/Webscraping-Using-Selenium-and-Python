from openpyxl import Workbook
import smtplib
from email.message import EmailMessage


def main():
    msg=EmailMessage()
    msg['To'] = 'deepakaldo47@gmail.com'
    msg['From'] = 'aldoenterprise'
    msg['Subject'] = "training invitation"
    msg.set_content("data")

    # with open('EmailTemplate.txt') as myfile:
    #     data=myfile.read()
    #     msg.set_content(data)
    #
    # with open("finalrecord.xlsx","rb") as f: #read as binary
    #     file_data=f.read()
    #     file_name=f.name
    #     msg.add_attachment(file_data,maintype="application",subtype="xlsx",filename=file_name)
    #
    with smtplib.SMTP_SSL('smtp.gmail.com',465) as server:
        server.login(user="aldoenterprise8@gmail.com", password="cxbs axln cmkl hmah")
        server.send_message(msg)

    print(" email sent!!")
main()

