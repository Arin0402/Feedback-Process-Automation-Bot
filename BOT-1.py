import smtplib
import os
import pandas as pd
import openpyxl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

sender_address = "Your Email Address Here"
sender_password = "Your Email Password here"

reciver_address = []
student_name = []
student_email_list =[]
student_registration_number = []
student_mobile_number =[]
student_organization = []
internship_start_list = []

def gettingALlEmailsFromExcelSheet():
    print("Sending Emails ...")
    os.getcwd()
    os.chdir(r"Excel file path")
    file = 'ExcelFile_input_1.xlsx'
    data = pd.ExcelFile(file)
    df = data.parse('Input File 2')
    df.head(10)

    ps = openpyxl.load_workbook('ExcelFile_input_1.xlsx')
    sheet = ps['Input File 2']

    for row in range(5, sheet.max_row + 1):
        email = sheet['P' + str(row)].value
        student_email = sheet['D' + str(row)].value
        name = sheet['B' + str(row)].value
        mobile_no = sheet['C' + str(row)].value
        registration_number = sheet['E' + str(row)].value
        orgainzation = sheet['J' + str(row)].value
        Int_date = sheet['H' + str(row)].value

        reciver_address.append(email)
        student_email_list.append(student_email)
        student_registration_number.append(registration_number)
        student_mobile_number.append(mobile_no)
        student_organization.append(orgainzation)
        internship_start_list.append(Int_date)
        student_name.append(name)

def sendingEmails():
    ind = 0
    for var in reciver_address:
        message = MIMEMultipart()
        message2 = MIMEMultipart()
        message['From'] = sender_address
        message['To'] = var
        message['Subject'] = "FeedBack Form"

        attatchment_file_name = "FeedbackForm.pdf"

        attatch_file = open(attatchment_file_name, 'rb')
        payload = MIMEBase('application', 'octate-stream')

        payload.set_payload(attatch_file.read())

        encoders.encode_base64(payload)
        payload.add_header('Content-Disposition', f'attatchment; filename={attatchment_file_name} ')
        message.attach(payload)

        message.attach(MIMEText("Student Name :" + student_name[ind] + "      "))
        message.attach(MIMEText("Student Registration number :" +student_registration_number[ind] + "      "))
        message.attach(MIMEText("Student Mobile Number :" + str(student_mobile_number[ind]) + "      "))
        message.attach(MIMEText("Student Email Id :" + student_email_list[ind] + "      "))
        message.attach(MIMEText("Internship Start Date :" + str(internship_start_list[ind]) + "      "))
        message.attach(MIMEText("Organization :" + str(student_organization[ind]) + "      "))


        session = smtplib.SMTP('smtp.gmail.com', 587)
        session.starttls()
        session.login(sender_address, sender_password)
        text = message.as_string()
        session.sendmail(sender_address, var, text )
        session.quit()
        print(str(ind+1) + ' mail sent' )

        ind += 1

    print("All the emails sent successfully !")

gettingALlEmailsFromExcelSheet()
sendingEmails()
