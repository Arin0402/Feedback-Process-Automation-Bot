import os
from imbox import Imbox
import traceback
import PyPDF2
import os
import re
import smtplib
import xlwt
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

print("Reading mails ...")
host = "imap.gmail.com"
username = "email id here"
password = 'Email password'
download_folder = "Folder path where forms to be downloaded"

if not os.path.isdir(download_folder):
    os.makedirs(download_folder, exist_ok=True)

mail = Imbox(host, username=username, password=password, ssl=True, ssl_context=None, starttls=False)
messages = mail.messages()

ind = 1
received_emails =[]
for (uid, message) in messages:
    mail.mark_seen(uid)  # optional, mark message as read

    # print(message.keys())
    for idx, attachment in enumerate(message.attachments):
        try:
            att_fn = attachment.get('filename')

            att_fn = att_fn[0 : 12] + str(ind) + '.pdf'
            ind+=1
            # print(att_fn)
            download_path = f"{download_folder}\{att_fn}"
            # print(download_path)
            with open(download_path, "wb") as fp:
                fp.write(attachment.get('content').read())
        except:
            print(traceback.print_exc())

mail.logout()

time.sleep(1)

print("Extracting Details from Form ...")

time.sleep(1)
from xlwt import Workbook
wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')

style = xlwt.easyxf('font: bold 1')

sheet1.write(0, 0, 'S No.', style )
sheet1.write(0, 1, 'Student Name', style )
sheet1.write(0, 2, 'Student Phone Number', style )
sheet1.write(0, 3, 'Student Email ID', style )
sheet1.write(0, 4, 'Student Registration No', style )
sheet1.write(0, 5, 'Branch ', style )
sheet1.write(0, 6, 'Father’s Mobile Number', style )
sheet1.write(0, 7, 'Internship Start Date ', style )
sheet1.write(0, 8, 'Stipend', style )
sheet1.write(0, 9, 'Organization Name', style )
sheet1.write(0, 10, 'Organization Address', style )
sheet1.write(0, 11, 'City', style )
sheet1.write(0, 12, 'Duration in Months', style )
sheet1.write(0, 13, 'Title of the Project', style )
sheet1.write(0, 14, 'Industry Guide Feedback ', style )
sheet1.write(0, 15, 'Industry Guide Rating on a scale of 10', style )
sheet1.write(0, 16, 'Industry Guide Name', style )
sheet1.write(0, 17, 'Industry Guide Email ID', style )
sheet1.write(0, 18, 'Designation', style )
sheet1.write(0, 19, 'Industry Guide Mobile Number', style )

for i in range(1,20):
    sheet1.col(i).width = 5000 # around 220 pixels

line_no = 1

print("Generating Excel File ...")

time.sleep(1)

def sendingEmails(faulty_mail):
    ind = 0
    message = MIMEMultipart()
    message2 = MIMEMultipart()
    message['From'] = username
    message['To'] = faulty_mail
    message['Subject'] = "Please fill the Feedback form correctly and send pdf format only"

    attatchment_file_name = "FeedbackForm.pdf"

    attatch_file = open(attatchment_file_name, 'rb')
    payload = MIMEBase('application', 'octate-stream')

    payload.set_payload(attatch_file.read())

    encoders.encode_base64(payload)
    payload.add_header('Content-Disposition', f'attatchment; filename={attatchment_file_name} ')
    message.attach(payload)


    session = smtplib.SMTP('smtp.gmail.com', 587)
    session.starttls()
    session.login(username, password)
    text = message.as_string()
    session.sendmail(username, faulty_mail, text )
    session.quit()
    print(str(ind+1) + ' mail sent' )

# try :
for file_name in os.listdir('C:\Feedback_form_received'):
    # print (file_name)
    try:
        load_pdf = open(r'C:\Feedback_form_received\\'+file_name, 'rb')

        read_pdf = PyPDF2.PdfFileReader(load_pdf)
        page_count= read_pdf.getNumPages ()
        first_page= read_pdf.getPage (0)
        page_content = first_page.extractText()
        page_content = page_content.replace('\n','')
        # print (page_content)

        Date = re.search('Performance Management Report(.*)Student Name:',page_content)
        Student_name = re.search('Student Name:(.*)Student Phone Number:',page_content)
        Student_phone_no = re.search('Student Phone Number:(.*)Student Email ID:', page_content)
        Student_email_id = re.search('Student Email ID:(.*)Student Registration No:', page_content)
        Student_reg_no = re.search('Student Registration No:(.*)Branch :', page_content)
        branch = re.search('Branch :(.*)Father Mobile Number:', page_content)
        father_mob_no = re.search('Father Mobile Number:(.*)Internship Start Date :', page_content)
        internship_start_date = re.search('Internship Start Date :(.*)Stipend:', page_content)
        Stipend = re.search('Stipend:(.*)Organization Name:', page_content)
        Org_name = re.search('Organization Name:(.*)Organization Address:', page_content)
        Organization_Address = re.search('Organization Address:(.*)City:', page_content)
        City = re.search('City:(.*)Duration in Months:', page_content)
        Duration_in_Months = re.search('Duration in Months:(.*)Internship Progress Report Title of the Project:', page_content)
        Title_of_project = re.search('Title of the Project:(.*)Industry Guide Feedback', page_content)
        Industry_guide_feedback = re.search('Resources provided/ Strengths/ Areas of Improvements(.*)Industry Guide Rating on a scale of 10', page_content)
        rating = re.search('10 being the best:(.*)Industry Guide Name:', page_content)
        Industry_Guide_Name = re.search('Industry Guide Name:(.*)Industry Guide Email ID:', page_content)
        IG_email = re.search('Industry Guide Email ID:(.*)Designation:', page_content)
        Designation = re.search('Designation:(.*)Industry Guide Mobile Number:', page_content)
        IG_mob_no = re.search('Industry Guide Mobile Number:(.*)Signature of Industry Guide', page_content)

        sheet1.write(line_no, 0, line_no )
        sheet1.write(line_no, 1, Student_name.group(1))
        sheet1.write(line_no, 2, Student_phone_no.group(1) )
        sheet1.write(line_no, 3, Student_email_id.group(1))
        sheet1.write(line_no, 4, Student_reg_no.group(1))
        sheet1.write(line_no, 5, branch.group(1))
        sheet1.write(line_no, 6, father_mob_no.group(1))
        sheet1.write(line_no, 7, internship_start_date.group(1) )
        sheet1.write(line_no, 8,  Stipend.group(1) )
        sheet1.write(line_no, 9, Org_name.group(1))
        sheet1.write(line_no, 10, Organization_Address.group(1))
        sheet1.write(line_no, 11, City.group(1) )
        sheet1.write(line_no, 12, Duration_in_Months.group(1))
        sheet1.write(line_no, 13, Title_of_project.group(1))
        sheet1.write(line_no, 14, Industry_guide_feedback.group(1)[2:])
        sheet1.write(line_no, 15, rating.group(1))
        sheet1.write(line_no, 16, Industry_Guide_Name.group(1))
        sheet1.write(line_no, 17, IG_email.group(1))
        sheet1.write(line_no, 18, Designation.group(1))
        sheet1.write(line_no, 19, IG_mob_no.group(1))

        line_no+=1
    except:

        print("Only PDF format is supported")
        sendingEmails("Emai id to which email to be send for correction")
wb.save('Internship Details.xls')
print("File Generated Successfully ...")
# except:
#
#     print("Error in File Generation ...")