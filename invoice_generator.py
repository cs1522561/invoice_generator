import os
import smtplib
import argparse
import pandas as pd
from fpdf import FPDF
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

INDIVIDUAL_RATE = 30
GROUP_RATE = 15
TODAY = date.today()

# Use smtp to send the emails
def SendEmail(userEmail, userPassword, parentEmail, subject, body, filePath):
    msg = MIMEMultipart()
    msg["From"] = userEmail
    msg["To"] = parentEmail
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain"))
    # add the invoice to the email
    with open(filePath, "rb") as file:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition", f"attachment; filename={os.path.basename(filePath)}"
        )
        msg.attach(part)
    # send the email
    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(userEmail, userPassword)
            server.send_message(msg)
        print(f"Email sent successfully to {parentEmail}.")
    except Exception as e:
        print(f"Error sending email to {parentEmail}: {e}")

#create the invoices
def GenerateInvoice(lessonType, studentInstrument, parentName, school, lessons, termNum, termAmount, lessonRate, invoiceFilename):
    pdf = FPDF()
    pdf.add_page()
    #cell(width, height, text, border, end line, align)

    pdf.set_font('Arial','',10)

    #vertical spacer
    pdf.cell(180 ,7,'',0,1)

    #personal address
    pdf.cell(150,5,'',0,0)
    pdf.cell(30 ,5,'Campbell Smith',0,1,'R')

    pdf.cell(150,5,'',0,0)
    pdf.cell(30 ,5,'246 Clyde Street',0,1,'R')

    pdf.cell(150,5,'',0,0)
    pdf.cell(30 ,5,'Hamilton 3216',0,1,'R')

    #vertical spacer
    pdf.cell(180 ,5,'',0,1)

    #phone number
    pdf.cell(150,5,'',0,0)
    pdf.cell(30 ,5,'021-0810-9807',0,1,'R')

    #subject
    pdf.cell(20 ,5,'',0,0)
    pdf.cell(125 ,5,f'To:  {parentName}',0,1)

    #invoice number and date
    pdf.cell(130,5,'',0,0)
    pdf.cell(20 ,5,'Invoice #:',0,0, 'R')
    pdf.cell(30 ,5,f'2300{termNum}',0,1,'R')

    pdf.cell(20,5,'',0,0)
    pdf.cell(90,5,f'Re:  {school} Brass Lessons',0,0)
    pdf.cell(20 ,5,'',0,0)
    pdf.cell(20 ,5,'Dated:',0,0, 'R')
    pdf.cell(30 ,5,f'{TODAY.strftime('%d/%m/%Y')}',0,1,'R')

    #vertical spacer
    pdf.cell(180 ,5,'',0,1)

    #invoice contents
    pdf.set_font('Arial','B',10)
    pdf.cell(10,5,'',0,0)
    pdf.cell(100,5,'Item',1,0)
    pdf.cell(20 ,5,'Quantity',1,0)
    pdf.cell(20 ,5,'Rate',1,0)
    pdf.cell(30 ,5,'Line total',1,1)

    pdf.set_font('Arial','',10)
    pdf.cell(10,5,'',0,0)
    pdf.cell(100,5,f'{lessonType} {studentInstrument} lesson',1,0)
    pdf.cell(20 ,5,f'{lessons}',1,0,'R')
    pdf.cell(20 ,5,f'${lessonRate}.00',1,0,'R')
    pdf.cell(30 ,5,f'${termAmount}.00',1,1,'R')

    pdf.cell(110,5,'',0,0)
    pdf.cell(40 ,5,'  Total',1,0,)
    pdf.cell(30 ,5,f'${termAmount}.00',1,1,'R')

    #vertical spacer
    pdf.cell(180 ,20,'',0,1)

    #bank number
    pdf.cell(10,5,'',0,0)
    pdf.cell(100,5,'Payments should be made to 06-0603-0098764-00',0,1)
        
    #vertical spacer
    pdf.cell(180 ,5,'',0,1)

    pdf.cell(20,5,'',0,0)
    pdf.cell(80,5,'With thanks',0,1)

    pdf.cell(40,5,'',0,0)
    pdf.cell(60,5,'Campbell Smith',0,1)

    pdf.output(invoiceFilename)
    print(f"Generated {invoiceFilename}")


#pull the information from the csv
def ExtractCsv(csvFilePath, userEmail, userPassword):
    df = pd.read_csv(csvFilePath)
    #iterate through the csv rows
    for index, row in df.iterrows():
        childName = row['Child Name']
        lessonType = row['Lesson Type']
        studentInstrument = row['Instrument']
        parentEmail = row['Email']
        parentName = row['Parent']
        school = row['School']
        lessons = row['Lessons Completed']
        termNum = row['Term']
        if lessonType == 'Individual':
            termAmount = lessons * INDIVIDUAL_RATE
            lessonRate = INDIVIDUAL_RATE
        else:
            termAmount = lessons * GROUP_RATE
            lessonRate = GROUP_RATE
        
        #create the pdf filename
        invoiceFilename = f'./term{termNum}_invoice_{parentName.replace(' ', '_')}.pdf'
        
        #call the function to create the invoice pdf
        GenerateInvoice(lessonType, studentInstrument, parentName, school, lessons, termNum, termAmount, lessonRate, invoiceFilename)

        #content of the email to send
        subject=f'{childName} term {termNum} invoice'
        body=f'Hi {parentName.split()[0]}\n\n I have attached the invoice for term {termNum} to this email. Please can this be completed before the start of term.\n\nKind regards\nCampbell Smith'

        #call the function to send the email        
        SendEmail(userEmail, userPassword, parentEmail, subject, body, invoiceFilename)



if __name__ == '__main__':
    #take the arguments and save as variables
    parser = argparse.ArgumentParser()
    parser.add_argument('csvFile', type=str)
    parser.add_argument('emailUser', type=str)
    parser.add_argument('emailPassword', type=str)
    args = parser.parse_args()

    csvFilePath = args.csvFile
    userEmail = args.emailUser
    userPassword = args.emailPassword
    
    #call the function to extract the csv data
    ExtractCsv(csvFilePath, userEmail, userPassword)
    
    print("done")