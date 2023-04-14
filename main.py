from operator import index
from random import randint, choices
from email import encoders
import pandas as pd
from pdf2image import convert_from_path
import os
from email.message import EmailMessage
import smtplib
import logging
import time
import sys
import os
import string
bodies = ['body1.txt']

Convert=False


logging.basicConfig(filename='mail.log', level=logging.DEBUG)
totalSend = 1
if(len(sys.argv) > 1):
    totalSend = int(sys.argv[1])
emaildf = pd.read_csv('senderSMTP.csv')
contactsData = pd.read_csv('receiver.csv')
subject = pd.read_csv('subjects.csv')


def convert_and_save(pdf_path):
    global Convert
    poopler_path=r"C:\Program Files\poppler-23.01.0\Library\bin"
    images = convert_from_path(pdf_path,50,poppler_path=poopler_path)
    for i in range(len(images)):
	# Save pages as images in the pdf
	    images[i].save('images/page'+ str(i) +'.jpg', 'JPEG')
    Convert=True    
    return len(images)
def remove(lenght):
    for i in range(lenght):
        os.remove('images/page'+ str(i) +'.jpg')  

def send_mail(name, email, emailId, password, subjects, senderName, images, bodyFile):
    newMessage = EmailMessage()

    # Invoice Number and Subject
    invoiceNo = randint(10000000, 99999999)
    randomString = ''.join(choices(string.ascii_uppercase, k=4))
    subject = subjects + "   #" + randomString + \
         str(invoiceNo)
    num = randint(1111, 9999)
    sender = f"{senderName}  #{randomString}{num}<{emailId}>"
    newMessage['Subject'] = subject
    newMessage['From'] = sender
    newMessage['To'] = email
    transaction_id = randint(10000000, 99999999)

    # Mail Body Content
    body = open(bodyFile, 'r').read()
    body = body.replace('$name', str(name))
    body = body.replace('$email', str(email))
    body = body.replace('$invoice_no', str(transaction_id))
    
    
    newMessage.set_content(body)

# file atteched section
# sending data will be imported here or link of any file also pasted here

    # with open(f"pdfFiles/{pdfFiles}.pdf","rb") as f:
    #     file_data=f.read()
    #     File_name=f"{pdfFiles}"
    #     newMessage.add_attachment(file_data,maintype="application",subtype="pdf",filename=File_name)
    for i in images:
        with open(f"images/{i}.jpg","rb") as f:
            file_data=f.read()
            File_name=f"{i}.jpg"
            newMessage.add_attachment(file_data,maintype="application",subtype="image",filename=File_name)  
   





    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
            smtp.starttls()
            smtp.login(emailId, password)
            smtp.send_message(newMessage)
            smtp.quit()


        print(f"send to {email} by {emailId} successfully : {totalSend}")
        logging.info(
            f"send to {email} by {emailId} successfully : {totalSend}")


    except smtplib.SMTPResponseException as e:
        error_code = e.smtp_code
        error_message = e.smtp_error
        print(f"send to {email} by {emailId} failed:{totalSend}")
        logging.info(f"send to {email}  by {emailId} failed:{totalSend}")
        print(f"error code: {error_code}")
        print(f"error message: {error_message}")
        logging.info(f"error code: {error_code}")
        logging.info(f"error message: {error_message}")

        remove_email(emailId, password)


def start_mail_system():
    global Convert
    global totalSend
    j = 0
    k = 0
    l = 0
    m = 0


    for i in range(len(contactsData)):
        emaildf = pd.read_csv('senderSMTP.csv')
        if(j >= len(emaildf)):
            j = 0
        
        # this is for subject
        subject = pd.read_csv('subjects.csv')
        if(l >= len(subject)):
            l = 0


        #this is pdf files section
        pdfFiles=pd.read_csv('pdfFiles.csv')
        if(m >= len(pdfFiles)):
            m = 0
        if Convert==False:    
            path1=("pdfFiles/"+str(pdfFiles.iloc[m]['pdfFiles']+".pdf"))
            lenght=convert_and_save(path1)  
            files=[]
            for a in range(lenght):
                files.append('page'+str(a))
     



        

        time.sleep(0.1)
        send_mail(contactsData.iloc[i]['name'], contactsData.iloc[i]['email'], emaildf.iloc[j]['email'],
                  emaildf.iloc[j]['password'], subject.iloc[l]['subject'], subject.iloc[l]['senderName'], files, bodies[k])
        # send_mail(contactsData.iloc[i]['name'], contactsData.iloc[i]['email'], emaildf.iloc[j]['email'],
        #           emaildf.iloc[j]['password'], subject.iloc[l]['subject'], subject.iloc[l]['senderName'], pdfFiles.iloc[m]['pdfFiles'], bodies[k])          
        totalSend += 1
        j = j + 1
        k = k + 1
        l = l + 1
        m = m + 1

        if j == len(emaildf):
            
            j = 0
        
        if k == len(bodies):
            k = 0
        
        if(l == len(subject)):
            l = 0
        
        if(m == len(pdfFiles)):
            m = 0

        
        

        print()    
    print("all mails send succesfully !")
    remove(lenght)
    quit()


   
    



def remove_email(emailId, password):
    df = pd.read_csv('senderSMTP.csv')
    index = df[df['email'] == emailId].index
    df.drop(index,inplace=True)
    df.to_csv('senderSMTP.csv', index=False)
    print(f"{emailId} removed from senderSMTP.csv")
    logging.info(f"{emailId} removed from senderSMTP.csv")


try:
    for i in range(6):
        start_mail_system()
except KeyboardInterrupt as e:
    print(f"\n\ncode stopped by user")
