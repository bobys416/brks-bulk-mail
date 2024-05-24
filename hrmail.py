import pandas as pd
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formataddr

# loading files 
excel_file='email.xlsx'
sheet_name='Sheet1'
df=pd.read_excel(excel_file,sheet_name=sheet_name)
your_name='Rishabh Kumar Sharma' 


gmail_user='rishabh.sharma41998@gmail.com'
gmail_password='qpwi fpqx yryq becl'
#EMail details
subject= "Application for Data Analyst / Business Analyst Position"
body="""
Dear Hiring Manager,
    I am excited to apply for the Data Analyst or Business Analyst position. With experience in data analysis, SQL, and machine learning, I am confident in my ability to contribute effectively to your team.

At Grid Infocom Pvt. Ltd., I led projects that improved process efficiency and provided critical insights. My skills in Python, Power BI, and SQL, along with my ability to work collaboratively, make me a strong fit for this role.

Please find my resume attached. I look forward to discussing how my background can benefit your team.

Thank you for considering my application.

Best regards,
Rishabh Kumar Sharma
+918853043355
rishabh.sharma41998@gmail.com
"""

attachment_path="Rishabh Kumar Sharma_Data Analyst.pdf"
# Counter for sent emails
emails_sent = 0
#Function
def send_email(to_email):
    msg=MIMEMultipart()
    msg['From']=formataddr((your_name,gmail_user))
    msg['To']=to_email
    msg['Subject']=subject
    
    #Attach the body
    msg.attach(MIMEText(body,'plain'))

    #open the file to be sent
    attachment=open(attachment_path, 'rb')  #'read binary'

    #Instance of MiMsbase
    p=MIMEBase('application','octet-stream')

    # To change the payload into encoded from
    p.set_payload(attachment.read())

    #encoder
    encoders.encode_base64(p)

    p.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(attachment_path)}")
    msg.attach(p)

    
     # Create SMTP session for sending the mail
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            print("SMTP connection started")
            server.login(gmail_user, gmail_password)
            print("Logged in to SMTP server")
            text = msg.as_string()
            server.sendmail(gmail_user, to_email, text)
            emails_sent += 1  # Increment the counter for sent emails
            print(f'Email sent to {to_email}')
    except Exception as e:
        print(f'Failed to send email to {to_email}: {e}')

# Iterate over the email addresses in the DataFrame
for index, row in df.iterrows():
    send_email(row['Email'])

# Print the total number of emails sent
print(f'Total emails sent: {emails_sent}')
