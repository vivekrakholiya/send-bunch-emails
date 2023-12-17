from email.message import EmailMessage
import ssl
import smtplib
import pandas as pd


# Strating Lines

print('----------------------------------------------------------')
print('----------------------------------------------------------')
print('Mailr script is running ')
print('----------------------------------------------------------')
print('----------------------------------------------------------')





# Load the Excel file
excel_file = 'Book1.xlsx'
data = pd.read_excel(excel_file)
# print(data)

data_list = data.values.tolist()

#  add mail details
email_sender = 'Your Email'
email_password = 'Add your pass key'
subject = 'Add Subject'
body = 'add body'

context = ssl.create_default_context()

with smtplib.SMTP_SSL('smtp.gmail.com',465,context=context) as smtp:
    smtp.login(email_sender,email_password)

 
    for index,mail in enumerate(data_list):
        email_receiver = str(''.join(mail))
        # print(email_receiver)
        em = EmailMessage()
        em['From'] = email_sender
        em['TO'] = email_receiver
        em['Subject'] = subject
        em.set_content(body)

        
        smtp.sendmail(email_sender,email_receiver,em.as_string())
        print(f'No.{index} Email Send To {email_receiver} successfully !!!')
        print('----------------------------------------------------------')

print('----------------------------------------------------------')
print('----------------------------------------------------------')
print('----------------------------------------------------------')
print(f'Total {len(data_list)} mails sends succesfully')
print('----------------------------------------------------------')
print('----------------------------------------------------------')
