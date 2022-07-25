import pandas as pd
import smtplib
from email.message import EmailMessage

df = pd.read_excel('physician_list.xlsx')
name_list = df['Name']
specialty = df['Specialty']
email_list = df['Email']
message = """

Dr. {}, 

My name is John Smith and I am a pre-med student in my senior year at Example University. I also work as a PCNA at the Local Hospital. I am currently looking for some opportunities to gain more shadowing experience, and I am interested to see what practicing as {} {} would be like. Would I be able to come by and follow you for a day? If your schedule would allow it, please let me know when you'd be available. I would greatly appreciate your help.

Thanks, 
John Smith
"""
vowels = "aeiou"

for i in range(1):
  if specialty[i][0].lower() in vowels:
    article = 'an'
  else:
    article = 'a'

  sent_from = 'my_email@gmail.com'
  password = 'googleapppassword'
  to = email_list[i]
  subject = 'Shadowing'

  msg = EmailMessage()
  msg['Subject'] = subject
  msg['To'] = to
  msg['From'] = sent_from
  msg.set_content(message.format(name_list[i], article, specialty[i]))

  try:
    server_ssl = smtplib.SMTP_SSL('smtp.gmail.com')
    server_ssl.login(sent_from, password) 
    server_ssl.send_message(msg)
    server_ssl.quit()
    print('sent email')
  except Exception as e:
    print(e)
