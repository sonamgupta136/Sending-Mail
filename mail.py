import pandas as p
import smtplib as sm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

data = p.read_excel("st.xlsx")
print(type(data))
Email_col=data.get("Email")
list_of_emails=list(Email_col)
print(list_of_emails)

try:
     server = sm.SMTP("smtp.gmail.com",587)
     server.starttls()
     server.login("sona.gupta2099@gmail.com","**********")
     from_="sona.gupta2099@gmail.com"
     to_=list_of_emails
     message=MIMEMultipart("alternative")
     message['Subject']="This is just testing message"
     message["from"]="sona.gupta2099@gmail.com"

     html='''
     <html>
     <head>
     
     </head>
     <body>
          <h1> This is sonam</h1>
          <h2> i am a b.tech student </h2>
          <p> i am testing my project </p>
          <button style="padding:20px;background:green;color:white;>Verify</button>
     </body>
     </html>
     '''
     text=MIMEText(html,"html")
     message.attach(text)
     server.sendmail(from_,to_,message.as_string())
     print("message has been send to emails")




except Exception as e:
     print(e)
