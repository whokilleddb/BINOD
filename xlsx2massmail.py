#!/bin/python3
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import sys
import xlrd

#Cause We All Loved It Here :""")
ted="""\
_________________________________________
/ “There is a word in German lebenslanger \\
| schicksalschatz. And the closest        |
| translation would be ‘lifelong treasure |
| of destiny’. And Victoria is wunderbar, |
| but she is not my lebenslanger          |
| schicksalschatz. She is my beinahe      |
| leidenschaft gegenstand. It means ‘the  |
| thing that is almost the thing that you |
\ want, but it’s not quite it."           /
 -----------------------------------------
   \\
    \\
        .--.
       |o_o |
       |:_/ |
      //   \ \\
     (|     | )
    /'\_   _/`\\
    \___)=(___/
"""
print(ted)

#The mail addresses and password
sender_address = input(str("[+] Enter Email ID : "))
sender_pass = input(str("[+] Enter Your Password : ")) #Can be substituted with getpass() if you do not want to echo the password
subject_line = input(str("[+] Enter Subject Line : "))
full_path = input(str("[+] Enter Full Path Of Excel Sheet : ")) #Or just the name if it's in the same directory

#Get Data from xlsx 
mailsheet=xlrd.open_workbook(full_path)
sheet=mailsheet.sheet_by_index(0)
names=list()
emails=list()
college=list()

#Fill the lists
i=1
try:
    while True:
         names.append(sheet.cell_value(i,1))
         emails.append(sheet.cell_value(i,2))
         college.append(sheet.cell_value(i,4))
         i=1+i
except:
     print("")


#Display Information
i=0
try:
     while True :
         print(f"Name = {names[i]} Colleges = {college[i]} Email = {emails[i]}")
         i=i+1
except IndexError:
     print("")

#Asking Confirmation
print("[+] Continue To Sending Mails ? ")
flag=input(str("[+] Enter Option (Yes/No) : "))
flag=flag.upper()
if flag=="NO" or flag=="N" :
     print("[+] Exiting")
     sys.exit(0)

#Send Mail
i=0
try:
     while True :
         receiver_address = emails[i]
         #Setup the MIME
         mail_content = f"Hello {names[i]} from {college[i]}\n\nBest One : https://www.youtube.com/watch?v=MYnhPtsRQjA\n\n~WhoKilledDB"
         #Set message accordingly
         message = MIMEMultipart()
         message['From'] = sender_address
         message['To'] = receiver_address
         message['Subject'] = subject_line  #The subject line
         
         #The body and the attachments for the mail
         message.attach(MIMEText(mail_content, 'plain'))
         
         #Create SMTP session for sending the mail
         session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
         session.starttls() #enable security
         session.login(sender_address, sender_pass) #login with mail_id and password
         text = message.as_string()
         session.sendmail(sender_address, receiver_address, text)
         session.quit()
         print(f'Mail Sent To {emails[i]}')
         i=i+1
except IndexError:
     print("\nDONE")
     sys.exit(0)