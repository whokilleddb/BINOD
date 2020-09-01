#!/bin/python3
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
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
attachment_to_send=input(str("[+] Enter Image To Be Attached : "))
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
         print(f"[+] Name = {names[i]} Colleges = {college[i]} Email = {emails[i]}")
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
         mail_content = "Dear Sir/Maam,\n    It is our distinct pleasure to cordially invite you to the 7th edition of IEM Model United Nations (IEM MUN 2020), which will be held from October 10th to October 11th 2020.\n    Since its inception, IEM MUN has offered its delegates an unrivaled Model UN experience by conducting highly-personalized, engaging and dynamic crisis committees. This year, we again look forward to running committees that will cover a wide array of geographical areas, and we are confident that every delegate will be able to find a IEM MUN committee that sparks his or her interest.\n    IEM MUN 2020 will be conducted through an online medium due to this crisis period. We very much look forward to the most intellectually stimulating and enjoyable conference that we know IEM MUN 2020 will be, and I hope that you do too. In the meantime, please refer to our website at https://iemmun.in/ for a more detailed description of the conference as well as the committees. Early Bird Registrations have now started !!\n    We will be continuously updating the website to reflect the progress we make as we plan for IEM MUN 2020, but if you have any pressing questions or concerns, please do not hesitate to contact the IEM MUN 2020 Secretariat at https://iemmun.in/contact/.\n    Thank you for your time, and we hope to see you in October!\nKind Regards,\nAditya Poddar\nHead of Delegate Affairs"
         #Set message accordingly
         message = MIMEMultipart()
         message['From'] = sender_address
         message['To'] = receiver_address
         message['Subject'] = subject_line  #The subject line
         
         #The body and the attachments for the mail
         message.attach(MIMEText(mail_content, 'plain'))
         with open(attachment_to_send, 'rb') as fp:
               img = MIMEImage(fp.read())
               img.add_header('Content-Disposition', 'attachment', filename=attachment_to_send)
               message.attach(img)
         #Create SMTP session for sending the mail
         session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
         session.starttls() #enable security
         session.login(sender_address, sender_pass) #login with mail_id and password
         text = message.as_string()
         session.sendmail(sender_address, receiver_address, text)
         session.quit()
         print(f'[+] Mail Sent To {emails[i]}')
         i=i+1
except IndexError:
     print("\nDONE")
     sys.exit(0)
