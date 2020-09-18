#!/bin/python3
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
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
print("""[+] Enter Type Of File You Want To Send :
[0] None
[1] Images
[2] Pdf/Ppt
[3] Text Files""")
attach_flag=int(input("[+] Enter Choice : "))
if attach_flag not in range(0,4):
     print("""[-] Option Not Available
     [-] Exiting""")
     sys.exit(-1)
if attach_flag != 0 :
    attachment_to_send=input(str("[+] Enter Path Of File To Be Attached : "))
#Get Data from xlsx
mailsheet=xlrd.open_workbook(full_path)
sheet=mailsheet.sheet_by_index(0)
#Change These
country=list()
emails=list()
committee=list()

#Fill the lists
i=0
try:
    while True:
         country.append(sheet.cell_value(i,2))#Change This
         emails.append(sheet.cell_value(i,0).lower())#Change This
         committee.append(sheet.cell_value(i,1))#Change This
         #Add Fields

         i=1+i
except:
     print("")


#Display Information
i=0
try:
     while True :
         if emails[i]=="":
              del emails[i]
              continue
         print(f"[+] Country = {country[i]} Email = {emails[i]} Commitee={committee[i]}") #Change This
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
         #Change This Message
         mail_content = """
Dear Madam/Sir,

Greetings from the IEM Model UN 2020 Secretariat.

It is with great pleasure that we inform you that your provisional allotment has been confirmed in the {} committee of the IEMMUN 2020
Allotted Country- {}


It is to kindly note that only after the successful payment of delegate fee and accommodation fee if opted for, this would be your final allotment and the participation will be confirmed.


You can view the allotments here:-

https://drive.google.com/file/d/1YBOde2vTvsxgfkWsej8C2-PFYLg3Gu08/view?usp=sharing

Fee Structure & Payment Procedure for IEM MUN 2020:

Registration Fee:
INR 500 (Indian Single Delegate Member) and INR 900 (Indian Double Delegate Member )
$12  (Foreign Single Delegate Member)
$22  (Foreign Double Delegate Member )


Country Upgrade Request will be entertained only once the Payment is done. The last date for payment is  31st August, 2020 for provisional first round allotment


For payment related queries, write to us at:  officialiemmun@gmail.com   or call us at
Vikash Gupta
+917980917161

Payments can be done according to above details through Online Payment :
https://www.explara.com/e/iem-model-united-nations
Paytm- +919330994191
PhonePay- +917980917161

For allotment related queries,

write to us at: officialiemmun@gmail.com

or call at us :
Subhajit Pal
+91 80173 63909

Vikash Gupta
+917980917161

Debjeet Baneerjee
+91 91639 80758

Aditya Poddar
+91 94322 80995
""".format(committee[i],country[i])
         #Set message accordingly
         if receiver_address=="":
              print("""[-] Empty Name List
[-] Exiting""")
              sys.exit(-1)
         message = MIMEMultipart()
         message['From'] = sender_address
         message['To'] = receiver_address
         message['Subject'] = subject_line  #The subject line

         #The body and the attachments for the mail
         message.attach(MIMEText(mail_content, 'plain'))
         if attach_flag==1 :
               with open(attachment_to_send, 'rb') as fp:
                    img = MIMEImage(fp.read())
                    img.add_header('Content-Disposition', 'attachment', filename=attachment_to_send)
                    message.attach(img)
         if attach_flag == 2 :

              pdf = MIMEApplication(open(attachment_to_send, 'rb').read())
              pdf.add_header('Content-Disposition','attachment',filename=attachment_to_send)
              message.attach(pdf)
         if attach_flag == 3:
              message.attach(MIMEText(open(attachment_to_send).read()))
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
