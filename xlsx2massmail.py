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
     printf("""[-] Option Not Available
     [-] Exiting""")
     sys.exit(-1)
attachment_to_send=input(str("[+] Enter Path Of File To Be Attached : "))
#Get Data from xlsx 
mailsheet=xlrd.open_workbook(full_path)
sheet=mailsheet.sheet_by_index(0)
names=list()
emails=list()


#Fill the lists
i=1
try:
    while True:
         names.append(sheet.cell_value(i,2))#Change This
         emails.append(sheet.cell_value(i,7).lower())#Change This
         #Add Fields
         
         i=1+i
except:
     print("")


#Display Information
i=0
try:
     while True :
         print(f"[+] Name = {names[i]} Email = {emails[i]}")
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
Greetings Delegates! 

It is our distinct pleasure to cordially invite you to the 7th edition of IEM Model United Nations (IEM MUN 2020), which will be held from October 10th to October 11th 2020.


As IEMMUN 2019 was a grandeur, this time we present to you the 7th edition of Model United Nations by IEM and this time we are going International . Yes you heard it right it’s an International MUN.

We are back to create and invite you all to witness grandeur as the Pheonix rises from it’s ashes, the 7th edition of IEMMUN we present to you with a brand new and INTERNATIONAL edition of MUN organised by the Institute of Engineering and Management, Kolkata called the IEMMUN 2020 (#ConcordOfDiscord), in collaboration with the United Nations Information Centre (India and Bhutan) and UN@75 motions, scheduled on this 10th and 11th October via online platform.

Last time with your presence, you helped us make our conference a grand success and again in this present time we would like to have this opportunity of honour to have you back on our stage, showcasing your expertise. Since it’s inception, IEM MUN has offered its delegates an unrivalled and paramount MUN experience by conducting highly-personalised, engaging and dynamic crisis committees. This year, we again look forward to running committees that will cover a wide array of geographical areas, and we are confident that every delegate will be able to find a IEM MUN committee that sparks his or her interest.


We are completely aware of the tensed situations around prevailing due to this Covid-19 pandemic, and we assure you that this upcoming event will not only bring in complete but also recreation and belief to strive towards betterment together amidst the distress. 
We very much look forward for your participation in the most intellectual, stimulating and enjoyable conference, IEM MUN 2020.
The Phoenix rose from the ashes, the flame that was ignited by the collaborative efforts of every member associated with this MUN and may remain glorious even in the negative situations. 

The committees which we are hosting and their respective agenda along with respective Executive Board are as follows: 
1.  UNGA DISEC - Reduction of Military Budget : Executive Board are: Aswath Komath and Anurag Sengupta
2. WHA (World Health Assembly) - Impact of Covid-19 pandemic on the Sustainable Development Goals, with special emphasis on SDG-3 : Executive Board are : Chairperson,Tannistha Sinha and Vice Chairperson, Muksetul Islam Alif

Delegation application link: https://forms.gle/5NnN4gGjSRYTda7T8

Thanking you, 
Team IEMMUN'20.
 
________


For more details check our:

FACEBOOK PAGE
Link: https://www.facebook.com/IEM-MUN-101186998348738

INSTAGRAM PAGE: 
@iemmun

WEBSITE
Link:  iemmun.in


FOR ANY FURTHER QUERY PLEASE FEEL FREE TO CONTACT

Subhajit Pal
(Secretary General)
Phone: +91 8017363909 



Vikash Gupta
(Deputy Secretary General)
Phone: +91 7980917161


Debjeet Banerjee
( Director General )
Phone : +91 9163980758
"""
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
