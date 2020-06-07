import smtplib
import xlrd 
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
mail_content = '''Hi,
Thanks for taking part of the COVID-19 quiz.
-Principal
'''
#The mail addresses and password
sender_address = 'sngcaspaingottoornoreply@gmail.com'
sender_pass = 'sngcasMG123!'


workbook = xlrd.open_workbook('covid_quiz.xlsx')
sheet = workbook.sheet_by_index(0)

for i in range(1,155) :
	#Setup the MIME
  	name= str(sheet.cell(i, 2).value)
  	email= str(sheet.cell(i, 0).value)


	message = MIMEMultipart()
	message['From'] = sender_address
	message['To'] = email
	message['Subject'] = 'Participation certificate of Covid-19 quiz.'
	#The subject line
	#The body and the attachments for the mail
	message.attach(MIMEText(mail_content, 'plain'))
	attach_file_name = 'certificates/'+name+'.png'
	attach_file = open(attach_file_name, 'rb') # Open the file as binary mode
	payload = MIMEBase('application',"pdf", Name= name+'.png')	
	payload.set_payload((attach_file).read())
	encoders.encode_base64(payload) #encode the attachment
	#add payload header with filename
	payload.add_header('Content-Decomposition',  "attachment", filename=name+".png")
	message.attach(payload)
	#Create SMTP session for sending the mail
	session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
	session.starttls() #enable security
	session.login(sender_address, sender_pass) #login with mail_id and password
	text = message.as_string()
	session.sendmail(sender_address, email, text)
	session.quit()

