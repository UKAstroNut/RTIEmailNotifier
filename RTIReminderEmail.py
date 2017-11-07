import smtplib
import gspread

from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from oauth2client.service_account import ServiceAccountCredentials



def ReadTodayUserList(): 
	# Use creds to create a client to interact with the Google Drive API
	scope = ['https://spreadsheets.google.com/feeds']
	creds = ServiceAccountCredentials.from_json_keyfile_name('FTRTIEmail.json', scope)
	client = gspread.authorize(creds)
 
	# Find a workbook by name and open the first sheet. Make sure you use the right name here.
	sheet = client.open("RTI Schedule").sheet1
 
	# Extract and print all of the values
	list_of_hashes = sheet.get_all_records()
	#print(list_of_hashes)

	# Get the indexes of the 4 cells with today's users email addresses in
	emailAddressIndexes = ['E5', 'F5', 'G5', 'H5']		# Indexes of email addresses for TOMORROWS observations
	emailAddressList = ['','']							# Contains the user email and one extra (e.g. mine for testing)
	emailAddressString = ""								# String of the email address, for the header

 	for x in range(0,4):
 		userEmailAddress = sheet.acell(emailAddressIndexes[x]).value
 		#print userEmailAddress
 		#SendReminderEmail(userEmailAddress)
 		emailAddressList[0] = userEmailAddress
 		emailAddressList[1] = "matthew.allen@faulkes-telescope.com"		# Also send a copy of the email to my email address

 		emailAddressString = userEmailAddress

		# Only send email if it contains an @ symbol - Doesn't falsely send if box contains "RTI OFFLINE"
		if(emailAddressList[0].find("@") >= 0):
 			SendReminderEmail(emailAddressList, emailAddressString)
 		else:
 			print "Email address isn't correct - \"" + emailAddressList[0] + "\""
	

def SendReminderEmail(emailAddressList, emailAddressString): 
	fromaddr = "faulkestelescope@gmail.com"
	toaddr = emailAddressList
	#bcc = emailAddressList
	msg = MIMEMultipart()
	msg['From'] = fromaddr
	msg['To'] = emailAddressString
	msg['Subject'] = "Reminder: Faulkes Telescope Observing Slot Tomorrow"
 
 	# The body text of the email
	body = "Dear Faulkes Telescope User,\n\n \
	This is to remind you that you have booked a Real Time observing slot on the Faulkes Telescope Project network of telescopes tomorrow. \
	To access the Real Time Interface and to begin your observations, login to your account at: https://observe.lco.global/ \n \
	If you need any help, you can find a guide to using the RTI and ask for support at the Faulkes Telescope Project website: www.faulkes-telescope.com \n \
	Good luck with your observations! \n \
	-The Faulkes Telescope Project Team \n \n \
	** Please do not reply to this email address. For support, please email support@faulkes-telescope.com **"
	msg.attach(MIMEText(body, 'plain'))
 
 	# Set all the variables for the email
 	toaddr = toaddr
	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.starttls()
	server.login(fromaddr, "Cardiff123")
	text = msg.as_string()
	server.sendmail(fromaddr, toaddr, text)
	server.quit()

	print "Email sent to user "+emailAddressString
	

ReadTodayUserList()






