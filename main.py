import sys
import smtplib

def getRandomRecipientsFromFile() :
	recipients = {};

	return recipients
	
def sendEmailToRecipients(server,user,recipients) :
	for r in recipients.keys() :
		msg = 'Oi {}! Estou testando algumas coisas...'.format(r)
		server.sendmail(user, recipients[r], msg)
		print('E-mail was sent to {} at {}'.format(r, recipients[r]))

def sendRandomEmails(srv) :
	recipients = getRandomRecipientsFromFile()
	sendEmailToRecipients(srv,recipients)
	
def authenticateOnGmail(usr,pwd) :
	server = smtplib.SMTP("smtp.gmail.com", 587)
	server.ehlo()
	server.starttls()
	server.login(usr,pwd)
	return server
	

def main(args):
    return 0



if __name__ == '__main__' :
	print("Starting...")
	
	user = str(sys.argv[1])
	password = str(sys.argv[2])
	
	server = authenticateOnGmail(user,password)
	recipients = getRandomRecipientsFromFile()
	sendEmailToRecipients(server,user,recipients)
	
	server.close()
	
	input("Press Enter to continue")
	sys.exit(main(sys.argv))
