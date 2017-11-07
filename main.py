import sys

def getRandomRecipientsFromFile() :
	recipients = {};
	recipients.update({"Mr. Vits" : "vszeni@gmail.com"})
	recipients.update({"Mrs. Pipi" : "fernanda@almeidamarcon.com"})
	
	return recipients
	
def sendEmailToRecipient(recipients) :
	for r in recipients.keys() :
		print('E-mail was sent to {} at {}'.format(r, recipients[r]))

def sendRandomEmail() :
	recipients = getRandomRecipientsFromFile()
	sendEmailToRecipient(recipients)
	

def main(args):
    return 0



if __name__ == '__main__' :
	print("Starting...")
	sendRandomEmail()
	input("Press Enter to continue")
	sys.exit(main(sys.argv))
