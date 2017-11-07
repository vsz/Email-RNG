import sys
import smtplib
import random
from openpyxl import Workbook
from openpyxl import load_workbook

	
def sendEmailToRecipient(server,user,recipients) :
	for r in recipients.keys() :
		msg = 'Oi {}! Estou testando algumas coisas! Este software deve estar funcional em pouco tempo :)'.format(r)
		server.sendmail(user, recipients[r], msg)
		print('E-mail was sent to {} at {}'.format(r, recipients[r]))
	
def authenticateOnGmail(usr,pwd) :
	server = smtplib.SMTP("smtp.gmail.com", 587)
	server.ehlo()
	server.starttls()
	server.login(usr,pwd)
	return server
	
def main(args):
    return 0

def getResponseByGender(ws) :
	m = ws['A3'].value
	f = ws['B3'].value
	'''
	if mmf >= 0 :
		m = 0
		f = mmf
	if mmf < 0 :
		m = -mmf
		f = 0
	'''
	return {'m':m , 'f':f}

def selectRecipients(ws, num) :
	# Get recipients that havent received an email yet
	available_recipients = {}
	
	for row in ws.iter_rows(min_row=2, min_col=1, max_row=ws.max_row, max_col=3) :
		if row[2].value == 0 :
			available_recipients[row[0].value] = row[1].value
			
	#Select random recipient from available
	selected_recipients = {}
	
	for _ in range(0,int(num)) :
		r = random.choice(list(available_recipients))
		selected_recipients[r] = available_recipients[r]
		del available_recipients[r]
		
	return selected_recipients

if __name__ == '__main__' :
	
	'''
	# Marks on excel table the selected recipients
	for row in ws_h.iter_rows(min_row=2, min_col=1, max_row=ws_h.max_row, max_col=3) :
		if row[0].value in selected_male_recipients :
			print("Achei")
			row[2].value = 1
			
	wb.save(filename)
	'''
	
	print("Iniciando...")
	
	# Parse login info
	user = str(sys.argv[1])
	password = str(sys.argv[2])
	
	
	# Load recipients file to memory
	print("Carregando arquivo...")
	filename = 'emails.xlsx'
	wb = load_workbook(filename)
	
	
	# Parse number of questionnaire responses by gender
	num = getResponseByGender(wb['Geral'])
	print("Há {} respostas de homens e {} respostas de mulheres.".format(num['m'], num['f']))


	# Parse amount of recipients
	num_male_recipients = input("Digite o numero de e-mails para mandar para homens: ")
	num_female_recipients = input("Digite o numero de e-mails para mandar para mulheres: ")
	input("Pressione Enter para mandar os e-mails!")
	
	
	# Select male recipients
	selected_male_recipients = selectRecipients(wb['H'], num_male_recipients)
	
	print("Selected male recipients")
	for r in selected_male_recipients :
		print("{} {}".format(r,selected_male_recipients[r]))
	
	
	# Select female recipients
	selected_female_recipients = selectRecipients(wb['M'], num_female_recipients)
	
	print("Selected female recipients")
	for r in selected_female_recipients :
		print("{} {}".format(r,selected_female_recipients[r]))
	
	
	server = authenticateOnGmail(user,password)
	#sendEmailToRecipients(server,user,recipients)
	
	server.close()
	
	input("Press Enter to continue")
	sys.exit(main(sys.argv))
