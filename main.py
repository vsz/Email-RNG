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

def getResponseByGenderCount(ws) :
	m = ws['A3'].value
	f = ws['B3'].value

	return {'m':m , 'f':f}
	
def getRecipientCount(m_ws,f_ws) :
	m = m_ws.max_row-1
	f = f_ws.max_row-1
	
	return {'m':m, 'f':f}


def getAvailableRecipients(ws) :
	# Get recipients that havent received an email yet
	available_recipients = {}
	
	for row in ws.iter_rows(min_row=2, min_col=1, max_row=ws.max_row, max_col=3) :
		if row[2].value == 0 :
			available_recipients[row[0].value] = row[1].value
			
	return available_recipients

def getAvailableRecipientCount(m_d,f_d) :
	m = len(m_d)
	f = len(f_d)
	return {'m':m, 'f':f}

def drawRandomRecipients(recipients,num) :
	#Select random recipient from available
	selected_recipients = {}
	
	for _ in range(0,int(num)) :
		r = random.choice(list(recipients))
		selected_recipients[r] = recipients[r]
		del recipients[r]
		
	return selected_recipients

def updateRecipientTable(fn,wb,ws,r) :
	# Marks on excel table the selected recipients
	for row in ws.iter_rows(min_row=2, min_col=1, max_row=ws.max_row, max_col=3) :
		if row[0].value == r :
			row[2].value = 1
			
	wb.save(fn)

def query_yes_no(msg) :
	yes = {'yes','y','s','sim'}
	no = {'no','n','nao'}
	
	while True :
		choice = input(msg + " s/n ? ").lower()
		
		if choice in yes :
			return True
		if choice in no :
			return False
		else :
			print("Escreva S ou N")

if __name__ == '__main__' :
	
	'''

	'''
	
	print("Iniciando...")
	'''
	# Parse login info
	user = str(sys.argv[1])
	password = str(sys.argv[2])
	'''
	
	# Load recipients file to memory
	print("Carregando arquivo...")
	filename = 'emails.xlsx'
	wb = load_workbook(filename)
	geral_sheet = wb['Geral']
	male_recipients_sheet = wb['H']
	female_recipients_sheet = wb['M']
	
	# Get available recipients
	available_male_recipients = getAvailableRecipients(male_recipients_sheet)
	available_female_recipients = getAvailableRecipients(female_recipients_sheet)
	
	# Stats
	num_response = getResponseByGenderCount(geral_sheet)
	num_recipients = getRecipientCount(male_recipients_sheet,female_recipients_sheet)
	num_available_recipients = getAvailableRecipientCount(available_male_recipients,available_female_recipients)
	
	print("Total cadastrados:             {0:4} Homens e {1:4} Mulheres.".format(num_recipients['m'], num_recipients['f']))
	print("Total cadastrados disponíveis: {0:4} Homens e {1:4} Mulheres.".format(num_available_recipients['m'], num_available_recipients['f']))
	print("Total respostas              : {0:4} Homens e {1:4} Mulheres.".format(num_response['m'], num_response['f']))

	# Draw phase
	num_desired_recipients = {}
	
	accepted = False
	while accepted == False :
		num_desired_recipients['m'] = int(input("Digite o numero de e-mails para mandar para homens: "))
		num_desired_recipients['f'] = int(input("Digite o numero de e-mails para mandar para mulheres: "))
		if num_desired_recipients['m']  <= num_available_recipients['m'] and  \
			num_desired_recipients['f'] <= num_available_recipients['f'] :
			accepted = True
		else :
			print("Escolha um número menor.")
			
	# Select female recipients
	accepted = False
	while accepted == False :
		selected_female_recipients = drawRandomRecipients(available_female_recipients.copy(), num_desired_recipients['f'])
		selected_male_recipients = drawRandomRecipients(available_male_recipients.copy(), num_desired_recipients['m'])
	
		print("Homens selecionados:")
		for r in selected_male_recipients :
			print("{0:30} {1:40}".format(r,selected_male_recipients[r]))
		
		print("")
		
		print("Selected female recipients")
		for r in selected_female_recipients :
			print("{0:30} {1:40}".format(r,selected_female_recipients[r]))
		
		accepted = query_yes_no("Você aceita a seleção?")
	
	input("Pressione qualquer tecla para enviar os e-mails aos sorteados.")	
	
	#server = authenticateOnGmail(user,password)
	
	# Send to female recipients
	print("Enviando para as mulheres...")
	for r in selected_female_recipients :
		#sendEmailToRecipients(server,user,r)
		updateRecipientTable(filename,wb,female_recipients_sheet,r)
		print("E-mail enviado para {}. Tabela atualizada!".format(r))
	
	print("Enviando para os homens...")
	for r in selected_male_recipients :
		#sendEmailToRecipients(server,user,r)
		updateRecipientTable(filename,wb,male_recipients_sheet,r)
		print("E-mail enviado para {}. Tabela atualizada!".format(r))
	
	
	#server.close()
	
	input("Press Enter to continue")
	sys.exit(main(sys.argv))
