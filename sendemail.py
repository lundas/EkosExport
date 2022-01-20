#!/usr/bin/env/ python

import os
import smtplib
import mimetypes
import logging
from email.MIMEMultipart import MIMEMultipart
from email import encoders
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage

# Set up logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
# Create file handler
fh = logging.FileHandler('DeliveryFormat/deliveryformat.log') # PATH to file on local machine
fh.setLevel(logging.INFO)
# Create formatter
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
# Add formatter to fh
fh.setFormatter(formatter)
# Add fh to logger
logger.addHandler(fh)

def send_gmail(
	self, 
	message, 
	subject, 
	email_to, 
	email_from, 
	password,
	reply_to='NoReply', 
	file_to_send=None
):
	'''Sends email with provided file as attachment from email_from to email_to.
	Username and Password provided for gmail acount that sends email. Only works
	with a gmail account.

	PARAMS
	---------------
	message : message to be included in the email

	subject : email subject

	email_to : email or list of emails of intended recipient(s)

	email_from : email that will be appear as the sender. Also used to log in to 
	email account using password provided

	password : password associated with email_from. Used to log in to email_from
	account in order to create and send email

	reply_to : email address to which all replies will be addressed

	file_to_send : attachment file
	'''

	logger.info('Sending Email')

	msg = MIMEMultipart()
	msg['From'] = email_from
	if type(email_to) == list:
		msg['To'] = ', '.join(email_to)
	else:
		msg['To'] = email_to
	msg['Reply-To'] = reply_to
	msg['Subject'] = subject

	body = MIMEText(message)
	msg.attach(body)

	if file_to_send == None:
		pass
	elif type(file_to_send) == list:	# Allows for multiple attachments
		for f in file_to_send:
			
			ctype, encoding = mimetypes.guess_type(f)
			if ctype is None or encoding is not None:
				ctype = 'application/octet-stream'

			maintype, subtype = ctype.split('/', 1)

			if maintype == 'application':
				fp = open(f, 'rb')
				att = MIMEApplication(fp.read(), _subtype=subtype)
				fp.close()
			elif maintype == 'text':
				fp = open(f)
				att = MIMEText(fp.read(), _subtype=subtype)
				fp.close()
			elif maintype == 'image':
				fp = open(f, 'rb')
				att = MIMEImage(fp.read(), _subtype=subtype)
				fp.close()
			elif maintype == 'audio':
				fp = open(f, 'rb')
				att = MIMEAudio(fp.read(), _subtype=subtype)
				fp.close()
			else:
				fp = open(f, 'rb')
				att = MIMEBase(maintype, subtype)
				att.set_payload(fp.read())
				fp.close()
				encoders.encode_base64(att)

			att.add_header('content-disposition', 'attachment', filename=os.path.basename(f))
			msg.attach(att)
	else:
		ctype, encoding = mimetypes.guess_type(file_to_send)
		if ctype is None or encoding is not None:
			ctype = 'application/octet-stream'

		maintype, subtype = ctype.split('/', 1)

		if maintype == 'application':
				fp = open(file_to_send, 'rb')
				att = MIMEApplication(fp.read(), _subtype=subtype)
				fp.close()
		elif maintype == 'text':
			fp = open(file_to_send)
			att = MIMEText(fp.read(), _subtype=subtype)
			fp.close()
		elif maintype == 'image':
			fp = open(file_to_send, 'rb')
			att = MIMEImage(fp.read(), _subtype=subtype)
			fp.close()
		elif maintype == 'audio':
			fp = open(file_to_send, 'rb')
			att = MIMEAudio(fp.read(), _subtype=subtype)
			fp.close()
		else:
			fp = open(file_to_send, 'rb')
			att = MIMEBase(maintype, subtype)
			att.set_payload(fp.read())
			fp.close()
			encoders.encode_base64(att)

		att.add_header('content-disposition', 'attachment', filename=os.path.basename(file_to_send))
		msg.attach(att)

	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.ehlo()
	server.starttls()
	server.ehlo()
	server.login(email_from, password)
	server.sendmail(email_from, email_to, msg.as_string())
	server.quit()

	return