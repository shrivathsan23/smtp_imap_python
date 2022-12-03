# Importing necessary modules
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
from email.message import EmailMessage
import smtplib
import imaplib
import email
import os
import time

# Attachment Directory

attachment_dir = 'F:/Mail Attachments'

# Creating SMTP & IMAP objects
smtp = None
imap = None
Gmail_Login = True

def objectCreation(email_var):
	global smtp
	global imap
	global Gmail_Login

	# Clearing previous user contents

	mail_list.clear()
	scrl.delete(1.0, END)
	listbox.delete(0, END)
	attach_label.config(text = "")
	no_of_mails.config(text = "")

	if email_var.get().split('@')[-1] == 'gmail.com':
		smtp = smtplib.SMTP_SSL('smtp.gmail.com', 465)
		imap = imaplib.IMAP4_SSL('imap.gmail.com')
		Gmail_Login = True
	else:
		smtp = smtplib.SMTP('localhost', 587)
		imap = imaplib.IMAP4('localhost', 143)
		Gmail_Login = False

# Creating Toggle function


def toggle(tog = [0]):
	tog[0] = not tog[0]
	if tog[0]:
		top.deiconify()
		top.state('zoomed')
		root.withdraw()
	else:
		top.withdraw()
		print(smtp.quit())
		print(imap.logout())
		root.deiconify()

# Creating Login Check function (SMTP)


def login_check_smtp(email_var, password_var):
	try:
		print(smtp.login(email_var.get(), password_var.get()))
		return 1
	except:
		return 0

# Creating Login Check function (IMAP)


def login_check_imap(email_var, password_var):
	try:
		print(imap.login(email_var.get(), password_var.get()))
		return 1
	except:
		return 0

# Creating Toggle & Login function


def toggle_and_login_check(email_var, password_var):
	if email_var.get() == '' or password_var.get() == '':
		print('Either empty Username or Password')
		root_info_label.config(text = 'Username or Password is Empty')

	else:
		objectCreation(email_var)
		if login_check_smtp(email_var, password_var) == 1 and login_check_imap(email_var, password_var) == 1:
			toggle()
		else:
			root_info_label.config(text = 'Login Unsuccessful')

# Creating Send Mail function


def send_mail(from_addr, to_var, cc_var, bcc_var, subject_var, body_var, attach, attachments = []):
	to = to_var.get()
	cc = cc_var.get()
	bcc = bcc_var.get()
	subject = subject_var.get()
	body = body_var.get(1.0, END).strip()
	print('To addr', to)
	print('Subject', subject)
	print('Body', body)
	if to == '' or subject == '' or body == '':
		messagebox.askokcancel('Error!!', 'Either Empty To address, Subject or Body')
	else:
		msg = EmailMessage()
		msg['To'] = str(to)
		msg['Subject'] = subject
		msg['From'] = from_addr
		msg['Cc'] = cc
		msg['Bcc'] = bcc
		msg['Date'] = time.ctime(time.time())
		msg.set_content(body)

		print('Before:', attachments)
		if attachments != []:
			print('After:', attachments)
			for value in attachments:
				with open(value, 'rb') as f:
					file_name = f.name.split('/')[-1]
					file_data = f.read()

				print('Filename', file_name)
				msg.add_attachment(file_data, maintype = 'application', subtype = 'octet-stream', filename = file_name)
		try:
			smtp.send_message(msg)
			print('Main sent')
			if Gmail_Login == False:
				imap.append('Sent', '\\Seen', imaplib.Time2Internaldate(time.time()), msg.as_string().encode('utf8'))
			messagebox.askokcancel('Success!!', 'Mail sent successfully!!')
			to_var.delete(0, END)
			cc_var.delete(0, END)
			bcc_var.delete(0, END)
			subject_var.delete(0, END)
			body_var.delete(1.0, END)
			attach.config(text = '')
			attachments = None
		except:
			messagebox.askokcancel('Error!!', 'Cannot send mail, check Inputs!!')


attachment_list = []
attachment_label = None
mail_list = []
mb = ""
dir_name = ""

# Creating Put Content in ListBox function


def put_content(l):
	listbox.delete(0, END)
	for index, value in enumerate(l[::-1]):
		str_value = 'Date: {}, Subject: {}, To: {}, From: {}'.format(value['Date'], value['Subject'], value['To'], value['From'])
		listbox.insert(index, str_value)
	no_of_mails.config(text = 'Total no. of mails: {}'.format(len(l)))

# Creating Print Content in ScrolledText function


def print_content(event, l):
	scrl.delete(1.0, END)
	print(listbox.curselection()[0])
	scrl.insert(1.0, l[::-1][listbox.curselection()[0]]['Body'])
	if l[::-1][listbox.curselection()[0]]['Attachment']:
		attach_label.config(text = 'Attachment(s) available!!')

	else:
		attach_label.config(text = '')

# Creating Get Body function


def get_body(msg):
	if msg.is_multipart():
		return get_body(msg.get_payload(0))
	else:
		return msg.get_payload(None, True)

# Creating Get Attachment(s) function


def get_attachments(msg):
	for part in msg.walk():
		if part.get_content_maintype() == 'multipart':
			continue

		if part.get('Content-Disposition') is None:
			continue

		filename = part.get_filename()

		if bool(filename):
			global dir_name
			dir_name = filedialog.askdirectory(initialdir = 'D:')
			filepath = os.path.join(dir_name, filename)
			if os.path.isfile(filepath) == False:
				with open(filepath, 'wb') as f:
					f.write(part.get_payload(decode = True))
			else:
				a = 1
				while True:
					temp = filename.split('.')
					alt_filename = temp[0] + ' ({})'.format(a) + '.' + temp[1]
					alt_filepath = os.path.join(dir_name, alt_filename)
					if os.path.isfile(alt_filepath) == False:
						with open(alt_filepath, 'wb') as f:
							f.write(part.get_payload(decode = True))
						break
					a += 1

# Creating Reverse List Mail order function


def download_attachment():
	_, search = imap.search(None, 'ALL')
	_, data = imap.fetch(search[0].split()[::-1][listbox.curselection()[0]], '(RFC822)')
	msg = email.message_from_bytes(data[0][1])
	if check_attachment(msg):
		get_attachments(msg)
		messagebox.askokcancel('Success!!', 'Attachment(s) downloaded!! in ' + dir_name)

	else:
		messagebox.askokcancel('Missing!!', "Attachment for this mail doesn't exist!!")

# Creating Check Attachment(s) function


def check_attachment(msg):
	for part in msg.walk():
		if part.get_content_maintype() == 'multipart':
			continue

		if part.get('Content-Disposition') is None:
			continue

		return bool(part.get_filename())

# Creating Delete Mail function


def delete_mail():
	global mail_list
	_, search = imap.search(None, 'ALL')
	####imap.store(search[0].split()[::-1][listbox.curselection()[0]], '+FLAGS', '"[Gmail]/Trash"')####
	if Gmail_Login == True:
		imap.store(search[0].split()[::-1][listbox.curselection()[0]], '+X-GM-LABELS', '\\Trash')
	else:
		_, s_d = imap.search(None, 'ALL')
		try:
			if imap.state == 'SELECTED':
				_, da = imap.fetch(s_d[0].split()[::-1][listbox.curselection()[0]], '(UID)')
			else:
				messagebox.askokcancel('Wait!!', 'No mail has been selected!!')
		except:
			messagebox.askokcancel('Wait!!', 'No mail has been selected!!')
		if mb != 'Trash':
			res = imap.uid('COPY', da[0].decode().split()[-1].split(')')[0], 'Trash')
			if res[0] == 'OK':
				_, da = imap.uid('STORE', da[0].decode().split()[-1].split(')')[0], '+FLAGS', '\\Deleted')
				imap.expunge()
		elif mb == 'Trash':
			_, da = imap.uid('STORE', da[0].decode().split()[-1].split(')')[0], '+FLAGS', '\\Deleted')
			imap.expunge()
	print(mail_list.pop(len(mail_list) - 1 - listbox.curselection()[0]))
	listbox.delete(listbox.curselection()[0])
	scrl.delete(1.0, END)
	no_of_mails.config(text = 'Total no. of mails: {}'.format(len(mail_list)))

# Creating Receive Mail function


def recv_mail(mailbox = 'INBOX'):
	if mailbox == 'Sent' and Gmail_Login == True:
		mailbox = '"[Gmail]/Sent Mail"'
	elif mailbox == 'Trash' and Gmail_Login == True:
		mailbox = '"[Gmail]/Bin"'
	mail_list.clear()
	scrl.delete(1.0, END)
	listbox.delete(0, END)
	if attach_label != None:
		attach_label.config(text = "")
	imap.select(mailbox)
	global mb
	mb = mailbox
	_, search_data = imap.search(None, 'ALL')

	for num in search_data[0].split():
		email_data = {}
		_, data = imap.fetch(num, '(RFC822)')
		email_msg = email.message_from_bytes(data[0][1])
		for header in ['Subject', 'To', 'From', 'Date']:
			email_data[header] = email_msg[header]
		email_data['Body'] = email.message_from_bytes(get_body(email_msg))
		email_data['Attachment'] = check_attachment(email_msg)

		mail_list.append(email_data)

	put_content(mail_list)

# Creating Add Attachment function


def put_attachment():
	global attachment_list
	temp = filedialog.askopenfilenames(initialdir = 'D:\\', title = 'Select Attachment(s)', filetypes = (('All files', '*.*'), ('JPEG files', '*.jpeg')))
	for k in temp:
		attachment_list.append(k)
	print('Inside function', attachment_list)
	attachment_str = []
	for value in attachment_list:
		attachment_str.append(value.split('/')[-1])
	attachment_label.config(text = ', '.join(attachment_str))

# Creating Compose Mail function


def compose_mail():
	# Creating Compose Mail Window
	com = Toplevel()
	com.title('Compose Mail')
	com.iconbitmap('./email.ico')

	# Creating Entry Variables
	to_entry_variable = StringVar()
	cc_entry_variable = StringVar()
	bcc_entry_variable = StringVar()
	subject_entry_variable = StringVar()
	scrolled_text_variable = StringVar()

	print('Outside function before calling', attachment_list)

	# Creating To Label
	to_label = Label(com, text = 'To', font = ('', 10))
	to_label.grid(row = 0, column = 0, padx = 10, pady = 10)

	# Creating CC Label
	cc_label = Label(com, text = 'CC', font = ('', 10))
	cc_label.grid(row = 1, column = 0, padx = 10, pady = 10)

	# Creating BCC Label
	bcc_label = Label(com, text = 'BCC', font = ('', 10))
	bcc_label.grid(row = 2, column = 0, padx = 10, pady = 10)

	# Creating Subject Label
	subject_label = Label(com, text = 'Subject', font = ('', 10))
	subject_label.grid(row = 3, column = 0, padx = 10, pady = 10)

	# Creating Body Label
	body_label = Label(com, text = 'Body', font = ('', 10))
	body_label.grid(row = 4, column = 0, padx = 10, pady = 10, sticky = N)

	# Creating To Entry
	to_entry = Entry(com, font = ('', 10), width = 60)
	to_entry.grid(row = 0, column = 1, padx = 10, pady = 10, sticky = W)

	# Creating CC Entry
	cc_entry = Entry(com, font = ('', 10), width = 60)
	cc_entry.grid(row = 1, column = 1, padx = 10, pady = 10, sticky = W)

	# Creating BCC Entry
	bcc_entry = Entry(com, font = ('', 10), width = 60)
	bcc_entry.grid(row = 2, column = 1, padx = 10, pady = 10, sticky = W)

	# Creating Subject Entry
	subject_entry = Entry(com, font = ('', 10), width = 60)
	subject_entry.grid(row = 3, column = 1, padx = 10, pady = 10, sticky = W)

	# Creating Body ScrolledText
	scrolled_text = ScrolledText(com, wrap = WORD, font = ('Verdana', 12))
	scrolled_text.grid(row = 4, column = 1, padx = 10, pady = 10)

	# Creating Attach button
	attachment_btn = Button(com, text = 'Attach', fg = 'white', bg = 'blue', font = ('', 10), command = put_attachment)
	attachment_btn.grid(row = 5, column = 1, sticky = E, padx = 10, pady = 10)

	# Creating Send Button
	send_btn = Button(com, text = 'Send', bg = 'blue', fg = 'white', font = ('', 12), command = lambda: send_mail(root_email_entry.get(), to_entry, cc_entry, bcc_entry, subject_entry, scrolled_text, attachment_label, attachment_list))
	send_btn.grid(row = 5, column = 0, padx = 10, pady = 10)

	# Creating Attachment Label
	global attachment_label
	attachment_label = Label(com, font = ('', 10))
	attachment_label.grid(row = 5, column = 1, sticky = W, padx = 10, pady = 10)

# Creating Root Window Destroy function


def on_closing():
	if messagebox.askyesno('Quit', 'Do you want to quit?'):
		root.destroy()

# Creating Login Window Destroy function


def on_closing_top(x):
	if messagebox.askyesno('Quit', 'Do you want to quit?'):
		x.destroy()
		root.deiconify()


# Creating Login Window (Main)
root = Tk()
root.title('SMTP & IMAP')
root.iconbitmap('./email.ico')
root.config(bg = 'cyan')

# Screen Details
app_width = 550
app_height = 400

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

x = (screen_width // 2) - (app_width // 2)
y = (screen_height // 2) - (app_height // 2)
root.geometry(f'{app_width}x{app_height}+{x}+{y}')

# Creating Login Frame
root_frame = Frame(root, bg = 'cyan')
root_frame.place(in_ = root, anchor = 'c', relx = .5, rely = .5)

# Creating Login page variables
root_email_variable = StringVar()
root_password_variable = StringVar()

# Creating Email Label
root_email_label = Label(root_frame, text = 'Email', font = ('', 16), bg = 'cyan')
root_email_label.grid(row = 0, column = 0, padx = 5, pady = 5)

# Creating Email Entry
root_email_entry = Entry(root_frame, font = ('Verdana', 16), textvariable = root_email_variable, width = 30)
root_email_entry.focus_set()
root_email_entry.grid(row = 0, column = 1, padx = 5, pady = 5)

# Creating Password Label
root_password_label = Label(root_frame, text = 'Password', font = ('', 16), bg = 'cyan')
root_password_label.grid(row = 1, column = 0, padx = 5, pady = 5)

# Creating Password Entry
root_password_entry = Entry(root_frame, font = ('Verdana', 16), show = '*', textvariable = root_password_variable, width = 30)
root_password_entry.grid(row = 1, column = 1, padx = 5, pady = 5)

# Creating Login Button
root_login_button = Button(root_frame, text = 'Login', font = ('', 16), command = lambda: toggle_and_login_check(root_email_variable, root_password_variable))
root_login_button.grid(columnspan = 2, padx = 5, pady = 5)

# Creating Info Label
root_info_label = Label(root_frame, text = '', fg = 'red', font = ('', 16), bg = 'cyan')
root_info_label.grid(columnspan = 2, padx = 5, pady = 5)

# Creating Mail sending Window
top = Toplevel()
top.withdraw()
top.iconbitmap('./email.ico')

# Creating Mail Window frames
left_frame = Frame(top, borderwidth = 5, highlightthickness = 3)
left_frame.pack(side = LEFT, expand = False, fill = Y)

right_frame = Frame(top)
right_frame.pack(side = LEFT, expand = True, fill = 'both')

mails_frame = Frame(right_frame)
mails_frame.pack(side = TOP, expand = True, fill = 'both')

view_frame = Frame(right_frame)
view_frame.pack(side = TOP, expand = True, fill = 'both')

# Creating Compose Button
compose_btn = Button(left_frame, text = 'Compose', font = ('', 16), command = compose_mail)
compose_btn.pack(side = TOP, padx = 30, pady = 20, fill = 'both')

# Creating Inbox Button
inbox_btn = Button(left_frame, text = 'Inbox', font = ('', 16), command = lambda: recv_mail('INBOX'))
inbox_btn.pack(side = TOP, padx = 30, pady = 20, fill = 'both')

# Creating Sent Mail Button
sent_mail_button = Button(left_frame, text = 'Sent', font = ('', 16), command = lambda: recv_mail('Sent'))
sent_mail_button.pack(side = TOP, padx = 30, pady = 20, fill = 'both')

# Creating Trash Button
trash_button = Button(left_frame, text = 'Trash', font = ('', 16), command = lambda: recv_mail('Trash'))
trash_button.pack(side = TOP, padx = 30, pady = 20, fill = 'both')

# Creating Delete Mail Button
delete_button = Button(left_frame, text = 'Delete', font = ('', 16), command = delete_mail)
delete_button.pack(side = TOP, padx = 30, pady = 20, fill = 'both')

# Creating Number of mails Label
no_of_mails = Label(left_frame, font = ('', 16))
no_of_mails.pack(side = TOP, padx = 30, pady = 20, fill = 'both')

# Creating Attachment Label
attach_label = Label(left_frame, font = ('', 16))
attach_label.pack(side = TOP, padx = 30, pady = 20, fill = 'both')

# Creating Logout Button
logout_btn = Button(left_frame, text = 'Logout', font = ('', 16), command = toggle)
logout_btn.pack(side = BOTTOM, padx = 30, pady = 10, fill = 'both')

# Creating Download Attachment Button
download_attach = Button(left_frame, text = 'Download', font = ('', 16), command = download_attachment)
download_attach.pack(side = BOTTOM, padx = 30, pady = 20, fill = 'both')

# Creating ListBox for Mails view (Scrollable)
scrl_x = Scrollbar(mails_frame, orient = 'horizontal')
scrl_x.pack(side = BOTTOM, fill = 'both', padx = 5)

listbox = Listbox(mails_frame, font = ('Verdana', 12))
listbox.pack(side = LEFT, padx = 5, pady = 5, fill = 'both', expand = True)
listbox.bind('<Double-1>', lambda event: print_content(event, mail_list))

scrl_y = Scrollbar(mails_frame)
scrl_y.pack(side = RIGHT, fill = 'both', pady = 5)

listbox.config(xscrollcommand = scrl_x.set, yscrollcommand = scrl_y.set)
scrl_x.config(command = listbox.xview)
scrl_y.config(command = listbox.yview)

# Creating Viewing area for Mails
scrl = ScrolledText(view_frame, width = 40, height = 10, font = ('Verdana', 14), wrap = WORD)
scrl.pack(side = LEFT, padx = 5, pady = 5, fill = 'both', expand = True)

# Closing Windows protocols
root.protocol("WM_DELETE_WINDOW", on_closing)
top.protocol("WM_DELETE_WINDOW", lambda: on_closing_top(top))
root.mainloop()
