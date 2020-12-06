import os
import sys
import smtplib
import random
import time
import json
import easyimap
import imaplib
import email
import email.message
import ctypes
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders
from string import digits, ascii_uppercase as letters

global send_from
global send_to
global server
global port
global username
global password
global use_tls
global enable_confirm
global host
global recievers_username
global recievers_password
global mailbox
global popup_report
global known_uid

global config_path
config_path = 'C:\\Windows\\System32' #This is the path to your config file

global config_name
config_name = 'ASMail_config' #This is the name of your config file

is_compiled = False
if is_compiled == True:
    config_path = os.getcwd()
    config_name = sys.argv[0].split('.')[0] + "_config"

def config(mode = 0):
    global send_from
    global send_to
    global server
    global port
    global username
    global password
    global use_tls
    global enable_confirm
    global host
    global recievers_username
    global recievers_password
    global mailbox
    global popup_report
    global known_uid

    os.chdir(config_path)

    try:
        if mode == 1:
            sys.exit()

        file = open(config_name, 'r')
        data = json.load(file)
        file.close()

        known_hosts = data[1]
        data = data[0]

        send_from = data[0]
        send_to = data[1]
        server = data[2]
        port = data[3]
        username = data[4]
        password = data[5]
        use_tls = data[6]
        enable_confirm = data[7]
        host = data[8]
        recievers_username = data[9]
        recievers_password = data[10]
        mailbox = data[11]
        popup_report = data[12]

    except:
        if mode == 0:
            print('No configuration file has been found!')
            print('Enter data to create it')
        if mode == 1:
            print('Configuration setup, enter new data')

        send_from = input('Enter the senders address: ')
        send_to = input('Enter the recievers address: ')
        server = input('Enter the servers SMTP address (senders server): ')
        port = input('Enter a port number (empty to use 587): ')
        username = input('Enter username (empty to use senders address): ')
        password = input('Enter password of the user above: ')
        use_tls = input('Use TLS: 1 to enable, 0 to disable (empty to set 1): ')
        enable_confirm = input('Confirm recieving: 1 to enable, 0 to disable (empty to set 1): ')
        host = input('Address of recievers IMAP server (not required if confirmation is disabled): ')
        recievers_username = input('Recievers username (not required if confirmation is disabled) (empty to set as recievers address): ')
        recievers_password = input('Recievers password(not required if confirmation is disabled): ')
        mailbox = input('Recievers mailbox(not required if confirmation is disabled) (empty to set INBOX): ')
        popup_report = input('Enable/disable pop ups: 1 to enable pop up reports, 0 to disable: ')

        if port == '':
            port = '587'

        if username == '':
            username = send_from

        if use_tls == '':
            use_tls = '1'

        if enable_confirm == '':
            enable_confirm = '1'

        if recievers_username == '':
            recievers_username = send_to

        if mailbox == '':
            mailbox = 'INBOX'

        def port_check():
            global port
            try:
                port = int(port)
                return port
            except:
                return None

        def tls_check():
            global use_tls
            try:
                use_tls = bool(int(use_tls))
                return use_tls
            except:
                return None

        def enable_confirm_check():
            global enable_confirm
            try:
                enable_confirm = bool(int(enable_confirm))
                return enable_confirm
            except:
                return None

        def popup_report_check():
            global popup_report
            try:
                popup_report = bool(int(popup_report))
                return popup_report
            except:
                return None

        while port_check() == None:
            port = input('Error, port number is not an integer, reenter port number: ')

        while tls_check() == None:
            use_tls = input('Error, tls should be 1 or 0, reenter tls value: ')

        while enable_confirm_check() == None:
            enable_confirm = input('Error, confirm recieving should be 1 or 0, reenter tls value: ')

        while popup_report_check() == None:
            popup_report = input('Error, popup report should be 1 or 0, reenter tls value: ')

        config_list = [[send_from, [send_to], server, port, username, password, use_tls, enable_confirm, host, recievers_username, recievers_password, mailbox, popup_report], []]

        file = open(config_name, 'w')
        file.write(json.dumps(config_list))
        file.close()

        if mode == 0:
            config()

def send_mail(send_from, send_to, subject, message, files, server, port, username, password, use_tls):
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(message))

    for path in files:
        part = MIMEBase('application', "octet-stream")
        with open(path, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        'attachment; filename="{}"'.format(Path(path).name))
        msg.attach(part)

    smtp = smtplib.SMTP(server, port)

    if use_tls:
        smtp.starttls()

    smtp.login(username, password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.quit()

def uid_generate(mode = 'create', uid = ''):
    file = open(config_name, 'r')
    data = json.load(file)
    file.close()

    known_hosts = data[1]

    if mode == 'create':
        def create():
            string = 'uid_'
            ltr = [random.choice(letters) for n in range(16)]
            nmb = [random.choice(digits) for n in range(16)]
            for item in ltr[0:8]:
                string += item
            string += '-'
            for item in nmb[0:8]:
                string += item
            string += '-'
            for item in ltr[8:16]:
                string += item
            string += '-'
            for item in nmb[8:16]:
                string += item
            return string

        string = create()

        while string in known_hosts:
            string = create()

        return string

    if mode == 'dump':
        file = open(config_name, 'r')
        data = json.load(file)
        file.close()

        data[1].append(uid)

        file = open(config_name, 'w')
        file.write(json.dumps(data))
        file.close()

def end(mode = 'standart'):
    if mode == 'standart':
        input('Press enter to exit: ')
        sys.exit()

    if mode == 'silent':
        sys.exit()

    if mode == 'params_success':
        print('Parameters has been changed successfully!', end = '')
        sys.exit()

    if mode == 'flag_error':
        print('Flag error occured!')
        sys.exit()

def check():
    global send_from
    global send_to
    global server
    global port
    global username
    global password
    global use_tls
    global enable_confirm
    global host
    global recievers_username
    global recievers_password
    global mailbox
    global popup_report
    global known_uid

    if len(sys.argv) == 1:
        config(1)
        end()

    else:
        config()

    if len(sys.argv) >= 2:

        if sys.argv[1] in ['-f', '-from', '-t', '-to', '-s', '-server', '-p', '-port', '-u', '-username', '-w', '-password', '-l', '-use_tls', '-c', '-print_config', '-r', '-clear_uids', '-h', '--help', '-i', '-print_uids', '-e', '-confirm', '-a', '-host', '-q', '-recievers_username', '-d', '-recievers_password', '-m', '-mailbox', '-o', '-popup_report']:
            
            file = open(config_name, 'r')
            data = json.load(file)
            file.close()

            known_hosts = data[1]
            data = data[0]

            send_from = data[0]
            send_to = data[1]
            server = data[2]
            port = data[3]
            username = data[4]
            password = data[5]
            use_tls = data[6]
            enable_confirm = data[7]
            host = data[8]
            recievers_username = data[9]
            recievers_password = data[10]
            mailbox = data[11]
            popup_report = data[12]

            if sys.argv[1] == '-c' or sys.argv[1] == '-print_config':
                print('\n---------------------------------------------\n---------------------------------------------\n           Configuration file\n---------------------------------------------\nSenders address: ' + send_from + '\nRecievers address: ' + send_to[0] + '\nSMTP senders server: ' + server + '\nPort number: ' + str(port) + '\nUsername: ' + username + '\nPassword: ' + password + '\nTLS: ' + str(use_tls) + '\nDelievery confirm: ' + str(enable_confirm) + '\nIMAP recievers server: ' + host + '\nRecievers username: ' + recievers_username + '\nRecievers password: ' + recievers_password + '\nDestination mailbox: ' + mailbox + '\nPop up report: ' + str(popup_report) + '\nConfig path: ' + config_path + '\nConfig name: ' + config_name + '\nis_compiled: ' + str(is_compiled) + '\n---------------------------------------------\n               End of file\n---------------------------------------------\n---------------------------------------------')
                end('silent')

            elif sys.argv[1] == '-h' or sys.argv[1] == '--help':
                print('List of all flags:\n1. -h or --help to view this menu\n2. -f or -from to change the senders address\n3. -t or -to to change the recievers address\n4. -p or -port to change the port number\n5. -u or -username to change the username\n6. -w or -password to change the password\n7. -l or -use_tls to enable/disable tls\n8. -c or -print_config to print configuration file\n9. -r or -clear_uids to delete records of used uids\n10. -e or -confirm to enable/disable delievery confirm\n11. -a or -host to change imap recievers host\n12. -q or -recievers_username to change recievers login\n13. -d or -recievers_password to change recievers password\n14. -m or -mailbox to change the destination mailbox\n15. -o or -popup_report to enable/diable pop up report', end = '')
                end('silent')

            elif sys.argv[1] == '-i' or sys.argv[1] == '-print_uids':
                print('\n---------------------------------------------\n---------------------------------------------\n                  Used UIDs\n---------------------------------------------')
                [print(item) for item in known_hosts]
                print('---------------------------------------------\n               End of the file\n---------------------------------------------\n---------------------------------------------')
                end('silent')
            else:

                try:
                    params = sys.argv[2]

                except:
                    if sys.argv[1] != '-r':
                        if sys.argv[1] != '-clear_uids':
                            params = input('Enter new parameter: ')

                if sys.argv[1] == '-f' or sys.argv[1] == '-from':
                    send_from = params
                elif sys.argv[1] == '-t' or sys.argv[1] == '-to':
                    send_to = [params]
                elif sys.argv[1] == '-s' or sys.argv[1] == '-server':
                    server = params
                elif sys.argv[1] == '-p' or sys.argv[1] == '-port':
                    port = params
                elif sys.argv[1] == '-u' or sys.argv[1] == '-username':
                    username = params
                elif sys.argv[1] == '-w' or sys.argv[1] == '-password':
                    password = params
                elif sys.argv[1] == '-l' or sys.argv[1] == '-use_tls':
                    use_tls = params
                elif sys.argv[1] == '-e' or sys.argv[1] == '-confirm':
                    enable_confirm = params
                elif sys.argv[1] == '-a' or sys.argv[1] == '-host':
                    host = params
                elif sys.argv[1] == '-q' or sys.argv[1] == '-recievers_username':
                    recievers_username = params
                elif sys.argv[1] == '-d' or sys.argv[1] == '-recievers_password':
                    recievers_password = params
                elif sys.argv[1] == '-m' or sys.argv[1] == '-mailbox':
                    mailbox = params
                elif sys.argv[1] == '-o' or sys.argv[1] == '-popup_report':
                    popup_report = params
                elif sys.argv[1] == '-r' or sys.argv[1] == '-clear_uids':
                    known_hosts = []

                else:
                    end('flag_error')

                if port == '':
                    port = '587'

                if username == '':
                    username = send_from

                if use_tls == '':
                    use_tls = '1'

                def port_check():
                    global port
                    try:
                        port = int(port)
                        return port
                    except:
                        return None

                def tls_check():
                    global use_tls
                    try:
                        use_tls = bool(int(use_tls))
                        return use_tls
                    except:
                        return None

                def enable_confirm_check():
                    global enable_confirm
                    try:
                        enable_confirm = bool(int(enable_confirm))
                        return enable_confirm
                    except:
                        return None

                def popup_report_check():
                    global popup_report
                    try:
                        popup_report = bool(int(popup_report))
                        return popup_report
                    except:
                        return None

                while port_check() == None:
                    port = input('Error, port number is not an integer, reenter port number: ')

                while tls_check() == None:
                    use_tls = input('Error, tls should be 1 or 0, reenter tls value: ')

                while enable_confirm_check() == None:
                    enable_confirm = input('Error, confirm recieving should be 1 or 0, reenter tls value: ')

                while popup_report_check() == None:
                    popup_report = input('Error, popup report should be 1 or 0, reenter tls value: ')

                config_list = [[send_from, send_to, server, port, username, password, use_tls, enable_confirm, host, recievers_username, recievers_password, mailbox, popup_report], []]

                file = open(config_name, 'w')
                file.write(json.dumps(config_list))
                file.close()

                end('params_success')

        else:
            None

def delivery_confirm(sent_uid, host, username, password, mailbox, report):
    imapper = easyimap.connect(host, username, password, mailbox)

    email_quantity = 10
    emails_from_your_mailbox = imapper.listids(limit=email_quantity)
    imapper.quit()

    imap = imaplib.IMAP4_SSL(host)
    imap.login(username, password)
    imap.select(mailbox)

    uid = 'NO'
    file = 'NO'

    for item in emails_from_your_mailbox:
        status, data = imap.fetch(item, '(RFC822)')
        msg = email.message_from_bytes(data[0][1],_class = email.message.EmailMessage)
        if msg['Subject'] == '***Automatically generated by ASMail***':
            payload = msg.get_payload()[ 0 ]
            recieved_uid = payload.get_payload()
            for part in msg.walk(): 
                    if part.get_content_maintype() == 'multipart':
                        continue
                    if part.get('Content-Disposition') is None:
                        continue
                    fileName = part.get_filename()
            sent_filename = Path(sys.argv[1]).name

            if sent_uid == recieved_uid:
                uid = 'YES'
                if sent_filename == fileName:
                    file = 'YES'

    if report == True:
        ctypes.windll.user32.MessageBoxW(0, 'UID status: ' + uid + '\nFilename status: ' + file, 'Delivery report', 0)

    print('UID status: ' + uid + '\nFilename status: ' + file)

    imap.logout()

def operator():
    if enable_confirm == True:
        uid = uid_generate()
    else:
        uid = ''

    send_mail(send_from, send_to, '***Automatically generated by ASMail***', uid, sys.argv[1 : len(sys.argv)], server, port, username, password, use_tls)
    
    if enable_confirm == True:
        uid_generate('dump', uid)
    else:
        None

    if enable_confirm == True:
        try:
            delivery_confirm(uid, host, recievers_username, recievers_password, mailbox, popup_report)
        except:
            try:
                delivery_confirm(uid, host, recievers_username, recievers_password, 'Spam', popup_report)
                if report == True:
                    ctypes.windll.user32.MessageBoxW(0, 'UID status: ' + uid + '\nFilename status: ' + file + '\nWarning! This letter has been found in Spam folder!', 'Delivery report', 0)

                print('UID status: ' + uid + '\nFilename status: ' + file + '\nWarning! This letter has been found in Spam folder!')

            except:
                if popup_report == True:
                    ctypes.windll.user32.MessageBoxW(0, 'Unable to confirm delivery' 'Delivery report', 0)
                print('Unable to confirm delivery')

def start():
    check()
    operator()
    end()

start()
