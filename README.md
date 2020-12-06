# ASMail
Script for sending an email with attachment.<br />
# Attachments
Files' full addresses has to be given as a parameters to an email<br />
A script.reg file will edit the windows registry value to enable drag & drop over python scripts<br />
It can also be launched with parameters from console e.g. python ASMail.py C://Path to file, C://Path to file ... etc<br />
# Filenames & paths to those files
All filenames and paths should be in english letters only, otherwise this will lead to program crash.
# Setup
The path to config file can be changed in line 36<br />
The name of the config file can be changed in line 39.
# Delivery confirm
If this option is enabled, script will try to connect to recievers mailbox and look for last 10 messages with generated subject. Each letter sended will contain a unique ID. If an email with same ID and same filename would be found, delivery report will comfirm delivery.
# Flags
Flags is used to print the config file data or to change its parameters separately<br />
Usage: ASMail.py -F new_value for:<br />
-h or --help to view help menu<br />
-f or -from to change the senders address<br />
-t or -to to change the recievers address<br />
-p or -port to change the port number of sender's SMTP server<br />
-u or -username to change the username (sender's address or login)<br />
-w or -password to change the password (sender's password)<br />
-l or -use_tls to enable/disable tls<br />
-c or -print_config to print configuration file<br />
-r or -clear_uids to delete records of used uids<br />
-e or -confirm to enable/disable delievery confirm<br />
-a or -host to change imap recievers host<br />
-q or -recievers_username to change recievers login<br />
-d or -recievers_password to change recievers password<br />
-m or -mailbox to change the destination mailbox<br />
-o or -popup_report to enable/diable pop up report<br />
(Second parameter after flag is not required, it can be entered in input field afterwards)<br />
# Configuration file
By default, the configuration file is created at C://Windows//System32, called ASMail_config<br />Program enters configuration mode if it is unable to find
ASMail_config file in the indicated directory.<br />The format of ASMail_config is a python list, where<br /><br />[[sender's address (string value), [receiver's address (string value)], sender's SMTP server address (string value), SMTP port (int value), sender's username (string value), sender's password  (string value), enable tls (bool value), enable delivary report (bool value), receiver's IMAP server address (string value), receiver's address (string value), receiver's password (string value), mailbox (string_value), enable pop-ups (bool value)], [list of all used uids (string values)]]<br /><br />This file will be created at the first launch, after that it can be edited manually or by the program.<br />To enter the configuration mode with ASMail_config existing, run ASMail without any parameters (python ASMail.py).
# .exe
There is a compiled version of the program, however it is unable to change the config file directory<br />
The path used is the current path where the program is stored<br />
It cannot be modified anyhow except recompiling the program.
