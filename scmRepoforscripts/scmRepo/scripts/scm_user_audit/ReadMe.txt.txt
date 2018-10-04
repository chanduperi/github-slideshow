scm_user_audit.py
=================
Python Script to query users list who are disabled in LDAP and send the list as an e-mail to "Engineering SCM" group. This will allow Application adminsitrators to quickly disable accounts

Queries data from:
Host: scmweb.extremenetworks.com
database.tables: toolUsers.toolUsers and toolUsers.displayNames
username: toolUsers
password: toolUsers

For mail uses: smtp.extremenetworks.com
For Teams uses Incoming Webhook of "Tool Users" channel: https://teams.microsoft.com/l/channel/19%3ae4eab443e99c4c57b0cfcec1afb10f29%40thread.skype/Tool%2520Users?groupId=8e646d4b-1fbe-461b-8ff4-cdeb94db188e&tenantId=fc8c2bf6-914d-4c1f-b352-46a9adb87030

Usage
================
python scm_user_audit.py

It asks for the application name for which report need to be generated. If need to be generated for everything, use "ALL"

Output will be Excel Sheet with details, currently hardcoded to mail to cperi@extremenetworks.com and please change it accordingly in future. 

To Do
================
1) Add code to automatically post to Microsoft Teams Channel by calling it's Incoming Web-Hook
2) There are concerns that, some Active users are listed in generated file (ClearCase and JIRA), need to look

Be Watch On
===============
If Microsoft Office upgrades to new version (say in 2022), will xlst xlsd Python libraries work? Do we need to pip install again?. Till then. Enjoyyy

Python Libraries used
=====================
requests
json
datetime
sys
time
urllib3 -- Not Used Now
mysql.connector
smtplib
xlwt
csv -- Not Used Now
xlrd import open_workbook
xlutils.copy import copy
email.mime.multipart import MIMEMultipart
email.mime.base import MIMEBase
email.mime.text import MIMEText
email import encoders
pprint import pprint
os.path

