import requests
import json
import datetime
import sys
import time
import urllib3
import mysql.connector
import smtplib
import xlwt
import csv
from xlrd import open_workbook
from xlutils.copy import copy
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from pprint import pprint
import os.path

requests.packages.urllib3.disable_warnings()

timestr = time.strftime("%Y%m%d")
xlsxFileName = "SCM_Tools_Disable_User_Audit_"+timestr+".xls"

##################################################################
### Connecting scmweb Database using mysql.connector #############
##################################################################
cnx = mysql.connector.connect(host='scmweb.extremenetworks.com', database='toolUsers', user='toolUsers', password='toolUsers', charset="utf8", use_unicode = True)
cursor = cnx.cursor()

##################################################################
###### Get all column names ######################################
##################################################################
query = ("describe toolUsers.toolUsers;")
cursor.execute(query)
result = cursor.fetchall()
# print (result)

#################################################################
############# Getting Column Names ##############################
#################################################################
tool_names = ("SHOW columns FROM toolUsers;")
cursor.execute(tool_names)
results_tools = [column[0] for column in cursor.fetchall()]
# print(results_tools)
user_input_of_tools = results_tools

##################################################################
######### Ask User Input whether to generate report for ALL tools#
######### OR for specicif Tool####################################
##################################################################
del user_input_of_tools[:3]
user_input_of_tools[0] = 'ALL'
print (" Asking User Inputs ")
print (user_input_of_tools)
which_tool_audit = input("Please enter for which application audit is required from the list: ")
# print (user_input_of_tools)
print ("Application Entered: " + which_tool_audit)

##############################################################################
########## Reading Known Service Accounts from File ##########################
########## Service Accounts are in known_service_accounts.txt in same folder #
##############################################################################
known_service_accounts = ''
with open('known_service_accounts.txt', 'r') as myfile:
    known_service_accounts=myfile.readlines()
# print ("Known Service Accounts")
known_service_accounts = [x.strip() for x in known_service_accounts]
# print (known_service_accounts)

# known_service_accounts = 'userName NOT IN (' + known_service_accounts + ')'

##################################################################
########### Generate query and fetch details #####################
##################################################################

##### Getting Display Names and Need to Map them
application_display_names_query = "select columnName, displayName from toolUsers.displayNames;"
application_display_query = (application_display_names_query)
cursor.execute(application_display_query)
header_names = [i[0] for i in cursor.description]
application_names_audit_result = [list(i) for i in cursor.fetchall()]
# application_names_audit_result.insert(0,header_names)
# print("===== Display Names ===========================")
# print(application_names_audit_result)
# print("====== Above are Display Names Mapping ========")
# display_names_dict = {x[1]:x[2] for x in application_names_audit_result}
display_names_dict = {}
for k in application_names_audit_result:
     # print(k)
     display_names_dict.update({k[0]:k[1]})

all_tools_query = ""
revuewboard_query = ""
jenkins_query = ""
git_query = ""
svn_query = ""
clearcase_query = ""
user_audit_result =[]
counting_tools = 0

######################################################################
##### Function to write to excel in worksheets #######################
######################################################################
def query_and_write_to_exclel(string_for_user_audit_query_f, which_tool_audit):
    print (string_for_user_audit_query_f)
    print (which_tool_audit)
    user_audit_query = (string_for_user_audit_query_f)
    which_tool_audit = which_tool_audit
    cursor.execute(user_audit_query)
    header = [i[0] for i in cursor.description]
    # print ("printing Header")
    # print (header)
    incrementer = 0
    for eachvalue in header:
        header[incrementer] = display_names_dict[eachvalue]
        incrementer = incrementer  + 1
    # print (header)
    # print ("printed Header")
    # user_audit_result = cursor.fetchall()
    user_audit_result = [list(i) for i in cursor.fetchall()]
    user_audit_result.insert(0,header)
    # print(user_audit_result)

    ############################################################
    ####### Remove Known Service Accounts from List ############
    ############################################################
    for known_user in known_service_accounts:
         row_number_users = 0
         for each_row in user_audit_result:
              if(each_row[1] == known_user):
                   user_audit_result.pop(row_number_users)
                   break
              row_number_users = row_number_users + 1

    #################################################################
    ##### Replace 0 with No and 1 with Yes in the SQL Query Result ##
    #################################################################
    row_number = 0
    for each_row in user_audit_result:
          column_number = 0
          remove_service_accounts = 0
          for x in each_row:
               if (x == 0):
                    user_audit_result[row_number][column_number] = 'No'
               if (x == 1):
                    user_audit_result[row_number][column_number] = 'Yes'
               column_number = column_number + 1
          row_number = row_number + 1


    # Check Dictionary check
    # print(display_names_dict['yn_jira'])


    ###############################################################
    ########### Writing Excel Workbook ############################
    ###############################################################
    # book = xlwt.Workbook()
    # xlsxFileName = "SCM_Tools_Disable_User_Audit_"+timestr+".xls"
    # filecreate = open(xlsxFileName,"w+")
    # filecreate.close()
    if os.path.isfile(xlsxFileName):
         rbook=open_workbook(xlsxFileName)
    else:
         workbook = xlwt.Workbook()
         newsheet = workbook.add_sheet('Users Audit')
         newsheet.write(1,1,'Please browse through Worksheets with respect to related application administration')
         workbook.save(xlsxFileName)
         rbook = open_workbook(xlsxFileName)
	
    wbook = copy(rbook)

    if (which_tool_audit == 'ALL'):
         sheet = wbook.add_sheet('ALL')
         # which_tool_audit = 'yn_jira'
    if (which_tool_audit == 'ReviewBoard'):
         sheet = wbook.add_sheet('ReviewBoard')
    if (which_tool_audit == 'Jenkins'):
         sheet = wbook.add_sheet('Jenkins')
    if (which_tool_audit == 'Git'):
         sheet = wbook.add_sheet('Git')
    if (which_tool_audit == 'Subversion'):
         sheet = wbook.add_sheet('Subversion')
    if (which_tool_audit == 'ClearCase'):
         sheet = wbook.add_sheet('ClearCase')
    if (which_tool_audit != 'ALL' and which_tool_audit != 'ReviewBoard'  and which_tool_audit != 'Jenkins'  and which_tool_audit != 'Git' and which_tool_audit != 'Subversion' and which_tool_audit != 'ClearCase'):
         sheet = wbook.add_sheet(display_names_dict[which_tool_audit])
    for i, l in enumerate(user_audit_result):
         for j, col in enumerate(l):
              sheet.write(i, j, str(col))
    # book.save('SCM_Tools_Disable_User_Audit.xls')
    wbook.save(xlsxFileName)

###########################################################################################
######## Building SQL Query for all tools and call function to write in Excel #############
###########################################################################################
if (which_tool_audit == 'ALL'):
    #### Currently it is querying individual tools and is commenting below to change to collected tools.
    for x in results_tools:
        counting_tools = counting_tools + 1
        if (counting_tools >= 5):
             all_tools_query = all_tools_query + x + " || " 
    all_tools_query = all_tools_query[:-3]   
    string_for_user_audit_query = "select * from toolUsers.toolUsers where not isnull(dateRemoved) and (" + all_tools_query + ");"
    query_and_write_to_exclel(string_for_user_audit_query, which_tool_audit)
	# revuewboard_query = ""
	# jenkins_query = ""
	# git_query = ""
	# svn_query = ""
	# clearcase_query = ""
	##### Query on grouped applications (All reviewboard, all Git, all jenkins)
	#### ReviewBoard
    counting_tools = 0
    reviewBoard_columns = ''
    for x in results_tools:
        counting_tools = counting_tools + 1
        if (counting_tools >= 5):
            if ('review' in x):
                revuewboard_query = revuewboard_query + x + " || "
                reviewBoard_columns = reviewBoard_columns + x + ", "
  
    revuewboard_query = revuewboard_query[:-3]
    reviewBoard_columns = reviewBoard_columns[:-2]
    string_for_user_audit_query = "select recordId, userName, dateRemoved, " + reviewBoard_columns + " from toolUsers.toolUsers where not isnull(dateRemoved) and (" + revuewboard_query + ");"
    print (string_for_user_audit_query)
    query_and_write_to_exclel(string_for_user_audit_query, 'ReviewBoard')
	
    #### Jenkins
    counting_tools = 0
    jenkins_columns = ''
    for x in results_tools:
        counting_tools = counting_tools + 1
        if (counting_tools >= 5):
            if ('jenkins' in x):
                jenkins_query = jenkins_query + x + " || "
                jenkins_columns = jenkins_columns + x + ", "
  
    jenkins_query = jenkins_query[:-3]
    jenkins_columns = jenkins_columns[:-2]
    string_for_user_audit_query = "select recordId, userName, dateRemoved, " + jenkins_columns + " from toolUsers.toolUsers where not isnull(dateRemoved) and (" + jenkins_query + ");"
    print (string_for_user_audit_query)
    query_and_write_to_exclel(string_for_user_audit_query, 'Jenkins')
	
    #### Git
    counting_tools = 0
    git_columns = ''
    for x in results_tools:
        counting_tools = counting_tools + 1
        if (counting_tools >= 5):
            if ('git' in x):
                git_query = git_query + x + " || "
                git_columns = git_columns + x + ", "
  
    git_query = git_query[:-3]
    git_columns = git_columns[:-2]
    string_for_user_audit_query = "select recordId, userName, dateRemoved, " + git_columns + " from toolUsers.toolUsers where not isnull(dateRemoved) and (" + git_query + ");"
    print (string_for_user_audit_query)
    query_and_write_to_exclel(string_for_user_audit_query, 'Git')
	    

    #### Subversion
    counting_tools = 0
    svn_columns = ''
    for x in results_tools:
        counting_tools = counting_tools + 1
        if (counting_tools >= 5):
            if ('svn' in x):
                svn_query = svn_query + x + " || "
                svn_columns = svn_columns + x + ", "
  
    svn_query = svn_query[:-3]
    svn_columns = svn_columns[:-2]
    string_for_user_audit_query = "select recordId, userName, dateRemoved, " + svn_columns + " from toolUsers.toolUsers where not isnull(dateRemoved) and (" + svn_query + ");"
    print (string_for_user_audit_query)
    query_and_write_to_exclel(string_for_user_audit_query, 'Subversion')
	
    #### ClearCase
    counting_tools = 0
    clearcase_columns = ''
    for x in results_tools:
        counting_tools = counting_tools + 1
        if (counting_tools >= 5):
            if ('clearcase' in x):
                clearcase_query = clearcase_query + x + " || "
                clearcase_columns = clearcase_columns + x + ", "
  
    clearcase_query = clearcase_query[:-3]
    clearcase_columns = clearcase_columns[:-2]
    string_for_user_audit_query = "select recordId, userName, dateRemoved, " + clearcase_columns + " from toolUsers.toolUsers where not isnull(dateRemoved) and (" + clearcase_query + ");"
    print (string_for_user_audit_query)
    query_and_write_to_exclel(string_for_user_audit_query, 'ClearCase')
		
    #### For all other individual applications
    counting_tools = 0
    clearcase_columns = ''
    for x in results_tools:
        counting_tools = counting_tools + 1
        if (counting_tools >= 3):
            if ('review' not in x and 'jenkins' not in x and 'git' not in x and 'svn' not in x and 'clearcase' not in x):
                all_tools_query = x + " || "
                all_tools_query = all_tools_query[:-3]
                string_for_user_audit_query = "select recordId, userName, dateRemoved, " + x + " from toolUsers.toolUsers where not isnull(dateRemoved) and (" + all_tools_query + ");"
                print (string_for_user_audit_query)
                query_and_write_to_exclel(string_for_user_audit_query, x)
		
		
		
########################################################################
####### Building SQL Query for the tool user is asking for #############
########################################################################
if (which_tool_audit != 'ALL'):
    all_tools_query = which_tool_audit + " || "
    all_tools_query = all_tools_query[:-3]
    string_for_user_audit_query = "select recordId, userName, dateRemoved, " + which_tool_audit + " from toolUsers.toolUsers where not isnull(dateRemoved) and (" + all_tools_query + ");"

print(all_tools_query)



####################################################
#### THIS CAN BE REMOVED ###########################
####################################################
field_names = [val[0] for val in cursor.description]
# print("=============Printing Columns of User Access Details ================")
# for user_access in user_audit_result:
      # print (user_access)

print ("Database Connected")
cursor.close()
cnx.close()

#####################################################
#### Converting to HTML table #######################
#### Write data to CSV File also ####################
#### As we are writing XLSX File Commenting CSV CODE#
#####################################################
list2d = user_audit_result
# outputfilename = "scm_user_audit_"+timestr+".csv"
# with open(outputfilename, 'w', newline='') as writeFile:
    # writer = csv.writer(writeFile)
    # writer.writerows(user_audit_result)
# writeFile.close()

#bold header
htable=u'<table border="1" bordercolor=000000 cellspacing="0" cellpadding="1" style="table-layout:fixed;vertical-align:bottom;font-size:13px;font-family:verdana,sans,sans-serif;border-collapse:collapse;border:1px solid rgb(130,130,130)" >'

##### DO NOT WRITE HTML FOR NOW
# list2d[0] = [u'<b>' + i + u'</b>' for i in list2d[0]] 
# for row in list2d:
     # newrow = u'<tr>' 
     # newrow += u' <td align="left" style="padding:1px 4px"> '+str(row[0])+u' </td> '
     # row.remove(row[0])
     # newrow = newrow + ''.join([u' <td align="right" style="padding:1px 4px"> ' + str(x) + u' </td> ' for x in row])  
     # newrow += '</tr>' 
     # htable+= newrow
# htable += '</table>'


###############################################################
######### Sending mail ########################################
###############################################################

sender = 'EngineeringSCM@extremenetworks.com'
receivers = 'cperi@extremenetworks.com'

msg = MIMEMultipart('alternative')
msg['Subject'] = "Audit Users -- Tools"
msg['From'] = sender
msg['To'] = receivers

text = "Engineering SCM Team,\nn Attached spread-sheet contains list of users who are not active but not disabled in our applications.\n\n The data is also available in attached spread-sheet for review against each application. Please review and disable the users accordingly. \n\n Engineering SCM Team\n\n"
html = """\
<html>
  <head></head>
  <body>
    <p>Engineering SCM Team,<br><br>
       Below is the list of users who are not active but not disabled in our applications.<br>
       The data is also available in attached spread-sheet for review against each application.<br>
       Please review and disable the users accordingly. <br><br>
    </p> 
  </body>
</html>
"""

part1 = MIMEText(text, 'plain')
part2 = MIMEText(html, 'html')
part3 = MIMEBase('application', "octet-stream")
part3.set_payload(open(xlsxFileName, "rb").read())
encoders.encode_base64(part3)
# part3.add_header('Content-Disposition', 'attachment', filename=outputfilename)

# Attaching Execel File
part3.add_header('Content-Disposition', 'attachment', filename=xlsxFileName)
msg.attach(part1)
msg.attach(part2)
msg.attach(part3)
message = """From: From Person cperi@extremenetworks.com
To: To Person cperi@extremenetworks.com
Subject: User Audit Results -- All Tools

This is a test e-mail message.
"""

try:
   smtpObj = smtplib.SMTP('smtp.extremenetworks.com')
   #### Commented For Now
   smtpObj.sendmail(sender, receivers, msg.as_string())         
   print ("Successfully sent email")
except SMTPException:
   print ("Error: unable to send email")

