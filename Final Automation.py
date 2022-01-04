#!/usr/bin/env python
# coding: utf-8

# In[ ]:

# this is dummy code
import pandas as pd
import numpy as np
import os as os
import datetime
import fnmatch
import calendar
import win32com.client
from collections import Counter
import re


# In[ ]:


import pysftp
import paramiko


# In[ ]:


pd.set_option('display.max_rows', 1000)


# In[ ]:


pysftp.__version__


# In[ ]:


# #Problem seems to occur if SSH-2.0-paramiko_2.6.0 client is connecting to 
# #SSH-2.0-srtSSHServer_11.00 server and agreed kex is diffie-hellman-group16-sha512.
# # run this comand for diabling server 512


# paramiko.Transport._preferred_kex = (        'ecdh-sha2-nistp256',
#         'ecdh-sha2-nistp384',
#         'ecdh-sha2-nistp521',
#         # 'diffie-hellman-group16-sha512',  # disable
#         'diffie-hellman-group-exchange-sha256',
#         'diffie-hellman-group14-sha256',
#         'diffie-hellman-group-exchange-sha1',
#         'diffie-hellman-group14-sha1',
#         'diffie-hellman-group1-sha1')


# # alternatively you can set the prefered kex as below

# # paramiko.Transport._preferred_kex = ('diffie-hellman-group-exchange-sha256',
# #                                      'diffie-hellman-group14-sha256',
# #                                      'diffie-hellman-group-exchange-sha1',
# #                                      'diffie-hellman-group14-sha1',
# #                                      'diffie-hellman-group1-sha1')


# In[ ]:





# #### Pysftp method - Establishing the connection

# In[ ]:



#change the local directory to the path where the files will be stored

#os.chdir('//infospace.emirates.com/newsites/Skywards/Sk/Advanced Analytics/Data Strategy')
#os.getcwd()


# In[ ]:


#change the local directoy to the path where the files will be stored
# Ensure you have access to this folder - Ask Unnati

os.chdir('//infospace.emirates.com/newsites/skywards_finance/ACFOUT files/')
os.getcwd()


# In[ ]:





# In[ ]:





# In[ ]:


# # TEMP CHANGING LOCAL DIR TO LOCAL MACHINE

# #change the local directoy to the path where the files will be stored

# os.chdir('C:/Users/S428545/OneDrive - emiratesgroup/Documents/ipython_stuff/skywards-jira/aa_ruchir/acfout_automation/output_files')
# os.getcwd()


# In[ ]:





# In[ ]:





# In[ ]:


# number of directories before running function

os.listdir(os.getcwd())


# In[ ]:


## make a directory within the drive if it does not exist




#os.mkdir('./Files Downloaded for {}'.format(date.today().year))

def my_folder(curr_dir):
    my_date = datetime.date.today().year
    if not os.path.exists(curr_dir + 'Files Downloaded for {}'.format(my_date)):
        os.mkdir(curr_dir + 'Files Downloaded for {}'.format(my_date))
        #curr_dir + 'Files Downloaded for {}'.format(my_date)
    return curr_dir + 'Files Downloaded for {}'.format(my_date)
    
    
# create a folder to save all files

aa = my_folder('./')

str(aa)


# In[ ]:


## make a directory within the drive if it does not exist

#os.mkdir('./Files Downloaded for {}'.format(date.today().year))

def my_folder_month(curr_dir):
    my_date = str((datetime.date.today().month)) + ' - ' + str(datetime.date.today().year)
    if not os.path.exists(curr_dir + str(aa) + '/' + 'Files Downloaded on {}'.format(my_date)):
        os.mkdir(curr_dir + str(aa) + '/'  + 'Files Downloaded on {}'.format(my_date))
        #curr_dir + 'Files Downloaded for {}'.format(my_date)
    return curr_dir + str(aa) + '/'  + 'Files Downloaded on {}'.format(my_date)
    
    
# create a folder to save all files

ab = my_folder_month('')

str(ab)


# In[ ]:


# list files in the current folder

print ('curent directories in drive:')
print
for xx in os.listdir('./' + str(aa)):
    print (xx)


# ### Connect to the SFTP

# In[ ]:


cnopts = pysftp.CnOpts()
cnopts.hostkeys = None


# In[ ]:


# # create ftp credentials


# # New details provided by Jeny - MAY 2021

# # myHost_name = 'doliadpv138.hq.emirates.com'

# myHost_name = 'hqliadpv138'
# # my_path = '/cislhome/crisuser/ssp/acfout/'

# my_path = '/cislhome/crisuser/ssp/acfout/'


# In[ ]:





# # Need SFTP details

# In[ ]:


# create ftp credentials as input prompt

 

# myHost_name = 'doliadpv138.hq.emirates.com'
myHost_name =  str(input("Type in server name within quotation marks eg 'doliadpv138.hq.emirates.com' : "))
                         
myUser_name =  str(input("Type in your Z ID within quotation marks : "))

myPasswd =  str(input("Type in your Z password within quotation marks : "))

my_path = '/cislhome/crisuser/ssp/acfout/'


# In[ ]:


myHost_name


# In[ ]:





# In[ ]:


# create ftp credentials


# New details provided by Varsha/Muna - JUL 2021

# myHost_name = 'doliadpv138.hq.emirates.com'

# myHost_name = 'DOLIADPV138'
# myHost_name_2 = 'HQLIADPV138'


# In[ ]:


# make the connection to the sftp

try:
    sftp = pysftp.Connection(host= myHost_name, username=  myUser_name, password=  myPasswd, default_path= my_path, cnopts = cnopts )
except:
    sftp = pysftp.Connection(host= myHost_name_2, username=  myUser_name, password=  myPasswd, default_path= my_path, cnopts = cnopts )


# In[ ]:


if sftp:
    print ("Connection sucess")


# In[ ]:


sftp.getcwd()


# In[ ]:


# sftp.pwd
# sftp.listdir()


# In[ ]:





# In[ ]:


# check if the drive has files

num_files_to_show = 10

print('first {} files in the folder:\n'.format(num_files_to_show))

for i in sftp.listdir('.')[: num_files_to_show]:
    print (i)


# In[ ]:




# validate the directories defined are the ones you need

print ('\nvalidate remote directory')
print (sftp.pwd)
# print('')
print ('\nvalidate local directory')
print (os.getcwd())


# In[ ]:


# format the directories for them to be used as path to download and transfer

remote_direct = sftp.pwd + '/'
local_direct = str(ab) + '/'


# current remote directorry & local dorectory
# define path variables


print ('define directories:')
print ('')
print ('remote directory:', remote_direct)

print ('local directory:', local_direct)


# In[ ]:


# create a list of all partners 

list_b = 'AS' # test list

list_a = ['EHG','ATLN','CORL','CRWN','HLTN','HYAT','ICH','JBH', 'FAB','MARR','MHG','ONOR','ROTN',
'RSAS','SFTL','SGL','RKTM','AAD','CREM','CTOD','DMS','EKH2','GYG','SKI','TFG','TRG','VALR',
'AMME','SWED','JLNZ','PNTS','EMAR','BCOM','NAMS','DTAG','ECAR','HRTZ','SIXT','ABG','CTLR','ABK',
'ADCB','CAUS','CBBH','CBD','CITI','EIB','ENBD','NBAD','NOOR','SCIN','SCPK','HSUS','ADIB','HSCN',
'AXPD','AXPI','DIB','CHSE','SABB','KLGO','AXPZ','HLB','RAK','EBI','EIB','LHRR', 'BUS','HSTW','BUS',
'SMAL','CRFE','EOS','HSCN', 'MSLF', 'DOLL', 'MSRF', 'CTRW', 'MASH', 'BILT', 'NAJM', 'HSHK',
         'HSTW','UTU','NBK']



list_c = ['AS','B6','CM','G3','JL','JQ','KE','MH','MK','QF','S7','SA','TP','PG']

list_d = ['EKH']


# In[ ]:





# In[ ]:


todaydate = datetime.date.today()

todaydate


# In[ ]:


lastDayLastMonth = todaydate - datetime.timedelta(days=todaydate.day)

lastDayLastMonth


# In[ ]:


local_direct + str("EHG") + '_ACFOUT_' +                           str.upper(calendar.month_name[lastDayLastMonth.month][:3])                           + str(lastDayLastMonth.year) +'.TXT'


# In[ ]:


os.path.isfile(local_direct + str("EHG") + '_ACFOUT_' +                           str.upper(calendar.month_name[lastDayLastMonth.month][:3])                           + str(lastDayLastMonth.year) +'.TXT')


# In[ ]:


# Notes 

# The SFTP file name has TODAY'S DATE
# The sharepoint file has LAST MONTH'S DATE


# In[ ]:


# loop through all files in the directory and transnfer each file

y = 0
for i in sftp.listdir(remote_direct):
    for x in list_a:
        if os.path.isfile( local_direct + str(x) + '_ACFOUT_' + 
                          str.upper(calendar.month_name[lastDayLastMonth.month][:3]) 
                          + str(lastDayLastMonth.year) +'.TXT'):
            continue
            
        if i.find(str(x)) == 0 and fnmatch.fnmatch(i,'*.txt')         and str(i[-12:-8])== str(todaydate.year) and         (todaydate.month) == int(i[-8:-6]):
            
            sftp.get(remote_direct + str(i) , local_direct +  str(x) + 
                         '_ACFOUT_' + str.upper(calendar.month_name[lastDayLastMonth.month][:3]) \
                     + str(lastDayLastMonth.year) +'.txt' )
            y = y + 1

print ('Addtional File/s Transferred = {}'.format( y))
print ('Total Files in Directory = {}'.format(len(os.listdir(local_direct))))


# In[ ]:



# loop for all AIRLINE partners only, since the suffix is only ywo characters

y1 = 0
for i1 in sftp.listdir(remote_direct):
    for x1 in list_c:
        if os.path.isfile( local_direct + str(x1) + '_ACFOUT_' + 
                          str.upper(calendar.month_name[lastDayLastMonth.month][:3]) 
                          + str(lastDayLastMonth.year ) +'.TXT'):
            continue
            
        if re.findall('\\b{}\\b'.format(x1), i1[0:2]) and i1[2:5] == 'ACF' and         str(i1[-12:-8])== str(todaydate.year) and (todaydate.month) == int(i1[-8:-6]):
            
            sftp.get(remote_direct + str(i1) , local_direct + str(x1) + 
                     '_ACFOUT_' + str.upper(calendar.month_name[lastDayLastMonth.month][:3]) 
                     + str(lastDayLastMonth.year) +'.txt' )
            y1 = y1 + 1

print ('Addtional File/s Transferred = {}'.format( y1))
print ('Total Files in Directory = {}'.format(len(os.listdir(local_direct))))
            

# print ('Addtional File/s Transferred = {}'.format( y1))
# print ('Total Files in Directory = {}'.format(len(os.listdir(local_direct))))


# In[ ]:



# loop for all EKH partner only, since EKH and EKH2 are very similar

y1 = 0
for i1 in sftp.listdir(remote_direct):
    for x1 in list_d:
        if os.path.isfile( local_direct + str(x1) + '_ACFOUT_' + 
                          str.upper(calendar.month_name[lastDayLastMonth.month][:3]) 
                          + str(lastDayLastMonth.year ) +'.TXT'):
            continue
            
        if re.findall('\\b{}\\b'.format(x1), i1[0:3]) and i1[3:6] == 'ACF' and         str(i1[-12:-8])== str(todaydate.year) and (todaydate.month) == int(i1[-8:-6]):
            
            sftp.get(remote_direct + str(i1) , local_direct + str(x1) + 
                     '_ACFOUT_' + str.upper(calendar.month_name[lastDayLastMonth.month][:3]) 
                     + str(lastDayLastMonth.year) +'.txt' )
            y1 = y1 + 1

print ('Addtional File/s Transferred = {}'.format( y1))
print ('Total Files in Directory = {}'.format(len(os.listdir(local_direct))))
            

# print ('Addtional File/s Transferred = {}'.format( y1))
# print ('Total Files in Directory = {}'.format(len(os.listdir(local_direct))))


# In[ ]:





# ### Send out the final email

# ## Send Automated Email

# In[ ]:


# create a list of all the files that have been transfered for sending an email

b = []
for i in os.listdir(local_direct):
    #a = (i[:-19])
    b.append(i)
tester = dict((i, b.count(i)) for i in b)
email_grid = pd.DataFrame(tester.items(), columns= ['Partner', 'FileTransfer_Count'])
email_grid.set_index('Partner', inplace = True)
len(email_grid)


# In[ ]:



email_grid.sort_index(inplace=True)


# In[ ]:


# send an email message via outlook

# Hard coded email subject
MAIL_SUBJECT = 'ACFOUT files transfered to SFTP'

# Hard coded email text
        
MAIL_BODY = 'Dear Team,\n\n''The ACFOUT file transfer for {} has been Successful.\n'.format(calendar.month_name[datetime.datetime.today().month] + str(datetime.datetime.today().year))
MAIL_BODY2 = '\nTotal count of files present in the folder are: {}\n'.format(len(email_grid))
MAIL_BODY3 = '\nThe number of files transferred can be found in the grid below\n\n {}'.format(email_grid)


def send_outlook_mail(recipients, subject='No Subject', body='Blank', body2 ='Blank'
                      ,body3 = 'Blank',send_or_display='Display', copies=None):
    """
    Send an Outlook Text email
    :param recipients: list of recipients' email addresses (list object)
    :param subject: subject of the email
    :param body: body of the email
    :param send_or_display: Send - send email automatically | Display - email gets created user have to click Send
    :param copies: list of CCs' email addresses
    :return: None
    """
    if len(recipients) > 0 and isinstance(recipient_list, list):
        outlook = win32com.client.Dispatch("Outlook.Application")

        ol_msg = outlook.CreateItem(0)

        str_to = ""
        for recipient in recipients:
            str_to += recipient + ";"

        ol_msg.To = str_to

        if copies is not None:
            str_cc = ""
            for cc in copies:
                str_cc += cc + ";"

            ol_msg.CC = str_cc

        ol_msg.Subject = subject
        ol_msg.Body = body + body2 + body3

        if send_or_display.upper() == 'SEND':
            ol_msg.Send()
        else:
            ol_msg.Display()
    else:
        print('Recipient email address - NOT FOUND')


# In[ ]:


recipient_list = ['sanjeevi.radhakrishnan@emirates.com',
                  'jeny.binil@emirates.com','sara.alshamsi@emirates.com',
                  'shabeena.aljariri@emirates.com', 'rajani.madhavan@emirates.com',
                 'masood.sarang@emirates.com', 
                  'beatrice.botelho@emirates.com', 'nazrine.nawabjan@emirates.com',
                 'skywards.systems@emirates.com' ]
copies_list = ['ruchir.varma@emirates.com', 'clyton.dcruz@emirates.com']

send_outlook_mail(recipients=recipient_list, subject=MAIL_SUBJECT, body= MAIL_BODY , 
                  body2= MAIL_BODY2 , body3= MAIL_BODY3, send_or_display='DISPLAY',copies=copies_list)


# In[ ]:




