#!/usr/bin/env python3

"""Create and email spreadsheet of patrons expiring next month

Author: Nina Acosta
"""

import psycopg2
import xlsxwriter
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from datetime import date, timedelta
next_month = date.today() + timedelta(32)
month_year = next_month.strftime ("%B %Y") #Formats date for subject and body of email
file_name = next_month.strftime ("%m%Y") #Formats date for filename of Excel attachment

#SQL Query :
q='''SELECT
CONCAT (patron_record_fullname.first_name, ' ', patron_record_fullname.middle_name, ' ', patron_record_fullname.last_name) AS "PATRON NAME",
barcode AS "PATRON BARCODE",
TO_CHAR(expiration_date_gmt, 'MM/DD/YYYY') AS "EXPIRATION DATE",
home_library_code AS "HOME LIBRARY",
field_content AS "EMAIL ADDRESS"

FROM
sierra_view.patron_view
JOIN sierra_view.patron_record_fullname
ON patron_record_fullname.patron_record_id = patron_view.id
JOIN sierra_view.varfield_view
ON varfield_view.record_id = patron_view.id

WHERE
varfield_type_code = 'z' AND
expiration_date_gmt >= DATE_TRUNC('month', now()) + interval '1 month' AND
expiration_date_gmt < DATE_TRUNC('month', now()) + interval '2 months'
--Finds all patrons with an email address on file that have expiration dates after the current month, and before 2 months from now

ORDER BY "HOME LIBRARY"
'''

#Name of Excel File
excelfile = "Expiring"+str(file_name)+".xlsx"

# These are variables for the email that will be sent.
# This code uses placeholders, please add your own email server info
emailhost = 'email.server.midhudson.org'
emailuser = 'emailaddress@midhudson.org'
emailpass = '*******'
emailport = '587'
emailsubject = 'Patrons Expiring Soon report : ' + str(month_year)
emailmessage = '''***This is an automated email***

The attached spreadsheet contains a list of patron records expiring in ''' + str(month_year) + '''.
Please update the report "Patrons Expiring Soon" in the Monthly/Quarterly Reports section of kb.midhudson.org'''
emailfrom= 'emailaddress@midhudson.org'
emailto = 'nacosta@midhudson.org'

#This code uses placeholder info to connect to Sierra SQL server, please replace with your own info
conn = psycopg2.connect("dbname='iii' user='*****' host='000.000.000.000' port='1032' password='*****' sslmode='require'")

#Open session and run query
cursor = conn.cursor()
cursor.execute(q)
rows = cursor.fetchall()
conn.close()

#Create Excel file
import xlsxwriter
workbook = xlsxwriter.Workbook("Expiring"+str(file_name)+".xlsx")#Adds relative date to the filename
worksheet = workbook.add_worksheet()

#Formatting our Excel worksheet
worksheet.set_landscape()
worksheet.hide_gridlines(0)

#Formatting Cells
eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'vcenter', 'bold': True})
bold = workbook.add_format({'valign': 'top','bold': True})


# Setting the column widths
worksheet.set_column(0,0,40.00)
worksheet.set_column(1,1,22.00)
worksheet.set_column(2,2,16.14)
worksheet.set_column(3,3,10.29)
worksheet.set_column(4,4,40.00)

# Adding column labels
worksheet.write(0,0,"PATRON NAME", eformatlabel)
worksheet.write(0,1,"BARCODE", eformatlabel)
worksheet.write(0,2,"EXPIRATION DATE", eformatlabel)
worksheet.write(0,3,"HOME LIBR", eformatlabel)
worksheet.write(0,4,"EMAIL ADDRESS", eformatlabel)

# Writing the report to the Excel worksheet
for rownum, row in enumerate(rows):
    worksheet.write(rownum+1,0,row[0], eformat)
    worksheet.write(rownum+1,1,row[1],eformat)
    worksheet.write(rownum+1,2,row[2],eformat)
    worksheet.write(rownum+1,3,row[3],eformat)
    worksheet.write(rownum+1,4,row[4],eformat)
workbook.close()

#Create an email with an attachement
msg = MIMEMultipart()
msg['From'] = emailfrom
if type(emailto) is list:
    msg['To'] = ', '.join(emailto)
else:
    msg['To'] = emailto
msg['Date'] = formatdate(localtime = True)
msg['Subject'] = emailsubject
msg.attach (MIMEText(emailmessage))
part = MIMEBase('application', "octet-stream")
part.set_payload(open(excelfile,"rb").read())
encoders.encode_base64(part)
part.add_header('Content-Disposition','attachment; filename=%s' % excelfile)
msg.attach(part)

#Send the email
smtp = smtplib.SMTP(emailhost, emailport)
#for Google connection
smtp.ehlo()
smtp.starttls()
smtp.login(emailuser, emailpass)
#end for Google connection
smtp.sendmail(emailfrom, emailto, msg.as_string())
smtp.quit()
