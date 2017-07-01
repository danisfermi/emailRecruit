#!/usr/bin/python
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders
import xlrd # For working with Excel

# Class to hold Recruiter Object
# Recruiter Object has following attributes:-
# First Name of Recruiter (String)
# Last Name of Recruiter (String)
# Company Name of Recruiter (String)
# Mail ID of Recruiter (String)
# Job Role of Opening (String)
# Job ID of Opening (String)
class Recruiter(object):
    def __init__(self, firstName, lastName, company, mailID, position, jobID):
        self.firstName = firstName
        self.lastName = lastName
        self.company = company
        self.mailID = mailID
        self.position = position
        self.jobID = jobID

    def __str__(self):
        return("Recruiter object:\n"
               "  firstName = {0}\n"
               "  lastName = {1}\n"
               "  Company = {2}\n"
               "  mailID = {3}\n"
               "  position = {4} \n"
               "  jobID = {5}"
               .format(self.firstName, self.lastName, self.company,
                       self.mailID, self.position, self.jobID))

# Function to read template mail from mailTemplate and replace generic strings with Company specific values
# Takes Name, Company, Position as arguments
# Returns Mail Body
def getBody(name, company, position):
    try:
        file = open("mailTemplate", "r")
        body = file.read()
        file.close()
        return body.replace("$insertName$", name).replace("$insertCompany$", company).replace("$insertPosition$", position)
    except IOError:
        print "Error: File does not seem to exist."
        return "Hello"

fromaddr = "danisfermijohn@gmail.com"

msg = MIMEMultipart()
msg['From'] = fromaddr

# Login to Gmail
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(fromaddr, "YOUR PASSWORD") # Replace with your password here

# Read from and Attach Resume to Mail
filename = "DanisFermiResume.pdf"
attachment = open(filename, "rb")
part = MIMEBase('application', 'octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
msg.attach(part)

# Opening and Parsing Excel data into list of Recruiter Objects
wb = xlrd.open_workbook("recruiters.xlsx")
for sheet in wb.sheets():
    number_of_rows = sheet.nrows
    number_of_columns = sheet.ncols
    recruiters = []

    rows = []
    for row in range(1, number_of_rows):
        values = []
        for col in range(number_of_columns):
            value  = (sheet.cell(row,col).value)
            try:
                value = str(int(value))
            except ValueError:
                pass
            finally:
                values.append(value)
        recruiter = Recruiter(*values)
        recruiters.append(recruiter)

# Iterate over recruiter list and send email
for recruiter in recruiters:
    msg['To'] = recruiter.mailID
    msg['Subject'] = "Danis Fermi-MS Student-Job Application for $insertPosition$ position".replace("$insertPosition$", recruiter.position)
    if recruiter.jobID != "":
        msg['Subject'].append("(Job Id: "+recruiter.jobID+")")

    body = getBody(recruiter.firstName, recruiter.company, recruiter.position)
    msg.attach(MIMEText(body, 'plain'))

    text = msg.as_string()
    server.sendmail(fromaddr, recruiter.mailID, text)
server.quit()