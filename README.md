# emailRecruit
A simple Python script to send Emails with Resume to Recruiters

### Sending Email as an Attachment

The script uses the native Python smtplib library. It takes the server-name (or IP address) and the port number as parameters. 
To send a clean email with proper subject line and receiver, email.MIMEMultipart and email.MIMEText modules will be used.
Attaching involves converting the attachment to Base64. This method works for text files, pdf files, images, audio and video files, just incase you wanna get more creative with your resumes :D

### Body of the Mail

The script uses the inbuilt string replace function to replace the generic string variables with company spcific values. Implemented in function getBody(), explained with inline comments.

### Reading and Parsing from Excel

The script uses Excel's xlrd Python library. A class is used to iterate over and store recruiter details. If you use a class, you can access all values via dot-notation (alternative would have been a dictionary).