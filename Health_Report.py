from Queries import DBQuery
import smtplib
import tabulate
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

# @ todo:make sure that 021 works with 1 or 2 rows - check num rows
path = Path(__file__).parent.absolute()
report = DBQuery(str(path) + '/Presentation/IBF_AUTOMATED_MS_Report.pptx')
msg = ""
msg2 = ""
msg2 = msg2 + report.ibs_kpi_020() # done
msg = msg + report.ibs_kpi_021() # done
report.ibs_ms_001() # table only
report.ibs_ms_003()
report.coverpage()

final_report = report
print(report)

# @todo: Mapping of IBF errors
# @todo: Email content needs to be sent


# server = smtplib.SMTP('10.10.10.1', 25)
#
# #Send the mail
# # The /n separates the message from the headers
# server.sendmail("MS_REPORT@globetom.com", "ritesh.doolabh@globetom.com", msg)


sender_email = "MS_REPORT@globetom.com"
receiver_email = "ritesh.doolabh@globetom.com"

message = MIMEMultipart("alternative")
message["Subject"] = "Daily Ad-Hoc Report"
message["From"] = sender_email
message["To"] = receiver_email

# Create the plain-text and HTML version of your message
text = """\
Hi, \n

MS_Team Daily Report \n \n
"""
html = msg + msg2

# Turn these into plain/html MIMEText objects
part1 = MIMEText(text, "plain")
part2 = MIMEText(html, "html")

# Add HTML/plain-text parts to MIMEMultipart message
# The email client will try to render the last part first
message.attach(part1)
message.attach(part2)



# Open PDF file in binary mode
with open(report.presentation_name, "rb") as attachment:
 # Add file as application/octet-stream
 # Email client can usually download this automatically as attachment
 part = MIMEBase("application", "octet-stream")
 part.set_payload(attachment.read())

# Encode file in ASCII characters to send by email
encoders.encode_base64(part)

# Add header as key/value pair to attachment part
part.add_header(
 "Content-Disposition",
 f"attachment; filename= {report.presentation_name}",
)

# Add attachment to message and convert message to string
message.attach(part)
text = message.as_string()

server = smtplib.SMTP('10.10.10.1', 25)
# server.sendmail("MS_REPORT@globetom.com", ["ritesh.doolabh@globetom.com","marne.meades@globetom.com","quintin.vorster@globetom.com","stephan.kruger@globetom.com"], message.as_string())
server.sendmail("IBF_MS_REPORT@globetom.com", ["ritesh.doolabh@globetom.com"], message.as_string())
