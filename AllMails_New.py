import imaplib
import email
from email.header import decode_header
import webbrowser
import os
from bs4 import BeautifulSoup
import pandas as pd
from csv import writer

# account credentials
username = 'prathap.sariputi@carbynetech.com'
password = 'sP@161993sp'
# use your email provider's IMAP server, you can look for your provider's IMAP server on Google
# or check this page: https://www.systoolsgroup.com/imap/
# for office 365, it's this:
imap_server = "outlook.office365.com"


def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)

# create an IMAP4 class with SSL, use your email provider's IMAP server
imap = imaplib.IMAP4_SSL(imap_server)
# authenticate
imap.login(username, password)

# select a mailbox (in this case, the inbox mailbox)
# use imap.list() to get the list of mailboxes
status, messages = imap.select("INBOX")

# total number of emails
messages = int(messages[0])
print(messages)


for i in range(messages, messages-1, -1):
    # fetch the email message by ID
    res, msg = imap.fetch(str(i), "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            # parse a bytes email into a message object
            msg = email.message_from_bytes(response[1])
            Date = msg["Date"]
            if msg.is_multipart():
                for part in msg.walk():
                    # extract content type of email
                    content_type = part.get_content_type()
                    print("1.",content_type)
                    content_disposition = str(part.get("Content-Disposition"))
                    try:
                        # get the email body
                        body = part.get_payload(decode=True).decode()
                    except:
                        pass
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        # print text/plain emails and skip attachments
                        #print(body)
                        body=body
            else:
                # extract content type of email
                content_type = msg.get_content_type()
                print("2.",content_type)
                # get the email body
                body = msg.get_payload(decode=True).decode()
                
            #print(body)
            soup = BeautifulSoup(body,  features='html.parser')

            # kill all script and style elements
            for script in soup(["script", "style"]):
                script.extract()    # rip it out

            # get text
            text = soup.get_text()

            # break into lines and remove leading and trailing space on each
            lines = (line.strip() for line in text.splitlines())
            # break multi-headlines into a line each
            chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
            # drop blank lines5
            text = '\n'.join(chunk for chunk in chunks if chunk)

            #print(text)
            text=text.encode("ascii", "ignore").decode()
            print(text)
            body=text[(text.find("ago")+3):text.find("Reject")].split("\n")
            name=body[1]
            currentCompany=body[3].split(" at ")[1]
            contactNumber=body[4]
            mailId=body[5]
            education=body[body.index("Education")+1]
            ctc=body[7]
            noticePeriod=body[body.index("Notice Period")+1]
            currentLocation=body[body.index("Location")+1].split("(")[0]
            preferedLocation = body[body.index("Location")+1].split("(")[1].split(" is ")[1][:-1]
            keySkills = body[body.index("Keyskills")+1]
            currentPosition = body[3].split(" at ")[0]
            list_data =[Date,name,contactNumber,mailId,education,ctc,noticePeriod,currentCompany,currentLocation,preferedLocation,keySkills,currentPosition]
            print(list_data)
            with open('file5.csv', 'a', newline='') as f_object:  
                # Pass the CSV  file object to the writer() function
                writer_object = writer(f_object)
                # Result - a writer object
                # Pass the data in the list as an argument into the writerow() function
                writer_object.writerow(list_data)  
                # Close the file object
                f_object.close()

                
            print("="*100)

# close the connection and logout
imap.close()
imap.logout()
            
        

