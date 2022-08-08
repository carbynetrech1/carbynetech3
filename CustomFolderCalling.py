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
print(imap)

for i in imap.list()[1]:
    l = i.decode().split(' "/" ')
    print(l[0] + " = " + l[1])

# select a mailbox (in this case, the inbox mailbox)
# use imap.list() to get the list of mailboxes
status, messages = imap.select('"IGL Alerts"')
print(status, messages )

# total number of emails
messages = int(messages[0])

for i in range(messages, messages-1, -1):
    # fetch the email message by ID
    res, msg = imap.fetch(str(i), "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            # parse a bytes email into a message object
            msg = email.message_from_bytes(response[1])
            From = msg["From"]
            To = msg["To"]
            Cc = msg["CC"]
            Bcc = msg["BCC"]
            Date = msg["Date"]
            subject = msg["Subject"]
            #print('From:',From,'To:',To,'Cc:',Cc,'Date:',Date,"Subject : ", subject,"BCC:",Bcc)
            # decode the email subject
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                # if it's a bytes, decode to str
                if subject!=None:
                    subject = subject.decode(encoding)
                
            # decode email sender
            From, encoding = decode_header(msg.get("From"))[0]
            if isinstance(From, bytes):
                From = From.decode(encoding)

            # if the email message is multipart
            if msg.is_multipart():
                # iterate over email parts
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
    
                   # elif "attachment" in content_disposition:
                        # download attachment
                        #filename = part.get_filename()
                        #if filename:
                           # folder_name = clean(subject)
                            #if not os.path.isdir(folder_name):
                                # make a folder for this email (named after the subject)
                                #os.mkdir(folder_name)
                           # filepath = os.path.join(folder_name, filename)
                            # download attachment and save it
                            #open(filepath, "wb").write(part.get_payload(decode=True))
            else:
                # extract content type of email
                content_type = msg.get_content_type()
                print("2.",content_type)
                # get the email body
                body = msg.get_payload(decode=True).decode()
                #if content_type == "text/plain":
                    # print only text email parts
                    #print(body)
            if content_type == "text/html":
                # if it's HTML, create a new HTML file and open it in browser
                folder_name = clean(subject)
                if not os.path.isdir(folder_name):
                    # make a folder for this email (named after the subject)
                    os.mkdir(folder_name)
                filename = "index.html"
                filepath = os.path.join(folder_name, filename)
                # write the file
                #open(filepath, "w").write(body)
                
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
            subject=subject.encode("ascii", "ignore").decode()
            
            list_data =[From,To,Cc,Bcc,Date,subject,text]
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


