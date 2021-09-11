###############################################################################
# Email Scraper Program
# Made by Jay Orton
# This program will connect to an email inbox and get the emails for todays date.
# It will then search each email for keywords provided in the keywords.txt file.
# Next, it will create an excel file with all of the keyword mathces and emails in a table like configuration.
# Lastly, it will email this excel file to a recipient.
###############################################################################
# Got help from the following, see for more info:
# https://humberto.io/blog/sending-and-receiving-emails-with-python/
# https://stackoverflow.com/questions/52054196/python-imaplib-search-email-with-date-and-time
# https://www.geeksforgeeks.org/working-csv-files-python/
# https://zetcode.com/python/smtplib/
# https://stackoverflow.com/questions/3362600/how-to-send-email-attachments
###############################################################################
import email
import imaplib
import smtplib
import datetime
import csv
import os
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from os.path import basename
import re
import pyap

#config
EMAIL = "orderrequests@impactvaluation.com"
PASSWORD = "Sum27822"
SERVER = "smtp.office365.com"
PORT = 587
RECIPIENTS = ["appraisalorders@impactvaluation.com"]
#RECIPIENTS = ["orderrequests@impactvaluation.com"]
DELETE = True
SEND = True

def main():
    print("Email Scraper: Getting keywords...");
    #Note that scraper needs to be case insensitive
    #Get keywords
    keywordfile = open(os.path.join(os.getcwd(),"keywords.txt"),"r");
    keywords = keywordfile.read().splitlines()
    keywords.sort(key=len, reverse=True)
    keywordfile.close();
    #parse keywords to remove blank strings
    while("" in keywords):
        keywords.remove("")
    
    print("Email Scraper: getting property addresses...");
    #Parse property address file, and put into list
    global addressfile
    addressfile = open(os.path.join(os.getcwd(),"property_addresses.txt"),"r+");
    global addresslist
    addresslist = addressfile.read().splitlines()
    addresslist.sort(key=len, reverse=True)
    #parse keywords to remove blank strings
    while("" in addresslist):
        addresslist.remove("") #Address list is a list of addresses.
    
    print("Email Scraper: Getting Client names...");
    #Note that scraper needs to be case insensitive
    #Get keywords
    clientfile = open(os.path.join(os.getcwd(),"clients.txt"),"r");
    clients = clientfile.read().splitlines()
    clients.sort(key=len, reverse=True)
    clientfile.close();
    #parse keywords to remove blank strings
    while("" in clients):
        clients.remove("")
    
    
    emailcount = {}
    keyword_hits = {}
    client_hits = {}
    client_keywords = {}
    #get emails
    print("Email Scraper: getting emails...");
    emails = getEmails()
    for email in emails:
        try:
            emailcount[email["From"]] += 1
        except KeyError:
            emailcount[email["From"]] = 1
    #print(emailcount)
    
    #for each email
    aaaa = 0;
    for email in emails:
        aaaa += 1;
        print("\nEmail Scraper: Scraper progress: email " + str(aaaa) + " of " + str(len(emails)) + "\n");
        #Attempt to get the property address in the email.
        #If there was no property address, read the email.
        #If there was a property address, check to see if its already in the list.
        #If not, read the email, and add the address to the list.
        #If it was already in the list, don't read that email (as its a duplicate)
        print("Email Scraper: Evaluating Property address of email " + str(aaaa));
        canEval = evalPropertyAddress(email);
        #if email address didnt exist in dict, initialize it.
        if(canEval):
            print("Email Scraper: Getting keyword hits...");
            try:
                keyword_hits[email["From"]]
            except:
                keyword_hits[email["From"]] = {}
            #temp_body
            try:
                temp_body = email["Body"].lower()
            except:
                temp_body = email["Body"] = ""
            try:
                temp_subject = email["Subject"].lower()
            except:
                temp_subject = email["Subject"] = ""
            #for each keyword
            getkeyword = True;
            for keyword in keywords:
                hits = 0;
                if(getkeyword):
                    #get hits of keyword
                    temp_body = temp_body.split(keyword.lower());
                    if(len(temp_body) > 1):
                        hits = 1;
                        getkeyword = False;
                    else:
                        hits = 0;
                    #rejoin body for when the next keyword goes
                    temp_body = " ".join(temp_body);
                    #if email.keyword didnt exist, initialize to 0, then add the hits
                    #if it did exist, just add the hits.
                    
                if(getkeyword):
                    #same as above but using the subject this time
                    #get hits of keyword
                    temp_subject = temp_subject.split(keyword.lower());
                    if(len(temp_subject) > 1): 
                        hits = 1;
                        getkeyword = False;
                    else:
                        hits = 0;
                    #rejoin subject for when the next keyword goes
                    temp_subject = " ".join(temp_subject);
                    #dict value is already initialized, just add hits
                try:
                    keyword_hits[email["From"]][keyword]
                except:
                    keyword_hits[email["From"]][keyword] = 0
                keyword_hits[email["From"]][keyword] += hits
                #if it found a keyword hit (hits is not 0), then we repeat this process, instead looking
                #for Client names. If it found one we associate the client to both the keyword and the
                #appraiser (email address)
                getclient = True;
                for client in clients:
                    clientsfound = 0;
                    try:
                        client_hits[email["From"]]
                    except:
                        client_hits[email["From"]] = {}
                        
                    try:
                        client_keywords[client]
                    except:
                        client_keywords[client] = {}
                    if(getclient and hits != 0):
                        #get hits of keyword
                        temp_body = temp_body.split(client.lower());
                        if(len(temp_body) > 1):
                            clientsfound = 1;
                            getclient = False;
                        else:
                            clientsfound = 0;
                        #rejoin body for when the next keyword goes
                        temp_body = " ".join(temp_body);
                        #if email.keyword didnt exist, initialize to 0, then add the hits
                        #if it did exist, just add the hits.
                            
                    if(getclient and hits != 0):
                        #same as above but using the subject this time
                        #get hits of keyword
                        temp_subject = temp_subject.split(client.lower());
                        if(len(temp_subject) > 1): 
                            clientsfound = 1;
                            getclient = False;
                        else:
                            clientsfound = 0;
                        #rejoin subject for when the next keyword goes
                        temp_subject = " ".join(temp_subject);
                        #dict value is already initialized, just add hits
                    try:
                        client_hits[email["From"]][client]
                        client_keywords[client][keyword]
                    except:
                        client_hits[email["From"]][client] = 0
                        client_keywords[client][keyword] = 0
                    client_hits[email["From"]][client] += clientsfound
                    client_keywords[client][keyword] += clientsfound
    #now we have all of the keyowrd hits associated with their email address.
    #now we need to create a csv file, and write the keywords as fields (and email address!),
    #and all of the keyword hits as rows.
    print("Email Scraper: keyword search complete. constructing csvs...");
    keywords.insert(0,"Email")
    fields = keywords
    rows = []
    for z,v in keyword_hits.items():
        #z is email address,
        #v is the items.
        temprow = []
        temprow.append(z)
        for key in v:
            temprow.append(v[key])
        rows.append(temprow)

    #now lets create the name of the file and write to it.
    csvfilename = "email_report_for_" + date.today().strftime('%Y-%m-%d') + "_appraiser-keywords.csv"
    # addressfile = open(os.path.join(os.getcwd(),"property_addresses.txt"),"r+");
    with open(os.path.join(os.getcwd(), csvfilename), 'w') as csvfile:
        # creating a csv writer object
        csvwriter = csv.writer(csvfile)
          
        # writing the fields
        csvwriter.writerow(fields)
          
        # writing the data rows
        csvwriter.writerows(rows)
    
    #Make the csv for Client - Keywords
    rows = []
    for z,v in client_keywords.items():
        #z is email address,
        #v is the items.
        temprow = []
        temprow.append(z)
        for key in v:
            temprow.append(v[key])
        rows.append(temprow)

    #now lets create the name of the file and write to it.
    csv_client_keywords = "email_report_for_" + date.today().strftime('%Y-%m-%d') + "-client_keywords.csv"
    # addressfile = open(os.path.join(os.getcwd(),"property_addresses.txt"),"r+");
    with open(os.path.join(os.getcwd(), csv_client_keywords), 'w') as csvfile:
        # creating a csv writer object
        csvwriter = csv.writer(csvfile)
          
        # writing the fields
        csvwriter.writerow(fields)
          
        # writing the data rows
        csvwriter.writerows(rows)
    
    clients.insert(0,"Client")
    fields = clients
    #Make the csv for Appraiser - Clients
    rows = []
    for z,v in client_hits.items():
        #z is email address,
        #v is the items.
        temprow = []
        temprow.append(z)
        for key in v:
            temprow.append(v[key])
        rows.append(temprow)

    #now lets create the name of the file and write to it.
    csv_appraiser_client = "email_report_for_" + date.today().strftime('%Y-%m-%d') + "-appraiser_clients.csv"
    # addressfile = open(os.path.join(os.getcwd(),"property_addresses.txt"),"r+");
    with open(os.path.join(os.getcwd(), csv_appraiser_client), 'w') as csvfile:
        # creating a csv writer object
        csvwriter = csv.writer(csvfile)
          
        # writing the fields
        csvwriter.writerow(fields)
          
        # writing the data rows
        csvwriter.writerows(rows)
        
    #Send CSV
    print("Email Scraper: Sending csv file...");
    if(SEND):
        sendMail([csvfilename,csv_client_keywords,csv_appraiser_client])
    addressfile.close();
#end of main

def getEmails():
    print("Email Scraper : start email retrieval...");
    email_return = []
    today = date.today()
    today = today.strftime('%d-%b-%Y') #06-Apr-2021
    print("getting inbox...");
    # connect to the server and go to its inbox
    mail = imaplib.IMAP4_SSL(SERVER)
    mail.login(EMAIL, PASSWORD)
    # we choose the inbox but you can select others
    mail.select('inbox')

    # we'll search using the ALL criteria to retrieve
    # every message inside the inbox
    # it will return with its status and a list of ids
    status, data = mail.search(None, 'ALL')
    # the list returned is a list of bytes separated
    # by white spaces on this format: [b'1 2 3', b'4 5 6']
    # so, to separate it first we create an empty list
    mail_ids = []
    # then we go through the list splitting its blocks
    # of bytes and appending to the mail_ids list
    data = data[0].split(b' ')
    for block in data:
        # the split function called without parameter
        # transforms the text or bytes into a list using
        # as separator the white spaces:
        # b'1 2 3'.split() => [b'1', b'2', b'3']
        mail_ids += block.split()

    # now for every id we'll fetch the email
    # to extract its content
    bbbb = 0;
    for i in mail_ids:
        bbbb += 1;
        print("fetching progress : email " + str(bbbb) + " of " + str(len(mail_ids)));
        # the fetch function fetch the email given its id
        # and format that you want the message to be
        status, data = mail.fetch(i, '(RFC822)')

        # the content data at the '(RFC822)' format comes on
        # a list with a tuple with header, content, and the closing
        # byte b')'
        for response_part in data:
            # so if its a tuple...
            if isinstance(response_part, tuple):
                # we go for the content at its second element
                # skipping the header at the first and the closing
                # at the third
                message = email.message_from_bytes(response_part[1])

                # with the content we can extract the info about
                # who sent the message and its subject
                mail_from = message['from']
                mail_subject = message['subject']

                # then for the text we have a little more work to do
                # because it can be in plain text or multipart
                # if its not plain text we need to separate the message
                # from its annexes to get the text
                if message.is_multipart():
                    mail_content = ''

                    # on multipart we have the text message and
                    # another things like annex, and html version
                    # of the message, in that case we loop through
                    # the email payload
                    for part in message.get_payload():
                        # if the content type is text/plain
                        # we extract it
                        if part.get_content_type() == 'text/plain':
                            mail_content += part.get_payload()
                else:
                    # if the message isn't multipart, just extract it
                    mail_content = message.get_payload()

                # and then let's show its result
                email_dict = {}
                email_dict["From"] = mail_from
                email_dict["Subject"] = mail_subject
                email_dict["Body"] = mail_content
                email_return.append(email_dict)
        #Send email to trash to keep inbox clean
        if(DELETE):
            mail.store(i, '+FLAGS', '\\Deleted');
    #Delete the items
    print("Email Scraper: Sending fetched emails to trash...");
    if(DELETE):
        mail.expunge();
    print("Email Scraper: Email Retrieval Complete");
    return email_return

def sendMail(files = []):
    print("Email Scraper: Building and sending csv...");
    mail_subject = "Email Report for " + date.today().strftime('%Y-%m-%d')
    
    mail_to_string = ', '.join(RECIPIENTS)
    
    mail_message = f'''Here are the email reports for today.'''
    msg = MIMEMultipart()
    msg["From"] = EMAIL
    msg["To"] = mail_to_string
    msg["Subject"] = mail_subject
    msg.attach(MIMEText(mail_message))
    
    for f in files or []:
        with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )
        # After the file is closed
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        msg.attach(part)

    print("Email Scraper: Sending email!");
    server = smtplib.SMTP(SERVER, PORT)
    server.starttls()
    server.login(EMAIL, PASSWORD)
    server.sendmail(EMAIL, RECIPIENTS, msg.as_string())
    server.quit()
    
def evalPropertyAddress(email = ""):
    property_address = email["Body"].split("Property Address")
    try:
        property_address[1]
    except:
        print("Email Scraper: No Property Address found.");
        return True;
    else:
        addresses = pyap.parse(property_address[1], country='US')
        try:
            addresses[0]
        except:
            print("Email Scraper: Unable to find Property Address.");
            return True; #Didnt find an address despite having property address field, return true as theres not much i can do.
        else:
            property_address = str(addresses[0]) #Finally have the property address
            #Add email address to check as well.
            property_address = property_address+" &EMAIL: " + email["From"]
            #Check to see if the property address exists in the list already.
            if(property_address in addresslist):
                print("Email Scraper: Property Address found. Not Scanning Email...");
                return False;
            else:
                print("Email Scraper: First instance of property address found. Scanning...");
                addressfile.write(property_address+"\n")
                addresslist.append(property_address)
                return True;
    return True;

main();
