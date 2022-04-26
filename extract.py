import os
import xml.etree.ElementTree as ET
import glob

folders = ["Inbox", "Sent Items"]
blacklist = ["asana.com", "reply", "noreply", "registration", "chime", "support@", "mailchimp", ".calendar", "=", "-_", "mailbox", "+","account","fatura","hizmetleri","musteri","marketing@", "help@", "sales"]

print("Starting to read file")
"""
1. Export your all e-mails using Export menu in Outlook.
2. Go to terminal and use `unzip Outlook\ for\ Mac\ Archive.olm` to unzip the archive file.
3. Put the `extract.py` file near `Local` directory.
4. Run `python extract.py` and obtain the emails.txt file
5. change the path below to match the folder with your extracted .olm file
"""
files = glob.glob('/Users/data/Desktop/emailextract/**/*.xml', 
                   recursive = True)

email_list = []

for file in files:
    #print(file)
    #print(len(files))

    if not file.endswith(".xml"):
        print(file + " is not xml")
        continue
    try:
        tree = ET.parse(file)
    except:
        print(file + " couldn't parsed")
    finally:
        root = tree.getroot()
        for item in root.iter('emailAddress'):
            if 'OPFContactEmailAddressAddress' in item.attrib:
                if any(to_check in item.attrib['OPFContactEmailAddressAddress'].lower() for to_check in blacklist):
                    #print(">>"+item.attrib['OPFContactEmailAddressAddress']+" is blacklisted")
                    continue
                if any(item.attrib['OPFContactEmailAddressAddress'].lower() in s for s in email_list):
                    #print(">>"+item.attrib['OPFContactEmailAddressAddress']+" exists")
                    continue
                if 'OPFContactEmailAddressName' in item.attrib:
                    to_append = item.attrib['OPFContactEmailAddressName']+" <"+item.attrib['OPFContactEmailAddressAddress'].lower()+">"
                else:
                    to_append = item.attrib['OPFContactEmailAddressAddress']+" <"+item.attrib['OPFContactEmailAddressAddress'].lower()+">"
                email_list.append(to_append)

# make every email unique
email_list = list(set(email_list))
#print(email_list)


textfile = open("emails.txt", "w")
for element in email_list:
    textfile.write(element + "\n")
textfile.close()

#https://github.com/eercanayar/outlook-olm-email-exporter/blob/master/extract.py
