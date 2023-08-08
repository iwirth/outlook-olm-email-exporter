import os
import xml.etree.ElementTree as ET
import glob
from xml.etree.ElementTree import ElementTree
from zipfile import ZipFile
import shutil
import re
import atexit


@atexit.register
def remove_tmp():
    shutil.rmtree('./tmp')


# RFC 5322 compliant regex
email_pattern = ("(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*|\"(?:["
                 "\\x01-\\x08\\x0b\\x0c\\x0e-\\x1f\\x21\\x23-\\x5b\\x5d-\\x7f]|\\\\["
                 "\\x01-\\x09\\x0b\\x0c\\x0e-\\x7f])*\")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\\.)+[a-z0-9](?:["
                 "a-z0-9-]*[a-z0-9])?|\\[(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.){3}(?:25[0-5]|2[0-4][0-9]|["
                 "01]?[0-9][0-9]?|[a-z0-9-]*[a-z0-9]:(?:[\\x01-\\x08\\x0b\\x0c\\x0e-\\x1f\\x21-\\x5a\\x53-\\x7f]|\\\\["
                 "\\x01-\\x09\\x0b\\x0c\\x0e-\\x7f])+)\\])")

# prompt for filepath and check if it exists
while True:
    olm_path = input("path to the olm file: ")
    if os.path.exists(olm_path):
        break

# check if tmp folder exists, if it does, delete it
if os.path.exists('./tmp'):
    while True:
        del_tmp = input('This will delete the current tmp folder. Continue? [y/n] ')
        if del_tmp == 'y':
            shutil.rmtree('./tmp')
            break
        else:
            print('Aborting.')
            exit()

print("Extracting olm File")
with ZipFile(olm_path, 'r') as zf:
    zf.extractall('./tmp')

# Select account
account_dir = 'tmp/Accounts'
accounts = [d for d in os.listdir(account_dir) if os.path.isdir(os.path.join(account_dir, d))]

index = 0
for account in accounts:
    print(f'{index}\t{account}')
    index += 1

selection = int(input('\nSelect an account: '))

account_name = accounts[selection]

xml_path = 'tmp/Accounts/' + account_name + '/**/*.xml'

files = glob.glob(xml_path, recursive=True)

email_list = []

for file in files:
    if not file.endswith(".xml"):
        print(file + " is not xml")
        continue
    try:
        tree: ElementTree = ET.parse(file)
    except:
        print(file + " couldn't parsed")
        continue
    finally:
        root = tree.getroot()
        for item in root.iter('emailAddress'):
            if 'OPFContactEmailAddressAddress' in item.attrib:
                # if any(to_check in item.attrib['OPFContactEmailAddressAddress'].lower() for to_check in blacklist):
                # print(">>"+item.attrib['OPFContactEmailAddressAddress']+" is blacklisted")
                # continue
                if any(item.attrib['OPFContactEmailAddressAddress'].lower() in s for s in email_list):
                    # print(">>"+item.attrib['OPFContactEmailAddressAddress']+" exists")
                    continue
                if 'OPFContactEmailAddressName' in item.attrib:
                    to_append = item.attrib['OPFContactEmailAddressName'] + " <" + item.attrib[
                        'OPFContactEmailAddressAddress'].lower() + ">"
                else:
                    to_append = item.attrib['OPFContactEmailAddressAddress'] + " <" + item.attrib[
                        'OPFContactEmailAddressAddress'].lower() + ">"
                email_list.append(to_append)

# make every email unique
email_list = list(set(email_list))

# output with sender name
textfile = open(f"raw_{account_name}_emails.txt", "w")
for element in email_list:
    textfile.write(element + "\n")
textfile.close()

# output mail addresses only
textfile = open(f"{account_name}_emails.txt", "w")
for element in email_list:
    email_addr = re.search(email_pattern, element)
    textfile.write(email_addr.group(0) + "\n")
textfile.close()

print('finished')
