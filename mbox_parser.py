#!/usr/bin/env python
# -*- coding: utf-8 -*-

import re
import os
from email.header import Header, decode_header, make_header
from datetime import datetime as dt


#--------------------------#
#       INITIALIZATION
#--------------------------#

mbox_files_path = '/PATH/TO/MBOX_FILE.MBOX'
path_to_write = '/PATH/TO/WRITE/EML_FILES'

if not os.path.isdir(path_to_write):
    os.mkdir(path_to_write)

regex = "From \d+@.+"
to_search = re.compile(regex)
counter = 0


#--------------------------#
#       SUBFUNCTIONS
#--------------------------#

# Get the date of the mail passed
# Return it formatted
def getDate(mail, format='%Y%m%d'):
    date_re = "\nDate: "
    date_to_search = re.compile(date_re, re.IGNORECASE)
    list_from_date = re.split(' +', re.split(date_to_search, mail)[1])
    try:
        if isinstance(int(list_from_date[0]), int):
            d = str(list_from_date[0]) + ' ' + str(list_from_date[1]) + ' ' + str(list_from_date[2])
    except:
        d = str(list_from_date[1]) + ' ' + str(list_from_date[2]) + ' ' + str(list_from_date[3])
    return dt.strptime(d, '%d %b %Y').strftime(format)


# Get the subject of the mail passed
# Return it formatted or 'no-subject'
def getSubject(mail):
    try:
        subject_re = "\nSubject:(?: |\n)"
        subject_to_search = re.compile(subject_re, re.IGNORECASE)
        list_from_subject = re.split(subject_to_search, mail)[1].split('\n')
        i = 1
        subject = list_from_subject[0]
        while list_from_subject[i] and list_from_subject[i][0] == ' ':
            subject += list_from_subject[i]
            i += 1
        subject = make_header(decode_header(subject))
    except:
        subject = ''
    return re.sub('(\W+)', '-', str(subject)).strip('-') or 'no-subject'


#--------------------------#
#       MAIN SCRIPT
#--------------------------#

# List objects in the given directory path
mbox_files = [f for f in os.listdir(mbox_files_path) if os.path.isfile(os.path.join(mbox_files_path, f))]

# Foreach object, ...
for mbox_file in mbox_files:
    mbox_path = os.path.join(mbox_files_path, mbox_file)
    # Test if it's an MBOX file
    if os.path.isfile(mbox_path) and os.path.splitext(mbox_file)[1] == '.mbox':
        # Create a directory for MBOX file if not exists
        if not os.path.isdir(path_to_write + mbox_file[:-5]):
            os.mkdir(path_to_write + mbox_file[:-5])
        # Read file
        with open(mbox_path, 'r', encoding="utf8") as f:
            new_file = f.read()
            # Split file into mails
            mails = re.split(regex, new_file)[1:]
            # Foreach mail, ...
            for mail in mails:
                mail = mail.strip('\n')
                name = path_to_write + mbox_file[:-5] + '/' + getDate(mail) + '_' + getSubject(mail) + '.eml'
                # Test if a mail already exist at this date with this subject
                if os.path.isfile(name):
                    name = name[:-4] + '-' + str(counter) + '.eml'
                print(name)
                # Write an EML file with
                with open(name, 'w') as message:
                    try:
                        message.write(mail)
                        counter += 1
                    except Exception as e:
                        print(e)
        print("Total msgs: " + str(counter))
    else:
        print("MBOX file not detected")