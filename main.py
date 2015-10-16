#!/bin/python
# -*- coding: utf-8 -*-
###############################################################################
#  Delay your email notification some hours - V1.0                            #
#  Copyright 2015 Benito Palacios (aka pleonex)                               #
#                                                                             #
#  Licensed under the Apache License, Version 2.0 (the "License");            #
#  you may not use this file except in compliance with the License.           #
#  You may obtain a copy of the License at                                    #
#                                                                             #
#      http://www.apache.org/licenses/LICENSE-2.0                             #
#                                                                             #
#  Unless required by applicable law or agreed to in writing, software        #
#  distributed under the License is distributed on an "AS IS" BASIS,          #
#  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.   #
#  See the License for the specific language governing permissions and        #
#  limitations under the License.                                             #
###############################################################################

import ProcImap.ImapServer as ImapServer
import ProcImap.ImapMailbox as ImapMailbox
import re
from getpass import getpass
from email.utils import parsedate
from time import mktime

CASE_REGEX = re.compile('Case Number ([0-9]{8})')


def login(server, user, pwd):
    """Create a connection to the server and log-in."""
    server = ImapServer.ImapServer(server, user, pwd)
    return server


def download_emails(mailbox):
    """Get the latest e-mails of interest.

    Use SEARCH command to get messages that does not contain the header
    "In-Reply-To", that is, they are not thread. Also discard first message,
    usually containing "from portal". Sort them by date so we start with the
    older one.
    """
    print("Searching unprocessed messages")
    emails = mailbox.search(
        'NOT HEADER In-Reply-To "" NOT SUBJECT "from portal"')
    sorted(emails,
           key=lambda e: mktime(parsedate(
               mailbox.get_fields(e, 'Date').get('Date'))))
    return emails


def polling(server):
    """Wait for new message and return the ID of the message to process."""
    unprocessed = download_emails(server)
    while unprocessed is None:
        print("Entering into IDLE")
        server.idle()
        unprocessed = download_emails(server)

    return unprocessed


def magic(mailbox, email_id, mailbox_name):
    """Do the magic."""
    print("Processing " + str(email_id) + " - " + email.get('Subject'))
    email = mailbox[email_id]

    # Get the case number from the subject
    case_number = re.search('Case Number ([0-9]{8})',
                            email.get('Subject'))
    if case_number is None:
        print("\tInvalid message")
        return
    else:
        case_number = case_number.group(1)

    # Search the root message
    parent = mailbox.search(
        'SUBJECT "Case Number %s" HEADER In-Reply-To ""' % case_number)
    sorted(parent, key=lambda e: mktime(parsedate(mailbox.get(e).get('Date'))))

    my_date = mktime(parsedate(email.get('Date')))
    parent_date = 0
    if len(parent) > 0:
        parent_date = mktime(parsedate(mailbox.get(parent[0]).get('Date')))
    if len(parent) == 0 or my_date < parent_date:
        parent = mailbox.search(
            'SUBJECT "New Case %s from portal"' % case_number)

        if len(parent) == 0:
            print("\tCannot find parent")
            return

    parent = parent[0]
    print("\tParent: " + mailbox.get(parent).get('Subject'))

    def replace_date(matches):
        minutes = int(matches.group(2)) + 1
        return matches.group(1) + ':' + str(minutes) + ':' + matches.group(3)

    date = email.get('Date')
    new_date = re.sub('(\d+):(\d+):(\d+)', replace_date, date)
    email.replace_header('Date', new_date)

    email.add_header('In-Reply-To',
                     mailbox.get_message(parent).get('Message-ID'))

    mailbox.move(email_id, mailbox_name + "/BackUp")  # Backup original message
    mailbox.add(email)


if __name__ == "__main__":
    server = raw_input('IMAP Server: ')
    email = raw_input('E-mail: ')
    pwd = getpass()
    server = login(server, email, pwd)

    # Go to mailbox
    mailbox_name = raw_input('Mailbox name: ')
    mailbox = ImapMailbox.ImapMailbox((server, mailbox_name), create=False)
    emails = polling(mailbox)
    print(emails)

    # print("Doing magic")
    # count = 0
    # for e in emails:
    #     print("\t%d/%d %f%%" % (count, len(emails), count * 100 / len(emails)))
    #     count = count + 1
    #     magic(mailbox, e, boxname)

    mailbox.close()
