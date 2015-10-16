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


def get_email_date(email):
    """Get the date of an email in float notation."""
    return mktime(parsedate(email.get('Date')))


def sort_email_by_date(mailbox, emails):
    """Sort a list of UID emails by their dates."""
    # Get only an email object with the Date header.
    sorted(emails, key=lambda e: get_email_date(mailbox.get_fields(e, 'Date')))
    emails.reverse()


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
    sort_email_by_date(mailbox, emails)
    return emails


def polling(server):
    """Wait for new message and return the ID of the message to process."""
    unprocessed = download_emails(server)
    while unprocessed is None:
        print("Entering into IDLE")
        server.idle(timeout=29*60)
        unprocessed = download_emails(server)

    return unprocessed


def search_parent_email(mailbox, email):
    """Search and return the ID of the parent email."""
    # Try to get the case number from the subject.
    case_number = re.search('Case Number ([0-9]{8})', email.get('Subject'))
    if case_number is None:
        print("\tInvalid message")
        return None
    else:
        case_number = case_number.group(1)

    # Search all the emails with the same case number already in threads.
    parent = mailbox.search(
        'SUBJECT "Case Number %s" HEADER In-Reply-To ""' % case_number)
    sort_email_by_date(mailbox, parent)

    # Get the email date to compare with the current parent.
    email_date = get_email_date(email)
    parent_date = 0
    if len(parent) > 0:
        # Get only an email object with the Date header.
        parent_date = get_email_date(mailbox.get_fields(parent[0], 'Date'))

    # If we haven't found any parent or our parent is before our message,
    # search for the first, New Case, email.
    if len(parent) == 0 or parent_date > email_date:
        parent = mailbox.search(
            'SUBJECT "New Case %s from portal"' % case_number)

        if len(parent) != 1:
            print("\tCannot find parent: " + str(len(parent)))
            return None

    # Get the first parent
    parent = parent[0]
    return parent


def thread_email(mailbox, email_id, mailbox_name):
    """Thread an email."""
    email = mailbox[email_id]
    if email.get('Subject') is None:
        print('Invalid subject')
        return False

    print("Processing " + str(email_id) + ": " + email.get('Subject'))

    # Search the parent email, not found do nothing, maybe it's the first msg.
    parent = search_parent_email(mailbox, email)
    if parent is None:
        return False

    # Add the In-Reply-To header, to create the threading behavior.
    parent_msgid = mailbox.get_fields(parent, 'Message-ID').get('Message-ID')
    email.add_header('In-Reply-To', parent_msgid)

    # In Gmail and other IMAP servers, we need to change the date if the
    # message are so similars. Just increment the minutes by one.
    def replace_date(matches):
        minutes = int(matches.group(2)) + 1
        return matches.group(1) + ':' + str(minutes) + ':' + matches.group(3)

    new_date = re.sub('(\d+):(\d+):(\d+)', replace_date, email.get('Date'))
    email.replace_header('Date', new_date)

    # Finally, move the original email to the BackUp folder and add our own.
    mailbox.move(email_id, mailbox_name + "/BackUp")
    mailbox.add(email)

    return True


if __name__ == "__main__":
    server = raw_input('IMAP Server: ')
    email = raw_input('E-mail: ')
    pwd = getpass()
    server = login(server, email, pwd)

    # Go to mailbox
    mailbox_name = raw_input('Mailbox name: ')
    mailbox = ImapMailbox.ImapMailbox((server, mailbox_name), create=False)

    # Enter into the infinite loop to check email and process it
    exit = False
    while not exit:
        emails = polling(mailbox)

        count = 0
        for e in emails:
            count += 1
            print("\t%d of %d" % (count, len(emails)))
            if thread_email(mailbox, e, mailbox_name):
                # exit = True
                # break
                pass

    mailbox.close()
