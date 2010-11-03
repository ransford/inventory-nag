#!/usr/bin/env python
import datetime
import nagger
from optparse import OptionParser
import sys

def itemgen (N, inputdict):
    """Takes a spreadsheet row as input and outputs a string representation of
       it.  Custom to each consumer of the nagger module."""
    return inputdict['itemdescription'] + '\n'

if __name__ == '__main__':
    op = OptionParser()
    op.add_option('-e', '--email', action='store_true', dest='send_by_email',
            help='Actually send email to recipients (default False)')
    op.add_option('-r', '--recipient', dest='recipient',
            help='Generate messages only to this recipient (-e to send)')
    (options, args) = nagger.GetOptions(op)

    N = nagger.nagger(EmailFromAddress='me@somewhere.com',
            GoogleSpreadsheetKey='',           # alphanumeric unique ID
#           GoogleWorksheetName='Sheet1',      # optional
            AuthUsername='USERNAME@gmail.com', # Google username
            AuthPassword='PASSWORD',           # corresponding password
            DoLogin=options.login,             # should be True for nonpublic
            DEBUG=options.debug)

    msg_template = """Dear $email,

According to the "$sstitle" spreadsheet (see below for URL), you have in your possession the following:

$items

Please contact $sender if you believe this is wrong.

Spreadsheet link (access restricted):
    https://spreadsheet.google.com/ccc?key=$sskey
"""

    msgs = N.GetMessages(MessageTemplate=msg_template, ItemStringGenerator=itemgen)
    if options.recipient:
        msgs = filter(lambda m: m['To']==options.recipient, msgs)

    # send email or print to stdout
    if options.send_by_email:
        N.SendMessages(msgs)
    else:
        for m in msgs:
            print m.as_string()
