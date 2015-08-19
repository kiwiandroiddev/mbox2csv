#!/usr/bin/env python

# mbox2csv
# Converts an email archive from mbox to csv format with columns: subject, from, date, message
# Dates are in ISO 8601 format (e.g. 2015-08-07T18:30:27Z)
#
# Usage:
# mbox2csv MBOX_FILE [CSV_FILE]

import sys, mailbox, csv
import dateutil.parser as parser

def print_progress(pct_progress):
    sys.stdout.write('\r[{0}] {1}%'.format('#'*(pct_progress/10), pct_progress))
    sys.stdout.flush()

mbox_file = sys.argv[1]
output_file = sys.argv[2] if len(sys.argv) > 2 else 'output.csv'

print 'Reading mbox file...'
messages = mailbox.mbox(mbox_file)
writer = csv.writer(open(output_file, "wb"))

writer.writerow(['subject', 'from', 'date', 'message'])
n = len(messages)
print 'Writing messages...'
for (i, msg) in enumerate(messages):
    body = msg.get_payload()[0].get_payload()

    # convert any date format to ISO!
    date = parser.parse(msg['date'])
    iso_date = date.isoformat()

    writer.writerow([msg['subject'], msg['from'], iso_date, body])

    # update progress bar on every 10th message for speed
    if i % 10 == 0:
        pct_complete = int(round(i/float(n) * 100.0))
        print_progress(pct_complete)
    
print
print 'Wrote %s' % output_file
