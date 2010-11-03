import datetime
from email.MIMEText import MIMEText
import gdata
import gdata.spreadsheet.service
import os.path
import re
import smtplib
from string import Template
import sys

_DEBUG = False

def GetOptions (OpParser=None):
    # Parse command-line options.  Without the '-e' argument, this script
    # emits mail messages to stdout rather than actually sending them.  The
    # '-l' argument causes this script to authenticate against Google Docs
    # in order to gain access to private spreadsheets.
    op = OpParser and OpParser or OptionParser()
    op.add_option('-l', '--login', action='store_true', dest='login',
            help='Attempt to authenticate with Google (default False)')
    op.add_option('-d', '--debug', action='store_true', dest='debug',
            help='Show debugging messages')
    (options, args) = op.parse_args()
    if not options.send_by_email:
        sys.stderr.write('WARNING: not actually sending email (use -e)\n')
    if not options.login:
        sys.stderr.write('WARNING: not trying to authenticate to Google (use -l)\n')
    return (options, args)

def StringToDate (datestr, rounddown=False):
    """Parse a date string like M/D/Y and return an appropriate datetime
       object."""
    try:
        m, d, y = map(int, datestr.split('/'))
        return datetime.date(y, m, d)
    except:
        if re.match(r'^\d{4}', datestr):
            if rounddown:
                return datetime.date(int(datestr), 1, 1)
            return datetime.date(int(datestr), 12, 31)
        else:
            m = re.match(r'^(\d\d?)/(\d{4})$', datestr)
            if m is None:
                return None
            return datetime.date(int(m.groups[1]), int(m.groups[0]), 28) # XXX

class nagger:
    gd_client = None

    def __init__ (self,
            EmailFromAddress,
            GoogleSpreadsheetKey,
            EmailColumnName='email',
            GoogleWorksheetName=[],
            AuthUsername=None, AuthPassword=None, DoLogin=False,
            SMTPServer='localhost',
            DEBUG=False):
        global _DEBUG
        _DEBUG = DEBUG
        self.EmailFromAddress = EmailFromAddress

        self.GoogleSSID = GoogleSpreadsheetKey
        if not isinstance(GoogleWorksheetName, list):
            GoogleWorksheetName = [GoogleWorksheetName]
        self.GoogleWorksheetName = GoogleWorksheetName
        self.EmailColumnName = EmailColumnName

        self.SMTPServer = SMTPServer

        self.gd_client = gdata.spreadsheet.service.SpreadsheetsService()
        if DoLogin:
            self.login(AuthUsername, AuthPassword)
            if _DEBUG: print "Successfully authenticated as %s" % AuthUsername

    def login (self, AuthUsername, AuthPassword):
        """Log into Google Docs using AuthUsername and AuthPassword as Google
           credentials."""
        self.gd_client.email = AuthUsername
        self.gd_client.password = AuthPassword
        self.gd_client.source = self.__script_name()
        self.gd_client.ProgrammaticLogin()

    def __script_name (self, show_full_path=False):
        """Return this script's filename, or full path including filename if
           show_full_path is True."""
        if show_full_path:
            return sys.argv[0]
        myfname = os.path.split(sys.argv[0])[-1]
        return os.path.splitext(myfname)[0]

    def _entries (self):
        """Returns a "list-based feed" in which each row represents an entry (so we
           can fetch whole rows at a time, rather than just cells).
           N.b.: according to
             http://code.google.com/apis/spreadsheets/data/3.0/ \
                developers_guide_protocol.html#ListFeeds,
           "The first blank row terminates the data set."  So make sure there are no
            blank rows in the data!"""
        myentries = []
        if self.GoogleWorksheetName:
            for ws in self.GoogleWorksheetName:
                # fetch the worksheet named by GoogleWorksheetName
                query = gdata.spreadsheet.service.DocumentQuery()
                query.title = ws
                worksheet_feed = self.gd_client.GetWorksheetsFeed(self.GoogleSSID,
                        query=query)
                if worksheet_feed.entry:
                    wsid = worksheet_feed.entry[0].id.text.rsplit('/', 1)[-1]
                    if _DEBUG: print 'Worksheet title is %s' % ws 
                    lf = self.gd_client.GetListFeed(self.GoogleSSID, wsid)
                    if lf:
                        myentries.extend(lf.entry)
            return myentries
        else:
            return gd_client.GetListFeed(self.GoogleSSID).entry # fetch default sheet

    def GetSpreadsheetTitle (self):
        global _DEBUG
        ssfeed = self.gd_client.GetSpreadsheetsFeed()
        for e in ssfeed.entry:
            for l in e.link:
                keyrx = r'\bkey=%s\b' % self.GoogleSSID
                if re.search(keyrx, l.href):
                    if _DEBUG: print 'Spreadsheet title is "%s"' % e.title.text
                    return e.title.text

    def GetPeopleItems (self):
        """Returns a map of [email address] -> [spreadsheet rows]"""
        people = {}

        # Iterate over the spreadsheet's rows.  Note that the first line is
        # interpreted as column headers and is therefore not included in the range
        # covered by this loop (i.e., we start at row 2).
        #for entry in self.GetSpreadsheetListFeed().entry:
        for entry in self._entries():
            # Google barely documents entry.custom, but its structure should be
            # obvious from what follows...
            c = entry.custom

            # grab the person's email address; skip rows.  fields like
            # 'blah@blah.org' and 'First Last <firstlast@somewhere.com>' (and
            # variations; basically any token that contains '@' and is optionally
            # surrounded by '<>' will be matched)
            if self.EmailColumnName not in c: continue
            email = c[self.EmailColumnName].text
            if email is None or '@' not in email: continue
            email = filter(lambda x: '@' in x, email.rsplit())[0].strip('<>')

            row = {}
            for column in entry.custom:
                row[column] = entry.custom[column].text

            people.setdefault(email, []).append(row)

        return people

    def GetMessages (self, MessageTemplate, ItemStringGenerator,
            SortItems=False, ContactPerson=None):
        msgs = []

        globaldict = dict(sstitle=self.GetSpreadsheetTitle(),
                sskey=self.GoogleSSID,
                sender=self.EmailFromAddress,
                contact=ContactPerson or self.EmailFromAddress)
        people_to_items_map = self.GetPeopleItems()

        for person in people_to_items_map:
            pd = dict(globaldict)
            pd['email'] = person
            pd['items'] = ''
            itemstmp = map(lambda x: ItemStringGenerator(self, x),
                    people_to_items_map[person])
            if SortItems: itemstmp = sorted(itemstmp)
            pd['items'] = ''.join(itemstmp).rstrip('\n')

            msgbody = Template(MessageTemplate).safe_substitute(pd)
            m = MIMEText(msgbody)
            m['To'] = person
            m['From'] = globaldict['sender']
            m['Subject'] = globaldict['sstitle']
            if ContactPerson:
                m['Reply-To'] = ContactPerson
            msgs.append(m)

        return msgs

    def SendMessages (self, Messages):
        smtp = smtplib.SMTP(self.SMTPServer)
        for m in Messages:
            smtp.sendmail(m['From'], m['To'], m.as_string())
