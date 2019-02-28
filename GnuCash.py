"""
Classes to hold details of GnuCash files

Tim Greening-Jackson 02 December 2017
"""

import logging
import datetime
import re
import gzip
import xml.etree.ElementTree as ET

class Event:
    """
    An event is a combination of a transaction (date, commentary) and a split (account, amount).

    Each account has a chronological list of Events starting with its opening balance
    """

    def __init__(self, split, balance=None):
        self.split = split
        self.transaction = split.parent
        for _, a in self.transaction.splits.items():
            if a.account != self.split.account:
                self.matching = a.account
                break
        else:
            logging.warning("Can't find matching transaction for {}".format(self.split))
            self.matching = None
        self.date = self.split.parent.dateposted
        self.description = self.split.parent.description
        self.value = self.split.value
        self.accountid = self.split.account
        self.balance = 0.0

    def SetBalance(self, amount):
        """
        Sets the account balance following a particular event
        :param amount:
        :return:
        """
        try:
            self.balance = amount
        except ValueError as e:
            logging.error('Error setting event {} value: {}'.self(description, e))
            self.balance = None

    def __repr__(self):
        if self.value < 0.0:
            return '{:%d/%m/%Y}: {:60} {:>12,.2f}            '.format(self.date, self.description, self.value)
        else:
            return '{:%d/%m/%Y}: {:60}             {:>12,.2f}'.format(self.date, self.description, self.value)

    def __lt__(self, other):
        return self.date < other.date

class Account:
    """

    """
    re = re.compile('^\{http://www.gnucash.org/XML/act\}(.*)$')

    def __init__(self, account):
        """
        Note that parent and parentid are two separate things. parentid is the
        internal Gnucash GUID. parent is a reference the parent object

        :param Gnu cash XML account object:
        """

        self.xml = account
        self.commodity = None
        self.commodityscu = None
        self.parent = None
        self.parentid = None
        self.slots = None
        self.description = None
        self.level = None
        self.children = []
        self.events = []

        for child in self.xml:
            m = Account.re.match(child.tag)
            if m:
                n = m.group(1)
                if n == 'parent':
                    setattr(self, 'parentid', child.text)
                else:
                    setattr(self, n, child.text)
            else:
                logging.error('Account: tag {} doesn\'t match'.format(child.tag))

    def SetLevel(self, level):
        """
        Sets the level of the account within the tree
        :param level: integer
        :return: None
        """
        self.level = level

    def GetLevel(self):
        """
        Return the level within the tree
        :return: integer
        """
        return self.level

    def AddEvent(self, split):
        """
        Adds an event from the splits of a transaction
        :param split:
        :return: None
        """
        self.events.append(Event(split))

    def SetParent(self, parent):
        """
        Sets the parent
        :return:
        """
        self.parent = parent

    def GetAllEvents(self):
        """
        Gets a list of the events on the account (essentially the self.events attribute)
        recursing through child accounts.
        :return: list of events
        """
        events = []
        for child in self.children:
            events += child.GetAllEvents()
        events += self.events
        return sorted(events)


    def __repr__(self):
        attributes = ['name', 'id', 'description', 'type', 'level']
        atrlist = ["{}: {}".format(a, getattr(self, a, None)) for a in attributes]
        return "; ".join(atrlist)

    def __lt__(self, other):
        if self.type == other.type:
            return self.name < other.name
        else:
            return self.type < other.type

class Split:
    """
    A transaction contains a series of splits, allocating its value between various accounts. The parent
    value is the transaction with which a split is associated so we can look up things like description, date etc.
    """

    def __init__(self, split, parent):
        self.xml = split
        self.parent = parent
        self.id = self.xml.find('{http://www.gnucash.org/XML/split}id').text
        self.account = self.xml.find('{http://www.gnucash.org/XML/split}account').text
        self.amttext = self.xml.find('{http://www.gnucash.org/XML/split}value').text
        try:
            self.value = eval(self.amttext)
        except TypeError as e:
            logging.error('Error converting {} in transaction {}: {}'.format(self.amttext, self.id, e))

    def __repr__(self):
        return "Split ID {}. Account ID {}. Amount {}".format(self.id, self.account, self.value)

class Transaction:
    """

    """
    transre = re.compile('^\{http://www.gnucash.org/XML/trn\}(.*)$')
    datere = re.compile('^([\d]{4}-[\d]{2}-[\d]{2}).*')

    def __init__(self, transaction):
        self.xml = transaction
        self.splits = {}
        for child in self.xml:
            m = Transaction.transre.match(child.tag)
            if m:
                n = m.group(1)
                if n == 'id':
                    self.id = child.text
                elif n == 'description':
                    self.description = child.text
                elif n == 'date-posted':
                    self.dateposted = self.ParseDateString(child.find('{http://www.gnucash.org/XML/ts}date').text)
                elif n == 'date-entered':
                    self.dateentered = self.ParseDateString(child.find('{http://www.gnucash.org/XML/ts}date').text)
                elif n == 'splits':
                    for split in child:
                        self.AddSplit(split)
                elif n == 'slots':
                    pass
                elif n == 'currency':
                    pass
                elif n == 'num':
                    self.num = child.text
                else:
                    logging.info('Got unexpected tat tag: {} text: {}'.format(n, child.text))
            else:
                logging.info('Transaction: tag {} doesn\'t match'.format(child.tag))

    def AddSplit(self, split):
        """
        Adds a split to ther transaction. Because we need the splits to be in a dictionary,
        the key to which refers to part of the split itself, we create the split first and
        then add it to the dictionary...
        :param split:
        :return: None
        """
        tmp = Split(split, self)
        self.splits[tmp.id] = tmp

    def ParseDateString(self, d):
        """

        :param d: Date string read in from XML file
        :return: datetime.date object
        """
        m = Transaction.datere.match(d)
        if m:
            return datetime.datetime.strptime(m.groups()[0], '%Y-%m-%d').date()

    def __repr__(self):
        attributes = ['id', 'description', 'dateposted', 'num']
        atrlist = ["{}: {}".format(a, getattr(self, a, None)) for a in attributes]
        return "; ".join(atrlist)

    def __lt__(self, other):
        return self.dateposted < other.dateposted

class Book:
    """
    Books appear to be the topmost element of the GnuCash hierarchy. They mainly consist
    of Accounts and Transactions
    """

    def __init__(self, name, **kwargs):

        self.accounts = {}
        self.transactions = []
        self.name = name
        self.writexml = kwargs.get('xml', False)
        self.contents = self.OpenGnuCashFile()
        root = ET.fromstring(self.contents)
        if root.tag != 'gnc-v2':
            logging.error('Got unexpected XML root tag {}. Is {} a GnuCash file?'.format(root.tag, ap.name))
        else:
            book = root.find('{http://www.gnucash.org/XML/gnc}book')
            if book:
                for child in book.getchildren():
                    if child.tag == '{http://www.gnucash.org/XML/gnc}account':
                        acc = Account(child)
                        self.accounts[acc.id] = acc
                    elif child.tag == '{http://www.gnucash.org/XML/gnc}transaction':
                        self.transactions.append(Transaction(child))
                    elif child.tag == '{http://www.gnucash.org/XML/book}id':
                        pass
                    else:
                        logging.warning('Book: Unknown tag {}'.format(child.tag))

                self.SetParents()
                self.categories = {a.type: a for a in self.root.children}

                if self.root:
                    self.SetLevels(self.root)
                else:
                    logging.error('No root account found!')

                self.SetEvents()
            else:
                logging.error("Can't find GnuCash book")

    def __repr__(self):
        return "{} accounts. {} transactions.".format(len(self.accounts), len(self.transactions))

    def SetLevels(self, account, level=0):
        """
        Recursively set the level of the various accounts in the hierarchy
        where root is level 0.
        :param level: integer
        :return: None
        """
        logging.debug('Setting level for account {} to {} ({} children)'.format(account.name, level, len(account.children)))
        account.SetLevel(level)
        for child in account.children:
            self.SetLevels(child, level+1)

    def OpenGnuCashFile(self):
        """
        Opens the file, unzips its contents and reads the XML in to memory
        :return: XML
        """
        with gzip.open(self.name, 'rb') as zf:
            contents = zf.read().decode('utf8')
            if self.writexml:
                try:
                    xmloutfile = self.name + '.xml'
                    with open(xmloutfile, 'w') as xf:
                        xf.write(contents)
                        logging.debug("Written {:,} bytes to {}".format(len(contents), xmloutfile))
                except PermissionError as e:
                    logging.error("Error writing output XML: {}".format(e))
            return contents

    def SetParents(self):
        """
        Create the parent/child relationship between the various accounts. All accounts are
        children, grandchildren etc. of the root account. (The root account is the only one with
        no parent.)
        """
        for key, account in self.accounts.items():
            if account.parentid:
                child = self.accounts[key]
                parentid = account.parentid
                account.SetParent(self.accounts[parentid])
                account.parent.children.append(child)
            else:
                account.SetParent(None)
                self.root = account

    def SetEvents(self):
        """
        Sets up a list of events and running balances on each account
        :return: None
        """
        for transaction in sorted(self.transactions):
            for _, split in transaction.splits.items():
                self.accounts[split.account].AddEvent(split)
        for key, category in self.categories.items():
            logging.debug('Calculating running balances for category {}'.format(key))
            for child in category.children:
                balance = 0.0
                for event in child.events:
                    balance += event.value
                    event.SetBalance(balance)