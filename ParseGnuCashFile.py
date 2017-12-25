"""
Open and parse a GnuCash file
"""

#import GnuCash
import gzip
import argparse
import logging
import xml.etree.ElementTree as ET
import collections
import re
import datetime

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
        self.opening = 0.0
        self.children = []

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

    def __repr__(self):
        attributes = ['name', 'id', 'description', 'type', 'level', 'parent']
        atrlist = ["{}: {}".format(a, getattr(self, a, None)) for a in attributes]
        return "; ".join(atrlist)

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
                    self.dateposted = self.ParseDateString(child.find('{http://www.gnucash.org/XML/ts}date').text)
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

    def __init__(self, book):

        self.accounts = {}
        self.transactions = []

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

        # Create the parent/child relationship between the various accounts. All accounts are
        # children, grandchildren etc. of the root account. (The root account is the only one with
        # no parent.)

        for account in self.accounts:
            if self.accounts[account].parentid:
                child = self.accounts[account]
                parentid = self.accounts[account].parentid
                parent = self.accounts[parentid]
                parent.children.append(child)
            else:
                self.root = self.accounts[account]

        # Having created the tree, now update each account with its level in the hierarchy

        if self.root:
            self.SetLevels(self.root)
        else:
            logging.error('No root account found!')

        # Now find the opening balances

        self.SetOpeningBalances()

    def __repr__(self):
        return "{} accounts. {} transactions.".format(len(self.accounts), len(self.transactions))

    def SetOpeningBalances(self):
        """
        Sets up the opening balance on those accounts where it is appropriate (i.e. opening
        balance was not 0.0
        :return:
        """
        for account in self.accounts:
            if self.accounts[account].type == 'EQUITY' and self.accounts[account].name == 'Opening Balances':
                balances = self.accounts[self.accounts[account].id]
                logging.info('Opening balances account id {}'.format(balances.id))
                for transaction in self.transactions:
                    for split in transaction.splits:
                        logging.debug('Found split {}'.format(split))
                break
        else:
            logging.warning('Can\'t find opening balances!')

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


if __name__ == '__main__':

    p = argparse.ArgumentParser(description='Read GnuCash files')
    p.add_argument('--debug', help='Debug mode', action='store_true')
    p.add_argument('--xml', help='Write XML file', action='store_true')
    p.add_argument('name', metavar='file', help='Name of GnuCash file')
    ap = p.parse_args()
    logging.basicConfig(level=logging.DEBUG if ap.debug else logging.INFO, format='%(asctime)s %(levelname)-10s %(message)s')

    try:
        with gzip.open(ap.name, 'rb') as zf:
            contents = zf.read().decode('utf8')
            if ap.xml:
                try:
                    xmloutfile = ap.name+'.xml'
                    with open(xmloutfile, 'w') as xf:
                        xf.write(contents)
                        logging.debug("Written {:,} bytes to {}".format(len(contents), xmloutfile))
                except PermissionError as e:
                    logging.error("Error writing output XML: {}".format(e))
    except (FileNotFoundError, OSError) as e:
        logging.error('Error opening {}: {}'.format(ap.name, e))
    else:
        root = ET.fromstring(contents)
        if root.tag != 'gnc-v2':
            logging.warning('Got unexpected XML root tag {}. Is {} a GnuCash file?'.format(root.tag, ap.name))
        else:
            book = root.find('{http://www.gnucash.org/XML/gnc}book')
            if book:
                accounts = Book(book)
            else:
                logging.error("Can't find GnuCash book")