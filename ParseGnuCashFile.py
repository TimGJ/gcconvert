"""
Open and parse a GnuCash file
"""

import GnuCash
import gzip
import argparse
import logging
import xml.etree.ElementTree as ET
import collections
import re
import datetime
import pandas as pd

if __name__ == '__main__':

    choices = {'debug': logging.DEBUG, 'info': logging.INFO, 'warning': logging.WARNING,
               'error': logging.ERROR, 'critical': logging.CRITICAL}
    p = argparse.ArgumentParser(description='Read GnuCash files')
    p.add_argument('--logging', metavar='level', help='Logging level', choices=choices, default='info')
    p.add_argument('--xml', help='Write XML file', action='store_true')
    p.add_argument('name', metavar='file', help='Name of GnuCash file')
    ap = p.parse_args()
    logging.basicConfig(level=choices[ap.logging], format='%(asctime)s %(levelname)-10s %(message)s')

    a = GnuCash.Book(**vars(ap))

    events = pd.DataFrame([{'date': event.date,
               'type': a.accounts[event.accountid].type,
               'account': a.accounts[event.accountid].name,
               'description':event.description,
               'value': event.value} for event in a.root.GetAllEvents()])

    events = events.set_index('date')
    typegroup = events.groupby([pd.TimeGrouper(freq='M'), 'type']).sum()
    print(typegroup.unstack('type'))
    accgroup =  events.groupby([pd.TimeGrouper(freq='M'), 'account']).sum()
    print(accgroup.unstack('date'))
