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
import xlsxwriter


def GenerateTransactionsReport(events, ap, accounts=['Cash', 'HSBC Current Account']):
    """
    Generates a chronological list of events covering a period (e.g. month) for a particular set
    of accounts and writes the output to an Excel spreadsheet
    :param events: List of events
    :param excelfile: Name of Excel output file
    :param accounts: List of accounts to report on (each will have its own Sheet within the Excel file)
    :return:
    """

    workbook = xlsxwriter.Workbook(ap.excel)
    h1 = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 14})
    h2 = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 12, 'italic': True})
    default = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 11})
    date = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 11, 'num_format': 'dd/mm/yy'})
    currency = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 11, 'num_format': '#,##0.00'})
    h2right = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 12, 'italic': True, 'align': 'right'})


    for account in accounts:
        worksheet = workbook.add_worksheet(account)
        worksheet.set_column('A:A', 9)
        worksheet.set_column('B:C', 24)
        worksheet.set_column('D:F', 9)
        worksheet.set_header('&L&"Times New Roman, 16, bold"{} Transactions&R&"Times New Roman, 16, bold"{:%d %b %y} to {:%d %b %y}'.format(account, ap.start, ap.end))
        worksheet.set_footer('&L&"Times New Roman, 14"Manchester Spiritualist Centre&C&"Times New Roman, 14, bold"CONFIDENTIAL&R&"Times New Roman, 14"Page &P of &N')
        acctevents = events[events['account']==account]
        row, col = 0, 0
        columnnames = ['Date', 'Description', 'Category', 'Debit', 'Credit', 'Balance']
        for col, name in enumerate(columnnames):
            worksheet.write(row, col, name, h2 if col == 1 or col == 2 else h2right)
        col = 0
        for record in acctevents.itertuples():
            row += 1
            worksheet.write(row, 0, record.date, date)
            worksheet.write(row, 1, record.description, default)
            worksheet.write(row, 2, record.matching, default)
            worksheet.write(row, 3 if record.value < 0 else 4, abs(record.value), currency)
            worksheet.write(row, 5, record.balance, currency)

    workbook.close()

def GetArgs():
    """
    Gets the command line arguments
    :return: argparse namespace
    """
    GetDate = lambda d: datetime.datetime.strptime(d, '%Y-%m-%d').date()
    reportchoices = ['transactions']
    p = argparse.ArgumentParser(description='Read GnuCash files')
    p.add_argument('--logging', metavar='level', help='Logging level', choices=logchoices, default='info')
    p.add_argument('--report', metavar='type', help='Type of report to generate', choices=reportchoices)
    p.add_argument('--start', metavar='date', help='Start of reporting period', type=GetDate)
    p.add_argument('--end', metavar='date', help='End of reporting period', type=GetDate)
    p.add_argument('--xml', help='Write XML file', action='store_true')
    p.add_argument('--excel', help='Write Excel (.xlsx) file', metavar = 'filename')
    p.add_argument('name', metavar='GnuCash file', help='Name of GnuCash file')

    ap = p.parse_args()
    if ap.start and ap.end and ap.end <= ap.start:
        p.error('Start date must be before end')

    now = datetime.date.today()
    if ap.start and ap.start > now:
        p.error('Start date {:%Y-%m-%d} is in the future'.format(ap.start))

    if ap.end and ap.end > now:
        p.error('End date {:%Y-%m-%d} is in the future'.format(ap.end))

    if ap.report and not ap.excel:
        p.error('Reports must have an Excel filename')
    return ap

if __name__ == '__main__':

    logchoices = {'debug': logging.DEBUG, 'info': logging.INFO, 'warning': logging.WARNING,
               'error': logging.ERROR, 'critical': logging.CRITICAL}
    ap = GetArgs()
    logging.basicConfig(level=logchoices[ap.logging], format='%(asctime)s %(levelname)-8s %(message)s')

    try:
        accounts = GnuCash.Book(**vars(ap))
    except (FileNotFoundError, OSError) as e:
        logging.critical('Error opening file: {}'.format(e))
    else:
        eventlist = accounts.root.GetAllEvents()
        events = pd.DataFrame([{'date': event.date,
                   'type': accounts.accounts[event.accountid].type,
                   'account': accounts.accounts[event.accountid].name,
                   'matching': accounts.accounts[event.matching].name,
                   'description':event.description,
                   'balance': event.balance,
                   'value': event.value} for event in eventlist])
        if not ap.start:
            ap.start = events.date.min()
        if not ap.end:
            ap.end = events.date.max()
        events = events[(events['date'] >= ap.start) & (events['date'] <= ap.end)]

        if ap.report and ap.report == 'transactions':
            GenerateTransactionsReport(events, ap)
