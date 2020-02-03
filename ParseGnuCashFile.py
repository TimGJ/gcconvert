"""
Open and parse a GnuCash file
"""

import GnuCash
import argparse
import logging
import datetime
import pandas as pd
import xlsxwriter
import collections
import csv


def GenerateSummaryReport(book, ap, categories=["INCOME", "EXPENSE"], minlevel=2, maxlevel=2):
    """
    Generates the monthly summaries for a particular period (e.g. a year) of income and
    expenditure
    :param accounts: The Account array as read by GnuCash
    :param categories: INCOME and EXPENSE
    :param minlevel: How far up the tree to go - i.e. where do we start reporting account
    :param maxlevel: How far down the tree to go - i.e. where do we rollup subaccounts to
    :return: None
    """
    # Get the level 1 Accounts corresponding to `categories`

    monthnames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    colnames = ["Category"]+monthnames
    logging.debug("Writing annual accounts to {}".format(ap.csv))
    with open(ap.csv, "w", newline="") as csvfile:
        writer = csv.DictWriter(csvfile, colnames)
        print("Manchester Spiritualist Centre Accounts {}".format(ap.year), file=csvfile)
        writer.writeheader()
        accountkeys = {k: v.id for k,v in book.categories.items() if k in categories}
        for category, acid in accountkeys.items():
            print("{category}".format(category=category), file=csvfile)
            for child in book.accounts[acid].children:
                events = [t for t in child.GetAllEvents() if t.date.year == ap.year]
                months = collections.defaultdict(float)
                for event in events:
                    months[monthnames[event.date.month-1]] += abs(event.value)
                if len(months) == 0:
                    continue
                months["Category"] = child.name
                writer.writerow(months)


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
    reportchoices = ['transactions', 'annual']
    p = argparse.ArgumentParser(description='Read GnuCash files')
    p.add_argument('--logging', metavar='level', help='Logging level', choices=logchoices, default='info')
    p.add_argument('--report', metavar='type', help='Type of report to generate', choices=reportchoices)
    p.add_argument('--start', metavar='date', help='Start of reporting period', type=GetDate)
    p.add_argument('--end', metavar='date', help='End of reporting period', type=GetDate)
    p.add_argument('--year', metavar='year', help='Annual report year', type=int, default=datetime.date.today().year-1)
    p.add_argument('--excel', help='Write Excel (.xlsx) file', metavar = 'filename')
    p.add_argument('--csv', help='Write CSV file', metavar = 'filename')
    p.add_argument('name', metavar='GnuCash file', help='Name of GnuCash file')

    ap = p.parse_args()
    if ap.start and ap.end and ap.end <= ap.start:
        p.error('Start date must be before end')

    now = datetime.date.today()

    if ap.csv and ap.report != "annual":
        p.error("CSV option only works for annual reports")

    if (ap.start or ap.end) and ap.report == 'annual':
        p.error('Annual report requires a year parameter not start or end')

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
        for event in eventlist:
            if not event.value:
                logging.error("Found suspicious event (zero or missing value) {} on {}".format(event.description, event.date))
        try:
            events = pd.DataFrame([{'date': event.date,
                       'type': accounts.accounts[event.accountid].type,
                       'account': accounts.accounts[event.accountid].name,
                       'matching': accounts.accounts[event.matching].name,
                       'description':event.description,
                       'balance': event.balance,
                       'value': event.value} for event in eventlist if event.value])
        except KeyError as e:
            logging.critical('Error loading transaction: {}'.format(e))
        else:
            if not ap.start:
                ap.start = events.date.min()
            if not ap.end:
                ap.end = events.date.max()
            events = events[(events['date'] >= ap.start) & (events['date'] <= ap.end)]

            if ap.report:
                if ap.report == 'transactions':
                    GenerateTransactionsReport(events, ap)
                elif ap.report == 'annual':
                    GenerateSummaryReport(accounts, ap)
                else:
                    logging.critical("Unknown report type {} specified".format(ap.report))
            else:
                logging.critical("No report type specified")


