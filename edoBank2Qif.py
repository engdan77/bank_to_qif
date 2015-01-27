#!/usr/bin/env python
#-*- coding: utf-8 -*-

import argparse

__version__ = "$Revision: 20150127.8"

dictCatBankDesc = {'Expenses:Car:Parking': ['Solna Stad'],
                   'Expenses:Work:Unemployment Fund': ['.*UNIONEN.*'],
                   'Expenses:Home:Services': ['.*Netflix.*'],
                   'Expenses:Home:Home Insurance': ['.*Hemf.rs.kring.*'],
                   'Expenses:Home:Computer Services': ['.*spideroak.*', '.*Evernote.*', '.*City Network.*', '.*iPeer.*', '.*JibJab.*'],
                   'Expenses:Health:Gym': ['.*Fitness 24.*'],
                   'Expenses:Health:Eye': ['.*Synsam.*'],
                   'Expenses:Car:Insurance': ['.*Falck.*'],
                   'Expenses:Car:Insurance': ['.*Bilf.rs.kring.*'],
                   'Assets:Current Assets:Savings Account:': ['.*besparing.*', '.*spara.*'],
                   'Expenses:Boat:Fees': ['.*Marinpool.*']}

dictCatBankCat = {'Assets:Current Assets:Retirement Savings': ['Pensionsparande'],
                  'Expenses:Car:Fees': ['Fordonsskatt', '.vrigt: Bil och transport', 'V.gtull'],
                  'Expenses:Car:Gas': ['Br.nsle'],
                  'Expenses:Car:Parking': ['Parkering'],
                  'Expenses:Car:Repair and Maintenance': ['Bilservice', 'Bilv.rd'],
                  'Expenses:Child': ['Barnartiklar'],
                  'Expenses:Clothes': ['Kl.der och skor'],
                  'Expenses:Computer': ['Dator och elektronik'],
                  'Expenses:Dining': ['Restaurang och kaf.*', 'Snabbmat'],
                  'Expenses:Entertainment:Alcohol': ['Alkohol'],
                  'Expenses:Entertainment:Travel': ['.vrigt: Semester och resor'],
                  'Expenses:Groceries': ['Livsmedel'],
                  'Expenses:Health:Gym': ['Friskv.rd och tr.ningskort'],
                  'Expenses:Home:Electric': ['El och v.rme'],
                  'Expenses:Home:Rent': ['Hyra', 'Boendeavgifter'],
                  'Expenses:Home:Services': ['Tv, telefoni och internet'],
                  'Expenses:Income:Insurance': ['Inkomstf.rs.kring'],
                  'Expenses:Interest:Mortgage Interest': ['Bol.n'],
                  'Expenses:Medical Expenses:Eye': ['Syn och h.rsel'],
                  'Expenses:Medical Expenses:Misc': ['L.kare, sjukv.rd, tandl.kare', 'Medicin'],
                  'Expenses:Public Transportation': ['Flyg, hyrbil och semestertransport', 'Taxi'],
                  'Expenses:Uncategorized': ['.vrigt: Okategoriserade utgifter', 'Uttagsautomat', 'Hotel och .vernattning', '.vrigt: Boende och hush.ll', '.vrigt: Shopping och service', 'M.bler och interi.r', 'Renovering och underh.ll', 'Bio, teater, konserter etc', 'Hobby', 'Nattklubb, dansst.lle, bar', 'Skidor och vintersport', 'Liv- och sjukf.rs.kring', 'G.vor', 'B.cker och spel', 'Smycken', 'Tr.dg.rd', 'K.p av konst', 'N.jen under semester', 'Prenumerationer och tidningar', '.vrigt: Restauranger och n.jen', 'Film, DVD etc', 'Tobak, snus, cigaretter etc', 'Leksaker', 'Kollektivtrafik', 'Bankavgifter', 'Bageri', '.vrigt: Fritid', 'Kiosker, glassbarer etc', 'Musik och instrument', 'Skol- och fritidshemavgifter', '.vrigt: Mat', 'Hemf.rs.kring', 'Sk.nhetsprodukter', 'Kosttillskott och vitaminer', 'Kurslitteratur och kontorsvaror', 'F.reningsliv', '.verf.ring mellan egna konton'],
                  'Expenses:Work:Unemployment Fund': ['A-Kassa'],
                  'Income:Other Income': ['.vrigt: Utl.gg och .terbetalda utl.gg', 'Okategoriserad inkomst'],
                  'Income:Salary': ['L.n'],
                  'Income:Tax Refund': ['Skatte.terb.ring'],
                  'Liabilities:Credit Card': ['Avbetalning konsumtionsl.n']}


def cleanSkandiaExcelXML(filename, output):
    ''' Clear out the <x: etc from the file '''
    import re

    f = open(filename, 'r')
    lines = f.readlines()

    w = open(output, 'wa')

    for line in lines:
        line = re.sub(r'<\w+:', '<', line)
        line = re.sub(r'</\w+:', '</', line)
        w.write(line)

    f.close()
    w.close()


def cleanList(listObject):
    ''' Remove header, get all columns and replace properly and create list like ...  [[u'TAXI J\xd6NK\xd6PING', u'Taxi', u'2010-10-12', u'-278.0000']]'''
    import re
    rows = listObject[0]
    rows.pop(0)
    result = []

    for row in rows:
        new_row = []
        for i, column in enumerate(row):
            column = re.sub(r'^\s+', '', column)
            column = re.sub(r'\\n\s+', '', column)
            column = re.sub(r'\\n\s+', '', column)
            column = re.sub(r'\s+$', '', column)
            column = column.replace('\n', '')
            if len(column) > 0:
                new_row.append(column)
        result.append(new_row)
    return result


def parseExcelXML(filename):
    ''' return a list of rows '''
    from xml.sax import parse
    from xml.sax import handler

    class ExcelHandler(handler.ContentHandler):
        def __init__(self):
            self.chars = []
            self.cells = []
            self.rows = []
            self.tables = []

        def characters(self, content):
            self.chars.append(content)

        def startElement(self, name, atts):
            if name == "Cell":
                self.chars = []
            elif name == "Row":
                self.cells = []
            elif name == "Table":
                self.rows = []

        def endElement(self, name):
            if name == "Cell":
                self.cells.append(''.join(self.chars))
            elif name == "Row":
                self.rows.append(self.cells)
            elif name == "Table":
                self.tables.append(self.rows)

    excelHandler = ExcelHandler()
    parse(filename, excelHandler)
    return excelHandler.tables


def checkInExisting(existingFile, date, amount):
    ''' Check if date and amount exists in file '''
    import re
    existingFile.seek(0)

    if existingFile:
        for line in existingFile.readlines():
            try:
                dateFile, dateAmount = re.match(r'([^;]+);[^;]+;[^;]+;([\d\-]+)$', line).groups()
            except:
                print "Wrong line format in %s: %s" % (existingFile.name, line)
            else:
                if str(date) == str(dateFile) and str(amount) == str(dateAmount):
                    return True
    # If none found
    return False


def addToExisting(existingFile, date, description, bankCategory, amount):
    ''' Adding a record to the existing file '''
    existingFile.write("%s;%s;%s;%s\n" % (date, description, bankCategory, amount))


def convertListByCat(bankList, dictCatBankDec, dictCatBankCat, **args):
    ''' Converting Categories based on dictionaries '''
    import re
    from datetime import datetime

    uncat = args.get('uncat', 'Expenses:Uncategorized')
    existing = args.get('existing', None)
    verbose = args.get('verbose', False)

    result = []

    # Iterate through all transactions
    for transaction in bankList:
        if len(transaction) == 5:
            description, bankCategory, date, amount, other = transaction
        elif len(transaction) == 4:
            description, bankCategory, date, amount = transaction
        elif len(transaction) == 3:
            description, date, amount = transaction
            bankCategory = uncat
        else:
            print "Wrong number of columns: %s" % (str(transaction))

        # Correct decimals in amount
        amount = re.sub(r'\.\d{4}$', '', amount)
        # Convert dateformat
        if re.match('\d{4}-\d{2}-\d{2}', date):
            date = datetime.strptime(date, '%Y-%m-%d').strftime('%d/%m/%Y')
        else:
            print "Wrong date format: %s" % (str(transaction))

        # First check if matching of category is found
        for category, expressions in dictCatBankCat.items():
            for expression in expressions:
                # Check if expression match with description
                if re.match(expression, bankCategory, flags=re.IGNORECASE):
                    # Found match replace/assign category
                    bankCategory = category
                    if verbose:
                        print "Found match of Category for %s - updated record |%s|%s|%s|%s|" % (expression, description, bankCategory, date, amount)

        # Second check for matching of description
        for category, expressions in dictCatBankDec.items():
            for expression in expressions:
                # Check if expression match with description
                if re.match(expression, description, flags=re.IGNORECASE):
                    # Found match replace/assign category
                    bankCategory = category
                    if verbose:
                        print "Found match in description for %s - updated record |%s|%s|%s|%s|" % (expression, description, bankCategory, date, amount)
        result.append((date, description, bankCategory, amount))
    return result


if __name__ == "__main__":
    ''' Main function '''
    import argparse
    TEMP = '/tmp/convert.tmp'

    parser = argparse.ArgumentParser(description="Tool to parse bank-exports into qif-format supported by GnuCash")
    parser.add_argument('input', metavar='file.xls', help="Input file to parse (Skandiabanken xls-file)", type=argparse.FileType('r'))
    parser.add_argument('output', metavar='file.qif', help="Output qif file", type=argparse.FileType('w'))
    parser.add_argument('--existing', metavar='file.csv', help="Create/Update/Read CSV file to only create qif of new changes", type=argparse.FileType('a+'))
    parser.add_argument('--verbose', help="Verbose mode", action='store_true')
    args = parser.parse_args()

    cleanSkandiaExcelXML(args.input.name, TEMP)
    skandiaList = parseExcelXML(TEMP)
    skandiaList = cleanList(skandiaList)
    skandiaList = convertListByCat(skandiaList, dictCatBankDesc, dictCatBankCat, existing=args.existing, verbose=args.verbose)

    row_no = 0
    existing_no = 0
    for row in skandiaList:
        row_no += 1
        date, description, category, amount = row
        # Check if row exist in existingfile and if not true
        if not checkInExisting(args.existing, date, amount):
            if args.verbose:
                print "Adding record to existing: %s, %s, %s" % (date, description, amount)
            existing_no += 1
            addToExisting(args.existing, date, description.encode('ascii', 'ignore'), category, amount)

    print "Processed %s in %s, added %s to %s" % (row_no, args.input.name, existing_no, args.existing.name)

    if args.existing:
        args.existing.close()
    # print skandiaList
