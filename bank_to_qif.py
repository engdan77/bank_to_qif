#!/usr/bin/env python
#-*- coding: utf-8 -*-

__version__ = "$Revision: 20150201.92"

dict_cat_bank_desc = {'Expenses:Car:Parking': ['Solna Stad'],
                   'Expenses:Work:Unemployment Fund': ['.*UNIONEN.*'],
                   'Expenses:Home:Services': ['.*Netflix.*'],
                   'Expenses:Home:Home Insurance': ['.*Hemf.rs.kring.*'],
                   'Expenses:Home:Mortgage Loan': ['.*VERFRING \d{11}.*', 'L.neavi\s\d+'],
                   'Expenses:Home:Computer Services': ['.*spideroak.*', '.*Evernote.*', '.*City Network.*', '.*iPeer.*', '.*JibJab.*'],
                   'Expenses:Health:Gym': ['.*Fitness 24.*'],
                      'Expenses:Health:Eye': ['.*Synsam.*'],
                      'Expenses:Car:Insurance': ['.*Falck.*', '.*Bilf.rs.kring.*'],
                      'Assets:Current Assets:Savings Account': ['.*besparing.*', '.*spara.*', '.*verf. 9159.*', '.*R.NTEKONTO*'],
                      'Expenses:Entertainment:Magazines': ['.*Datormagazin.*'],
                      'Expenses:Credit Card': ['.*eurocard.*'],
                      'Income:Salary': ['^L.n\b', '.*salary.*', '.*\(\)\d{9}.*', 'LN.*', '.*T\sL.N$'],
                      'Expenses:Uncategorized': ['.*Paypal.*'],
                      'Expenses:Boat:Fees': ['.*Marinpool.*']}

dir_cat_bank_cat = {'Assets:Current Assets:Retirement Savings': ['Pensionsparande'],
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
                    'Expenses:Work:Income Insurance': ['Inkomstf.rs.kring'],
                    'Expenses:Interest:Mortgage Interest': ['Bol.n'],
                    'Expenses:Helth:Eye': ['Syn och h.rsel'],
                    'Expenses:Health:Misc': ['L.kare, sjukv.rd, tandl.kare', 'Medicin'],
                    'Expenses:Public Transportation': ['Flyg, hyrbil och semestertransport', 'Taxi'],
                    'Expenses:Uncategorized': ['.vrigt: Okategoriserade utgifter', 'Uttagsautomat', 'Hotel och .vernattning', '.vrigt: Boende och hush.ll', '.vrigt: Shopping och service', 'M.bler och interi.r', 'Renovering och underh.ll', 'Bio, teater, konserter etc', 'Hobby', 'Nattklubb, dansst.lle, bar', 'Skidor och vintersport', 'Liv- och sjukf.rs.kring', 'G.vor', 'B.cker och spel', 'Smycken', 'Tr.dg.rd', 'K.p av konst', 'N.jen under semester', 'Prenumerationer och tidningar', '.vrigt: Restauranger och n.jen', 'Film, DVD etc', 'Tobak, snus, cigaretter etc', 'Leksaker', 'Kollektivtrafik', 'Bankavgifter', 'Bageri', '.vrigt: Fritid', 'Kiosker, glassbarer etc', 'Musik och instrument', 'Skol- och fritidshemavgifter', '.vrigt: Mat', 'Hemf.rs.kring', 'Sk.nhetsprodukter', 'Kosttillskott och vitaminer', 'Kurslitteratur och kontorsvaror', 'F.reningsliv', '.verf.ring mellan egna konton'],
                    'Expenses:Work:Unemployment Fund': ['A-Kassa'],
                    'Income:Other Income': ['.vrigt: Utl.gg och .terbetalda utl.gg', 'Okategoriserad inkomst'],
                    'Income:Tax Refund': ['Skatte.terb.ring'],
                    'Expenses:Credit Card': ['Avbetalning konsumtionsl.n']}


def clean_skandia_excel_xml(filename, output):
    """ Clear out the <x: etc from the file """
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


def clean_list(list_object):
    ''' Remove header, get all columns and replace properly and create list like ...  [[u'TAXI J\xd6NK\xd6PING', u'Taxi', u'2010-10-12', u'-278.0000']]'''
    import re
    rows = list_object[0]
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


def parse_excel_xml(filename):
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

    excel_handler = ExcelHandler()
    parse(filename, excel_handler)
    return excel_handler.tables


def checj_in_existing(existing_file, date, amount):
    ''' Check if date and amount exists in file '''
    import re
    existing_file.seek(0)

    if existing_file:
        for line in existing_file.readlines():
            try:
                date_file, date_amount = re.match(r'([^;]+);[^;]+;[^;]+;([\d\-]+)$', line).groups()
            except:
                print "Wrong line format in %s: %s" % (existing_file.name, line)
            else:
                if str(date) == str(date_file) and str(amount) == str(date_amount):
                    return True
    # If none found
    return False


def add_to_existing(existing_file, date, description, bank_category, amount):
    ''' Adding a record to the existing file '''
    existing_file.write("%s;%s;%s;%s\n" % (date, description.encode('ascii', 'ignore'), bank_category, amount))


def convert_list_by_cat(bank_list, dict_cat_bank_dec, dict_cat_bank_cat, **args):
    ''' Converting Categories based on dictionaries '''
    import re
    from datetime import datetime

    uncat = args.get('uncat', 'Expenses:Uncategorized')
    verbose = args.get('verbose', False)

    result = []

    # Iterate through all transactions
    for transaction in bank_list:
        if len(transaction) == 5:
            description, bank_category, date, amount, other = transaction
        elif len(transaction) == 4:
            description, bank_category, date, amount = transaction
        elif len(transaction) == 3:
            description, date, amount = transaction
            bank_category = uncat
        else:
            print "Wrong number of columns: %s" % (str(transaction))

        # Correct decimals in amount
        amount = re.sub(r'\.\d{4}$', '', amount)
        # Convert dateformat
        if re.match('\d{4}-\d{2}-\d{2}', date):
            date = datetime.strptime(date, '%Y-%m-%d').strftime('%d/%m/%Y')
        else:
            print "Wrong date format: %s" % (str(transaction))

        cat_found = False

        # Second check for matching of description
        for category, expressions in dict_cat_bank_dec.items():
            for expression in expressions:
                # Check if expression match with description
                if re.match(expression, description, flags=re.IGNORECASE) and not cat_found:
                    # Found match replace/assign category
                    bank_category = category
                    cat_found = True
                    if verbose:
                        print "Found match in description for %s - updated record |%s|%s|%s|%s|" % (expression, description.encode('ascii', 'ignore'), bank_category, date, amount)

        # First check if matching of category is found
        for category, expressions in dict_cat_bank_cat.items():
            for expression in expressions:
                # Check if expression match with description
                if re.match(expression, bank_category, flags=re.IGNORECASE) and not cat_found:
                    # Found match replace/assign category
                    bank_category = category
                    cat_found = True
                    if verbose:
                        print "Found match of Category for %s - updated record |%s|%s|%s|%s|" % (expression, description.encode('ascii', 'ignore'), bank_category, date, amount)

        # If not catefories found
        if not cat_found:
                bank_category = uncat

        result.append((date, description, bank_category, amount))
    return result


def create_qif_header(output, account):
    ''' Function to create QIF Header '''
    output.write('!Account\n')
    output.write("N%s\n" % args.account)
    output.write('TBank\n')
    output.write('^\n')


def add_qif_record(output, date, description, category, amount):
    ''' Function to create QIF file '''
    description = description.encode('ascii', 'ignore')
    output.write('!Type:Bank\n')
    output.write('D%s\n' % date)
    output.write('P%s\n' % description)
    output.write('T%s\n' % amount)
    output.write('L%s\n' % category)
    output.write('^\n')


if __name__ == "__main__":
    ''' Main function '''
    import argparse
    TEMP = '/tmp/convert.tmp'

    parser = argparse.ArgumentParser(description="Tool to parse bank-exports into qif-format supported by GnuCash")
    parser.add_argument('input', metavar='file.xls', help="Input file to parse (Skandiabanken xls-file)", type=argparse.FileType('r'))
    parser.add_argument('output', metavar='file.qif', help="Output qif file", type=argparse.FileType('w'))
    parser.add_argument('--existing', metavar='file.csv', help="Create/Update/Read CSV file to only create qif of new changes", type=argparse.FileType('a+'))
    parser.add_argument('--account', metavar='Default_Qif_Account', help="The BankAccount to associate QIF output to", default='Assets:Current Assets:Checking account')
    parser.add_argument('--verbose', help="Verbose mode", action='store_true')
    args = parser.parse_args()

    clean_skandia_excel_xml(args.input.name, TEMP)
    skandiaList = parse_excel_xml(TEMP)
    skandiaList = clean_list(skandiaList)
    skandiaList = convert_list_by_cat(skandiaList, dict_cat_bank_desc, dir_cat_bank_cat, existing=args.existing, verbose=args.verbose)

    row_no = 0
    existing_no = 0
    duplicate_sum = 0

    if args.input and args.output:
        # Adding header to output
        create_qif_header(args.output, args.account)

        # Parse through all records in input
        for row in skandiaList:
            row_no += 1
            date, description, category, amount = row
            # Check if row exist in existingfile and if not true
            if not checj_in_existing(args.existing, date, amount):
                if args.verbose:
                    print "Adding record to existing: %s, %s, %s, %s" % (date, description.encode('ascii', 'ignore'), amount, category)
                existing_no += 1
                # Adding record to existing
                add_to_existing(args.existing, date, description, category, amount)
                # Add record to QIF output
                add_qif_record(args.output, date, description, category, amount)
            else:
                if args.verbose:
                    print "Already or duplicate record in %s: %s" % (args.existing.name, row)
                duplicate_sum += int(amount)
        print "Processed %s records in %s, added %s to %s" % (row_no, args.input.name, existing_no, args.existing.name)
        print "Total summary of duplicates: %s" % (duplicate_sum,)

    if args.existing:
        args.existing.close()
    if args.output:
        args.output.close()
    if args.input:
        args.input.close()
