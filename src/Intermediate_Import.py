import csv
from datetime import datetime
from decimal import Decimal
from os.path import join, dirname

from dateutil.parser import parser
from openpyxl import load_workbook

from ExcelUtils import find_value

PARSER = parser()


def transplusminus(invalue):
    if invalue.lower() == 'af':
        return '-'
    return '+'


def setamount(inamount):
    if not '.' in inamount and not ',' in inamount:
        inamount = inamount[:-2] + '.' + inamount[-2:]
    return Decimal(inamount)


ING_mapping = [(0, 'transactiondate', lambda x: PARSER.parse(x)),
               (1, 'naam', None),
               (2, 'ownaccountnr', None),
               (3, 'accountnr', None),
               (5, 'plusminus', lambda x: transplusminus(x)),
               (6, 'amount', lambda x: setamount(x)),
               ]


class importline:
    def __init__(self, parent, inlist):
        self.transactiondate = datetime.now()
        self.parent = parent
        self.plusminus = '+'
        self.amount = Decimal('0')
        self.accountnr = '0'
        self.naam = 'Onbekend'
        if not parent.bank_sjabloon:
            parent.bank_sjabloon = ING_mapping
        for colnr, fieldname, conversion in parent.bank_sjabloon:
            setattr(self, fieldname, conversion(inlist[colnr]) if conversion else inlist[colnr])
        self.grootboekrekening = self.checkgrbrek()
        self.amount = self.amount if self.plusminus == '+' else self.amount * -1

    def checkgrbrek(self):
        zwaarte = 100
        ok_recognized = {}
        for grootboekrekening, grbkwaarde in self.parent.rekeningschema.items():
            ok_recognized[grootboekrekening] = [0,0,0,0]
            idx = -1
            minwaarde = Decimal('0')
            for herkenningsitem, waarde in grbkwaarde['herkenningswaarden'].items():
                idx += 1
                if not waarde:
                    continue
                if isinstance(waarde, str):
                    listwaarde = [w.strip().lower() for w in waarde.split(',')]
                else:
                    listwaarde = [waarde]
                if herkenningsitem in self.__dict__:
                    for lw in listwaarde:
                        currentval = getattr(self, herkenningsitem)
                        if lw == '*' or (lw and currentval and lw in currentval.lower()):
                            ok_recognized[grootboekrekening][idx] = zwaarte
                    zwaarte = int(zwaarte / 2)
                elif idx == 2 and waarde <= self.amount:
                    minwaarde = waarde
                    ok_recognized[grootboekrekening][idx] = zwaarte
                elif idx == 3 and waarde >= self.amount >= minwaarde:
                    ok_recognized[grootboekrekening][idx] = zwaarte
        found_grootboekrekening = None
        max_waarde = 0
        for grootboekrekening, found_values in ok_recognized.items():
            if found_values[0] and found_values[1] and max_waarde < sum(found_values):
                max_waarde = sum(found_values)
                found_grootboekrekening = grootboekrekening
        if max_waarde > 120:
            return found_grootboekrekening

    def is_contributie(self, contributiebedrag):
        return self.plusminus == '+' and self.amount and (self.amount % contributiebedrag) < 2

    def addledenrow(self):
        return [self.accountnr, self.naam, "", "", "", self.transactiondate.strftime('%d-%m-%Y'), "", ""]

    def addtransrow(self):
        return [self.transactiondate, self.grootboekrekening, self.plusminus, self.amount, self.naam]


class importcsv:
    def __init__(self, importfile, bank_sjabloon: list = None, administratie: str = None):
        self.rowlist = []
        if administratie:
            self.administratie = administratie
        else:
            self.administratie = join(dirname(dirname(__file__)), 'resources', 'Administratie_sjabloon.xlsx')
        self.wb = load_workbook(self.administratie, keep_vba=True)
        common_ws = self.wb['Standaardwaarden']
        cell = find_value('Contributiebedrag', common_ws)
        self.rekeningschema = self.buildrekschema(self.wb['Rekeningschema'])
        self.contributiebedrag = Decimal(common_ws['B' + str(cell.column + 1)].value)

        self.importfile = importfile
        self.bank_sjabloon = bank_sjabloon

    def buildrekschema(self, rekws):
        rekeningschema = {}
        for rekeningschemarij in rekws.rows:
            v = rekeningschemarij[0].value
            if v:
                try:
                    rekeningnr = int(v)
                except ValueError:
                    continue
                rekeningschema[rekeningnr] = {
                    'omschrijving': rekeningschemarij[2].value,
                    'herkenningswaarden': {
                        'plusminus': '+' if len(rekeningschemarij) > 3 and rekeningschemarij[3].value and rekeningschemarij[3].value.lower() == 'bij' else '-',
                        'naam': rekeningschemarij[4].value.lower() if len(rekeningschemarij) > 4 and rekeningschemarij[4].value else None,
                        'min_bedrag': Decimal(rekeningschemarij[5].value) if len(rekeningschemarij) > 5 and rekeningschemarij[5].value else None,
                        'max_bedrag': Decimal(rekeningschemarij[6].value) if len(rekeningschemarij) > 6 and rekeningschemarij[6].value else None}}
        return rekeningschema

    def process_importfile(self):
        with open(self.importfile) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            for line_count, row in enumerate(csv_reader):
                if line_count == 0:  # HEADER
                    print('Import van csv bestand %s loopt' % csv_file)
                else:  # DATALINES
                    row_object = importline(self, row)
                    self.rowlist.append(row_object)

    def process_leden(self):
        ledenws = self.wb['Leden']
        for newrow in self.rowlist:
            accountcell = find_value(newrow.accountnr, ledenws, max_col=1)
            if accountcell:
                self.update_lid(newrow, list(ledenws[accountcell.row]))
            elif newrow.is_contributie(self.contributiebedrag):
                ledenws.append(newrow.addledenrow())

    def update_lid(self, newrow, ledenrow):
        if newrow.transactiondate > PARSER.parse(ledenrow[5].value):
            ledenrow[5] = newrow.transactiondate

    def process_transactions(self):
        transws = self.wb['Transacties']
        for newrow in self.rowlist:
            accountcell = find_value([newrow.accountnr, newrow.amount, newrow.transactiondate], transws)
            if not accountcell:
                transws.append(newrow.addtransrow())
