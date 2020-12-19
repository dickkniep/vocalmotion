from openpyxl import load_workbook
from os.path import isfile
from os import rename, remove
from shutil import copyfile
from locale import atof, setlocale, LC_NUMERIC, LC_MONETARY


def reverse_sign(f):
    setlocale(LC_NUMERIC, 'nl_NL')
    value = atof(f)
    value = value * -1
    return value

rabo_matchfields = {'A': ('Volgnr', True, int), 'B': ('Rentedatum', False, None),
                    'C': ('Naam tegenpartij', False, None), 'D': ('Omschrijving-1', False, None),
                    'J': ('Bedrag', False, reverse_sign)}
start_row_data = 10

class BankCSV:
    def __init__(self, yeartransactions: dict, filename: str):
        self.filename = filename
        self.yeartransactions = yeartransactions
        self.rekeningcodes = {}
        self.boekhouding = None
        self.werkbalans = None
        self.namen = []
        self.accounts = []

    def get_accounts(self):
        self.rekeningcodes = { self.boekhouding['M5'].value: self.boekhouding['M4'].value,
                          self.boekhouding['N6'].value.replace('.',''): self.boekhouding['N4'].value,
                          self.boekhouding['O6'].value.replace('.',''): self.boekhouding['O4'].value }

    def get_bankcode(self, rekening):
        rek = rekening[8:]
        return self.rekeningcodes[rek]

    def backup_file(self):
        orig_backupname = backupname = self.filename + '.bak'
        counter = 1
        while isfile(backupname):
            backupname = orig_backupname + '_' + str(counter)
            counter += 1

        counter -= 1

        if counter == 9:
            remove(orig_backupname + '_' + str(counter-1))
            backupname = orig_backupname + '_' + str(counter-1)
            counter -= 1

        while counter > 0:
            counter -= 1
            if counter > 0:
                oldbackupname = orig_backupname + '_' + str(counter)
            else:
                oldbackupname = orig_backupname
            rename(oldbackupname, backupname)
            backupname = oldbackupname

        copyfile(self.filename, orig_backupname)

    def getresidents(self, bewoners):
        for rownr, row in enumerate(bewoners):
            bewonernaam = self.get_sheet_value('E', rownr+1, bewoners)
            if bewonernaam:
                self.namen.append(bewonernaam)
            if rownr > 18:
                break

    def getaccounts(self, werkbalans):
        for rownr, row in enumerate(werkbalans):
            account = self.get_sheet_value('A', rownr+1, werkbalans)
            if account:
                try:
                    int(account)
                except ValueError:
                    pass
                else:
                    self.accounts.append(account)


    def process(self):
        self.backup_file()
        wb = load_workbook(self.filename, keep_links=True, keep_vba= True)
        self.boekhouding = wb['boekhouding']
        self.werkbalans = wb['werkbalans']
        self.getresidents(wb['bewoners'])

        self.get_accounts()

        for row in self.yeartransactions:
            matched_rownr, full_match = self.already_entered(row)
            if full_match: # No update necessary
                print('matched:')
                print(row)
            elif matched_rownr:
                self.updrow(row, matched_rownr)
            else:
                self.addrow(row)

        wb.save(self.filename)

    def get_sheet_value(self, column: str, rownr: int, sheet=None):
        if not sheet:
            sheet = self.boekhouding
        col_id = column + str(rownr)
        return sheet[col_id].value

    def set_sheet_value(self, column: str, new_value, rownr: int = None, sheet=None, func=None):
        if not sheet:
            sheet = self.boekhouding
        if rownr:
            col_id = column + str(rownr)
        else:
            col_id = column

        if func:
            new_value = func(new_value)

        sheet[col_id].value = new_value

    def already_entered(self, row) -> (int, bool):
        matched_rownr = None
        fullmatch = False
        for rownr, bhr in enumerate(self.boekhouding):
            if rownr < start_row_data or self.get_sheet_value('H', rownr) == 1:        # Skip balansrekeningen
                continue

            fullmatch = True
            matched = True
            for column, field_data in rabo_matchfields.items():
                field, totalmatch, func = field_data
                if row[field] == self.get_sheet_value(column, rownr):
                    if totalmatch:
                        matched_rownr = rownr
                else:
                    fullmatch = False
                    matched = False
            if matched:
                matched_rownr = rownr
            if fullmatch or matched_rownr :
                break
        return matched_rownr, fullmatch

    def updrow(self, row, matched_rownr: int):
        for col, field_data in rabo_matchfields.items():
            field, totalmatch, func = field_data
            self.set_sheet_value(col, row[field], matched_rownr, func=func)

        self.set_sheet_value('H', 2, matched_rownr)
        self.set_sheet_value('E', None, matched_rownr)
        self.set_sheet_value('F', None, matched_rownr)
        self.set_sheet_value('I', self.get_grbk_rek(row), matched_rownr)

    def addrow(self, row):
        rownr = self.get_first_empty_row()
        self.updrow(row, rownr)

    def get_first_empty_row(self) -> int:
        rownr = start_row_data
        for rownr, bhr in enumerate(self.boekhouding):
            if rownr < start_row_data or self.get_sheet_value('H', rownr+1) == 1 or self.get_sheet_value('A', rownr+1):
                continue
            return rownr + 1
        else:
            added_rownr = rownr + 1
            for col in [cell.column for cell in self.boekhouding[1]]:
                self.boekhouding[col + str(added_rownr)].value = self.get_sheet_value(col, rownr+1)
            return rownr + 1

    def get_grbk_rek(self, row):
        if 'kosten' == row['Naam tegenpartij'].lower():
            return 418

        for naam in self.namen:
            if naam.lower() in row['Naam tegenpartij'].lower():
                return 101

        if 'tuinonderhoud' in row['Naam tegenpartij'].lower():
            return 420

        if 'nuon' in row['Naam tegenpartij'].lower():
            return 413

        if 'medisol' in row['Naam tegenpartij'].lower():
            return 419



