import sys
from os.path import isdir, dirname, join, isfile, expanduser
from os import mkdir
from argparse import ArgumentParser
from datetime import datetime
from Intermediate_Import import ImportCsv, ImportXls

FILENAME = 'ADM%s.xlsx'

if __name__ == '__main__':
    parser = ArgumentParser(description='Start Vocalmotion administratie')
    parser.add_argument('-c', '--csv', type=str, required=True,
                        help='Selecteer waar de import van de csv staat')
    parser.add_argument('-j', '--nieuwjaar', type=int,
                        help='Maak nieuw administratie aan voor dit jaar')
    parser.add_argument('-l', '--locatie', type=str, default=None,
                        help='Specificeer locatie waar de spreadsheet opgeslagen wordt')

    args = parser.parse_args()
    if not args.locatie:
        args.locatie = expanduser(join('~', 'Documents', 'Vocalmotion administratie'))
    if not isdir(args.locatie):
        if isdir(dirname(args.locatie)):
            mkdir(args.locatie)
        else:
            sys.exit('Locatie voor opslag gespecificeerd, maar die bestaat niet')

    new_in_filename = args.csv
    fullcsv = args.csv
    checklist = ('.csv', '.xlsx', 'xls')
    idx = 0
    for ext in checklist:
        if not args.csv.endswith(checklist[idx]):
            new_in_filename = args.csv + ext
        if new_in_filename and not isfile(new_in_filename):
            fullcsv = join(args.locatie, new_in_filename)
            if not isfile(fullcsv):
                continue
        args.csv = fullcsv
        break
    else:
        sys.exit('De CSV file of XLSX met de transacties %s werd niet gevonden in %s' % (args.csv, args.location) )

    if args.nieuwjaar:
        jaar = args.nieuwjaar
    else:
        jaar = datetime.now().year
    targetfile = join(args.locatie, FILENAME % jaar)
    vorigjaar = join(args.locatie, FILENAME % (jaar - 1))
    if isfile(vorigjaar) and not isfile(targetfile):
        vm_obj = ImportCsv(args.csv, administratie=targetfile, verwerkingsjaar=args.nieuwjaar, vorigjaar=vorigjaar)
        vm_obj.bouw_vanuit_vorigjaar()
        import_all = True
    else:
        import_all = False
        if ext == '.csv':
            vm_obj = ImportCsv(args.csv, administratie=targetfile, verwerkingsjaar=args.nieuwjaar)
        else:
            vm_obj = ImportXls(args.csv, administratie=targetfile, verwerkingsjaar=args.nieuwjaar)

    vm_obj.process_importfile()
    vm_obj.process_leden()
    vm_obj.process_transactions()
    vm_obj.proc_debiteuren()
    vm_obj.save()
