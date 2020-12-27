import sys
from os.path import isdir, dirname, join, isfile, expanduser
from os import mkdir
from argparse import ArgumentParser
from datetime import datetime
from Intermediate_Import import importcsv

FILENAME = 'ADM%s.xlsx'

if __name__ == '__main__':
    parser = ArgumentParser(description='Start Vocalmotion administratie')
    parser.add_argument('-c', '--csv', type=str,
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

    if not args.csv.endswith('.csv'):
        args.csv = args.csv + '.csv'
    if args.csv and not isfile(args.csv):
        fullcsv = join(args.locatie, args.csv)
        if not isfile(fullcsv):
            sys.exit('De CSV file %s werd niet gevonden' % args.csv)
        else:
            args.csv = fullcsv

    if args.nieuwjaar:
        jaar = args.nieuwjaar
    else:
        jaar = datetime.now().year
    targetfile = join(args.locatie, FILENAME % jaar)
    vorigjaar = join(args.locatie, FILENAME % (jaar - 1))
    if isfile(vorigjaar):
        vm_obj = importcsv(args.csv, administratie=targetfile, verwerkingsjaar=args.nieuwjaar, vorigjaar=vorigjaar)
        vm_obj.bouw_vanuit_vorigjaar()
    else:
        vm_obj = importcsv(args.csv, administratie=targetfile, verwerkingsjaar=args.nieuwjaar)

    vm_obj.process_importfile()
    vm_obj.process_leden()
    vm_obj.process_transactions()
    vm_obj.proc_debiteuren()
    vm_obj.save()
