from Intermediate_Import import ImportCsv
from os.path import dirname, join

fixturedir = join(dirname(__file__), 'fixtures')

def test_importcsv():
    i_obj = ImportCsv(join(fixturedir, 'ing_test.csv'))
    i_obj.process_importfile()
    assert(len(i_obj.rowlist) == 12)

    ledenws = i_obj.process_leden()

    transws = i_obj.process_transactions()

    debws = i_obj.proc_debiteuren()



