from Intermediate_Import import importcsv
from os.path import dirname, join

fixturedir = join(dirname(__file__), 'fixtures')

def test_importcsv():
    i_obj = importcsv(join(fixturedir, 'ing_test.csv'))
    i_obj.process_importfile()
    assert(len(i_obj.rowlist) == 12)

    i_obj.process_leden()

    i_obj.process_transactions()