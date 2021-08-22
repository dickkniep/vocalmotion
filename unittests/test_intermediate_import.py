from Intermediate_Import import ImportCsv
from os.path import dirname, join
from os import makedirs
from shutil import copyfile, rmtree
from tempfile import gettempdir
from pytest import fixture
fixturedir = join(dirname(__file__), 'fixtures')

@fixture
def bld_tempadm():
    testdir = join(gettempdir(), 'vocaltest', '2019')
    rmtree(join(gettempdir(), 'vocaltest'), ignore_errors=True)
    makedirs(testdir)
    copyfile(join(dirname(dirname(__file__)), 'resources', 'Administratie_sjabl.xlsx'),
             join(testdir, 'adm2019.xlsx'))
    yield testdir
    rmtree(join(gettempdir(), 'vocaltest'))


def test_importcsv(bld_tempadm):

    i_obj = ImportCsv(join(fixturedir, 'ing_test.csv'), join(bld_tempadm, 'adm2019.xlsx'))
    i_obj.process_importfile()
    assert(len(i_obj.rowlist) == 9)
    i_obj.save()

    i_obj = ImportCsv(join(fixturedir, 'ing_test2.csv'), join(bld_tempadm, 'adm2019.xlsx'))
    i_obj.process_importfile()
    assert(len(i_obj.rowlist) == 8)
    i_obj.save()

    i_obj.process_leden()
    assert(list(i_obj.wb['Leden'][4])[0].value == 'NL12INGB0002229737')
    assert(list(i_obj.wb['Leden'][3])[0].value == 'NL03ABNA0981912591')

    i_obj.save()

    vm_obj = ImportCsv(join(fixturedir, 'ing_test_2020.csv'), join(dirname(bld_tempadm), '2020', 'adm2020.xlsx'),
                       vorigjaar=join(bld_tempadm, 'adm2019.xlsx'))
    vm_obj.bouw_vanuit_vorigjaar()

    vm_obj.save()


