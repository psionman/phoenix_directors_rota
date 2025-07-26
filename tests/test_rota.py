from workbooky import Workbook
from directors_rota.process import get_directors
from pathlib import Path


from directors_rota.config import read_config
config = read_config()

VALID_WORKBOOK_PATH = Path('tests', 'test_data', 'directors-rota.xlsx')


def test_workbook_exists():
    try:
        Workbook(VALID_WORKBOOK_PATH)
        assert True
    except FileNotFoundError:
        assert False


def test_worksheet_exists():
    try:
        workbook = Workbook(VALID_WORKBOOK_PATH)
        worksheet = workbook.worksheets['Main']
        worksheet = workbook.worksheets['Directors']
        assert True
    except FileNotFoundError:
        assert False


def test_get_directors():

    workbook = Workbook(VALID_WORKBOOK_PATH)
    directors_sheet = workbook.worksheets['Directors']
    directors = get_directors(config, directors_sheet)
    assert len(directors) == 9
