"""Provide outputs for the director's rota."""

import datetime
from pathlib import Path
from dateutil.relativedelta import relativedelta  # type: ignore
from workbooky import Workbook, Worksheet
import asyncio
from typing import NamedTuple

from directors_rota.config import read_config

import directors_rota.text as text

status = {
    'OK': 0,
    'FILE_MISSING': 1,
    'SHEET_MISSING': 2,
}


class DirectorData(NamedTuple):
    initials: str
    name: str
    email: str
    username: str
    active: bool


class Director():
    def __init__(self, data: DirectorData) -> None:
        self.initials = ''
        self.name = ''
        self.email = ''
        self.dates = []
        self._dollars = 0
        self.first_name = self._get_first_name()
        self.username = ''
        self.active = False

        if data:
            self.initials = data.initials
            self.name = data.name
            self.email = data.email
            self.username = data.username
            self.active = data.active

    def __repr__(self) -> str:
        return f'{self.initials} {self.name}'

    def _get_first_name(self) -> str:
        return self.name.split(' ')[0]


def generate_rota(month: datetime) -> None:
    config = read_config()
    path = Path(config.workbook_dir, config.workbook_file_name)
    workbook = _get_workbook(path)
    if workbook == status['FILE_MISSING']:
        return (status['FILE_MISSING'],)

    main_sheet = asyncio.run(_get_sheet(workbook, config.main_sheet))
    if main_sheet == status['SHEET_MISSING']:
        return (status['SHEET_MISSING'], config.main_sheet)

    directors_sheet = asyncio.run(_get_sheet(workbook,
                                             config.directors_sheet))
    if directors_sheet == status['SHEET_MISSING']:
        return (status['SHEET_MISSING'], config.directors_sheet)

    directors = get_directors(config, directors_sheet)
    (start_date, end_date) = _date_limits(month)
    return (
        status['OK'],
        _get_rota(start_date, end_date, config, main_sheet, directors),
        directors)


def _get_workbook(path):
    """Return the workbook from the path."""
    try:
        workbook = Workbook(path)
        return workbook
    except FileNotFoundError:
        return status['FILE_MISSING']


async def _get_sheet(workbook, sheet_name):
    """Return a sheet from the workbook."""
    try:
        sheet = await workbook.get_worksheet(sheet_name)
        return sheet
    except KeyError:
        return status['SHEET_MISSING']
    except Exception as err:
        print(f'{sheet_name, err}')


def get_directors(config, worksheet: object) -> dict[str, Director]:
    # Return a dict of Director objects keyed on initials
    directors = {}
    for row in worksheet.iter_rows(values_only=True):
        if row[0] and row[0] != 'Initials':
            data = DirectorData(
                row[config.initials_col],
                row[config.name_col],
                row[config.email_col],
                row[config.username_col],
                row[config.active_col],
            )
            director = Director(data)
            directors[director.initials] = director
    return directors


def _date_limits(month: datetime) -> tuple[datetime.datetime,
                                           datetime.datetime]:
    """Return start and end dates for the rota."""
    year = month.year
    start_date = datetime.datetime(year, month.month, 1)
    end_date = start_date + relativedelta(months=+1, seconds=-1)
    return (start_date, end_date)


def _get_rota(start_date: datetime,
              end_date: datetime,
              config: dict,
              main_sheet: object,
              directors: dict[str, Director]) -> list[str]:
    """Print the rota."""
    mon_rota = _get_rota_dates(config.mon_date_col, start_date, end_date,
                               main_sheet, directors)
    wed_rota = _get_rota_dates(config.wed_date_col, start_date, end_date,
                               main_sheet, directors)
    rota = _generate_rota_list(mon_rota, wed_rota)
    return _create_rota_email(config, start_date, rota)


def _get_rota_dates(date_col: int,
                    start_date: datetime,
                    end_date: datetime,
                    worksheet: Worksheet,
                    directors: dict[str, Director]) -> list[str]:
    """Return a list of dates and directors names."""
    rota = []
    dir_col = date_col + 1
    for row in worksheet.iter_rows(values_only=True):
        rota_date = row[date_col]
        if isinstance(rota_date, datetime.date):
            if start_date <= rota_date < end_date:
                err_date = f'{rota_date:%d %b %Y}'
                dir_inits = row[dir_col]
                if not dir_inits:
                    print(f'{text.NO_DIRECTOR} {err_date}')
                    continue
                if dir_inits not in directors:
                    print(f'{text.NO_DIRECTOR_NAME} {dir_inits} on {err_date}')
                    continue
                director = directors[dir_inits]
                rota.append(f'{rota_date:%d/%m/%y}, {director.name}')
    return rota


def _generate_rota_list(mon_rota: list[str], wed_rota: list[str]) -> list[str]:
    # Return the rota as a list of strings
    rota = []
    rota.append('Mondays')
    if not mon_rota:
        rota.append(text.NO_DATES)
    else:
        for item in mon_rota:
            rota.append(item)
    rota.append('')
    rota.append('Wednesdays')
    if not wed_rota:
        rota.append(text.NO_DATES)
    else:
        for item in wed_rota:
            rota.append(item)
    return rota


def _create_rota_email(config: dict,
                       start_date: datetime, rota: list[str]) -> None:
    # Print out the rota email text
    template_path = config.email_template
    try:
        with open(template_path, 'r') as f_email:
            email_text = f_email.read()
    except FileNotFoundError:
        print(f"{text.NO_TEMPLATE} {template_path}")
        quit()

    email_text = email_text.replace('<month>', f'{start_date:%b %Y}')
    email_text = email_text.replace('<rota>', '\n'.join(rota))
    return email_text
