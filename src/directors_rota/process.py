"""Provide outputs for the director's rota."""

import asyncio
import datetime
from dataclasses import dataclass
from pathlib import Path
from typing import NamedTuple

from dateutil.relativedelta import relativedelta
from workbooky import Workbook

from directors_rota import logger
from directors_rota.config import read_config
from directors_rota.text import Text

status = {
    "OK": 0,
    "FILE_MISSING": 1,
    "SHEET_MISSING": 2,
}

txt = Text()


class DirectorData(NamedTuple):
    initials: str
    name: str
    email: str
    username: str
    active: bool
    send_reminder: bool


class Director:
    def __init__(self, data: DirectorData) -> None:
        self.initials = ""
        self.name = ""
        self.email = ""
        self.dates = []
        self._dollars = 0
        self.first_name = self._get_first_name()
        self.username = ""
        self.active = False

        if data:
            self.initials = data.initials
            self.name = data.name
            self.email = data.email
            self.username = data.username
            self.active = data.active
            self.send_reminder = data.send_reminder

    def __repr__(self) -> str:
        return f"{self.initials} {self.name}"

    def _get_first_name(self) -> str:
        return self.name.split(" ")[0]


@dataclass
class DayRota:
    day: str
    day_rota: list[str]


class RotaData(NamedTuple):
    start_date: datetime
    end_date: datetime
    main_sheet: object
    directors: dict[str, Director]


def generate_rota(month: datetime) -> tuple | None:
    """Return the rota adn directors as a tuple."""
    # pylint: disable=no-member)
    config = read_config()
    path = Path(config.workbook_dir, config.workbook_file_name)
    workbook = _get_workbook(path)
    if workbook == status["FILE_MISSING"]:
        logger.error(f"Workbook not found at {path}")
        return

    main_sheet = asyncio.run(_get_sheet(workbook, config.main_sheet))
    if main_sheet == status["SHEET_MISSING"]:
        logger.error(f"Sheet '{config.main_sheet}' missing in Workbook")
        return

    directors_sheet = asyncio.run(_get_sheet(workbook, config.directors_sheet))
    if directors_sheet == status["SHEET_MISSING"]:
        logger.error(f"Sheet '{config.directors_sheet}' missing in Workbook")
        return

    directors = get_directors(config, directors_sheet)
    rota_email = _get_rota(month, config, main_sheet, directors)
    return (rota_email, directors)


def _get_workbook(path):
    """Return the workbook from the path."""
    try:
        return Workbook(path)
    except FileNotFoundError:
        return status["FILE_MISSING"]


async def _get_sheet(workbook, sheet_name):
    """Return a sheet from the workbook."""

    try:
        return await workbook.get_worksheet(sheet_name)
    except KeyError:
        logger.error(f"Sheet '{sheet_name}' missing in Workbook")
    except Exception as err:
        logger.error(f"Unexpected error: sheet '{sheet_name}' {err}")
    return


def get_directors(config, worksheet: object) -> dict[str, Director]:
    """Return a dict of Director objects keyed on initials."""
    directors = {}
    for row in worksheet.iter_rows(values_only=True):
        if row[0] and row[0] != "Initials":
            data = DirectorData(
                row[config.initials_col],
                row[config.name_col],
                row[config.email_col],
                row[config.username_col],
                row[config.active_col],
                row[config.send_reminder_col],
            )
            director = Director(data)
            directors[director.initials] = director

    return directors


def _date_limits(
    month: datetime,
) -> tuple[datetime.datetime, datetime.datetime]:
    """Return start and end dates for the rota."""
    year = month.year
    start_date = datetime.datetime(year, month.month, 1)
    end_date = start_date + relativedelta(months=+1, seconds=-1)
    return (start_date, end_date)


def _get_rota(
    month: datetime,
    config: dict,
    main_sheet: object,
    directors: dict[str, Director],
) -> list[str]:
    """Print the rota."""
    (start_date, end_date) = _date_limits(month)
    rota_data = RotaData(start_date, end_date, main_sheet, directors)
    day_rotas = []
    day_rotas.append(
        DayRota("Monday", _get_rota_dates(config.mon_date_col, rota_data))
    )
    day_rotas.append(
        DayRota("Wednesday", _get_rota_dates(config.wed_date_col, rota_data))
    )
    day_rotas.append(
        DayRota("Thursday", _get_rota_dates(config.thurs_date_col, rota_data))
    )
    rota = _generate_rota_list(day_rotas)
    return _create_rota_email(config, start_date, rota)


def _get_rota_dates(date_col: int, rota_data: RotaData) -> list[str]:
    """Return a list of dates and directors names."""
    rota = []
    dir_col = date_col + 1
    no_dates = True
    for row in rota_data.main_sheet.iter_rows(values_only=True):
        rota_date = row[date_col]
        if (
            rota_date
            and isinstance(rota_date, datetime.date)
            and rota_data.start_date <= rota_date < rota_data.end_date
        ):
            no_dates = False
            dir_inits = row[dir_col]
            if not dir_inits:
                logger.warning(f"{txt.NO_DIRECTOR} for {rota_date:%d %b %Y}")
                continue
            if dir_inits not in rota_data.directors:
                logger.warning(
                    f"{txt.INVALID_DIRECTOR} '{dir_inits}' "
                    f"for {rota_date:%d %b %Y}"
                )
                continue
            director = rota_data.directors[dir_inits]
            rota.append(f"{rota_date:%d/%m/%y}, {director.name}")
            logger.info(
                f"Rota data added for {rota_date:%d %b %Y}, {director.name}"
            )
    if no_dates:
        logger.warning(
            f"No dates in this period "
            f"{rota_data.start_date:%d/%m/%y} to {rota_data.end_date:%d/%m/%y}"
        )
    return rota


def _generate_rota_list(day_rotas) -> list[str]:
    # Return the rota as a list of strings
    rota = []
    for day_rota in day_rotas:
        rota.append("")
        rota.append(day_rota.day)
        rota.extend(iter(day_rota.day_rota))
    return rota


def _create_rota_email(
    config: dict, start_date: datetime, rota: list[str]
) -> str | None:
    """Return the rota email text."""
    template_path = config.email_template
    try:
        with open(template_path, encoding="utf-8") as f_email:
            email_text = f_email.read()
    except FileNotFoundError:
        logger.error(f"{txt.NO_TEMPLATE} {template_path}")
        return None

    email_text = email_text.replace("<month>", f"{start_date:%b %Y}")
    email_text = email_text.replace("<rota>", "\n".join(rota))
    return email_text
