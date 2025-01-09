from datetime import datetime
from pathlib import Path
from typing import NamedTuple


class Date(NamedTuple):
    dt: datetime
    long: str
    short: str
    colvir: str

    def as_dict(self) -> dict[str, str]:
        return {
            "dt": self.dt.isoformat(),
            "long": self.long,
            "short": self.short,
            "colvir": self.colvir,
        }

    @classmethod
    def to_date(cls, dt: datetime) -> "Date":
        return Date(
            dt=dt,
            long=dt.strftime("%d.%m.%Y"),
            short=dt.strftime("%d.%m.%y"),
            colvir=dt.strftime("%d%m%y"),
        )


class TimeRange(NamedTuple):
    start: Date
    end: Date


class Reports(NamedTuple):
    report_root_folder: Path
    docs_folder: Path
    credit_contracts_fpath: Path
    zbrk_l_deashd4_fpath: Path
    zbrk_l_deashd4_xlsx_fpath: Path
