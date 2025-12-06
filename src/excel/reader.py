from __future__ import annotations

from dataclasses import dataclass
from contextlib import contextmanager
from pathlib import Path
from typing import Dict, List, Optional
import warnings

import pandas as pd


@dataclass
class SheetPreview:
    name: str
    columns: List[str]
    sample: pd.DataFrame


class ExcelReader:
    """Loads Excel files and exposes sheet metadata and previews."""

    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        if not self.path.exists():
            raise FileNotFoundError(self.path)
        self._sheets_cache: Dict[str, SheetPreview] = {}

    def sheet_names(self) -> List[str]:
        with self._suppress_openpyxl_print_area_warning():
            return pd.ExcelFile(self.path).sheet_names

    def load_sheet_preview(
        self, sheet_name: str, header_row: int | None = None, start_row: int | None = None, end_row: int | None = None
    ) -> SheetPreview:
        if sheet_name in self._sheets_cache and header_row is None and start_row is None and end_row is None:
            return self._sheets_cache[sheet_name]

        kwargs: Dict[str, object] = {"sheet_name": sheet_name}
        if header_row is not None:
            kwargs["header"] = header_row
        else:
            kwargs["header"] = 0
        if start_row is not None:
            kwargs["skiprows"] = max(start_row - 1, 0)
        if end_row is not None:
            kwargs["nrows"] = max(end_row - (kwargs.get("skiprows", 0) or 0), 0)

        with self._suppress_openpyxl_print_area_warning():
            df = pd.read_excel(self.path, **kwargs)
        df = df.dropna(how="all")  # remove empty rows
        columns = [str(col) for col in df.columns]
        preview = SheetPreview(name=sheet_name, columns=columns, sample=df.head(10))

        if header_row is None and start_row is None and end_row is None:
            self._sheets_cache[sheet_name] = preview
        return preview

    def read_records(
        self,
        sheet_name: str,
        column_mapping: Dict[str, str],
        header_row: int,
        start_row: Optional[int] = None,
        end_row: Optional[int] = None,
    ) -> List[Dict[str, object]]:
        kwargs: Dict[str, object] = {"sheet_name": sheet_name}
        if header_row is not None:
            kwargs["header"] = header_row
        else:
            kwargs["header"] = 0
        if start_row is not None:
            kwargs["skiprows"] = max(start_row - 1, 0)
        if end_row is not None:
            kwargs["nrows"] = max(end_row - (kwargs.get("skiprows", 0) or 0), 0)

        with self._suppress_openpyxl_print_area_warning():
            df = pd.read_excel(self.path, **kwargs)
        df = df.dropna(how="all")

        records: List[Dict[str, object]] = []
        for _, row in df.iterrows():
            mapped_row = {db_col: row[sheet_col] for sheet_col, db_col in column_mapping.items() if sheet_col in row}
            records.append(mapped_row)
        return records

    @contextmanager
    def _suppress_openpyxl_print_area_warning(self):
        with warnings.catch_warnings():
            warnings.filterwarnings(
                "ignore",
                message=r"Print area cannot be set to Defined name:.*",
                category=UserWarning,
                module="openpyxl.reader.workbook",
            )
            yield
