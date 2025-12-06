from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import warnings
from typing import Dict, List, Optional, Sequence

import pandas as pd


@dataclass
class SheetPreview:
    name: str
    columns: List[str]
    sample: pd.DataFrame
    header_row: int


class ExcelReader:
    """Loads Excel files and exposes sheet metadata and previews."""

    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        if not self.path.exists():
            raise FileNotFoundError(self.path)
        self._sheets_cache: Dict[str, SheetPreview] = {}
        # Some workbooks use Excel table names as print areas, which triggers noisy openpyxl warnings.
        warnings.filterwarnings(
            "ignore",
            message="Print area cannot be set to Defined name.*",
            category=UserWarning,
        )

    def sheet_names(self) -> List[str]:
        return pd.ExcelFile(self.path).sheet_names

    def _normalize_columns(self, columns: Sequence[object]) -> List[str]:
        normalized: List[str] = []
        seen: Dict[str, int] = {}
        for idx, col in enumerate(columns, start=1):
            name = "" if col is None else str(col).strip()
            if not name or name.lower().startswith("unnamed:"):
                name = f"Coluna_{idx}"
            if name in seen:
                seen[name] += 1
                name = f"{name}_{seen[name]}"
            else:
                seen[name] = 1
            normalized.append(name)
        return normalized

    def _read_dataframe(
        self,
        sheet_name: str,
        header_row: int,
        data_start_row: Optional[int] = None,
        data_end_row: Optional[int] = None,
        col_start: Optional[int] = None,
        col_end: Optional[int] = None,
    ) -> pd.DataFrame:
        # header_row is 1-based Excel row number
        skiprows: Optional[Sequence[int]] = None
        if header_row > 1:
            skiprows = list(range(header_row - 1))

        kwargs: Dict[str, object] = {"sheet_name": sheet_name, "dtype": object, "header": 0}
        if skiprows:
            kwargs["skiprows"] = skiprows

        if data_end_row is not None:
            # include header row in nrows calculation
            kwargs["nrows"] = max(data_end_row - (header_row - 1), 0)

        df = pd.read_excel(self.path, **kwargs)
        df = df.dropna(how="all")

        if col_start is not None or col_end is not None:
            start_idx = (col_start or 1) - 1
            end_idx = col_end if col_end else None
            df = df.iloc[:, start_idx:end_idx]

        df.columns = self._normalize_columns(df.columns)

        first_data_row = header_row + 1
        if data_start_row and data_start_row > first_data_row:
            drop_count = data_start_row - first_data_row
            df = df.iloc[drop_count:]

        df = df.reset_index(drop=True)
        return df

    def _normalize_cell(self, value: object) -> object:
        """Normalize pandas cell values for DB insertion."""
        if pd.isna(value):
            return None
        if isinstance(value, pd.Timestamp):
            return value.to_pydatetime()
        return value

    def load_sheet_preview(
        self,
        sheet_name: str,
        header_row: int,
        data_start_row: Optional[int] = None,
        data_end_row: Optional[int] = None,
        col_start: Optional[int] = None,
        col_end: Optional[int] = None,
    ) -> SheetPreview:
        df = self._read_dataframe(
            sheet_name,
            header_row=header_row,
            data_start_row=data_start_row,
            data_end_row=data_end_row,
            col_start=col_start,
            col_end=col_end,
        )
        columns = list(df.columns)
        preview = SheetPreview(name=sheet_name, columns=columns, sample=df.head(30), header_row=header_row)
        return preview

    def read_records(
        self,
        sheet_name: str,
        column_mapping: Dict[str, str],
        header_row: int,
        start_row: Optional[int] = None,
        end_row: Optional[int] = None,
        col_start: Optional[int] = None,
        col_end: Optional[int] = None,
    ) -> List[Dict[str, object]]:
        df = self._read_dataframe(
            sheet_name,
            header_row,
            data_start_row=start_row,
            data_end_row=end_row,
            col_start=col_start,
            col_end=col_end,
        )
        records: List[Dict[str, object]] = []
        for _, row in df.iterrows():
            mapped_row = {
                db_col: self._normalize_cell(row[sheet_col])
                for sheet_col, db_col in column_mapping.items()
                if sheet_col in row
            }
            records.append(mapped_row)
        return records
