from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional


from src.db.provider import ColumnInfo


@dataclass
class ForeignKeyLookup:
    target_column: str
    excel_column: str
    foreign_table: str
    foreign_id_column: str
    foreign_label_column: str


@dataclass
class MappingSelection:
    sheet_name: str
    table_name: str
    header_row: int
    start_column: Optional[int]
    end_column: Optional[int]
    column_mapping: List[tuple[str, str]]
    default_values: Dict[str, object]
    operation: str
    join_column: Optional[str]
    primary_key: Optional[str]
    autogenerate_pk: bool
    fk_lookups: List[ForeignKeyLookup]
    remove_duplicate_rows: bool
    duplicate_check_column: Optional[str]

    def mapped_table_columns(self, columns: List[ColumnInfo]) -> List[str]:
        mapped: List[str] = []
        mapped_cols = {table_col for _, table_col in self.column_mapping}
        for col in columns:
            if col.name in mapped_cols:
                mapped.append(col.name)
        return mapped
