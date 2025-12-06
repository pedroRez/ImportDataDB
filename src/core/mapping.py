from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional


from src.db.provider import ColumnInfo



@dataclass
class MappingSelection:
    sheet_name: str
    table_name: str
    header_row: int
    start_column: Optional[int]
    end_column: Optional[int]
    column_mapping: Dict[str, str]
    default_values: Dict[str, object]
    operation: str
    join_column: Optional[str]
    primary_key: Optional[str]
    autogenerate_pk: bool

    def mapped_table_columns(self, columns: List[ColumnInfo]) -> List[str]:
        mapped = []
        for col in columns:
            if col.name in self.column_mapping.values():
                mapped.append(col.name)
        return mapped
