from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional

from importdatadb.db.provider import ColumnInfo


@dataclass
class MappingSelection:
    sheet_name: str
    table_name: str
    header_row: int
    start_row: Optional[int]
    end_row: Optional[int]
    column_mapping: Dict[str, str]
    operation: str
    join_column: Optional[str]

    def mapped_table_columns(self, columns: List[ColumnInfo]) -> List[str]:
        mapped = []
        for col in columns:
            if col.name in self.column_mapping.values():
                mapped.append(col.name)
        return mapped
