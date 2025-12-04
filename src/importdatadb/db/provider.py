from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional

from sqlalchemy import create_engine, inspect, text
from sqlalchemy.engine import Engine


@dataclass
class ColumnInfo:
    name: str
    type: str
    nullable: bool
    primary_key: bool


class DatabaseProvider:
    def __init__(self) -> None:
        self.engine: Optional[Engine] = None

    def connect(self, host: str, port: int, database: str, user: str, password: str) -> None:
        url = f"postgresql+psycopg2://{user}:{password}@{host}:{port}/{database}"
        self.engine = create_engine(url, future=True)
        # quick test connection
        with self.engine.connect() as conn:
            conn.execute(text("SELECT 1"))

    def list_tables(self, schema: str = "public") -> List[str]:
        if not self.engine:
            return []
        inspector = inspect(self.engine)
        return inspector.get_table_names(schema=schema)

    def get_columns(self, table_name: str, schema: str = "public") -> List[ColumnInfo]:
        if not self.engine:
            return []
        inspector = inspect(self.engine)
        pk = set(inspector.get_pk_constraint(table_name, schema=schema).get("constrained_columns", []))
        columns = []
        for col in inspector.get_columns(table_name, schema=schema):
            columns.append(
                ColumnInfo(
                    name=col["name"],
                    type=str(col.get("type", "")),
                    nullable=bool(col.get("nullable", True)),
                    primary_key=col["name"] in pk,
                )
            )
        return columns

    def execute_insert(
        self,
        table: str,
        records: List[Dict[str, object]],
        schema: str = "public",
        autogenerate_pk: bool = False,
        primary_key: Optional[str] = None,
    ) -> int:
        if not self.engine or not records:
            return 0

        records_to_use = records
        if autogenerate_pk and primary_key:
            records_to_use = [{k: v for k, v in record.items() if k != primary_key} for record in records]

        with self.engine.begin() as conn:
            placeholders = ", ".join(f":{col}" for col in records_to_use[0].keys())
            columns = ", ".join(records_to_use[0].keys())
            stmt = text(f"INSERT INTO {schema}.{table} ({columns}) VALUES ({placeholders})")
            conn.execute(stmt, records_to_use)
        return len(records_to_use)

    def execute_update(self, table: str, records: List[Dict[str, object]], join_column: str, schema: str = "public") -> int:
        if not self.engine or not records:
            return 0
        with self.engine.begin() as conn:
            set_clause = ", ".join(f"{col} = :{col}" for col in records[0].keys() if col != join_column)
            stmt = text(f"UPDATE {schema}.{table} SET {set_clause} WHERE {join_column} = :{join_column}")
            conn.execute(stmt, records)
        return len(records)
