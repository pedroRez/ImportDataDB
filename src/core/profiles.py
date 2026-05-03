from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional


ROOT_DIR = Path(__file__).resolve().parents[2]
PROFILE_DIR = ROOT_DIR / "profiles"


@dataclass
class ImportProfile:
    id: str
    name: str
    target_type: str
    sheet_name: str
    header_row: int
    data_start_row: Optional[int] = None
    data_end_row: Optional[int] = None
    col_start: Optional[int] = None
    col_end: Optional[int] = None
    filial_id: Optional[int] = None
    usuario_id: Optional[int] = None
    source_key_strategy: List[str] = field(default_factory=list)
    field_map: Dict[str, Any] = field(default_factory=dict)
    defaults: Dict[str, Any] = field(default_factory=dict)
    filters: Dict[str, Any] = field(default_factory=dict)
    unit_rules: List[Dict[str, Any]] = field(default_factory=list)
    table_name: Optional[str] = None
    operation: Optional[str] = None
    description: Optional[str] = None

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "ImportProfile":
        return cls(
            id=str(data["id"]).strip(),
            name=str(data["name"]).strip(),
            target_type=str(data["target_type"]).strip(),
            sheet_name=str(data["sheet_name"]).strip(),
            header_row=int(data["header_row"]),
            data_start_row=_optional_int(data.get("data_start_row")),
            data_end_row=_optional_int(data.get("data_end_row")),
            col_start=_optional_int(data.get("col_start")),
            col_end=_optional_int(data.get("col_end")),
            filial_id=_optional_int(data.get("filial_id")),
            usuario_id=_optional_int(data.get("usuario_id")),
            source_key_strategy=[str(value).strip() for value in data.get("source_key_strategy", []) if str(value).strip()],
            field_map=dict(data.get("field_map", {})),
            defaults=dict(data.get("defaults", {})),
            filters=dict(data.get("filters", {})),
            unit_rules=list(data.get("unit_rules", [])),
            table_name=_optional_str(data.get("table_name")),
            operation=_optional_str(data.get("operation")),
            description=_optional_str(data.get("description")),
        )

    def to_dict(self) -> Dict[str, Any]:
        data = asdict(self)
        return {key: value for key, value in data.items() if value is not None}

    @property
    def summary(self) -> str:
        parts = [self.name, f"alvo={self.target_type}", f"aba={self.sheet_name}", f"cabecalho={self.header_row}"]
        if self.data_start_row:
            parts.append(f"dados={self.data_start_row}-{self.data_end_row or 'fim'}")
        if self.col_start:
            parts.append(f"colunas={self.col_start}-{self.col_end or 'fim'}")
        if self.filial_id is not None:
            parts.append(f"filial={self.filial_id}")
        return " | ".join(parts)


def _optional_int(value: Any) -> Optional[int]:
    if value is None or value == "":
        return None
    return int(value)


def _optional_str(value: Any) -> Optional[str]:
    if value is None:
        return None
    normalized = str(value).strip()
    return normalized or None


def ensure_profile_dir() -> Path:
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    return PROFILE_DIR


def list_profiles() -> List[ImportProfile]:
    directory = ensure_profile_dir()
    profiles: List[ImportProfile] = []
    for path in sorted(directory.glob("*.json")):
        with path.open("r", encoding="utf-8") as handle:
            profiles.append(ImportProfile.from_dict(json.load(handle)))
    return sorted(profiles, key=lambda item: (item.name.lower(), item.id.lower()))


def load_profile(profile_id: str) -> ImportProfile:
    path = ensure_profile_dir() / f"{profile_id}.json"
    with path.open("r", encoding="utf-8") as handle:
        return ImportProfile.from_dict(json.load(handle))


def save_profile(profile: ImportProfile) -> Path:
    if not profile.id.strip():
        raise ValueError("Profile id is required.")
    ensure_profile_dir()
    path = PROFILE_DIR / f"{profile.id}.json"
    with path.open("w", encoding="utf-8") as handle:
        json.dump(profile.to_dict(), handle, indent=2, ensure_ascii=True)
        handle.write("\n")
    return path
