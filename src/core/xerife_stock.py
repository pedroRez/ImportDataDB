from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence

import pandas as pd

from src.core.profiles import ImportProfile
from src.excel.reader import ExcelReader


UNIT_DIMENSIONS = {
    "UN": "UNIDADE",
    "PAR": "UNIDADE",
    "CT": "UNIDADE",
    "KG": "MASSA",
    "G": "MASSA",
    "L": "VOLUME",
    "ML": "VOLUME",
    "M": "COMPRIMENTO",
    "M2": "AREA",
    "M3": "VOLUME",
}

DEFAULT_REPORT_UNIT_BY_DIMENSION = {
    "UNIDADE": "UN",
    "MASSA": "KG",
    "VOLUME": "L",
    "COMPRIMENTO": "M",
    "AREA": "M2",
}


@dataclass
class ValidationIssue:
    row_number: int
    code: str
    message: str


@dataclass
class XerifeValidationResult:
    profile: ImportProfile
    total_rows: int
    importable_rows: int
    prepared_items: List[Dict[str, Any]] = field(default_factory=list)
    issues: List[ValidationIssue] = field(default_factory=list)
    ignored_issues: List[ValidationIssue] = field(default_factory=list)
    skipped_rows: int = 0

    @property
    def can_import(self) -> bool:
        return self.importable_rows > 0 and not self.issues

    def preview_text(self, max_rows: int = 12) -> str:
        lines = [
            f"Perfil: {self.profile.name}",
            f"Alvo: {self.profile.target_type}",
            f"Linhas lidas: {self.total_rows}",
            f"Linhas importaveis: {self.importable_rows}",
            f"Linhas descartadas: {self.skipped_rows}",
            f"Erros bloqueantes: {len(self.issues)}",
            f"Descartes por regra: {len(self.ignored_issues)}",
        ]
        if self.prepared_items:
            sample = pd.DataFrame(self.prepared_items[:max_rows])
            lines.extend(["", "Amostra preparada:", sample.to_string(index=False)])
        if self.issues:
            lines.extend(["", "Primeiros erros:"])
            for issue in self.issues[:10]:
                lines.append(f"Linha {issue.row_number}: [{issue.code}] {issue.message}")
            remaining = len(self.issues) - min(len(self.issues), 10)
            if remaining > 0:
                lines.append(f"...mais {remaining} erro(s).")
        if self.ignored_issues:
            lines.extend(["", "Primeiras linhas descartadas:"])
            for issue in self.ignored_issues[:10]:
                lines.append(f"Linha {issue.row_number}: [{issue.code}] {issue.message}")
            remaining = len(self.ignored_issues) - min(len(self.ignored_issues), 10)
            if remaining > 0:
                lines.append(f"...mais {remaining} descarte(s).")
        return "\n".join(lines)


class XerifeStockImporter:
    def __init__(self, reader: ExcelReader, profile: ImportProfile) -> None:
        self.reader = reader
        self.profile = profile
        self._unit_rules = self._parse_unit_rules(profile.unit_rules)
        self._skip_issue_codes = {
            str(code).strip()
            for code in self.profile.filters.get("skip_issue_codes", [])
            if str(code).strip()
        }

    def validate(self) -> XerifeValidationResult:
        data_start_row = self.profile.data_start_row or (self.profile.header_row + 1)
        dataframe = self.reader._read_dataframe(
            self.profile.sheet_name,
            header_row=self.profile.header_row,
            data_start_row=data_start_row,
            data_end_row=self.profile.data_end_row,
            col_start=self.profile.col_start,
            col_end=self.profile.col_end,
        )

        result = XerifeValidationResult(
            profile=self.profile,
            total_rows=len(dataframe.index),
            importable_rows=0,
        )
        group_dimensions: Dict[str, tuple[str, int]] = {}

        for row_index, (_, series) in enumerate(dataframe.iterrows()):
            row_number = data_start_row + row_index
            prepared = self._prepare_row(series, row_number)
            if prepared is None:
                result.skipped_rows += 1
                continue
            if isinstance(prepared, ValidationIssue):
                if self._should_skip_issue(prepared.code):
                    result.ignored_issues.append(prepared)
                    result.skipped_rows += 1
                else:
                    result.issues.append(prepared)
                continue

            prepared["grupo_produto"] = self._resolve_group_name(
                prepared["grupo_produto"],
                prepared["grupo_produto_dimensao"],
                prepared["grupo_produto_unidade_relatorio"],
            )
            group_name = prepared["grupo_produto"]
            group_dimension = prepared["grupo_produto_dimensao"]
            if group_name in group_dimensions:
                previous_dimension, previous_row = group_dimensions[group_name]
                if previous_dimension != group_dimension:
                    issue = ValidationIssue(
                        row_number=row_number,
                        code="group_dimension_conflict",
                        message=(
                            f"Grupo '{group_name}' tambem apareceu na linha {previous_row} "
                            f"com dimensao '{previous_dimension}', mas esta linha exige '{group_dimension}'."
                        ),
                    )
                    if self._should_skip_issue(issue.code):
                        result.ignored_issues.append(issue)
                        result.skipped_rows += 1
                    else:
                        result.issues.append(issue)
                    continue
            else:
                group_dimensions[group_name] = (group_dimension, row_number)

            result.prepared_items.append(prepared)

        result.importable_rows = len(result.prepared_items)
        return result

    def _prepare_row(self, row: pd.Series, row_number: int) -> Dict[str, Any] | ValidationIssue | None:
        key_candidates = [self._string_value(row.get(column)) for column in self.profile.source_key_strategy]
        key_candidates = [value for value in key_candidates if value]
        if not key_candidates:
            return ValidationIssue(
                row_number=row_number,
                code="missing_key",
                message="Linha sem chave de origem em nenhuma coluna configurada.",
            )

        nome = self._source_value(row, "nome")
        if not nome:
            return ValidationIssue(
                row_number=row_number,
                code="missing_name",
                message="Campo de descricao/nome vazio.",
            )

        grupo_produto = self._source_value(row, "grupo_produto")
        if not grupo_produto:
            return ValidationIssue(
                row_number=row_number,
                code="missing_group",
                message="Grupo do produto vazio.",
            )

        tipo = self._source_value(row, "tipo")
        raw_unit = self._source_value(row, "source_unit")
        if not raw_unit:
            return ValidationIssue(
                row_number=row_number,
                code="missing_unit",
                message="Sigla de unidade vazia.",
            )

        unit_resolution = self._resolve_unit(
            raw_unit,
            nome,
            tipo=tipo,
            grupo_produto=grupo_produto,
        )
        if isinstance(unit_resolution, str):
            return ValidationIssue(
                row_number=row_number,
                code="invalid_unit",
                message=unit_resolution,
            )

        estoque_atual = self._numeric_value(row, "estoque_atual")
        valor_medio = self._numeric_value(row, "valor_medio_unitario")
        estoque_minimo = self._default_numeric("estoque_minimo")

        item = {
            "row_number": row_number,
            "nome": nome,
            "tipo": tipo,
            "codigo_peca": key_candidates[0],
            "source_lookup_candidates": key_candidates,
            "grupo_produto": grupo_produto,
            "grupo_produto_descricao": self.profile.defaults.get("grupo_produto_descricao"),
            "grupo_produto_dimensao": unit_resolution["dimension"],
            "grupo_produto_unidade_relatorio": unit_resolution["report_unit"],
            "unidade_informada": unit_resolution["base_unit"],
            "conteudo_por_unidade": unit_resolution["content_quantity"],
            "unidade_conteudo": unit_resolution["content_unit"],
            "estoque_atual": estoque_atual,
            "estoque_minimo": estoque_minimo,
            "valor_medio_unitario": valor_medio,
            "filial_id": self.profile.filial_id,
            "usuario_id": self.profile.usuario_id or self.profile.defaults.get("usuario_id"),
        }

        optional_fields = [
            "fabricante",
            "fabricante_id",
            "fabricante_aux",
            "classificacao_id",
            "classificacao_aux",
            "aplicacao",
            "aplicacao_id",
            "aplicacao_aux",
            "data_lancamento",
        ]
        for field_name in optional_fields:
            value = self._source_value(row, field_name)
            if value is None:
                value = self.profile.defaults.get(field_name)
            if value not in (None, ""):
                item[field_name] = value

        return item

    def _parse_unit_rules(self, unit_rules: Sequence[Dict[str, Any]]) -> Dict[str, Any]:
        config = {
            "direct": set(),
            "aliases": {},
            "package_rules": [],
        }
        for rule in unit_rules:
            mode = str(rule.get("mode", "")).strip().lower()
            if mode == "direct":
                config["direct"].update(self._normalize_unit(value) for value in rule.get("values", []))
            elif mode == "alias":
                config["aliases"].update(
                    {
                        self._normalize_unit(source): self._normalize_unit(target)
                        for source, target in dict(rule.get("map", {})).items()
                    }
                )
            elif mode == "package_regex":
                pattern_text = str(rule.get("pattern", "")).strip()
                if not pattern_text:
                    continue
                package_rule = {
                    "units": {
                        self._normalize_unit(value)
                        for value in rule.get("units", [])
                        if self._normalize_unit(value)
                    },
                    "pattern": re.compile(pattern_text, re.IGNORECASE),
                    "content_aliases": {
                        self._normalize_unit(source): self._normalize_unit(target)
                        for source, target in dict(rule.get("content_aliases", {})).items()
                    },
                    "group_patterns": [
                        re.compile(str(pattern), re.IGNORECASE)
                        for pattern in rule.get("group_patterns", [])
                        if str(pattern).strip()
                    ],
                    "type_patterns": [
                        re.compile(str(pattern), re.IGNORECASE)
                        for pattern in rule.get("type_patterns", [])
                        if str(pattern).strip()
                    ],
                    "base_unit": self._normalize_unit(rule.get("base_unit") or "UN"),
                    "required": bool(rule.get("required", True)),
                }
                if package_rule["units"]:
                    config["package_rules"].append(package_rule)
        return config

    def _resolve_unit(
        self,
        raw_unit: str,
        nome: str,
        *,
        tipo: Optional[str] = None,
        grupo_produto: Optional[str] = None,
    ) -> Dict[str, Any] | str:
        normalized_unit = self._normalize_unit(raw_unit)
        if normalized_unit in self._unit_rules["aliases"]:
            normalized_unit = self._unit_rules["aliases"][normalized_unit]

        last_package_error: Optional[str] = None
        for package_rule in self._unit_rules["package_rules"]:
            if normalized_unit not in package_rule["units"]:
                continue
            if not self._matches_package_rule(package_rule, tipo=tipo, grupo_produto=grupo_produto):
                continue
            package_resolution = self._resolve_package_rule(package_rule, nome)
            if isinstance(package_resolution, dict):
                return package_resolution
            if package_rule["required"]:
                return package_resolution
            last_package_error = package_resolution

        if normalized_unit in self._unit_rules["direct"]:
            dimension = self._dimension_for_unit(normalized_unit)
            return {
                "base_unit": normalized_unit,
                "content_quantity": 1,
                "content_unit": normalized_unit,
                "dimension": dimension,
                "report_unit": DEFAULT_REPORT_UNIT_BY_DIMENSION[dimension],
            }

        if last_package_error:
            return last_package_error
        return f"Sigla '{normalized_unit}' nao possui regra de importacao."

    def _source_value(self, row: pd.Series, field_name: str) -> Optional[str]:
        source_column = self.profile.field_map.get(field_name)
        if not source_column:
            default_value = self.profile.defaults.get(field_name)
            return self._string_value(default_value) if default_value is not None else None
        return self._string_value(row.get(str(source_column)))

    def _matches_package_rule(
        self,
        package_rule: Dict[str, Any],
        *,
        tipo: Optional[str],
        grupo_produto: Optional[str],
    ) -> bool:
        if package_rule["group_patterns"]:
            group_text = grupo_produto or ""
            if not any(pattern.search(group_text) for pattern in package_rule["group_patterns"]):
                return False
        if package_rule["type_patterns"]:
            type_text = tipo or ""
            if not any(pattern.search(type_text) for pattern in package_rule["type_patterns"]):
                return False
        return True

    def _resolve_package_rule(self, package_rule: Dict[str, Any], nome: str) -> Dict[str, Any] | str:
        match = package_rule["pattern"].search(nome) if package_rule["pattern"] else None
        if not match:
            return (
                "Sigla exige conteudo parseavel em NOME FANTASIA, "
                "mas o padrao nao foi encontrado."
            )
        raw_content_unit = self._normalize_unit(match.group("unit"))
        content_unit = package_rule["content_aliases"].get(raw_content_unit, raw_content_unit)
        if content_unit not in UNIT_DIMENSIONS:
            return f"Unidade de conteudo '{content_unit}' nao e suportada."
        quantity = self._parse_decimal(match.group("qty"))
        if quantity is None or quantity <= 0:
            return f"Conteudo por unidade invalido em '{nome}'."
        dimension = self._dimension_for_unit(content_unit)
        return {
            "base_unit": package_rule["base_unit"],
            "content_quantity": quantity,
            "content_unit": content_unit,
            "dimension": dimension,
            "report_unit": DEFAULT_REPORT_UNIT_BY_DIMENSION[dimension],
        }

    def _resolve_group_name(self, group_name: str, dimension: str, report_unit: str) -> str:
        strategy = str(self.profile.filters.get("group_dimension_strategy", "strict")).strip().lower()
        if strategy == "suffix_dimension":
            return f"{group_name} [{dimension}]"
        if strategy == "suffix_report_unit":
            return f"{group_name} [{report_unit}]"
        return group_name

    def _should_skip_issue(self, code: str) -> bool:
        return code in self._skip_issue_codes

    def _numeric_value(self, row: pd.Series, field_name: str) -> float:
        source_column = self.profile.field_map.get(field_name)
        if not source_column:
            return self._default_numeric(field_name)
        parsed = self._parse_decimal(row.get(str(source_column)))
        if parsed is None:
            return self._default_numeric(field_name)
        return parsed

    def _default_numeric(self, field_name: str) -> float:
        parsed = self._parse_decimal(self.profile.defaults.get(field_name))
        return parsed if parsed is not None else 0.0

    def _parse_decimal(self, value: Any) -> Optional[float]:
        if value is None:
            return None
        if isinstance(value, (int, float)):
            if isinstance(value, float) and math.isnan(value):
                return None
            return float(value)
        text = str(value).strip()
        if not text or text.lower() == "nan":
            return None
        text = text.replace(".", "").replace(",", ".") if "," in text and "." in text else text.replace(",", ".")
        try:
            return float(text)
        except ValueError:
            return None

    def _string_value(self, value: Any) -> Optional[str]:
        if value is None:
            return None
        if isinstance(value, float) and math.isnan(value):
            return None
        text = str(value).strip()
        if not text or text.lower() == "nan":
            return None
        return text

    def _normalize_unit(self, value: Any) -> str:
        return str(value or "").strip().upper()

    def _dimension_for_unit(self, unit_sigla: str) -> str:
        if unit_sigla not in UNIT_DIMENSIONS:
            raise ValueError(f"Unsupported unit dimension lookup: {unit_sigla}")
        return UNIT_DIMENSIONS[unit_sigla]
