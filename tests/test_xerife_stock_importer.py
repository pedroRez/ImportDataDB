from __future__ import annotations

import unittest

import pandas as pd

from src.core.profiles import ImportProfile
from src.core.xerife_stock import XerifeStockImporter


class FakeReader:
    def __init__(self, dataframe: pd.DataFrame) -> None:
        self.dataframe = dataframe

    def _read_dataframe(self, *_args, **_kwargs) -> pd.DataFrame:
        return self.dataframe.copy()


def build_profile() -> ImportProfile:
    return ImportProfile(
        id="arcos_xerife_stock",
        name="Arcos",
        target_type="xerife_stock",
        sheet_name="Inventario",
        header_row=5,
        data_start_row=6,
        filial_id=1,
        usuario_id=1,
        source_key_strategy=["COD PRODUTO", "IDPRD"],
        field_map={
            "nome": "NOME FANTASIA",
            "tipo": "FAMILIA N1",
            "grupo_produto": "FAMILIA N2",
            "source_unit": "UNIDADE CONTROLE",
            "estoque_atual": "QUANTIDADE",
            "valor_medio_unitario": "CUSTO MEDIO",
        },
        defaults={"estoque_minimo": 0, "usuario_id": 1},
        unit_rules=[
            {"mode": "direct", "values": ["UN", "KG", "L", "M", "M2", "M3", "PAR", "CT"]},
            {"mode": "alias", "map": {"LT": "L", "MT": "M"}},
            {
                "mode": "package_regex",
                "units": ["BD", "BD20", "CX", "FR", "GL", "LA", "PCT", "TAM"],
                "pattern": r"(?P<qty>\d+(?:[.,]\d+)?)\s*(?P<unit>KG|G|GR|L|LT|ML|M3|M2|M|MT)\b",
                "content_aliases": {"GR": "G", "LT": "L", "MT": "M"},
            },
        ],
    )


class XerifeStockImporterTests(unittest.TestCase):
    def test_resolves_direct_alias_and_package_units(self) -> None:
        dataframe = pd.DataFrame(
            [
                {
                    "COD PRODUTO": "A1",
                    "IDPRD": "ALT-A1",
                    "NOME FANTASIA": "Adesivo estrutural",
                    "FAMILIA N1": "Quimicos",
                    "FAMILIA N2": "Adesivos",
                    "UNIDADE CONTROLE": "LT",
                    "QUANTIDADE": "2",
                    "CUSTO MEDIO": "10,50",
                },
                {
                    "COD PRODUTO": "B1",
                    "IDPRD": "",
                    "NOME FANTASIA": "Perfil metalico",
                    "FAMILIA N1": "Ferragens",
                    "FAMILIA N2": "Perfis",
                    "UNIDADE CONTROLE": "MT",
                    "QUANTIDADE": "8",
                    "CUSTO MEDIO": "22,10",
                },
                {
                    "COD PRODUTO": "C1",
                    "IDPRD": "",
                    "NOME FANTASIA": "Selador premium 3,60L",
                    "FAMILIA N1": "Tintas",
                    "FAMILIA N2": "Seladores",
                    "UNIDADE CONTROLE": "GL",
                    "QUANTIDADE": "4",
                    "CUSTO MEDIO": "91,30",
                },
            ]
        )

        result = XerifeStockImporter(FakeReader(dataframe), build_profile()).validate()

        self.assertEqual(result.importable_rows, 3)
        self.assertFalse(result.issues)
        self.assertEqual(result.prepared_items[0]["unidade_informada"], "L")
        self.assertEqual(result.prepared_items[1]["unidade_informada"], "M")
        self.assertEqual(result.prepared_items[2]["unidade_informada"], "UN")
        self.assertEqual(result.prepared_items[2]["conteudo_por_unidade"], 3.6)
        self.assertEqual(result.prepared_items[2]["unidade_conteudo"], "L")

    def test_blocks_group_dimension_conflict(self) -> None:
        dataframe = pd.DataFrame(
            [
                {
                    "COD PRODUTO": "A1",
                    "IDPRD": "",
                    "NOME FANTASIA": "Produto massa",
                    "FAMILIA N1": "Categoria",
                    "FAMILIA N2": "Grupo misto",
                    "UNIDADE CONTROLE": "KG",
                    "QUANTIDADE": "1",
                    "CUSTO MEDIO": "10",
                },
                {
                    "COD PRODUTO": "B1",
                    "IDPRD": "",
                    "NOME FANTASIA": "Produto comprimento",
                    "FAMILIA N1": "Categoria",
                    "FAMILIA N2": "Grupo misto",
                    "UNIDADE CONTROLE": "M",
                    "QUANTIDADE": "1",
                    "CUSTO MEDIO": "10",
                },
            ]
        )

        result = XerifeStockImporter(FakeReader(dataframe), build_profile()).validate()

        self.assertEqual(result.importable_rows, 1)
        self.assertEqual(len(result.issues), 1)
        self.assertEqual(result.issues[0].code, "group_dimension_conflict")


if __name__ == "__main__":
    unittest.main()
