from __future__ import annotations

import traceback
from pathlib import Path
from typing import Dict, List

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from importdatadb.core.mapping import MappingSelection
from importdatadb.db.provider import ColumnInfo, DatabaseProvider
from importdatadb.excel.reader import ExcelReader


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("ImportDataDB")
        self.resize(1200, 800)

        self.database = DatabaseProvider()
        self.excel_reader: ExcelReader | None = None
        self.table_columns: List[ColumnInfo] = []

        self._build_menu()
        self._build_layout()

    def _build_menu(self) -> None:
        open_action = QAction("Abrir Excel", self)
        open_action.triggered.connect(self._choose_excel)
        menubar = self.menuBar()
        file_menu = menubar.addMenu("Arquivo")
        file_menu.addAction(open_action)

    def _build_layout(self) -> None:
        central = QWidget()
        central_layout = QVBoxLayout(central)

        splitter = QSplitter(Qt.Horizontal)

        splitter.addWidget(self._build_excel_panel())
        splitter.addWidget(self._build_database_panel())
        splitter.addWidget(self._build_mapping_panel())
        splitter.addWidget(self._build_preview_panel())

        central_layout.addWidget(splitter)
        self.setCentralWidget(central)

    # Excel panel
    def _build_excel_panel(self) -> QWidget:
        panel = QGroupBox("Excel")
        layout = QVBoxLayout(panel)

        self.excel_path_label = QLabel("Nenhum arquivo selecionado")
        layout.addWidget(self.excel_path_label)

        self.sheet_list = QListWidget()
        self.sheet_list.itemSelectionChanged.connect(self._on_sheet_selected)
        layout.addWidget(QLabel("Abas"))
        layout.addWidget(self.sheet_list)

        header_layout = QHBoxLayout()
        self.header_row_spin = QSpinBox()
        self.header_row_spin.setMinimum(1)
        self.header_row_spin.setValue(1)
        header_layout.addWidget(QLabel("Linha do cabeçalho"))
        header_layout.addWidget(self.header_row_spin)
        layout.addLayout(header_layout)

        range_layout = QHBoxLayout()
        self.start_row_spin = QSpinBox()
        self.start_row_spin.setMinimum(1)
        self.start_row_spin.setValue(2)
        self.end_row_spin = QSpinBox()
        self.end_row_spin.setMinimum(0)
        self.end_row_spin.setValue(0)
        range_layout.addWidget(QLabel("Início"))
        range_layout.addWidget(self.start_row_spin)
        range_layout.addWidget(QLabel("Fim (0 = até o fim)"))
        range_layout.addWidget(self.end_row_spin)
        layout.addLayout(range_layout)

        self.refresh_sheet_btn = QPushButton("Pré-visualizar")
        self.refresh_sheet_btn.clicked.connect(self._refresh_sheet_preview)
        layout.addWidget(self.refresh_sheet_btn)

        self.sheet_preview = QTextEdit()
        self.sheet_preview.setReadOnly(True)
        layout.addWidget(self.sheet_preview)
        return panel

    # Database panel
    def _build_database_panel(self) -> QWidget:
        panel = QGroupBox("Banco de Dados")
        layout = QVBoxLayout(panel)

        grid = QGridLayout()
        self.host_edit = QLineEdit("localhost")
        self.port_edit = QLineEdit("5432")
        self.db_edit = QLineEdit()
        self.user_edit = QLineEdit()
        self.pwd_edit = QLineEdit()
        self.pwd_edit.setEchoMode(QLineEdit.Password)

        grid.addWidget(QLabel("Host"), 0, 0)
        grid.addWidget(self.host_edit, 0, 1)
        grid.addWidget(QLabel("Porta"), 1, 0)
        grid.addWidget(self.port_edit, 1, 1)
        grid.addWidget(QLabel("Database"), 2, 0)
        grid.addWidget(self.db_edit, 2, 1)
        grid.addWidget(QLabel("Usuário"), 3, 0)
        grid.addWidget(self.user_edit, 3, 1)
        grid.addWidget(QLabel("Senha"), 4, 0)
        grid.addWidget(self.pwd_edit, 4, 1)

        layout.addLayout(grid)

        self.connect_btn = QPushButton("Conectar")
        self.connect_btn.clicked.connect(self._connect_db)
        layout.addWidget(self.connect_btn)

        self.table_list = QListWidget()
        self.table_list.itemSelectionChanged.connect(self._on_table_selected)
        layout.addWidget(QLabel("Tabelas"))
        layout.addWidget(self.table_list)

        self.columns_list = QListWidget()
        layout.addWidget(QLabel("Colunas da tabela"))
        layout.addWidget(self.columns_list)
        return panel

    # Mapping panel
    def _build_mapping_panel(self) -> QWidget:
        panel = QGroupBox("Mapeamento")
        layout = QVBoxLayout(panel)

        lists_layout = QHBoxLayout()
        self.sheet_columns_list = QListWidget()
        self.table_columns_list = QListWidget()
        lists_layout.addWidget(self.sheet_columns_list)
        lists_layout.addWidget(self.table_columns_list)
        layout.addLayout(lists_layout)

        map_buttons = QHBoxLayout()
        self.add_mapping_btn = QPushButton("Mapear →")
        self.add_mapping_btn.clicked.connect(self._add_mapping)
        map_buttons.addWidget(self.add_mapping_btn)
        layout.addLayout(map_buttons)

        self.mapping_table = QTableWidget(0, 2)
        self.mapping_table.setHorizontalHeaderLabels(["Coluna Excel", "Coluna Tabela"])
        layout.addWidget(self.mapping_table)

        operation_layout = QHBoxLayout()
        self.insert_radio = QRadioButton("INSERT")
        self.insert_radio.setChecked(True)
        self.update_radio = QRadioButton("UPDATE")
        operation_layout.addWidget(self.insert_radio)
        operation_layout.addWidget(self.update_radio)
        layout.addLayout(operation_layout)

        join_layout = QHBoxLayout()
        join_layout.addWidget(QLabel("Coluna de junção (UPDATE)"))
        self.join_combo = QComboBox()
        join_layout.addWidget(self.join_combo)
        layout.addLayout(join_layout)

        self.generate_sql_btn = QPushButton("Gerar pré-visualização")
        self.generate_sql_btn.clicked.connect(self._generate_preview)
        layout.addWidget(self.generate_sql_btn)

        self.execute_btn = QPushButton("Executar")
        self.execute_btn.clicked.connect(self._execute)
        layout.addWidget(self.execute_btn)

        return panel

    # Preview panel
    def _build_preview_panel(self) -> QWidget:
        panel = QGroupBox("Pré-visualização")
        layout = QVBoxLayout(panel)
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        layout.addWidget(self.preview_text)
        return panel

    # Actions
    def _choose_excel(self) -> None:
        file_name, _ = QFileDialog.getOpenFileName(self, "Selecione o arquivo Excel", str(Path.home()), "Planilhas (*.xlsx)")
        if not file_name:
            return
        self.excel_path_label.setText(file_name)
        try:
            self.excel_reader = ExcelReader(file_name)
            self.sheet_list.clear()
            for name in self.excel_reader.sheet_names():
                self.sheet_list.addItem(name)
            self.sheet_preview.setPlainText("")
            self.sheet_columns_list.clear()
        except Exception as exc:  # noqa: BLE001
            self._show_error("Erro ao abrir Excel", exc)

    def _connect_db(self) -> None:
        try:
            host = self.host_edit.text().strip()
            port = int(self.port_edit.text())
            database = self.db_edit.text().strip()
            user = self.user_edit.text().strip()
            pwd = self.pwd_edit.text()
            self.database.connect(host, port, database, user, pwd)
            self._load_tables()
            QMessageBox.information(self, "Banco", "Conexão realizada com sucesso")
        except Exception as exc:  # noqa: BLE001
            self._show_error("Erro ao conectar", exc)

    def _load_tables(self) -> None:
        self.table_list.clear()
        for table in self.database.list_tables():
            self.table_list.addItem(table)

    def _on_sheet_selected(self) -> None:
        self._refresh_sheet_preview()

    def _refresh_sheet_preview(self) -> None:
        if not self.excel_reader:
            return
        items = self.sheet_list.selectedItems()
        if not items:
            return
        sheet_name = items[0].text()
        header = self.header_row_spin.value() - 1
        start = self.start_row_spin.value()
        end = self.end_row_spin.value() or None
        preview = self.excel_reader.load_sheet_preview(sheet_name, header_row=header, start_row=start, end_row=end)
        self.sheet_preview.setPlainText(preview.sample.to_string())
        self.sheet_columns_list.clear()
        for col in preview.columns:
            self.sheet_columns_list.addItem(col)

    def _on_table_selected(self) -> None:
        items = self.table_list.selectedItems()
        if not items:
            return
        table = items[0].text()
        self.table_columns = self.database.get_columns(table)
        self.columns_list.clear()
        self.table_columns_list.clear()
        self.join_combo.clear()
        for col in self.table_columns:
            label = f"{col.name} ({col.type})"
            if col.primary_key:
                label += " [PK]"
            self.columns_list.addItem(label)
            self.table_columns_list.addItem(col.name)
            self.join_combo.addItem(col.name)

    def _add_mapping(self) -> None:
        sheet_items = self.sheet_columns_list.selectedItems()
        table_items = self.table_columns_list.selectedItems()
        if not sheet_items or not table_items:
            return
        sheet_col = sheet_items[0].text()
        table_col = table_items[0].text()
        row = self.mapping_table.rowCount()
        self.mapping_table.insertRow(row)
        self.mapping_table.setItem(row, 0, QTableWidgetItem(sheet_col))
        self.mapping_table.setItem(row, 1, QTableWidgetItem(table_col))

    def _collect_mapping(self) -> MappingSelection | None:
        sheet_items = self.sheet_list.selectedItems()
        table_items = self.table_list.selectedItems()
        if not sheet_items or not table_items:
            QMessageBox.warning(self, "Mapeamento", "Selecione uma aba e uma tabela")
            return None
        mapping: Dict[str, str] = {}
        for row in range(self.mapping_table.rowCount()):
            sheet_col_item = self.mapping_table.item(row, 0)
            table_col_item = self.mapping_table.item(row, 1)
            if sheet_col_item and table_col_item:
                mapping[sheet_col_item.text()] = table_col_item.text()
        if not mapping:
            QMessageBox.warning(self, "Mapeamento", "Adicione ao menos um mapeamento de coluna")
            return None
        return MappingSelection(
            sheet_name=sheet_items[0].text(),
            table_name=table_items[0].text(),
            header_row=self.header_row_spin.value() - 1,
            start_row=self.start_row_spin.value(),
            end_row=self.end_row_spin.value() or None,
            column_mapping=mapping,
            operation="UPDATE" if self.update_radio.isChecked() else "INSERT",
            join_column=self.join_combo.currentText() if self.update_radio.isChecked() else None,
        )

    def _generate_preview(self) -> None:
        selection = self._collect_mapping()
        if not selection or not self.excel_reader:
            return
        try:
            preview = self.excel_reader.load_sheet_preview(
                selection.sheet_name,
                header_row=selection.header_row,
                start_row=selection.start_row,
                end_row=selection.end_row,
            )
            sql_example = self._build_sql_example(selection)
            text = ["Pré-visualização de dados:", preview.sample.head().to_string(), "\nSQL estimado:", sql_example]
            self.preview_text.setPlainText("\n".join(text))
        except Exception as exc:  # noqa: BLE001
            self._show_error("Erro ao pré-visualizar", exc)

    def _build_sql_example(self, selection: MappingSelection) -> str:
        cols = list(selection.column_mapping.values())
        if selection.operation == "INSERT":
            placeholders = ", ".join(f":{c}" for c in cols)
            return f"INSERT INTO {selection.table_name} ({', '.join(cols)}) VALUES ({placeholders});"
        else:
            set_clause = ", ".join(f"{c} = :{c}" for c in cols if c != selection.join_column)
            return f"UPDATE {selection.table_name} SET {set_clause} WHERE {selection.join_column} = :{selection.join_column};"

    def _execute(self) -> None:
        selection = self._collect_mapping()
        if not selection or not self.excel_reader:
            return
        try:
            records = self.excel_reader.read_records(
                selection.sheet_name,
                selection.column_mapping,
                header_row=selection.header_row,
                start_row=selection.start_row,
                end_row=selection.end_row,
            )
            affected = 0
            if selection.operation == "INSERT":
                affected = self.database.execute_insert(selection.table_name, records)
            else:
                if not selection.join_column:
                    QMessageBox.warning(self, "UPDATE", "Selecione uma coluna de junção")
                    return
                affected = self.database.execute_update(selection.table_name, records, selection.join_column)
            QMessageBox.information(self, "Importação", f"Registros processados: {affected}")
        except Exception as exc:  # noqa: BLE001
            self._show_error("Erro na importação", exc)

    def _show_error(self, title: str, exc: Exception) -> None:
        traceback.print_exc()
        QMessageBox.critical(self, title, f"{exc}\n\n{traceback.format_exc()}")


def main() -> int:
    app = QApplication([])
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
