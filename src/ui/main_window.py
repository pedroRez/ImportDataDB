from __future__ import annotations

import traceback
from datetime import date
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import pandas as pd
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
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
    QHeaderView,
    QRadioButton,
    QSpinBox,
    QSplitter,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


from src.core.mapping import ForeignKeyLookup, MappingSelection
from src.db.provider import ColumnInfo, DatabaseProvider
from src.excel.reader import ExcelReader, SheetPreview



class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("ImportDataDB")
        self.resize(1200, 800)

        self.database = DatabaseProvider()
        self.excel_reader: ExcelReader | None = None
        self.table_columns: List[ColumnInfo] = []
        self.primary_key_column: str | None = None
        self._current_header_excel_row_value = 1
        self._foreign_columns_cache: dict[str, List[ColumnInfo]] = {}
        self._pre_validation_remove_duplicates = False
        self._pre_validation_column: str | None = None
        self._pre_validation_last_result: tuple[int, int] | None = None

        self.host_edit = QLineEdit("localhost")
        self.port_edit = QLineEdit("5432")
        self.db_edit = QLineEdit()
        self.user_edit = QLineEdit()
        self.pwd_edit = QLineEdit()
        self.pwd_edit.setEchoMode(QLineEdit.Password)

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

        tabs = QTabWidget()
        tabs.addTab(self._build_excel_tab(), "Excel")
        tabs.addTab(self._build_database_tab(), "Banco de Dados")
        tabs.addTab(self._build_mapping_tab(), "Mapeamento")

        central_layout.addWidget(tabs)
        self.setCentralWidget(central)

    def _build_excel_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.addWidget(self._build_excel_panel())
        return tab

    def _build_database_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        connection_group = QGroupBox("Conexão")
        connection_layout = QHBoxLayout(connection_group)
        self.connection_btn = QPushButton("Conectar ao banco...")
        self.connection_btn.clicked.connect(self._open_connection_dialog)
        connection_layout.addWidget(self.connection_btn)

        self.connection_status_label = QLabel("Não conectado")
        self.connection_status_label.setWordWrap(True)
        connection_layout.addWidget(self.connection_status_label, 1)
        layout.addWidget(connection_group)

        layout.addWidget(self._build_database_panel())
        return tab

    def _build_mapping_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        splitter = QSplitter(Qt.Vertical)
        splitter.addWidget(self._build_mapping_panel())
        splitter.addWidget(self._build_preview_panel())
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)

        layout.addWidget(splitter)
        return tab

    # Excel panel
    def _build_excel_panel(self) -> QWidget:
        panel = QGroupBox("Excel")
        layout = QVBoxLayout(panel)

        path_layout = QHBoxLayout()
        self.excel_path_label = QLabel("Nenhum arquivo selecionado")
        self.excel_path_label.setWordWrap(True)
        path_layout.addWidget(self.excel_path_label)

        self.import_excel_btn = QPushButton("Importar Excel")
        self.import_excel_btn.clicked.connect(self._choose_excel)
        path_layout.addWidget(self.import_excel_btn)
        layout.addLayout(path_layout)

        self.sheet_list = QListWidget()
        self.sheet_list.itemSelectionChanged.connect(self._on_sheet_selected)
        layout.addWidget(QLabel("Abas"))
        layout.addWidget(self.sheet_list)

        header_layout = QHBoxLayout()
        self.header_row_spin = QSpinBox()
        self.header_row_spin.setMinimum(1)
        self.header_row_spin.setValue(1)
        header_layout.addWidget(QLabel("Linha do cabecalho"))
        header_layout.addWidget(self.header_row_spin)
        layout.addLayout(header_layout)

        range_layout = QHBoxLayout()
        self.col_start_spin = QSpinBox()
        self.col_start_spin.setMinimum(1)
        self.col_start_spin.setValue(1)
        self.col_end_spin = QSpinBox()
        self.col_end_spin.setMinimum(0)
        self.col_end_spin.setValue(0)
        range_layout.addWidget(QLabel("Coluna inicial"))
        range_layout.addWidget(self.col_start_spin)
        range_layout.addWidget(QLabel("Coluna final (0 = ate o fim)"))
        range_layout.addWidget(self.col_end_spin)
        layout.addLayout(range_layout)

        preview_buttons = QHBoxLayout()
        self.refresh_sheet_btn = QPushButton("Pre-visualizar")
        self.refresh_sheet_btn.clicked.connect(self._refresh_sheet_preview)
        preview_buttons.addWidget(self.refresh_sheet_btn)

        self.use_range_btn = QPushButton("Usar selecao p/ Inicio/Fim")
        self.use_range_btn.clicked.connect(self._apply_selection_to_range)
        preview_buttons.addWidget(self.use_range_btn)

        self.use_header_btn = QPushButton("Usar selecao como cabecalho")
        self.use_header_btn.clicked.connect(self._apply_selection_to_header)
        preview_buttons.addWidget(self.use_header_btn)
        layout.addLayout(preview_buttons)

        self.selection_info_label = QLabel(self._selection_hint_text())
        self.selection_info_label.setWordWrap(True)
        layout.addWidget(self.selection_info_label)

        self.sheet_preview_table = QTableWidget()
        self.sheet_preview_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.sheet_preview_table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.sheet_preview_table.setAlternatingRowColors(True)
        self.sheet_preview_table.itemSelectionChanged.connect(self._update_selection_info)
        layout.addWidget(self.sheet_preview_table)
        return panel

    # Database panel
    def _build_database_panel(self) -> QWidget:
        panel = QGroupBox("Estrutura do Banco")
        layout = QVBoxLayout(panel)

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
        layout.setSpacing(12)

        splitter = QSplitter(Qt.Vertical)
        splitter.setChildrenCollapsible(False)

        splitter.addWidget(self._build_mapping_section())
        splitter.addWidget(self._build_defaults_and_fk_section())
        splitter.addWidget(self._build_operations_section())

        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        splitter.setStretchFactor(2, 1)

        layout.addWidget(splitter)

        return panel

    def _build_mapping_section(self) -> QWidget:
        section = QWidget()
        section_layout = QVBoxLayout(section)
        section_layout.setSpacing(10)

        lists_layout = QHBoxLayout()
        self.sheet_columns_list = QListWidget()
        self.table_columns_list = QListWidget()
        lists_layout.addWidget(self.sheet_columns_list)
        lists_layout.addWidget(self.table_columns_list)
        section_layout.addLayout(lists_layout)

        map_buttons = QHBoxLayout()
        self.add_mapping_btn = QPushButton("Mapear →")
        self.add_mapping_btn.clicked.connect(self._add_mapping)
        map_buttons.addWidget(self.add_mapping_btn)
        self.remove_mapping_btn = QPushButton("Remover selecionado")
        self.remove_mapping_btn.clicked.connect(self._remove_mapping)
        map_buttons.addWidget(self.remove_mapping_btn)
        section_layout.addLayout(map_buttons)

        self.mapping_table = QTableWidget(0, 2)
        self.mapping_table.setHorizontalHeaderLabels(["Coluna Excel", "Coluna Tabela"])
        section_layout.addWidget(self.mapping_table)

        self.required_columns_label = QLabel("Campos obrigatorios: --")
        self.required_columns_label.setWordWrap(True)
        section_layout.addWidget(self.required_columns_label)

        return section

    def _build_defaults_and_fk_section(self) -> QWidget:
        section = QWidget()
        section_layout = QVBoxLayout(section)
        section_layout.setSpacing(12)

        defaults_group = QGroupBox("Valores padrao para colunas nao mapeadas")
        defaults_layout = QVBoxLayout(defaults_group)

        default_form = QHBoxLayout()
        default_form.addWidget(QLabel("Coluna tabela"))
        self.default_column_combo = QComboBox()
        self.default_column_combo.setMinimumContentsLength(12)
        self.default_column_combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.default_column_combo.currentTextChanged.connect(self._on_default_column_changed)
        self.default_column_combo.currentTextChanged.connect(
            lambda text: self._set_combo_tooltip(self.default_column_combo, text)
        )
        default_form.addWidget(self.default_column_combo)

        default_form.addWidget(QLabel("Valor padrao"))
        self.default_value_line = QLineEdit()
        self.default_bool_combo = QComboBox()
        self.default_bool_combo.addItem("True", True)
        self.default_bool_combo.addItem("False", False)
        self.default_date_edit = QDateEdit(QDate.currentDate())
        self.default_date_edit.setCalendarPopup(True)
        self.default_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.default_value_holder = QWidget()
        self.default_value_layout = QVBoxLayout(self.default_value_holder)
        self.default_value_layout.setContentsMargins(0, 0, 0, 0)
        self.default_value_layout.addWidget(self.default_value_line)
        default_form.addWidget(self.default_value_holder)

        self.add_default_btn = QPushButton("Definir valor")
        self.add_default_btn.clicked.connect(self._add_default_value)
        default_form.addWidget(self.add_default_btn)
        defaults_layout.addLayout(default_form)

        self.defaults_table = QTableWidget(0, 2)
        self.defaults_table.setHorizontalHeaderLabels(["Coluna Tabela", "Valor padrao"])
        self.defaults_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        defaults_layout.addWidget(self.defaults_table)

        self.remove_default_btn = QPushButton("Remover valor selecionado")
        self.remove_default_btn.clicked.connect(self._remove_default_value)
        defaults_layout.addWidget(self.remove_default_btn)

        section_layout.addWidget(defaults_group)

        fk_group = QGroupBox("Relacionamentos (FK por descrição)")
        fk_layout = QVBoxLayout(fk_group)

        fk_row1 = QHBoxLayout()
        fk_row1.addWidget(QLabel("Coluna FK na tabela"))
        self.fk_target_combo = QComboBox()
        fk_row1.addWidget(self.fk_target_combo)

        fk_row1.addWidget(QLabel("Coluna Excel (descrição)"))
        self.fk_excel_combo = QComboBox()
        fk_row1.addWidget(self.fk_excel_combo)
        fk_layout.addLayout(fk_row1)

        fk_row2 = QHBoxLayout()
        fk_row2.addWidget(QLabel("Tabela estrangeira"))
        self.fk_table_combo = QComboBox()
        self.fk_table_combo.currentTextChanged.connect(self._on_fk_table_changed)
        fk_row2.addWidget(self.fk_table_combo)

        fk_row2.addWidget(QLabel("Coluna ID"))
        self.fk_id_combo = QComboBox()
        fk_row2.addWidget(self.fk_id_combo)

        fk_row2.addWidget(QLabel("Coluna descrição"))
        self.fk_label_combo = QComboBox()
        fk_row2.addWidget(self.fk_label_combo)
        fk_layout.addLayout(fk_row2)

        self.add_fk_btn = QPushButton("Adicionar relacionamento")
        self.add_fk_btn.clicked.connect(self._add_fk_lookup)
        fk_layout.addWidget(self.add_fk_btn)

        self.fk_table = QTableWidget(0, 5)
        self.fk_table.setHorizontalHeaderLabels(
            ["Coluna Tabela", "Coluna Excel", "Tabela FK", "Coluna ID", "Coluna Descrição"]
        )
        self.fk_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        fk_layout.addWidget(self.fk_table)

        self.remove_fk_btn = QPushButton("Remover relacionamento selecionado")
        self.remove_fk_btn.clicked.connect(self._remove_fk_lookup)
        fk_layout.addWidget(self.remove_fk_btn)

        section_layout.addWidget(fk_group)
        return section

    def _build_operations_section(self) -> QWidget:
        section = QWidget()
        section_layout = QVBoxLayout(section)
        section_layout.setSpacing(10)

        self.pk_auto_checkbox = QCheckBox("PK gerada pelo banco (auto-incremento)")
        self.pk_auto_checkbox.setEnabled(False)
        self.pk_auto_checkbox.toggled.connect(self._on_pk_auto_toggled)
        section_layout.addWidget(self.pk_auto_checkbox)

        operation_layout = QHBoxLayout()
        self.insert_radio = QRadioButton("INSERT")
        self.insert_radio.setChecked(True)
        self.update_radio = QRadioButton("UPDATE")
        operation_layout.addWidget(self.insert_radio)
        operation_layout.addWidget(self.update_radio)
        section_layout.addLayout(operation_layout)

        join_layout = QHBoxLayout()
        join_layout.addWidget(QLabel("Coluna de junção (UPDATE)"))
        self.join_combo = QComboBox()
        self.join_combo.setMinimumContentsLength(12)
        self.join_combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.join_combo.currentTextChanged.connect(lambda text: self._set_combo_tooltip(self.join_combo, text))
        join_layout.addWidget(self.join_combo)
        section_layout.addLayout(join_layout)

        pre_validation_layout = QHBoxLayout()
        self.pre_validation_btn = QPushButton("Pré-validação...")
        self.pre_validation_btn.clicked.connect(self._open_pre_validation)
        pre_validation_layout.addWidget(self.pre_validation_btn)
        self.pre_validation_status = QLabel("Pré-validação: sem regras")
        self.pre_validation_status.setWordWrap(True)
        pre_validation_layout.addWidget(self.pre_validation_status)
        pre_validation_layout.addStretch()
        section_layout.addLayout(pre_validation_layout)

        self.generate_sql_btn = QPushButton("Gerar pré-visualização")
        self.generate_sql_btn.clicked.connect(self._generate_preview)
        section_layout.addWidget(self.generate_sql_btn)

        self.execute_btn = QPushButton("Executar")
        self.execute_btn.clicked.connect(self._execute)
        section_layout.addWidget(self.execute_btn)

        return section

    # Preview panel
    def _build_preview_panel(self) -> QWidget:
        panel = QGroupBox("Pré-visualização")
        layout = QVBoxLayout(panel)
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        layout.addWidget(self.preview_text)

        export_buttons = QHBoxLayout()
        self.export_csv_btn = QPushButton("Exportar CSV mapeado")
        self.export_csv_btn.clicked.connect(lambda: self._export_mapped_data("csv"))
        export_buttons.addWidget(self.export_csv_btn)
        self.export_excel_btn = QPushButton("Exportar Excel mapeado")
        self.export_excel_btn.clicked.connect(lambda: self._export_mapped_data("excel"))
        export_buttons.addWidget(self.export_excel_btn)
        layout.addLayout(export_buttons)
        return panel

    # Actions
    def _choose_excel(self) -> None:
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Selecione o arquivo Excel", str(Path.home()), "Planilhas (*.xlsx *.xlsm)"
        )
        if not file_name:
            return
        self.excel_path_label.setText(file_name)
        try:
            self.excel_reader = ExcelReader(file_name)
            self.sheet_list.clear()
            for name in self.excel_reader.sheet_names():
                self.sheet_list.addItem(name)
            self.sheet_preview_table.clear()
            self.sheet_preview_table.setRowCount(0)
            self.sheet_preview_table.setColumnCount(0)
            self.selection_info_label.setText(self._selection_hint_text())
            self.sheet_columns_list.clear()
            self._refresh_fk_excel_options()
            self._clear_pre_validation_state()
        except Exception as exc:  # noqa: BLE001
            self._show_error("Erro ao abrir Excel", exc)

    def _open_connection_dialog(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Conectar ao banco")
        layout = QVBoxLayout(dialog)

        grid = QGridLayout()
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

        buttons = QHBoxLayout()
        buttons.addStretch()
        connect_btn = QPushButton("Conectar")
        connect_btn.clicked.connect(lambda: self._connect_and_close(dialog))
        cancel_btn = QPushButton("Cancelar")
        cancel_btn.clicked.connect(dialog.reject)
        buttons.addWidget(connect_btn)
        buttons.addWidget(cancel_btn)
        layout.addLayout(buttons)

        dialog.exec()

    def _connect_and_close(self, dialog: QDialog) -> None:
        if self._connect_db():
            dialog.accept()

    def _connect_db(self) -> bool:
        try:
            host = self.host_edit.text().strip()
            port = int(self.port_edit.text())
            database = self.db_edit.text().strip()
            user = self.user_edit.text().strip()
            pwd = self.pwd_edit.text()
            self.database.connect(host, port, database, user, pwd)
            self._foreign_columns_cache = {}
            self._load_tables()
            connection_text = f"Conectado: {user or 'usuário'}@{host}:{port}/{database}"
            self.connection_status_label.setText(connection_text)
            QMessageBox.information(self, "Banco", "Conexão realizada com sucesso")
            return True
        except Exception as exc:  # noqa: BLE001
            self._show_error("Erro ao conectar", exc)
            self.connection_status_label.setText("Erro ao conectar")
            return False

    def _load_tables(self) -> None:
        self.table_list.clear()
        for table in self.database.list_tables():
            self.table_list.addItem(table)
        self._refresh_fk_table_options()

    def _on_sheet_selected(self) -> None:
        self._clear_pre_validation_state()
        self._refresh_sheet_preview()

    def _refresh_sheet_preview(self) -> None:
        if not self.excel_reader:
            return
        items = self.sheet_list.selectedItems()
        if not items:
            return
        sheet_name = items[0].text()
        col_start = self.col_start_spin.value()
        col_end = self.col_end_spin.value() or None
        if col_end is not None and col_end < col_start:
            col_end = None
            self.col_end_spin.setValue(0)
        header_excel_row = self._current_header_excel_row()
        try:
            preview = self.excel_reader.load_sheet_preview(
                sheet_name,
                header_row=header_excel_row,
                col_start=col_start,
                col_end=col_end,
            )
            if self._maybe_promote_first_data_row_as_header(preview, header_excel_row):
                return
            first_data_row = header_excel_row + 1
            self._populate_sheet_preview(preview, first_data_row)
            self.sheet_columns_list.clear()
            for col in preview.columns:
                self.sheet_columns_list.addItem(col)
            self._refresh_fk_excel_options()
        except Exception as exc:  # noqa: BLE001
            self._show_error("Erro ao pre-visualizar", exc)

    def _maybe_promote_first_data_row_as_header(self, preview: SheetPreview, current_header_row: int) -> bool:
        """If columns are placeholders and first data row seems to be the real header, move header down by 1."""
        if not preview.columns or not preview.sample.shape[0]:
            return False
        placeholders = all(str(col).startswith("Coluna_") for col in preview.columns)
        if not placeholders:
            return False
        first_row = preview.sample.iloc[0]
        values = [str(v).strip() for v in first_row.tolist() if v is not None and str(v).strip() != ""]
        if not values:
            return False
        if all(val.startswith("Coluna_") for val in values):
            return False
        new_header_row = current_header_row + 1
        if new_header_row == current_header_row:
            return False
        # Update spin and rerun preview with the next row as header.
        self.header_row_spin.setValue(new_header_row)
        self._refresh_sheet_preview()
        return True

    def _populate_sheet_preview(self, preview: SheetPreview, header_row_excel: int) -> None:
        df = preview.sample
        first_data_row = header_row_excel + 1
        self._current_header_excel_row_value = header_row_excel
        self._current_first_data_row = first_data_row

        self.sheet_preview_table.clear()
        total_rows = len(df.index)
        if total_rows <= 0:
            self.sheet_preview_table.setRowCount(0)
            self.sheet_preview_table.setColumnCount(0)
            self.selection_info_label.setText("Nenhum dado no intervalo atual. Ajuste cabecalho/colunas.")
            return

        self.sheet_preview_table.setRowCount(total_rows)
        self.sheet_preview_table.setColumnCount(len(preview.columns))
        self.sheet_preview_table.setHorizontalHeaderLabels(preview.columns)

        row_labels = [str(first_data_row + idx) for idx in range(len(df.index))]
        self.sheet_preview_table.setVerticalHeaderLabels(row_labels)

        for row_idx, (_, row) in enumerate(df.iterrows(), start=0):
            for col_idx, value in enumerate(row):
                item = QTableWidgetItem("" if value is None else str(value))
                self.sheet_preview_table.setItem(row_idx, col_idx, item)

        header = self.sheet_preview_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setStretchLastSection(True)
        self.sheet_preview_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self._update_selection_info()

    def _apply_selection_to_range(self) -> None:
        ranges = self.sheet_preview_table.selectedRanges()
        if not ranges:
            return
        current_col_start = self.col_start_spin.value()
        min_col = min(r.leftColumn() for r in ranges)
        max_col = max(r.rightColumn() for r in ranges)
        new_start = current_col_start + min_col
        new_end = current_col_start + max_col
        self.col_start_spin.setValue(new_start)
        self.col_end_spin.setValue(new_end)
        self._refresh_sheet_preview()

    def _apply_selection_to_header(self) -> None:
        ranges = self.sheet_preview_table.selectedRanges()
        if not ranges:
            return
        min_row = min(r.topRow() for r in ranges)
        header_item = self.sheet_preview_table.verticalHeaderItem(min_row)
        if header_item is not None:
            try:
                excel_row = int(header_item.text())
            except ValueError:
                excel_row = self._excel_row_from_table_row(min_row)
        else:
            excel_row = self._excel_row_from_table_row(min_row)
        target_header = max(1, excel_row)
        self.header_row_spin.setValue(target_header)
        self._refresh_sheet_preview()

    def _update_selection_info(self) -> None:
        ranges = self.sheet_preview_table.selectedRanges()
        if not ranges:
            self.selection_info_label.setText(self._selection_hint_text())
            self.selection_info_label.setToolTip("")
            return
        min_row = min(r.topRow() for r in ranges)
        max_row = max(r.bottomRow() for r in ranges)
        min_row_excel = self._excel_row_from_table_row(min_row)
        max_row_excel = self._excel_row_from_table_row(max_row)

        cols = sorted({col for r in ranges for col in range(r.leftColumn(), r.rightColumn() + 1)})
        col_start_excel = self.col_start_spin.value()
        col_labels = []
        for col in cols:
            header_item = self.sheet_preview_table.horizontalHeaderItem(col)
            label = header_item.text() if header_item else str(col_start_excel + col)
            col_number = col_start_excel + col
            col_labels.append(f"{label} (col {col_number})")
        col_brief = self._compact_columns_text(col_labels, limit=6)
        text = (
            f"Linhas Excel: {min_row_excel} - {max_row_excel} | "
            f"Colunas ({len(col_labels)}): {col_brief}"
        )
        self.selection_info_label.setText(text)
        self.selection_info_label.setToolTip(", ".join(col_labels))

    def _current_header_excel_row(self) -> int:
        """Return header row in Excel (1-based), clamped to at least 1."""
        header_excel_row = self.header_row_spin.value()
        if header_excel_row < 1:
            header_excel_row = 1
            self.header_row_spin.setValue(header_excel_row)
        return header_excel_row

    def _selection_hint_text(self) -> str:
        return (
            "Selecione celulas na pre-visualizacao (Shift permite multiplas) e use os botoes acima para ajustar cabecalho e intervalo de colunas."
        )

    def _excel_row_from_table_row(self, row_idx: int) -> int:
        first_data_row = getattr(self, "_current_first_data_row", 2)
        return first_data_row + row_idx

    def _normalize_lookup_key(self, value: object) -> str:
        if value is None:
            return ""
        try:
            if pd.isna(value):
                return ""
        except Exception:
            # pd.isna may not support the value type; ignore and continue.
            pass
        return str(value).strip().casefold()

    def _current_fk_columns(self) -> set[str]:
        cols: set[str] = set()
        for row in range(self.fk_table.rowCount()):
            item = self.fk_table.item(row, 0)
            if item:
                cols.add(item.text())
        return cols

    def _refresh_fk_target_options(self) -> None:
        mapped = self._current_mapped_columns()
        defaults = self._current_default_columns()
        fks = self._current_fk_columns()
        blocked = mapped | defaults | fks
        if self.pk_auto_checkbox.isChecked() and self.primary_key_column:
            blocked.add(self.primary_key_column)
        available = [col for col in self.table_columns if col.name not in blocked]
        self.fk_target_combo.clear()
        for col in available:
            label = f"{col.name} ({col.type})"
            if not col.nullable:
                label += " [Obrigatorio]"
            self.fk_target_combo.addItem(label, col.name)
        self.add_fk_btn.setEnabled(bool(available) and self.fk_excel_combo.count() > 0)

    def _refresh_fk_excel_options(self) -> None:
        self.fk_excel_combo.clear()
        for idx in range(self.sheet_columns_list.count()):
            item = self.sheet_columns_list.item(idx)
            if item:
                self.fk_excel_combo.addItem(item.text())
        enable = self.fk_target_combo.count() > 0 and self.fk_excel_combo.count() > 0
        self.add_fk_btn.setEnabled(enable)

    def _refresh_fk_table_options(self) -> None:
        tables = self.database.list_tables()
        self.fk_table_combo.blockSignals(True)
        self.fk_table_combo.clear()
        for table in tables:
            self.fk_table_combo.addItem(table)
        self.fk_table_combo.blockSignals(False)
        if tables:
            self.fk_table_combo.setCurrentIndex(0)
            self._on_fk_table_changed()
        else:
            self.fk_id_combo.clear()
            self.fk_label_combo.clear()

    def _on_fk_table_changed(self) -> None:
        table = self.fk_table_combo.currentText()
        self.fk_id_combo.clear()
        self.fk_label_combo.clear()
        if not table:
            return
        columns = self._foreign_columns_cache.get(table)
        if columns is None:
            columns = self.database.get_columns(table)
            self._foreign_columns_cache[table] = columns
        for col in columns:
            label = f"{col.name} ({col.type})"
            self.fk_id_combo.addItem(label, col.name)
            self.fk_label_combo.addItem(label, col.name)

    def _add_fk_lookup(self) -> None:
        target = self.fk_target_combo.currentData()
        excel_col = self.fk_excel_combo.currentText()
        foreign_table = self.fk_table_combo.currentText()
        foreign_id = self.fk_id_combo.currentData()
        foreign_label = self.fk_label_combo.currentData()
        if not target or not excel_col or not foreign_table or not foreign_id or not foreign_label:
            QMessageBox.warning(self, "Relacionamento", "Selecione todos os campos para adicionar o relacionamento.")
            return
        if target in self._current_mapped_columns() or target in self._current_default_columns() or target in self._current_fk_columns():
            QMessageBox.warning(self, "Relacionamento", "Coluna de tabela já utilizada no mapeamento, valor padrão ou outro relacionamento.")
            return
        row = self.fk_table.rowCount()
        self.fk_table.insertRow(row)
        self.fk_table.setItem(row, 0, QTableWidgetItem(target))
        self.fk_table.setItem(row, 1, QTableWidgetItem(excel_col))
        self.fk_table.setItem(row, 2, QTableWidgetItem(foreign_table))
        self.fk_table.setItem(row, 3, QTableWidgetItem(foreign_id))
        self.fk_table.setItem(row, 4, QTableWidgetItem(foreign_label))
        self._refresh_fk_target_options()
        self._refresh_default_column_options()
        self._refresh_required_columns_hint()

    def _remove_fk_lookup(self) -> None:
        selection_model = self.fk_table.selectionModel()
        if not selection_model:
            return
        rows = sorted({index.row() for index in selection_model.selectedRows()}, reverse=True)
        for row in rows:
            self.fk_table.removeRow(row)
        self._refresh_fk_target_options()
        self._refresh_default_column_options()
        self._refresh_required_columns_hint()

    def _collect_fk_lookups(self) -> List[ForeignKeyLookup]:
        lookups: List[ForeignKeyLookup] = []
        for row in range(self.fk_table.rowCount()):
            target_item = self.fk_table.item(row, 0)
            excel_item = self.fk_table.item(row, 1)
            table_item = self.fk_table.item(row, 2)
            id_item = self.fk_table.item(row, 3)
            label_item = self.fk_table.item(row, 4)
            if not target_item or not excel_item or not table_item or not id_item or not label_item:
                continue
            lookups.append(
                ForeignKeyLookup(
                    target_column=target_item.text(),
                    excel_column=excel_item.text(),
                    foreign_table=table_item.text(),
                    foreign_id_column=id_item.text(),
                    foreign_label_column=label_item.text(),
                )
            )
        return lookups

    def _set_default_input_widget(self, widget: QWidget) -> None:
        while self.default_value_layout.count():
            item = self.default_value_layout.takeAt(0)
            old = item.widget()
            if old:
                old.setParent(None)
        self.default_value_layout.addWidget(widget)

    def _set_combo_tooltip(self, combo: QComboBox, text: str | None = None) -> None:
        combo.setToolTip(text if text is not None else combo.currentText())

    def _column_kind(self, column: ColumnInfo | None) -> str:
        if not column:
            return "text"
        normalized = column.type.lower()
        if "bool" in normalized:
            return "bool"
        if "date" in normalized or "time" in normalized:
            return "date"
        return "text"

    def _find_column_info(self, name: str) -> ColumnInfo | None:
        for col in self.table_columns:
            if col.name == name:
                return col
        return None

    def _refresh_default_column_options(self) -> None:
        mapped = self._current_mapped_columns()
        defaults = self._current_default_columns()
        blocked = mapped | defaults
        if self.pk_auto_checkbox.isChecked() and self.primary_key_column:
            blocked.add(self.primary_key_column)
        available = [col for col in self.table_columns if col.name not in blocked]
        self.default_column_combo.blockSignals(True)
        self.default_column_combo.clear()
        for col in available:
            label = f"{col.name} ({col.type})"
            if not col.nullable:
                label += " [Obrigatorio]"
            self.default_column_combo.addItem(label, col.name)
        self.default_column_combo.blockSignals(False)
        has_options = bool(available)
        self.add_default_btn.setEnabled(has_options)
        self.default_value_holder.setEnabled(has_options)
        if has_options:
            self.default_column_combo.setCurrentIndex(0)
            self._on_default_column_changed()
        else:
            self._set_default_input_widget(self.default_value_line)
        self._refresh_fk_target_options()
        self._set_combo_tooltip(self.default_column_combo)

    def _on_default_column_changed(self) -> None:
        column_name = self.default_column_combo.currentData()
        column_info = self._find_column_info(column_name) if column_name else None
        kind = self._column_kind(column_info)
        if kind == "bool":
            self.default_bool_combo.setCurrentIndex(0)
            self._set_default_input_widget(self.default_bool_combo)
        elif kind == "date":
            self.default_date_edit.setDate(QDate.currentDate())
            self._set_default_input_widget(self.default_date_edit)
        else:
            self.default_value_line.clear()
            self._set_default_input_widget(self.default_value_line)
        self._set_combo_tooltip(self.default_column_combo)

    def _add_default_value(self) -> None:
        column_name = self.default_column_combo.currentData()
        if not column_name:
            return
        if (
            column_name in self._current_mapped_columns()
            or column_name in self._current_default_columns()
            or column_name in self._current_fk_columns()
        ):
            QMessageBox.warning(self, "Valor padrao", "Coluna ja mapeada, com valor padrao ou relacionamento definido")
            return
        value, display = self._read_default_input(column_name)
        row = self.defaults_table.rowCount()
        self.defaults_table.insertRow(row)
        self.defaults_table.setItem(row, 0, QTableWidgetItem(column_name))
        value_item = QTableWidgetItem(display)
        value_item.setData(Qt.UserRole, value)
        self.defaults_table.setItem(row, 1, value_item)
        self._refresh_default_column_options()
        self._refresh_required_columns_hint()

    def _remove_default_value(self) -> None:
        selection_model = self.defaults_table.selectionModel()
        if not selection_model:
            return
        rows = sorted({index.row() for index in selection_model.selectedRows()}, reverse=True)
        for row in rows:
            self.defaults_table.removeRow(row)
        self._refresh_default_column_options()
        self._refresh_required_columns_hint()

    def _read_default_input(self, column_name: str) -> tuple[object, str]:
        column_info = self._find_column_info(column_name)
        kind = self._column_kind(column_info)
        if kind == "bool":
            value = bool(self.default_bool_combo.currentData())
            return value, "true" if value else "false"
        if kind == "date":
            qdate = self.default_date_edit.date()
            value = date(qdate.year(), qdate.month(), qdate.day())
            return value, value.isoformat()
        value = self.default_value_line.text()
        return value, value

    def _current_mapped_columns(self) -> set[str]:
        cols: set[str] = set()
        for row in range(self.mapping_table.rowCount()):
            table_col_item = self.mapping_table.item(row, 1)
            if table_col_item:
                cols.add(table_col_item.text())
        return cols

    def _current_default_columns(self) -> set[str]:
        cols: set[str] = set()
        for row in range(self.defaults_table.rowCount()):
            col_item = self.defaults_table.item(row, 0)
            if col_item:
                cols.add(col_item.text())
        return cols

    def _collect_default_values(self) -> Dict[str, object]:
        values: Dict[str, object] = {}
        for row in range(self.defaults_table.rowCount()):
            col_item = self.defaults_table.item(row, 0)
            val_item = self.defaults_table.item(row, 1)
            if col_item and val_item:
                col_name = col_item.text()
                stored = val_item.data(Qt.UserRole)
                values[col_name] = stored if stored is not None else val_item.text()
        return values

    def _missing_required_columns(
        self,
        mapping_pairs: List[tuple[str, str]],
        defaults: Dict[str, object],
        autogenerate_pk: bool,
        fk_lookups: List[ForeignKeyLookup],
    ) -> List[str]:
        covered = {table_col for _, table_col in mapping_pairs} | set(defaults.keys()) | {
            fk.target_column for fk in fk_lookups
        }
        missing = []
        for col in self.table_columns:
            if col.nullable:
                continue
            if col.primary_key and autogenerate_pk:
                continue
            if col.name not in covered:
                missing.append(col.name)
        return missing

    def _compact_columns_text(self, columns: List[str], limit: int = 4) -> str:
        if not columns:
            return "--"
        if len(columns) <= limit:
            return ", ".join(columns)
        remaining = len(columns) - limit
        return f"{', '.join(columns[:limit])} +{remaining}"

    def _refresh_required_columns_hint(self) -> None:
        if not self.table_columns:
            self.required_columns_label.setText("Campos obrigatorios: --")
            self.required_columns_label.setToolTip("")
            return
        required = [
            col.name
            for col in self.table_columns
            if not col.nullable and not (col.primary_key and self.pk_auto_checkbox.isChecked())
        ]
        mapping_cols = self._current_mapped_columns()
        default_cols = self._current_default_columns()
        fk_cols = self._current_fk_columns()
        missing = [name for name in required if name not in mapping_cols and name not in default_cols and name not in fk_cols]
        if not required:
            self.required_columns_label.setText("Campos obrigatorios: nenhum")
            self.required_columns_label.setToolTip("")
            return
        required_brief = self._compact_columns_text(required)
        tooltip_lines = [f"Obrigatorios ({len(required)}): {', '.join(required)}"]
        if missing:
            missing_brief = self._compact_columns_text(missing)
            text = (
                f"Obrigatorios ({len(required)}): {required_brief} | "
                f"Faltando ({len(missing)}): {missing_brief}"
            )
            tooltip_lines.append(f"Faltando ({len(missing)}): {', '.join(missing)}")
        else:
            text = f"Obrigatorios ({len(required)}): {required_brief} | Todos atendidos"
        self.required_columns_label.setText(text)
        self.required_columns_label.setToolTip("\n".join(tooltip_lines))

    def _clear_pre_validation_state(self) -> None:
        self._pre_validation_remove_duplicates = False
        self._pre_validation_column = None
        self._pre_validation_last_result = None
        self._refresh_pre_validation_hint()

    def _refresh_pre_validation_hint(self) -> None:
        if not getattr(self, "pre_validation_status", None):
            return
        if not (self._pre_validation_remove_duplicates and self._pre_validation_column):
            self.pre_validation_status.setText("Pré-validação: sem regras")
            self.pre_validation_status.setToolTip("")
            return
        text = f"Pré-validação: remover duplicados em '{self._pre_validation_column}'"
        tooltip = text
        if self._pre_validation_last_result:
            total, unique = self._pre_validation_last_result
            removed = max(total - unique, 0)
            text += f" (previsto remover {removed}/{total})"
            tooltip += f"\nTotal: {total} | Removidos: {removed} | Restantes: {unique}"
        self.pre_validation_status.setText(text)
        self.pre_validation_status.setToolTip(tooltip)

    def _reset_after_execute(self) -> None:
        self.mapping_table.setRowCount(0)
        self.defaults_table.setRowCount(0)
        self.fk_table.setRowCount(0)
        self.preview_text.clear()
        self.default_value_line.clear()
        self.default_bool_combo.setCurrentIndex(0)
        self.default_date_edit.setDate(QDate.currentDate())
        self.pk_auto_checkbox.setChecked(False)
        self.insert_radio.setChecked(True)
        self.update_radio.setChecked(False)
        self.sheet_columns_list.clearSelection()
        self.table_columns_list.clearSelection()
        self.mapping_table.clearSelection()
        self.defaults_table.clearSelection()
        self.fk_table.clearSelection()
        if self.join_combo.count() > 0:
            self.join_combo.setCurrentIndex(0)
        self._refresh_default_column_options()
        self._refresh_fk_target_options()
        self._refresh_required_columns_hint()
        self._clear_pre_validation_state()

    def _on_table_selected(self) -> None:
        items = self.table_list.selectedItems()
        if not items:
            return
        table = items[0].text()
        self.table_columns = self.database.get_columns(table)
        self.columns_list.clear()
        self.table_columns_list.clear()
        self.join_combo.clear()
        self.mapping_table.setRowCount(0)
        self.defaults_table.setRowCount(0)
        self.fk_table.setRowCount(0)
        self.primary_key_column = None
        for col in self.table_columns:
            label = f"{col.name} ({col.type})"
            if col.primary_key:
                label += " [PK]"
                if not self.primary_key_column:
                    self.primary_key_column = col.name
            if not col.nullable:
                label += " [Obrigatorio]"
            self.columns_list.addItem(label)
            self.table_columns_list.addItem(col.name)
            self.join_combo.addItem(col.name)
        self._set_combo_tooltip(self.join_combo)
        if self.primary_key_column:
            self.pk_auto_checkbox.setEnabled(True)
            self.pk_auto_checkbox.setChecked(False)
            self.pk_auto_checkbox.setText(
                f"PK {self.primary_key_column} gerada pelo banco (auto-incremento)"
            )
        else:
            self.pk_auto_checkbox.setEnabled(False)
            self.pk_auto_checkbox.setChecked(False)
            self.pk_auto_checkbox.setText("PK gerada pelo banco (auto-incremento)")
        self._refresh_default_column_options()
        self._refresh_fk_target_options()
        self._refresh_required_columns_hint()

    def _add_mapping(self) -> None:
        sheet_items = self.sheet_columns_list.selectedItems()
        table_items = self.table_columns_list.selectedItems()
        if not sheet_items or not table_items:
            return
        sheet_col = sheet_items[0].text()
        table_col = table_items[0].text()
        if (
            table_col in self._current_mapped_columns()
            or table_col in self._current_default_columns()
            or table_col in self._current_fk_columns()
        ):
            QMessageBox.warning(self, "Mapeamento", "Coluna de tabela já utilizada no mapeamento, valor padrão ou relacionamento")
            return
        row = self.mapping_table.rowCount()
        self.mapping_table.insertRow(row)
        self.mapping_table.setItem(row, 0, QTableWidgetItem(sheet_col))
        self.mapping_table.setItem(row, 1, QTableWidgetItem(table_col))
        self._refresh_default_column_options()
        self._refresh_required_columns_hint()

    def _remove_mapping(self) -> None:
        selection_model = self.mapping_table.selectionModel()
        if not selection_model:
            return
        rows = sorted({index.row() for index in selection_model.selectedRows()}, reverse=True)
        for row in rows:
            self.mapping_table.removeRow(row)
        self._refresh_default_column_options()
        self._refresh_required_columns_hint()

    def _on_pk_auto_toggled(self, checked: bool) -> None:  # noqa: ARG002
        if checked and self.primary_key_column:
            for row in range(self.defaults_table.rowCount() - 1, -1, -1):
                col_item = self.defaults_table.item(row, 0)
                if col_item and col_item.text() == self.primary_key_column:
                    self.defaults_table.removeRow(row)
        self._refresh_default_column_options()
        self._refresh_required_columns_hint()

    def _collect_mapping(self) -> MappingSelection | None:
        sheet_items = self.sheet_list.selectedItems()
        table_items = self.table_list.selectedItems()
        if not sheet_items or not table_items:
            QMessageBox.warning(self, "Mapeamento", "Selecione uma aba e uma tabela")
            return None
        mapping: List[tuple[str, str]] = []
        for row in range(self.mapping_table.rowCount()):
            sheet_col_item = self.mapping_table.item(row, 0)
            table_col_item = self.mapping_table.item(row, 1)
            if sheet_col_item and table_col_item:
                mapping.append((sheet_col_item.text(), table_col_item.text()))
        defaults = self._collect_default_values()
        fk_lookups = self._collect_fk_lookups()
        join_column = self.join_combo.currentText() if self.update_radio.isChecked() else None
        mapping_targets = {table_col for _, table_col in mapping}
        covered_join = mapping_targets | set(defaults.keys()) | {fk.target_column for fk in fk_lookups}
        if join_column and join_column not in covered_join:
            QMessageBox.warning(
                self,
                "Mapeamento",
                "Para UPDATE, a coluna de junção precisa estar mapeada/definida para evitar falhas",
            )
            return None
        autogenerate_pk = bool(
            self.pk_auto_checkbox.isChecked()
            and self.primary_key_column
            and self.primary_key_column not in mapping_targets
        )
        if autogenerate_pk and self.primary_key_column:
            defaults.pop(self.primary_key_column, None)
        if not mapping and not defaults and not fk_lookups:
            QMessageBox.warning(self, "Mapeamento", "Adicione ao menos um mapeamento ou valor padrão")
            return None
        missing_required = self._missing_required_columns(mapping, defaults, autogenerate_pk, fk_lookups)
        if missing_required:
            QMessageBox.warning(
                self,
                "Mapeamento",
                f"Preencha os campos obrigatórios: {', '.join(missing_required)}",
            )
            return None
        header_excel_row = self._current_header_excel_row()
        remove_duplicates = self._pre_validation_remove_duplicates and bool(self._pre_validation_column)
        duplicate_column = self._pre_validation_column if remove_duplicates else None
        return MappingSelection(
            sheet_name=sheet_items[0].text(),
            table_name=table_items[0].text(),
            header_row=header_excel_row,
            start_column=self.col_start_spin.value(),
            end_column=self.col_end_spin.value() or None,
            column_mapping=mapping,
            default_values=defaults,
            fk_lookups=fk_lookups,
            operation="UPDATE" if self.update_radio.isChecked() else "INSERT",
            join_column=join_column,
            primary_key=self.primary_key_column,
            autogenerate_pk=autogenerate_pk,
            remove_duplicate_rows=remove_duplicates,
            duplicate_check_column=duplicate_column,
        )

    def _current_sheet_columns(self) -> List[str]:
        columns: List[str] = []
        for idx in range(self.sheet_columns_list.count()):
            item = self.sheet_columns_list.item(idx)
            if item:
                columns.append(item.text())
        return columns

    def _calculate_duplicate_stats(self, selection: MappingSelection, column: str) -> tuple[int, int]:
        if not self.excel_reader:
            raise ValueError("Nenhuma planilha carregada")
        df = self.excel_reader._read_dataframe(
            selection.sheet_name,
            selection.header_row,
            data_start_row=None,
            data_end_row=None,
            col_start=selection.start_column,
            col_end=selection.end_column,
        )
        if column not in df.columns:
            raise ValueError(f"Coluna '{column}' não encontrada na seleção atual")
        df[column] = df[column].map(self.excel_reader._normalize_cell)
        total_rows = len(df.index)
        unique_rows = len(df.drop_duplicates(subset=[column], keep="first").index)
        return total_rows, unique_rows

    def _open_pre_validation(self) -> None:
        selection = self._collect_mapping()
        if not selection or not self.excel_reader:
            return
        columns = self._current_sheet_columns()
        if not columns:
            QMessageBox.warning(self, "Pré-validação", "Nenhuma coluna disponível para análise")
            return

        dialog = PreValidationDialog(
            self,
            columns=columns,
            remove_duplicates=self._pre_validation_remove_duplicates,
            selected_column=self._pre_validation_column,
            last_result=self._pre_validation_last_result,
            run_check=lambda col: self._calculate_duplicate_stats(selection, col),
        )
        if dialog.exec():
            self._pre_validation_remove_duplicates = dialog.remove_duplicates
            self._pre_validation_column = dialog.selected_column if dialog.remove_duplicates else None
            self._pre_validation_last_result = dialog.last_result if dialog.remove_duplicates else None
            self._refresh_pre_validation_hint()

    def _generate_preview(self) -> None:
        selection = self._collect_mapping()
        if not selection or not self.excel_reader:
            return
        try:
            preview = self.excel_reader.load_sheet_preview(
                selection.sheet_name,
                header_row=selection.header_row,
                col_start=selection.start_column,
                col_end=selection.end_column,
            )
            sql_example = self._build_sql_example(selection)
            text = ["Pré-visualização de dados:", preview.sample.head().to_string()]
            if selection.default_values:
                text.append("")
                text.append("Valores padrão aplicados:")
                for col, value in selection.default_values.items():
                    text.append(f"- {col}: {value}")
            if selection.fk_lookups:
                text.append("")
                text.append("Relacionamentos (descrição -> ID):")
                for fk in selection.fk_lookups:
                    text.append(
                        f"- {fk.target_column} <= {fk.foreign_table}.{fk.foreign_id_column} via {fk.foreign_label_column} = Excel[{fk.excel_column}]"
                    )
            if selection.remove_duplicate_rows and selection.duplicate_check_column:
                text.append("")
                duplicate_summary = "Remover duplicados em " + selection.duplicate_check_column
                try:
                    stats = self._calculate_duplicate_stats(selection, selection.duplicate_check_column)
                    self._pre_validation_last_result = stats
                except Exception:  # noqa: BLE001
                    stats = self._pre_validation_last_result
                if stats:
                    total, unique = stats
                    removed = max(total - unique, 0)
                    duplicate_summary += f" (previsto remover {removed} de {total} linhas)"
                text.append(duplicate_summary)
            text.extend(["", "SQL estimado:", sql_example])
            self.preview_text.setPlainText("\n".join(text))
        except Exception as exc:  # noqa: BLE001
            self._show_error("Erro ao pré-visualizar", exc)

    def _build_sql_example(self, selection: MappingSelection) -> str:
        cols: List[str] = []
        for _, c in selection.column_mapping:
            if selection.autogenerate_pk and selection.primary_key == c:
                continue
            if c not in cols:
                cols.append(c)
        for c in selection.default_values.keys():
            if selection.autogenerate_pk and selection.primary_key == c:
                continue
            if c not in cols:
                cols.append(c)
        for fk in selection.fk_lookups:
            c = fk.target_column
            if selection.autogenerate_pk and selection.primary_key == c:
                continue
            if c not in cols:
                cols.append(c)
        if selection.operation == "INSERT":
            placeholders = ", ".join(f":{c}" for c in cols)
            return f"INSERT INTO {selection.table_name} ({', '.join(cols)}) VALUES ({placeholders});"
        else:
            set_clause = ", ".join(f"{c} = :{c}" for c in cols if c != selection.join_column)
            return f"UPDATE {selection.table_name} SET {set_clause} WHERE {selection.join_column} = :{selection.join_column};"

    def _validate_record_lengths(self, records: List[Dict[str, object]], selection: MappingSelection) -> None:
        limits = {col.name: col.max_length for col in self.table_columns if col.max_length}
        if not limits:
            return
        first_excel_row = selection.header_row + 1
        too_long: List[tuple[int, str, int, str]] = []
        for idx, record in enumerate(records):
            for col_name, max_len in limits.items():
                if col_name not in record:
                    continue
                value = record[col_name]
                if value is None:
                    continue
                text_value = str(value)
                if len(text_value) > max_len:
                    excel_row = first_excel_row + idx
                    too_long.append((excel_row, col_name, max_len, text_value))
        if too_long:
            lines = []
            for excel_row, col_name, max_len, text_value in too_long[:5]:
                preview = text_value if len(text_value) <= 40 else f"{text_value[:37]}..."
                lines.append(
                    f"Linha {excel_row} coluna '{col_name}': {len(text_value)} > {max_len} caracteres (valor: {preview})"
                )
            remaining = len(too_long) - len(lines)
            if remaining > 0:
                lines.append(f"...mais {remaining} ocorrencias com tamanho acima do limite.")
            raise ValueError("Valores excedem o tamanho permitido para a coluna:\n" + "\n".join(lines))

    def _execute(self) -> None:
        selection = self._collect_mapping()
        if not selection or not self.excel_reader:
            return
        try:
            records = self._build_records_for_selection(selection)
            self._validate_record_lengths(records, selection)
            affected = 0
            if selection.operation == "INSERT":
                affected = self.database.execute_insert(
                    selection.table_name,
                    records,
                    autogenerate_pk=selection.autogenerate_pk,
                    primary_key=selection.primary_key,
                )
            else:
                if not selection.join_column:
                    QMessageBox.warning(self, "UPDATE", "Selecione uma coluna de junção")
                    return
                affected = self.database.execute_update(selection.table_name, records, selection.join_column)
            QMessageBox.information(self, "Importação", f"Registros processados: {affected}")
            self._reset_after_execute()
        except Exception as exc:  # noqa: BLE001
            self._show_error("Erro na importação", exc)

    def _show_error(self, title: str, exc: Exception) -> None:
        traceback.print_exc()
        details = traceback.format_exc()
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.resize(800, 500)

        layout = QVBoxLayout(dialog)

        msg_label = QLabel(str(exc) if str(exc) else exc.__class__.__name__)
        msg_label.setWordWrap(True)
        layout.addWidget(msg_label)

        details_edit = QTextEdit()
        details_edit.setReadOnly(True)
        details_edit.setLineWrapMode(QTextEdit.NoWrap)
        details_edit.setPlainText(details)
        layout.addWidget(details_edit, 1)

        buttons = QHBoxLayout()
        copy_btn = QPushButton("Copiar detalhes")
        close_btn = QPushButton("Fechar")
        copy_btn.clicked.connect(lambda: QApplication.clipboard().setText(details))
        close_btn.clicked.connect(dialog.accept)
        buttons.addWidget(copy_btn)
        buttons.addStretch()
        buttons.addWidget(close_btn)
        layout.addLayout(buttons)

        dialog.exec()

    def _build_records_for_selection(self, selection: MappingSelection) -> List[Dict[str, object]]:
        # Carrega todas as colunas necessárias (mapeamento + lookups de FK)
        df = self.excel_reader._read_dataframe(
            selection.sheet_name,
            selection.header_row,
            data_start_row=None,
            data_end_row=None,
            col_start=selection.start_column,
            col_end=selection.end_column,
        )
        if selection.remove_duplicate_rows and selection.duplicate_check_column:
            if selection.duplicate_check_column not in df.columns:
                raise ValueError(
                    f"Coluna '{selection.duplicate_check_column}' não encontrada para remover duplicados"
                )
            df[selection.duplicate_check_column] = df[selection.duplicate_check_column].map(
                self.excel_reader._normalize_cell
            )
            total_rows = len(df.index)
            df = df.drop_duplicates(subset=[selection.duplicate_check_column], keep="first")
            df = df.reset_index(drop=True)
            self._pre_validation_last_result = (total_rows, len(df.index))
        column_mapping = selection.column_mapping
        if selection.autogenerate_pk and selection.primary_key:
            column_mapping = [(s, t) for s, t in selection.column_mapping if t != selection.primary_key]
        needed_excel_cols = {s for s, _ in column_mapping} | {fk.excel_column for fk in selection.fk_lookups}
        missing_excel = [col for col in needed_excel_cols if col not in df.columns]
        if missing_excel:
            raise ValueError(f"Colunas da planilha não encontradas: {', '.join(missing_excel)}")

        records: List[Dict[str, object]] = []
        for _, row in df.iterrows():
            record: Dict[str, object] = {}
            for sheet_col, table_col in column_mapping:
                if sheet_col in row:
                    record[table_col] = self.excel_reader._normalize_cell(row[sheet_col])
            records.append(record)

        # Aplica valores padrão
        if selection.default_values:
            for record in records:
                for col, value in selection.default_values.items():
                    record.setdefault(col, value)

        # Aplica lookups de FK (descrição -> ID)
        if selection.fk_lookups:
            lookup_cache: Dict[tuple[str, str, str], Dict[str, object]] = {}
            for fk in selection.fk_lookups:
                key = (fk.foreign_table, fk.foreign_id_column, fk.foreign_label_column)
                if key not in lookup_cache:
                    cache: Dict[str, object] = {}
                    duplicates: List[str] = []
                    for ident, label in self.database.fetch_lookup_values(
                        fk.foreign_table, fk.foreign_id_column, fk.foreign_label_column
                    ):
                        normalized = self._normalize_lookup_key(label)
                        if not normalized:
                            continue
                        existing = cache.get(normalized)
                        if existing is None:
                            cache[normalized] = ident
                        elif existing != ident:
                            duplicates.append(str(label))
                    if duplicates:
                        raise ValueError(
                            f"Valores duplicados na tabela {fk.foreign_table} para a coluna de descrição "
                            f"{fk.foreign_label_column}: {', '.join(sorted(set(duplicates)))}"
                        )
                    lookup_cache[key] = cache
            unresolved: List[str] = []
            first_excel_row = selection.header_row + 1
            for idx, (_, row) in enumerate(df.iterrows()):
                excel_row = first_excel_row + idx
                for fk in selection.fk_lookups:
                    raw_value = row.get(fk.excel_column)
                    normalized = self._normalize_lookup_key(raw_value)
                    if not normalized:
                        unresolved.append(
                            f"Linha {excel_row} coluna '{fk.excel_column}' vazia para preencher {fk.target_column}"
                        )
                        continue
                    key = (fk.foreign_table, fk.foreign_id_column, fk.foreign_label_column)
                    mapped = lookup_cache.get(key, {}).get(normalized)
                    if mapped is None:
                        preview = str(raw_value)
                        unresolved.append(
                            f"Linha {excel_row}: valor '{preview}' não encontrado em "
                            f"{fk.foreign_table}.{fk.foreign_label_column} para preencher {fk.target_column}"
                        )
                    else:
                        records[idx][fk.target_column] = mapped
            if unresolved:
                details = "\n".join(unresolved[:5])
                remaining = len(unresolved) - len(unresolved[:5])
                if remaining > 0:
                    details += f"\n...mais {remaining} ocorrências sem correspondência."
                raise ValueError("Não foi possível resolver os relacionamentos FK:\n" + details)

        # Remove PK se marcada como auto-gerada
        if selection.autogenerate_pk and selection.primary_key:
            records = [{k: v for k, v in record.items() if k != selection.primary_key} for record in records]
        return records

    def _export_mapped_data(self, kind: str) -> None:
        selection = self._collect_mapping()
        if not selection or not self.excel_reader:
            return
        try:
            records = self._build_records_for_selection(selection)
            if not records:
                QMessageBox.information(self, "Exportar", "Nenhum registro encontrado no intervalo selecionado.")
                return
            export_records = [
                {k: self._normalize_value_for_export(v) for k, v in record.items()} for record in records
            ]
            df = pd.DataFrame(export_records)
            if kind == "csv":
                path, _ = QFileDialog.getSaveFileName(
                    self, "Salvar CSV mapeado", str(Path.home() / "dados_mapeados.csv"), "CSV (*.csv)"
                )
                if not path:
                    return
                df.to_csv(path, index=False)
            else:
                path, _ = QFileDialog.getSaveFileName(
                    self, "Salvar Excel mapeado", str(Path.home() / "dados_mapeados.xlsx"), "Excel (*.xlsx)"
                )
                if not path:
                    return
                df.to_excel(path, index=False)
            QMessageBox.information(self, "Exportar", f"Arquivo salvo em:\n{path}")
        except Exception as exc:  # noqa: BLE001
            self._show_error("Erro ao exportar", exc)

    def _normalize_value_for_export(self, value: object) -> object:
        if isinstance(value, bool):
            return "true" if value else "false"
        if isinstance(value, (date,)):
            return value.isoformat()
        return value


class PreValidationDialog(QDialog):
    def __init__(
        self,
        parent: QWidget,
        *,
        columns: List[str],
        remove_duplicates: bool,
        selected_column: Optional[str],
        last_result: Optional[Tuple[int, int]],
        run_check: Callable[[str], tuple[int, int]],
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Pré-validação")
        self._run_check = run_check
        self._last_result = last_result

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Configure verificações antes de importar os dados."))

        self.remove_duplicates_checkbox = QCheckBox("Remover linhas com valores duplicados")
        self.remove_duplicates_checkbox.setChecked(remove_duplicates)
        layout.addWidget(self.remove_duplicates_checkbox)

        column_layout = QHBoxLayout()
        column_layout.addWidget(QLabel("Coluna para checar:"))
        self.duplicate_column_combo = QComboBox()
        for col in columns:
            self.duplicate_column_combo.addItem(col)
        if selected_column and selected_column in columns:
            self.duplicate_column_combo.setCurrentText(selected_column)
        column_layout.addWidget(self.duplicate_column_combo)
        layout.addLayout(column_layout)

        actions_layout = QHBoxLayout()
        self.check_duplicates_btn = QPushButton("Checar duplicados")
        self.check_duplicates_btn.clicked.connect(self._on_check_duplicates)
        actions_layout.addWidget(self.check_duplicates_btn)
        actions_layout.addStretch()
        layout.addLayout(actions_layout)

        self.summary_label = QLabel(
            "Selecione uma coluna e clique em Checar duplicados para ver o impacto da limpeza."
        )
        self.summary_label.setWordWrap(True)
        layout.addWidget(self.summary_label)

        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()
        self.apply_btn = QPushButton("Aplicar")
        self.apply_btn.clicked.connect(self._on_accept)
        self.apply_btn.setEnabled(True)
        cancel_btn = QPushButton("Cancelar")
        cancel_btn.clicked.connect(self.reject)
        buttons_layout.addWidget(self.apply_btn)
        buttons_layout.addWidget(cancel_btn)
        layout.addLayout(buttons_layout)

        if self._last_result:
            self._update_summary_text(self._last_result)

    @property
    def remove_duplicates(self) -> bool:
        return self.remove_duplicates_checkbox.isChecked()

    @property
    def selected_column(self) -> Optional[str]:
        if self.duplicate_column_combo.count() == 0:
            return None
        return self.duplicate_column_combo.currentText()

    @property
    def last_result(self) -> Optional[Tuple[int, int]]:
        return self._last_result if self.remove_duplicates else None

    def _on_check_duplicates(self) -> None:
        column = self.selected_column
        if not column:
            QMessageBox.warning(self, "Pré-validação", "Selecione uma coluna para checar duplicados")
            return
        try:
            self._last_result = self._run_check(column)
            self._update_summary_text(self._last_result)
            self.apply_btn.setEnabled(True)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "Pré-validação", str(exc))

    def _update_summary_text(self, result: Tuple[int, int]) -> None:
        total, unique = result
        removed = max(total - unique, 0)
        self.summary_label.setText(
            f"Total de linhas: {total} | Duplicadas removíveis: {removed} | Restantes: {unique}"
        )

    def _on_accept(self) -> None:
        if self.remove_duplicates and self._last_result is None:
            QMessageBox.warning(self, "Pré-validação", "Execute a checagem antes de aplicar a remoção de duplicados")
            return
        self.accept()


def main() -> int:
    app = QApplication([])
    window = MainWindow()
    window.showMaximized()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
