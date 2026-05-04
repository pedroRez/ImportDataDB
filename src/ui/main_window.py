from __future__ import annotations

import traceback
from collections import Counter
from datetime import date
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple
import unicodedata
from uuid import uuid4

import pandas as pd
from PySide6.QtCore import QDate, QItemSelectionModel, QThread, Qt
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QFileDialog,
    QFrame,
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
    QScrollArea,
    QHeaderView,
    QRadioButton,
    QProgressDialog,
    QSpinBox,
    QSplitter,
    QStackedWidget,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QToolButton,
    QVBoxLayout,
    QWidget,
)


from src.core.mapping import ForeignKeyLookup, MappingSelection
from src.core.profiles import ImportProfile, list_profiles, load_profile, save_profile
from src.core.xerife_bridge import run_xerife_stock_batch
from src.core.xerife_stock import XerifeStockImporter, XerifeValidationResult
from src.db.provider import ColumnInfo, DatabaseProvider
from src.excel.reader import ExcelReader
from src.ui.excel_selection_dialog import ExcelSelectionDialog
from src.ui.excel_workspace import (
    ExcelGridModel,
    ExcelLoadWorker,
    ExcelSelectionResult,
    ExcelSelectionWorkspace,
)
from src.ui.quick_import import ConnectionStatusBadge, QuickImportPage
from src.ui.theme import build_app_stylesheet
from src.version import APP_NAME, __version__



class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} {__version__}")
        self.resize(1200, 800)

        self.database = DatabaseProvider()
        self.excel_reader: ExcelReader | None = None
        self.table_columns: List[ColumnInfo] = []
        self.primary_key_column: str | None = None
        self._current_header_excel_row_value = 1
        self._foreign_columns_cache: dict[str, List[ColumnInfo]] = {}
        self._pre_validation_remove_duplicates = False
        self._pre_validation_trim_whitespace = False
        self._pre_validation_column: str | None = None
        self._pre_validation_last_result: tuple[int, int] | None = None
        self._split_enabled = False
        self._split_column: str | None = None
        self._split_operator: str = "gt"
        self._split_length: int | None = None
        self._split_extra_name: str | None = None
        self._virtual_extra_columns: set[str] = set()
        self._fk_trim_whitespace = True
        self._similarity_replacements: Dict[str, Dict[str, str]] = {}
        self._cancel_requested = False
        self._last_skipped_null_rows = 0
        self.excel_file_path: Path | None = None
        self._last_conversion_file: Path | None = None
        self._relation_conversions: Dict[str, Dict[str, str]] = {}

        self._excel_step_ready = False
        self._db_step_ready = False
        self._available_profiles: List[ImportProfile] = []
        self._last_validation_result: XerifeValidationResult | None = None
        self._last_import_result: Dict[str, Any] | None = None
        self._quick_mode = "advanced"
        self._excel_grid_model = ExcelGridModel()
        self._excel_grid_selection_model = QItemSelectionModel(self._excel_grid_model)
        self._excel_dataframe: pd.DataFrame | None = None
        self._excel_grid_headers: list[str] = []
        self._excel_grid_first_data_row = 2
        self._excel_grid_header_row = 1
        self._excel_load_thread: QThread | None = None
        self._excel_load_worker: ExcelLoadWorker | None = None
        self._excel_load_job_id = 0
        self._applied_excel_selection: ExcelSelectionResult | None = None
        self._active_excel_workspace: ExcelSelectionWorkspace | None = None
        self._excel_selection_dialog: ExcelSelectionDialog | None = None

        self.host_edit = QLineEdit("localhost")
        self.port_edit = QLineEdit("5432")
        self.db_edit = QLineEdit()
        self.user_edit = QLineEdit()
        self.pwd_edit = QLineEdit()
        self.pwd_edit.setEchoMode(QLineEdit.Password)
        self.loaded_profile_id: str | None = None
        self._last_profile_preview: str = ""

        self._build_menu()
        self._build_layout()
        self._excel_grid_selection_model.selectionChanged.connect(
            lambda _selected, _deselected: self._update_selection_info()
        )
        self.setStyleSheet(build_app_stylesheet())
        self._set_mode("advanced")
        self._sync_quick_workflow_state()

    def _build_menu(self) -> None:
        open_action = QAction("Abrir Excel", self)
        open_action.triggered.connect(self._choose_excel)
        menubar = self.menuBar()
        file_menu = menubar.addMenu("Arquivo")
        file_menu.addAction(open_action)

    def _build_layout(self) -> None:
        central = QWidget()
        central.setObjectName("AppShell")
        central_layout = QVBoxLayout(central)
        central_layout.setContentsMargins(24, 24, 24, 24)
        central_layout.setSpacing(16)

        central_layout.addWidget(self._build_shell_header())
        central_layout.addWidget(self._build_mode_switch())

        self.content_stack = QStackedWidget()
        self.quick_page = self._build_quick_import_page()
        self.advanced_page = self._build_advanced_page()
        self.quick_page_container = self._wrap_mode_page(self.quick_page)
        self.advanced_page_container = self._wrap_mode_page(self.advanced_page, allow_horizontal=True)
        self.content_stack.addWidget(self.quick_page_container)
        self.content_stack.addWidget(self.advanced_page_container)
        central_layout.addWidget(self.content_stack, 1)

        self.setCentralWidget(central)
        self._update_step_progress()

    def _wrap_mode_page(self, page: QWidget, *, allow_horizontal: bool = False) -> QScrollArea:
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded if allow_horizontal else Qt.ScrollBarAlwaysOff)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setWidget(page)
        return scroll_area

    def _build_shell_header(self) -> QWidget:
        shell = QFrame()
        shell.setProperty("card", True)
        shell.setProperty("modeShell", True)
        layout = QHBoxLayout(shell)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        title_block = QFrame()
        title_layout = QVBoxLayout(title_block)
        title_layout.setContentsMargins(18, 14, 18, 14)
        title_layout.setSpacing(4)

        eyebrow = QLabel(f"{APP_NAME} {__version__}")
        eyebrow.setProperty("role", "eyebrow")
        title_layout.addWidget(eyebrow)

        self.shell_title_label = QLabel("Importacao generica de Excel para banco")
        self.shell_title_label.setProperty("role", "section-title")
        title_layout.addWidget(self.shell_title_label)

        self.shell_subtitle_label = QLabel(
            "Mapeie colunas de/para, resolva FK por descricao, valide os dados e execute INSERT ou UPDATE."
        )
        self.shell_subtitle_label.setWordWrap(True)
        self.shell_subtitle_label.setProperty("role", "muted")
        title_layout.addWidget(self.shell_subtitle_label)

        layout.addWidget(title_block, 1)

        status_block = QFrame()
        status_block.setProperty("card", True)
        status_layout = QHBoxLayout(status_block)
        status_layout.setContentsMargins(18, 14, 18, 14)
        status_layout.setSpacing(12)

        status_title = QLabel("Conexao do banco")
        status_title.setProperty("role", "card-title")
        status_layout.addWidget(status_title)

        self.connection_status_label = ConnectionStatusBadge()
        status_layout.addWidget(self.connection_status_label)

        self.connection_btn = QPushButton("Configurar conexao")
        self.connection_btn.setProperty("variant", "primary")
        self.connection_btn.clicked.connect(self._open_connection_dialog)
        status_layout.addWidget(self.connection_btn)

        layout.addWidget(status_block, 0)
        return shell

    def _build_mode_switch(self) -> QWidget:
        container = QFrame()
        container.setProperty("card", True)
        layout = QHBoxLayout(container)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(10)

        text = QLabel("Fluxo de trabalho:")
        text.setProperty("role", "card-title")
        layout.addWidget(text)

        self.mode_button_group = QButtonGroup(self)
        self.mode_button_group.setExclusive(True)

        self.advanced_mode_btn = QToolButton()
        self.advanced_mode_btn.setText("Importacao generica")
        self.advanced_mode_btn.setCheckable(True)
        self.advanced_mode_btn.setProperty("modeButton", True)
        self.advanced_mode_btn.clicked.connect(lambda: self._set_mode("advanced"))
        self.mode_button_group.addButton(self.advanced_mode_btn)
        layout.addWidget(self.advanced_mode_btn)

        self.quick_mode_btn = QToolButton()
        self.quick_mode_btn.setText("Perfil Xerife")
        self.quick_mode_btn.setCheckable(True)
        self.quick_mode_btn.setProperty("modeButton", True)
        self.quick_mode_btn.clicked.connect(lambda: self._set_mode("quick"))
        self.mode_button_group.addButton(self.quick_mode_btn)
        layout.addWidget(self.quick_mode_btn)
        layout.addStretch()
        return container

    def _build_quick_import_page(self) -> QuickImportPage:
        page = QuickImportPage()
        for index, button in enumerate(page.stepper.buttons):
            button.clicked.connect(lambda _checked=False, step=index: page.set_step(step))
        page.use_profile_btn.clicked.connect(self._apply_selected_quick_profile)
        page.new_profile_btn.clicked.connect(self._open_advanced_profile_editor)
        page.model_next_btn.clicked.connect(lambda: page.set_step(1))
        page.spreadsheet_back_btn.clicked.connect(lambda: page.set_step(0))
        page.spreadsheet_next_btn.clicked.connect(lambda: page.set_step(2))
        page.validation_back_btn.clicked.connect(lambda: page.set_step(1))
        page.validation_next_btn.clicked.connect(lambda: page.set_step(3))
        page.import_back_btn.clicked.connect(lambda: page.set_step(2))
        page.profile_list.currentItemChanged.connect(lambda _current, _previous: self._sync_quick_workflow_state())
        page.open_selection_btn.clicked.connect(self._open_excel_selection_dialog)
        page.reopen_file_btn.clicked.connect(self._choose_excel)
        page.validate_btn.clicked.connect(self._validate_profile_only)
        page.import_btn.clicked.connect(self._import_validated_profile)
        page.validation_details_btn.clicked.connect(self._toggle_validation_details)
        page.set_step(0)
        return page

    def _build_advanced_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        intro = QFrame()
        intro.setProperty("card", True)
        intro_layout = QVBoxLayout(intro)
        intro_layout.setContentsMargins(18, 16, 18, 16)
        intro_layout.setSpacing(6)

        intro_eyebrow = QLabel("Fluxo principal")
        intro_eyebrow.setProperty("role", "eyebrow")
        intro_layout.addWidget(intro_eyebrow)

        intro_title = QLabel("Mapeamento de/para, FK e validacoes")
        intro_title.setProperty("role", "section-title")
        intro_layout.addWidget(intro_title)

        intro_text = QLabel(
            "Use este fluxo para importar qualquer planilha para qualquer tabela conectada, "
            "com valores padrao, relacionamentos FK por descricao e INSERT ou UPDATE."
        )
        intro_text.setWordWrap(True)
        intro_text.setProperty("role", "muted")
        intro_layout.addWidget(intro_text)
        layout.addWidget(intro)

        self.step_header_label = QLabel()
        self.step_hint_label = QLabel()
        self.step_hint_label.setWordWrap(True)

        progress_card = QFrame()
        progress_card.setProperty("card", True)
        progress_layout = QVBoxLayout(progress_card)
        progress_layout.setContentsMargins(20, 18, 20, 18)
        progress_layout.setSpacing(6)
        progress_layout.addWidget(self.step_header_label)
        progress_layout.addWidget(self.step_hint_label)
        layout.addWidget(progress_card)

        self.advanced_profile_toggle = QToolButton()
        self.advanced_profile_toggle.setText("Mostrar editor de modelo")
        self.advanced_profile_toggle.setCheckable(True)
        self.advanced_profile_toggle.setProperty("variant", "ghost")
        self.advanced_profile_toggle.toggled.connect(self._toggle_advanced_profile_panel)
        layout.addWidget(self.advanced_profile_toggle, alignment=Qt.AlignLeft)

        self.advanced_profile_panel = self._build_profile_panel()
        self.advanced_profile_panel.setVisible(False)
        layout.addWidget(self.advanced_profile_panel)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_excel_tab(), "1. Excel")
        self.tabs.addTab(self._build_database_tab(), "2. Banco de dados")
        self.tabs.addTab(self._build_mapping_tab(), "3. Mapeamento")
        layout.addWidget(self.tabs, 1)
        return page

    def _build_profile_panel(self) -> QWidget:
        panel = QGroupBox("Modelos e parametros do fluxo rapido")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(18, 20, 18, 18)
        layout.setSpacing(16)

        saved_group = QGroupBox("Usar modelo salvo")
        saved_layout = QVBoxLayout(saved_group)
        saved_layout.setSpacing(14)

        saved_grid = QGridLayout()
        saved_grid.setHorizontalSpacing(12)
        saved_grid.setVerticalSpacing(12)
        saved_grid.addWidget(QLabel("Modelo"), 0, 0)
        self.profile_combo = QComboBox()
        self.profile_combo.currentIndexChanged.connect(self._on_profile_combo_changed)
        saved_grid.addWidget(self.profile_combo, 0, 1, 1, 2)
        self.profile_load_btn = QPushButton("Aplicar modelo")
        self.profile_load_btn.clicked.connect(self._apply_selected_profile)
        saved_grid.addWidget(self.profile_load_btn, 0, 3)
        self.profile_refresh_btn = QPushButton("Atualizar lista")
        self.profile_refresh_btn.clicked.connect(self._refresh_profile_options)
        saved_grid.addWidget(self.profile_refresh_btn, 0, 4)
        saved_layout.addLayout(saved_grid)

        action_row = QGridLayout()
        action_row.setHorizontalSpacing(12)
        action_row.setVerticalSpacing(12)
        self.quick_choose_file_btn = QPushButton("Escolher planilha")
        self.quick_choose_file_btn.clicked.connect(self._choose_excel)
        action_row.addWidget(self.quick_choose_file_btn, 0, 0)
        self.quick_validate_btn = QPushButton("Validar e importar")
        self.quick_validate_btn.setProperty("variant", "primary")
        self.quick_validate_btn.clicked.connect(self._validate_and_import_profile)
        action_row.addWidget(self.quick_validate_btn, 0, 1)
        saved_layout.addLayout(action_row)

        self.profile_summary_label = QLabel("Selecione um modelo salvo para carregar o mapeamento e executar o fluxo rapido.")
        self.profile_summary_label.setWordWrap(True)
        saved_layout.addWidget(self.profile_summary_label)
        layout.addWidget(saved_group)

        edit_group = QGroupBox("Criar/editar modelo")
        edit_layout = QVBoxLayout(edit_group)
        edit_layout.setSpacing(14)

        id_grid = QGridLayout()
        id_grid.setHorizontalSpacing(12)
        id_grid.setVerticalSpacing(12)
        id_grid.addWidget(QLabel("ID"), 0, 0)
        self.profile_id_edit = QLineEdit()
        id_grid.addWidget(self.profile_id_edit, 0, 1)
        id_grid.addWidget(QLabel("Nome"), 0, 2)
        self.profile_name_edit = QLineEdit()
        id_grid.addWidget(self.profile_name_edit, 0, 3)
        edit_layout.addLayout(id_grid)

        meta_grid = QGridLayout()
        meta_grid.setHorizontalSpacing(12)
        meta_grid.setVerticalSpacing(12)
        meta_grid.addWidget(QLabel("Alvo"), 0, 0)
        self.profile_target_combo = QComboBox()
        self.profile_target_combo.addItem("xerife_stock")
        meta_grid.addWidget(self.profile_target_combo, 0, 1)
        meta_grid.addWidget(QLabel("Filial"), 0, 2)
        self.profile_filial_spin = QSpinBox()
        self.profile_filial_spin.setMinimum(0)
        self.profile_filial_spin.setValue(1)
        meta_grid.addWidget(self.profile_filial_spin, 0, 3)
        meta_grid.addWidget(QLabel("Usuario"), 1, 0)
        self.profile_usuario_spin = QSpinBox()
        self.profile_usuario_spin.setMinimum(0)
        self.profile_usuario_spin.setValue(1)
        meta_grid.addWidget(self.profile_usuario_spin, 1, 1)
        self.profile_save_btn = QPushButton("Salvar modelo atual")
        self.profile_save_btn.setProperty("variant", "primary")
        self.profile_save_btn.clicked.connect(self._save_current_profile)
        meta_grid.addWidget(self.profile_save_btn, 1, 3)
        edit_layout.addLayout(meta_grid)

        helper = QLabel(
            "O modelo salvo usa a aba, cabecalho, faixa de linhas/colunas e o alvo atual. "
            "No fluxo Xerife, o importador aplica as regras de unidade e envia em lote para o ERP."
        )
        helper.setWordWrap(True)
        edit_layout.addWidget(helper)
        layout.addWidget(edit_group)

        self._refresh_profile_options()
        return panel

    def _configure_excel_workspace(self, workspace: ExcelSelectionWorkspace) -> None:
        workspace.bind_grid(self._excel_grid_model, self._excel_grid_selection_model)
        workspace.chooseFileRequested.connect(self._choose_excel)
        workspace.fileDropped.connect(self._load_excel_file_from_drop)
        workspace.sheetChanged.connect(self._on_workspace_sheet_changed)
        workspace.reloadRequested.connect(self._refresh_sheet_preview)
        workspace.applySelectionRequested.connect(self._apply_selection_to_range)
        workspace.useSelectionAsHeaderRequested.connect(self._apply_selection_to_header)
        workspace.clearSelectionRequested.connect(self._clear_excel_selection)

    def _all_excel_workspaces(self) -> list[ExcelSelectionWorkspace]:
        workspaces: list[ExcelSelectionWorkspace] = []
        if self._active_excel_workspace is not None:
            workspaces.append(self._active_excel_workspace)
        return workspaces

    def _set_mode(self, mode: str) -> None:
        self._quick_mode = mode
        is_quick = mode == "quick"
        self.quick_mode_btn.setChecked(is_quick)
        self.advanced_mode_btn.setChecked(not is_quick)
        if hasattr(self, "shell_title_label"):
            if is_quick:
                self.shell_title_label.setText("Importacao por perfil Xerife")
                self.shell_subtitle_label.setText(
                    "Selecione um modelo salvo, confirme o Excel em uma janela dedicada e envie o lote."
                )
            else:
                self.shell_title_label.setText("Importacao generica de Excel para banco")
                self.shell_subtitle_label.setText(
                    "Mapeie colunas de/para, resolva FK por descricao, valide os dados e execute INSERT ou UPDATE."
                )
        self._refresh_widget_style(self.quick_mode_btn)
        self._refresh_widget_style(self.advanced_mode_btn)
        self.content_stack.setCurrentWidget(
            self.quick_page_container if is_quick else self.advanced_page_container
        )

    def _toggle_advanced_profile_panel(self, visible: bool) -> None:
        if hasattr(self, "advanced_profile_panel"):
            self.advanced_profile_panel.setVisible(visible)
        if hasattr(self, "advanced_profile_toggle"):
            self.advanced_profile_toggle.setText(
                "Ocultar editor de modelo" if visible else "Mostrar editor de modelo"
            )

    def _selected_quick_profile_id(self) -> str | None:
        if not hasattr(self, "quick_page"):
            return None
        return self.quick_page.selected_profile_id()

    def _apply_selected_quick_profile(self) -> None:
        profile_id = self._selected_quick_profile_id()
        if not profile_id:
            QMessageBox.warning(self, "Modelo", "Selecione um modelo salvo na lista.")
            return
        profile = load_profile(profile_id)
        self._apply_profile(profile)
        self.quick_page.set_step(1)

    def _open_advanced_profile_editor(self) -> None:
        self._set_mode("advanced")
        if hasattr(self, "advanced_profile_toggle"):
            self.advanced_profile_toggle.setChecked(True)
        self.profile_id_edit.setFocus()

    def _toggle_validation_details(self) -> None:
        visible = not self.quick_page.validation_details.isVisible()
        self.quick_page.validation_details.setVisible(visible)
        self.quick_page.validation_details_btn.setText("Ocultar detalhes" if visible else "Ver detalhes")

    def _load_excel_file_from_drop(self, file_name: str) -> None:
        self._load_excel_file(Path(file_name), open_dialog=True)

    def _clone_excel_selection(self, selection: ExcelSelectionResult | None) -> ExcelSelectionResult | None:
        if selection is None:
            return None
        return ExcelSelectionResult(
            header_row=selection.header_row,
            data_start_row=selection.data_start_row,
            data_end_row=selection.data_end_row,
            col_start=selection.col_start,
            col_end=selection.col_end,
            selected_columns=list(selection.selected_columns),
        )

    def _capture_excel_dialog_snapshot(self) -> dict[str, Any]:
        selected_items = self.sheet_list.selectedItems() if hasattr(self, "sheet_list") else []
        selected_sheet = selected_items[0].text() if selected_items else None
        return {
            "sheet_name": selected_sheet,
            "header_row": self.header_row_spin.value() if hasattr(self, "header_row_spin") else 1,
            "row_start": self.row_start_spin.value() if hasattr(self, "row_start_spin") else 0,
            "row_end": self.row_end_spin.value() if hasattr(self, "row_end_spin") else 0,
            "col_start": self.col_start_spin.value() if hasattr(self, "col_start_spin") else 1,
            "col_end": self.col_end_spin.value() if hasattr(self, "col_end_spin") else 0,
            "applied_selection": self._clone_excel_selection(self._applied_excel_selection),
            "validation_result": self._last_validation_result,
            "import_result": self._last_import_result,
            "profile_preview": self._last_profile_preview,
            "excel_ready": self._excel_step_ready,
        }

    def _restore_excel_dialog_snapshot(self, snapshot: dict[str, Any]) -> None:
        self.header_row_spin.setValue(int(snapshot.get("header_row", 1)))
        self.row_start_spin.setValue(int(snapshot.get("row_start", 0)))
        self.row_end_spin.setValue(int(snapshot.get("row_end", 0)))
        self.col_start_spin.setValue(int(snapshot.get("col_start", 1)))
        self.col_end_spin.setValue(int(snapshot.get("col_end", 0)))
        self._applied_excel_selection = self._clone_excel_selection(snapshot.get("applied_selection"))
        self._last_validation_result = snapshot.get("validation_result")
        self._last_import_result = snapshot.get("import_result")
        self._last_profile_preview = str(snapshot.get("profile_preview", "") or "")
        self._set_excel_step_ready(bool(snapshot.get("excel_ready", False)))

        selected_sheet = snapshot.get("sheet_name")
        if selected_sheet:
            self._set_selected_sheet(str(selected_sheet))
        else:
            self._refresh_sheet_preview()
        self._sync_quick_workflow_state()

    def _default_excel_selection(self) -> ExcelSelectionResult | None:
        if not self.excel_reader:
            return None
        selected_items = self.sheet_list.selectedItems()
        if not selected_items:
            return None
        data_start_row = self._current_data_start_excel_row()
        if self._excel_grid_model.rowCount() > 0:
            last_excel_row = self._excel_row_from_table_row(self._excel_grid_model.rowCount() - 1)
        else:
            last_excel_row = data_start_row
        data_end_row = self.row_end_spin.value() or last_excel_row
        col_start = self.col_start_spin.value() or 1
        col_end = self.col_end_spin.value() or max(col_start, col_start + max(len(self._excel_grid_headers) - 1, 0))
        selected_columns = list(self._excel_grid_headers)
        return ExcelSelectionResult(
            header_row=self._current_header_excel_row(),
            data_start_row=data_start_row,
            data_end_row=data_end_row,
            col_start=col_start,
            col_end=col_end,
            selected_columns=selected_columns,
        )

    def _confirm_excel_dialog_selection(self, *, use_selected_header: bool) -> None:
        selection = self._current_excel_selection() or self._default_excel_selection()
        if selection is None:
            QMessageBox.warning(self, "Excel", "Nao foi possivel confirmar o recorte atual.")
            return

        if use_selected_header and self._current_excel_selection() is not None:
            new_header_row = max(1, selection.data_start_row)
            self.header_row_spin.setValue(new_header_row)
            self.row_start_spin.setValue(new_header_row + 1)
            self.row_end_spin.setValue(selection.data_end_row)
            self.col_start_spin.setValue(selection.col_start)
            self.col_end_spin.setValue(selection.col_end)
            self._applied_excel_selection = ExcelSelectionResult(
                header_row=new_header_row,
                data_start_row=new_header_row + 1,
                data_end_row=selection.data_end_row,
                col_start=selection.col_start,
                col_end=selection.col_end,
                selected_columns=[],
            )
            self._refresh_sheet_preview()
        else:
            self._applied_excel_selection = self._clone_excel_selection(selection)
            self.header_row_spin.setValue(selection.header_row)
            self.row_start_spin.setValue(selection.data_start_row)
            self.row_end_spin.setValue(selection.data_end_row)
            self.col_start_spin.setValue(selection.col_start)
            self.col_end_spin.setValue(selection.col_end)

        self._set_excel_step_ready(True)
        self._last_validation_result = None
        self._last_import_result = None
        self._last_profile_preview = ""
        self._update_profile_summary()
        self._sync_quick_workflow_state()

    def _open_excel_selection_dialog(self) -> None:
        if not self.excel_reader:
            QMessageBox.warning(self, "Excel", "Escolha a planilha antes de abrir o seletor.")
            return
        if not self.sheet_list.selectedItems():
            QMessageBox.warning(self, "Excel", "Selecione uma aba da planilha antes de abrir o seletor.")
            return

        snapshot = self._capture_excel_dialog_snapshot()
        dialog = ExcelSelectionDialog(parent=self)
        self._excel_selection_dialog = dialog
        self._active_excel_workspace = dialog.workspace
        self._configure_excel_workspace(dialog.workspace)
        dialog.set_file_hint(self.excel_file_path)
        dialog.show_maximized()
        self._sync_excel_workspaces()
        if self._excel_dataframe is None:
            dialog.workspace.set_loading("Carregando planilha...", "Montando grade da aba selecionada.")
        elif self._excel_grid_model.rowCount() > 0:
            dialog.workspace.show_table(
                status_text=(
                    f"Grade pronta: {self._excel_grid_model.rowCount()} linhas x "
                    f"{self._excel_grid_model.columnCount()} colunas."
                ),
                tone="success",
            )
        if self._excel_dataframe is None and self._excel_load_thread is None:
            self._refresh_sheet_preview()
        result = QDialog.Rejected
        use_selected_header = False
        try:
            result = dialog.exec()
            use_selected_header = dialog.header_checkbox.isChecked()
        finally:
            self._active_excel_workspace = None
            self._excel_selection_dialog = None
            dialog.deleteLater()

        if result == QDialog.Accepted:
            self._confirm_excel_dialog_selection(use_selected_header=use_selected_header)
        else:
            self._restore_excel_dialog_snapshot(snapshot)

    def _set_connection_status(self, text: str, *, connected: bool, tone: str | None = None) -> None:
        self.connection_status_label.set_status(connected, text, tone=tone)
        if hasattr(self, "quick_page"):
            self.quick_page.connection_hint.setText(text)

    def _refresh_widget_style(self, widget: QWidget) -> None:
        widget.style().unpolish(widget)
        widget.style().polish(widget)
        widget.update()

    def _set_selected_sheet(self, sheet_name: str) -> None:
        if not self.excel_reader or not sheet_name:
            return
        matching_items = self.sheet_list.findItems(sheet_name, Qt.MatchExactly)
        if matching_items:
            self.sheet_list.blockSignals(True)
            self.sheet_list.setCurrentItem(matching_items[0])
            self.sheet_list.blockSignals(False)
            self._on_sheet_selected()

    def _on_workspace_sheet_changed(self, sheet_name: str) -> None:
        if not sheet_name or not self.excel_reader:
            return
        selected_items = self.sheet_list.selectedItems()
        if selected_items and selected_items[0].text() == sheet_name:
            return
        self._set_selected_sheet(sheet_name)

    def _sync_excel_workspaces(self) -> None:
        if not hasattr(self, "sheet_list"):
            return
        sheet_names = [self.sheet_list.item(index).text() for index in range(self.sheet_list.count())]
        selected_items = self.sheet_list.selectedItems()
        current_sheet = selected_items[0].text() if selected_items else None
        selection_text = self.selection_info_label.text() if hasattr(self, "selection_info_label") else self._selection_hint_text()
        applied_text = self._applied_selection_summary()
        if hasattr(self, "advanced_file_label"):
            self.advanced_file_label.setText(f"Arquivo atual: {self.excel_file_path or 'nenhum'}")
        if hasattr(self, "advanced_sheet_label"):
            self.advanced_sheet_label.setText(f"Aba atual: {current_sheet or 'nenhuma'}")
        if hasattr(self, "advanced_selection_label"):
            self.advanced_selection_label.setText(f"Recorte confirmado: {applied_text}")
        if hasattr(self, "open_advanced_selection_btn"):
            self.open_advanced_selection_btn.setEnabled(bool(self.excel_reader))
        if hasattr(self, "use_range_btn"):
            self.use_range_btn.setText("Recorte confirmado")
            self.use_range_btn.setToolTip(applied_text)
        for workspace in self._all_excel_workspaces():
            workspace.set_file_path(self.excel_file_path)
            workspace.set_sheet_options(sheet_names, current_sheet)
            workspace.set_selection_summary(selection_text)
            workspace.set_applied_summary(applied_text)
        if self._excel_selection_dialog is not None:
            self._excel_selection_dialog.set_footer_hint(
                "Selecione um bloco ou confirme o intervalo atual." if self._excel_dataframe is not None else "Aguarde a grade carregar."
            )

    def _summarize_range(self) -> str:
        if not all(
            hasattr(self, attr)
            for attr in ("header_row_spin", "row_start_spin", "row_end_spin", "col_start_spin", "col_end_spin")
        ):
            return "Aguardando configuracao do recorte."
        header_row = self._current_header_excel_row()
        data_start_row = self._current_data_start_excel_row(header_row)
        data_end_row = self.row_end_spin.value() or 0
        col_start = self.col_start_spin.value() or 1
        col_end = self.col_end_spin.value() or 0
        row_text = f"linhas {data_start_row} até {data_end_row or 'fim'}"
        col_text = f"colunas {col_start} até {col_end or 'fim'}"
        return f"Cabeçalho na linha {header_row}; {row_text}; {col_text}."

    def _sync_quick_workflow_state(self) -> None:
        if not hasattr(self, "quick_page"):
            return

        profile_items = [(profile.id, profile.name, profile.summary) for profile in self._available_profiles]
        self.quick_page.set_profile_items(profile_items)
        self.quick_page.profile_list.blockSignals(True)
        self.quick_page.select_profile(self.loaded_profile_id or self._selected_quick_profile_id())
        self.quick_page.profile_list.blockSignals(False)

        selected_profile = None
        selected_profile_id = self._selected_quick_profile_id() or self.profile_combo.currentData()
        if selected_profile_id:
            try:
                selected_profile = load_profile(str(selected_profile_id))
            except FileNotFoundError:
                selected_profile = None

        applied_profile = None
        if self.loaded_profile_id:
            try:
                applied_profile = load_profile(self.loaded_profile_id)
            except FileNotFoundError:
                applied_profile = None

        if applied_profile:
            model_detail = f"Modelo aplicado: {applied_profile.summary}"
        elif selected_profile:
            model_detail = f"Selecionado: {selected_profile.summary}. Clique em 'Usar este modelo' para aplicar."
        else:
            model_detail = "Nenhum modelo aplicado."
        self.quick_page.profile_detail_label.setText(model_detail)
        model_value = applied_profile.name if applied_profile else (selected_profile.name if selected_profile else "Nenhum modelo")
        model_state_detail = "Aplicado ao fluxo atual." if applied_profile else "Selecione e aplique um modelo salvo."
        self.quick_page.set_model_status(model_value, model_state_detail)

        file_value = self.excel_file_path.name if self.excel_file_path else "Nenhuma planilha"
        current_sheet = ""
        if hasattr(self, "sheet_list"):
            selected_items = self.sheet_list.selectedItems()
            if selected_items:
                current_sheet = selected_items[0].text()
        file_detail = f"Aba: {current_sheet}" if current_sheet else "Escolha um arquivo Excel."
        self.quick_page.set_file_status(file_value, file_detail)
        self.quick_page.file_name_label.setText(str(self.excel_file_path) if self.excel_file_path else "Nenhum arquivo selecionado.")
        self.quick_page.sheet_name_label.setText(current_sheet or "Nenhuma aba confirmada.")

        selection_summary = self._applied_selection_summary()
        selection_value = "Confirmado" if self._applied_excel_selection else ("Arquivo carregado" if self.excel_reader else "Pendente")
        selection_detail = selection_summary
        self.quick_page.set_selection_status(selection_value, selection_detail)
        self.quick_page.range_summary_label.setText(selection_summary)
        self.quick_page.spreadsheet_status_label.setProperty(
            "badgeTone",
            "success" if self._applied_excel_selection else ("warning" if self.excel_reader else "neutral"),
        )
        self.quick_page.spreadsheet_status_label.setText(
            "Recorte confirmado para validacao."
            if self._applied_excel_selection
            else ("Abra o seletor do Excel para revisar o recorte." if self.excel_reader else "Escolha um arquivo Excel.")
        )
        self._refresh_widget_style(self.quick_page.spreadsheet_status_label)
        self.quick_page.open_selection_btn.setEnabled(bool(self.excel_reader))
        self.quick_page.spreadsheet_next_btn.setEnabled(bool(self._applied_excel_selection))

        self._sync_excel_workspaces()

        validation = self._last_validation_result
        if validation:
            blocking = len(validation.issues)
            skipped = validation.skipped_rows
            self.quick_page.validation_cards["total"].set_value(str(validation.total_rows))
            self.quick_page.validation_cards["importable"].set_value(
                str(validation.importable_rows),
                tone="success" if validation.importable_rows else "warning",
            )
            self.quick_page.validation_cards["skipped"].set_value(str(skipped), tone="warning" if skipped else "neutral")
            self.quick_page.validation_cards["blocking"].set_value(
                str(blocking),
                tone="danger" if blocking else "success",
            )
            if blocking:
                self.quick_page.validation_status_label.setProperty("badgeTone", "danger")
                self.quick_page.validation_status_label.setText(
                    f"Foram encontrados {blocking} erro(s) bloqueantes. Ajuste o perfil ou use o modo avançado."
                )
                self._refresh_widget_style(self.quick_page.validation_status_label)
            else:
                tone = "warning" if skipped else "success"
                self.quick_page.validation_status_label.setProperty("badgeTone", tone)
                self.quick_page.validation_status_label.setText(
                    f"Validação pronta. Importáveis: {validation.importable_rows}. Descartadas: {validation.skipped_rows}."
                )
                self._refresh_widget_style(self.quick_page.validation_status_label)
            self.quick_page.validation_details.setPlainText(validation.preview_text())
            issue_rows = self._build_quick_issue_rows(validation)
            self.quick_page.populate_validation_rows(issue_rows)
        else:
            self.quick_page.validation_cards["total"].set_value("--")
            self.quick_page.validation_cards["importable"].set_value("--")
            self.quick_page.validation_cards["skipped"].set_value("--")
            self.quick_page.validation_cards["blocking"].set_value("--")
            self.quick_page.validation_status_label.setProperty("badgeTone", "neutral")
            self.quick_page.validation_status_label.setText("Validação ainda não executada.")
            self._refresh_widget_style(self.quick_page.validation_status_label)
            self.quick_page.validation_details.setPlainText(self._last_profile_preview or "")
            self.quick_page.populate_validation_rows([])

        _refresh_import = self._last_import_result
        skipped_rows = validation.skipped_rows if validation else 0
        self.quick_page.result_card.set_result(_refresh_import, skipped_rows=skipped_rows)

        connection_ready = bool(
            self.db_edit.text().strip() and self.user_edit.text().strip() and self.host_edit.text().strip()
        )
        checklist_lines = [
            f"Modelo aplicado: {'sim' if self.loaded_profile_id else 'nao'}",
            f"Planilha carregada: {'sim' if self.excel_reader else 'nao'}",
            f"Conexao configurada: {'sim' if connection_ready else 'nao'}",
            f"Validacao pronta: {'sim' if validation and validation.can_import else 'nao'}",
        ]
        self.quick_page.import_checklist_label.setText("\n".join(checklist_lines))
        can_import = bool(validation and validation.can_import and connection_ready)
        self.quick_page.import_btn.setEnabled(can_import)
        if can_import:
            self.quick_page.import_status_label.setProperty("badgeTone", "success")
            self.quick_page.import_status_label.setText("Tudo pronto para importar o lote.")
            self._refresh_widget_style(self.quick_page.import_status_label)
        elif validation and validation.can_import and not connection_ready:
            self.quick_page.import_status_label.setProperty("badgeTone", "warning")
            self.quick_page.import_status_label.setText("A validacao terminou, mas faltam os dados da conexao.")
            self._refresh_widget_style(self.quick_page.import_status_label)
        else:
            self.quick_page.import_status_label.setProperty("badgeTone", "warning")
            self.quick_page.import_status_label.setText("Ainda faltam itens obrigatorios antes do envio.")
            self._refresh_widget_style(self.quick_page.import_status_label)

        self._sync_quick_stepper(applied_profile, validation)
        self.quick_page.validation_next_btn.setEnabled(bool(validation and validation.can_import))

    def _summarize_range(self) -> str:
        if not all(
            hasattr(self, attr)
            for attr in ("header_row_spin", "row_start_spin", "row_end_spin", "col_start_spin", "col_end_spin")
        ):
            return "Aguardando configuracao do recorte."
        header_row = self._current_header_excel_row()
        data_start_row = self._current_data_start_excel_row(header_row)
        data_end_row = self.row_end_spin.value() or 0
        col_start = self.col_start_spin.value() or 1
        col_end = self.col_end_spin.value() or 0
        row_text = f"linhas {data_start_row} ate {data_end_row or 'fim'}"
        col_text = f"colunas {col_start} ate {col_end or 'fim'}"
        return f"Cabecalho na linha {header_row}; {row_text}; {col_text}."

    def _applied_selection_summary(self) -> str:
        if not self._applied_excel_selection:
            return f"Nenhum recorte aplicado. Intervalo atual: {self._summarize_range()}"
        selection = self._applied_excel_selection
        columns_text = ", ".join(selection.selected_columns[:4]) if selection.selected_columns else "colunas confirmadas"
        if len(selection.selected_columns) > 4:
            columns_text += f" ... (+{len(selection.selected_columns) - 4})"
        return (
            "Recorte aplicado: "
            f"cabecalho={selection.header_row}, linhas {selection.data_start_row}-{selection.data_end_row}, "
            f"colunas {selection.col_start}-{selection.col_end} | {columns_text}"
        )

    def _sync_quick_stepper(
        self,
        current_profile: ImportProfile | None,
        validation: XerifeValidationResult | None,
    ) -> None:
        step0_state = "ready" if current_profile else "active"
        step1_state = "ready" if self._applied_excel_selection else ("warning" if self.excel_reader else "pending")
        step2_state = "pending"
        step3_state = "pending"

        if validation:
            if validation.issues:
                step2_state = "error"
            elif validation.skipped_rows:
                step2_state = "warning"
            else:
                step2_state = "ready"
            step3_state = "ready" if self._last_import_result else "warning"

        self.quick_page.stepper.set_step_state(
            0,
            title="Modelo",
            detail=current_profile.name if current_profile else "Selecione e aplique um modelo salvo.",
            state=step0_state,
        )
        self.quick_page.stepper.set_step_state(
            1,
            title="Planilha",
            detail=(
                self.excel_file_path.name if self._applied_excel_selection and self.excel_file_path
                else "Escolha o arquivo e confirme o recorte na janela do Excel."
            ),
            state=step1_state,
        )
        self.quick_page.stepper.set_step_state(
            2,
            title="Revisão",
            detail=(
                f"{validation.importable_rows} importáveis, {validation.skipped_rows} descartadas."
                if validation
                else "Execute a validação do perfil."
            ),
            state=step2_state,
        )
        self.quick_page.stepper.set_step_state(
            3,
            title="Importar",
            detail=(
                "Lote já enviado nesta sessão."
                if self._last_import_result
                else "Envie o lote após uma validação sem bloqueios."
            ),
            state=step3_state,
        )

    def _build_quick_issue_rows(self, validation: XerifeValidationResult) -> list[tuple[str, str, str, str]]:
        rows: list[tuple[str, str, str, str]] = []
        for issue in validation.issues[:8]:
            rows.append((str(issue.row_number), "Bloqueante", issue.code, issue.message))
        for issue in validation.ignored_issues[:8]:
            rows.append((str(issue.row_number), "Descartada", issue.code, issue.message))
        return rows

    def _update_step_progress(self) -> None:
        step1 = f"{'[x]' if self._excel_step_ready else '[ ]'} Step 1 - Excel"
        step2 = f"{'[x]' if self._db_step_ready else '[ ]'} Step 2 - Banco de Dados"
        step3 = "[ ] Step 3 - Mapeamento"
        self.step_header_label.setText("  >  ".join([step1, step2, step3]))

        if not self._excel_step_ready:
            hint = "Step 1: importe o Excel e confirme o recorte para liberar o proximo passo."
        elif not self._db_step_ready:
            hint = "Step 2: va para Banco de Dados, conecte e escolha a tabela; depois siga para o mapeamento."
        else:
            hint = "Step 3: finalize o mapeamento, gere a pre-visualizacao ou execute a importacao."
        self.step_hint_label.setText(hint)

        if getattr(self, "excel_next_btn", None):
            self.excel_next_btn.setEnabled(self._excel_step_ready)
        if getattr(self, "db_next_btn", None):
            self.db_next_btn.setEnabled(self._db_step_ready)
        if getattr(self, "excel_step_hint", None):
            if self._excel_step_ready:
                self.excel_step_hint.setText("Recorte confirmado. Clique em Proximo para ir para Banco de Dados (Step 2).")
            else:
                self.excel_step_hint.setText("Escolha o arquivo e confirme o recorte do Excel para habilitar o proximo passo.")
        if getattr(self, "db_step_hint", None):
            if self._db_step_ready:
                self.db_step_hint.setText("Tabela selecionada. Clique em Proximo para seguir para o mapeamento (Step 3).")
            else:
                self.db_step_hint.setText("Conecte no banco e selecione uma tabela para habilitar o proximo passo.")

    def _set_excel_step_ready(self, ready: bool) -> None:
        self._excel_step_ready = ready
        self._update_step_progress()

    def _set_db_step_ready(self, ready: bool) -> None:
        self._db_step_ready = ready
        self._update_step_progress()

    def _build_excel_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.addWidget(self._build_excel_panel())

        excel_controls = QHBoxLayout()
        self.excel_step_hint = QLabel("Importe o Excel e confirme o recorte para seguir para o banco de dados.")
        self.excel_step_hint.setWordWrap(True)
        excel_controls.addWidget(self.excel_step_hint, 1)
        self.excel_next_btn = QPushButton("Proximo: Banco de Dados")
        self.excel_next_btn.setEnabled(False)
        self.excel_next_btn.clicked.connect(lambda: self.tabs.setCurrentIndex(1))
        excel_controls.addWidget(self.excel_next_btn)
        layout.addLayout(excel_controls)
        return tab

    def _build_database_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        connection_group = QGroupBox("Conexao")
        connection_layout = QVBoxLayout(connection_group)
        helper = QLabel(
            "A conexao principal fica na barra superior. "
            "Use este painel para confirmar o status e seguir para a escolha da tabela."
        )
        helper.setWordWrap(True)
        connection_layout.addWidget(helper)

        self.database_status_copy_label = QLabel("Sem conexao ativa.")
        self.database_status_copy_label.setWordWrap(True)
        connection_layout.addWidget(self.database_status_copy_label)
        layout.addWidget(connection_group)

        layout.addWidget(self._build_database_panel())

        db_controls = QHBoxLayout()
        self.db_step_hint = QLabel("Conecte e escolha a tabela para seguir ao mapeamento.")
        self.db_step_hint.setWordWrap(True)
        db_controls.addWidget(self.db_step_hint, 1)
        self.db_next_btn = QPushButton("Proximo: Mapeamento")
        self.db_next_btn.setEnabled(False)
        self.db_next_btn.clicked.connect(lambda: self.tabs.setCurrentIndex(2))
        db_controls.addWidget(self.db_next_btn)
        layout.addLayout(db_controls)
        return tab

    def _build_mapping_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        splitter = QSplitter(Qt.Horizontal)
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
        layout.setSpacing(16)

        self.sheet_list = QListWidget()
        self.sheet_list.hide()
        self.sheet_list.itemSelectionChanged.connect(self._on_sheet_selected)
        layout.addWidget(self.sheet_list)

        selection_card = QFrame()
        selection_card.setProperty("card", True)
        selection_layout = QVBoxLayout(selection_card)
        selection_layout.setContentsMargins(18, 18, 18, 18)
        selection_layout.setSpacing(10)

        selection_title = QLabel("Selecao atual da planilha")
        selection_title.setProperty("role", "card-title")
        selection_layout.addWidget(selection_title)

        selection_text = QLabel(
            "Use a janela dedicada do Excel para definir aba, cabecalho, linhas e colunas."
        )
        selection_text.setProperty("role", "muted")
        selection_text.setWordWrap(True)
        selection_layout.addWidget(selection_text)

        self.advanced_file_label = QLabel("Arquivo atual: nenhum")
        self.advanced_file_label.setWordWrap(True)
        selection_layout.addWidget(self.advanced_file_label)

        self.advanced_sheet_label = QLabel("Aba atual: nenhuma")
        self.advanced_sheet_label.setWordWrap(True)
        selection_layout.addWidget(self.advanced_sheet_label)

        self.advanced_selection_label = QLabel("Recorte confirmado: nenhum")
        self.advanced_selection_label.setWordWrap(True)
        selection_layout.addWidget(self.advanced_selection_label)

        selection_actions = QHBoxLayout()
        self.choose_advanced_file_btn = QPushButton("Escolher planilha")
        self.choose_advanced_file_btn.clicked.connect(self._choose_excel)
        selection_actions.addWidget(self.choose_advanced_file_btn)

        self.open_advanced_selection_btn = QPushButton("Abrir seletor do Excel")
        self.open_advanced_selection_btn.setProperty("variant", "primary")
        self.open_advanced_selection_btn.clicked.connect(self._open_excel_selection_dialog)
        selection_actions.addWidget(self.open_advanced_selection_btn)

        self.refresh_sheet_btn = QPushButton("Recarregar aba")
        self.refresh_sheet_btn.clicked.connect(self._refresh_sheet_preview)
        selection_actions.addWidget(self.refresh_sheet_btn)
        selection_actions.addStretch()
        selection_layout.addLayout(selection_actions)
        layout.addWidget(selection_card)

        manual_group = QGroupBox("Ajustes manuais do intervalo")
        manual_layout = QVBoxLayout(manual_group)
        manual_layout.setSpacing(12)

        header_grid = QGridLayout()
        header_grid.setHorizontalSpacing(12)
        header_grid.setVerticalSpacing(12)
        self.header_row_spin = QSpinBox()
        self.header_row_spin.setMinimum(1)
        self.header_row_spin.setValue(1)
        header_grid.addWidget(QLabel("Linha do cabecalho"), 0, 0)
        header_grid.addWidget(self.header_row_spin, 0, 1)

        self.row_start_spin = QSpinBox()
        self.row_start_spin.setMinimum(0)
        self.row_start_spin.setValue(0)
        header_grid.addWidget(QLabel("Linha inicial dos dados"), 0, 2)
        header_grid.addWidget(self.row_start_spin, 0, 3)

        self.row_end_spin = QSpinBox()
        self.row_end_spin.setMinimum(0)
        self.row_end_spin.setValue(0)
        header_grid.addWidget(QLabel("Linha final dos dados"), 1, 0)
        header_grid.addWidget(self.row_end_spin, 1, 1)

        self.col_start_spin = QSpinBox()
        self.col_start_spin.setMinimum(1)
        self.col_start_spin.setValue(1)
        header_grid.addWidget(QLabel("Coluna inicial"), 1, 2)
        header_grid.addWidget(self.col_start_spin, 1, 3)

        self.col_end_spin = QSpinBox()
        self.col_end_spin.setMinimum(0)
        self.col_end_spin.setValue(0)
        header_grid.addWidget(QLabel("Coluna final"), 1, 4)
        header_grid.addWidget(self.col_end_spin, 1, 5)
        manual_layout.addLayout(header_grid)

        self.selection_info_label = QLabel(self._selection_hint_text())
        self.selection_info_label.setWordWrap(True)
        manual_layout.addWidget(self.selection_info_label)

        manual_actions = QHBoxLayout()
        self.use_range_btn = QPushButton("Resumo do recorte")
        self.use_range_btn.setEnabled(False)
        manual_actions.addWidget(self.use_range_btn)
        manual_actions.addStretch()
        manual_layout.addLayout(manual_actions)

        layout.addWidget(manual_group)
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

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setSpacing(12)

        content_layout.addWidget(self._build_mapping_section())
        content_layout.addWidget(self._build_defaults_and_fk_section())
        content_layout.addWidget(self._build_operations_section())
        content_layout.addStretch()

        scroll_area.setWidget(content)
        layout.addWidget(scroll_area)

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

        self.pk_auto_checkbox = QCheckBox("PK gerada pelo banco (auto-incremento)")
        self.pk_auto_checkbox.setEnabled(False)
        self.pk_auto_checkbox.toggled.connect(self._on_pk_auto_toggled)
        section_layout.addWidget(self.pk_auto_checkbox)

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

        self.fk_trim_checkbox = QCheckBox("Remover espaços em branco no início/fim ao comparar (Excel vs banco)")
        self.fk_trim_checkbox.setChecked(True)
        self.fk_trim_checkbox.toggled.connect(lambda checked: setattr(self, "_fk_trim_whitespace", bool(checked)))
        fk_layout.addWidget(self.fk_trim_checkbox)

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

        conversion_layout = QHBoxLayout()
        self.fk_conversion_btn = QPushButton("Importar arquivo de conversões (FK)")
        self.fk_conversion_btn.clicked.connect(self._load_fk_conversion_file)
        conversion_layout.addWidget(self.fk_conversion_btn)
        self.fk_conversion_status = QLabel("Sem conversões carregadas")
        self.fk_conversion_status.setWordWrap(True)
        conversion_layout.addWidget(self.fk_conversion_status, 1)
        fk_layout.addLayout(conversion_layout)

        section_layout.addWidget(fk_group)
        return section

    def _build_operations_section(self) -> QWidget:
        section = QWidget()
        section_layout = QVBoxLayout(section)
        section_layout.setSpacing(10)

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

        similarity_layout = QHBoxLayout()
        self.similarity_btn = QPushButton("Verificar palavras parecidas...")
        self.similarity_btn.clicked.connect(self._open_similarity_validation)
        similarity_layout.addWidget(self.similarity_btn)
        self.similarity_status = QLabel("Padronização: inativa")
        self.similarity_status.setWordWrap(True)
        similarity_layout.addWidget(self.similarity_status, 1)
        similarity_layout.addStretch()
        section_layout.addLayout(similarity_layout)

        conversion_layout = QHBoxLayout()
        self.export_conversion_btn = QPushButton("Gerar arquivo de conversões")
        self.export_conversion_btn.setEnabled(False)
        self.export_conversion_btn.clicked.connect(self._export_similarity_file_action)
        conversion_layout.addWidget(self.export_conversion_btn)
        conversion_layout.addStretch()
        section_layout.addLayout(conversion_layout)

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
            self, "Selecione o arquivo Excel", str(Path.home()), "Planilhas (*.xls *.xlsx *.xlsm)"
        )
        if not file_name:
            return
        self._load_excel_file(Path(file_name), open_dialog=True)

    def _load_excel_file(self, file_path: Path, *, open_dialog: bool = False) -> None:
        self.excel_file_path = file_path
        try:
            self.excel_reader = ExcelReader(str(file_path))
            self._relation_conversions = {}
            self._refresh_fk_conversion_hint()
            self.sheet_list.clear()
            for name in self.excel_reader.sheet_names():
                self.sheet_list.addItem(name)
            self._set_excel_step_ready(False)
            self._excel_grid_model.clear()
            self._excel_dataframe = None
            self._excel_grid_headers = []
            self._excel_grid_selection_model.clearSelection()
            self.selection_info_label.setText(self._selection_hint_text())
            self.sheet_columns_list.clear()
            self._refresh_fk_excel_options()
            self._clear_pre_validation_state()
            self._last_validation_result = None
            self._last_import_result = None
            self._last_profile_preview = ""
            for workspace in self._all_excel_workspaces():
                workspace.set_file_path(file_path)
                workspace.set_loading("Carregando planilha...", "Preparando abas e intervalo inicial.")

            selected_profile: ImportProfile | None = None
            target_sheet = None
            selected_profile_id = self._selected_profile_id()
            if selected_profile_id:
                try:
                    selected_profile = load_profile(selected_profile_id)
                    target_sheet = selected_profile.sheet_name
                except FileNotFoundError:
                    target_sheet = None
                    selected_profile = None

            if selected_profile is not None:
                self.header_row_spin.setValue(selected_profile.header_row)
                self.row_start_spin.setValue(selected_profile.data_start_row or 0)
                self.row_end_spin.setValue(selected_profile.data_end_row or 0)
                self.col_start_spin.setValue(selected_profile.col_start or 1)
                self.col_end_spin.setValue(selected_profile.col_end or 0)
                self._applied_excel_selection = ExcelSelectionResult(
                    header_row=selected_profile.header_row,
                    data_start_row=selected_profile.data_start_row or (selected_profile.header_row + 1),
                    data_end_row=selected_profile.data_end_row or (selected_profile.data_start_row or (selected_profile.header_row + 1)),
                    col_start=selected_profile.col_start or 1,
                    col_end=selected_profile.col_end or (selected_profile.col_start or 1),
                    selected_columns=[],
                )
            else:
                self._applied_excel_selection = None

            self.sheet_list.blockSignals(True)
            if target_sheet:
                matching_items = self.sheet_list.findItems(target_sheet, Qt.MatchExactly)
                if matching_items:
                    self.sheet_list.setCurrentItem(matching_items[0])
            if not self.sheet_list.selectedItems() and self.sheet_list.count() > 0:
                self.sheet_list.setCurrentRow(0)
            self.sheet_list.blockSignals(False)
            if self.sheet_list.selectedItems():
                self._on_sheet_selected()
            self._update_profile_summary()
            self._sync_quick_workflow_state()
            if open_dialog:
                self._open_excel_selection_dialog()
        except Exception as exc:  # noqa: BLE001
            self._show_error("Erro ao abrir Excel", exc)

    def _refresh_profile_options(self) -> None:
        current_profile_id = self.profile_combo.currentData() if hasattr(self, "profile_combo") else None
        profiles = list_profiles()
        self._available_profiles = profiles
        self.profile_combo.clear()
        for profile in profiles:
            self.profile_combo.addItem(profile.name, profile.id)
        if current_profile_id:
            for index in range(self.profile_combo.count()):
                if self.profile_combo.itemData(index) == current_profile_id:
                    self.profile_combo.setCurrentIndex(index)
                    break
        self._update_profile_summary()
        self._sync_quick_workflow_state()

    def _on_profile_combo_changed(self) -> None:
        current_profile_id = self.profile_combo.currentData()
        if self.loaded_profile_id and current_profile_id and str(current_profile_id) != self.loaded_profile_id:
            self.loaded_profile_id = None
        self._update_profile_summary()
        self._sync_quick_workflow_state()

    def _selected_profile_id(self) -> str | None:
        profile_id = self.loaded_profile_id or self.profile_combo.currentData() or self._selected_quick_profile_id()
        if not profile_id:
            return None
        return str(profile_id)

    def _apply_selected_profile(self) -> None:
        profile_id = self.profile_combo.currentData()
        if not profile_id:
            QMessageBox.warning(self, "Modelo", "Nenhum modelo salvo selecionado.")
            return
        profile = load_profile(str(profile_id))
        self._apply_profile(profile)

    def _apply_profile(self, profile: ImportProfile) -> None:
        self.loaded_profile_id = profile.id
        self.profile_id_edit.setText(profile.id)
        self.profile_name_edit.setText(profile.name)
        self.profile_target_combo.setCurrentText(profile.target_type)
        self.profile_filial_spin.setValue(profile.filial_id or 0)
        self.profile_usuario_spin.setValue(profile.usuario_id or 0)
        self.header_row_spin.setValue(profile.header_row)
        self.row_start_spin.setValue(profile.data_start_row or 0)
        self.row_end_spin.setValue(profile.data_end_row or 0)
        self.col_start_spin.setValue(profile.col_start or 1)
        self.col_end_spin.setValue(profile.col_end or 0)
        self._applied_excel_selection = ExcelSelectionResult(
            header_row=profile.header_row,
            data_start_row=profile.data_start_row or (profile.header_row + 1),
            data_end_row=profile.data_end_row or (profile.data_start_row or (profile.header_row + 1)),
            col_start=profile.col_start or 1,
            col_end=profile.col_end or (profile.col_start or 1),
            selected_columns=[],
        )
        self._last_validation_result = None
        self._last_import_result = None
        self._last_profile_preview = ""

        if self.excel_reader:
            matching_items = self.sheet_list.findItems(profile.sheet_name, Qt.MatchExactly)
            if matching_items:
                self.sheet_list.setCurrentItem(matching_items[0])
            self._refresh_sheet_preview()
        self._update_profile_summary(profile)
        self._sync_quick_workflow_state()

    def _build_profile_from_ui(self) -> ImportProfile | None:
        profile_id = self.profile_id_edit.text().strip()
        profile_name = self.profile_name_edit.text().strip()
        base_profile: ImportProfile | None = None
        selected_profile_id = self._selected_profile_id()
        if selected_profile_id:
            try:
                base_profile = load_profile(selected_profile_id)
            except FileNotFoundError:
                base_profile = None

        sheet_name = base_profile.sheet_name if base_profile else ""
        selected_items = self.sheet_list.selectedItems()
        if selected_items:
            sheet_name = selected_items[0].text()

        if not sheet_name:
            QMessageBox.warning(self, "Modelo", "Selecione uma aba da planilha para salvar o modelo.")
            return None
        if not profile_id:
            QMessageBox.warning(self, "Modelo", "Informe o ID do modelo.")
            return None
        if not profile_name:
            QMessageBox.warning(self, "Modelo", "Informe o nome do modelo.")
            return None

        defaults = dict(base_profile.defaults if base_profile else {})
        defaults["usuario_id"] = self.profile_usuario_spin.value() or None

        return ImportProfile(
            id=profile_id,
            name=profile_name,
            target_type=self.profile_target_combo.currentText(),
            description=base_profile.description if base_profile else None,
            sheet_name=sheet_name,
            header_row=self._current_header_excel_row(),
            data_start_row=self.row_start_spin.value() or None,
            data_end_row=self.row_end_spin.value() or None,
            col_start=self.col_start_spin.value() or None,
            col_end=self.col_end_spin.value() or None,
            filial_id=self.profile_filial_spin.value() or None,
            usuario_id=self.profile_usuario_spin.value() or None,
            source_key_strategy=list(base_profile.source_key_strategy if base_profile else []),
            field_map=dict(base_profile.field_map if base_profile else {}),
            defaults=defaults,
            filters=dict(base_profile.filters if base_profile else {}),
            unit_rules=list(base_profile.unit_rules if base_profile else []),
            table_name=base_profile.table_name if base_profile else None,
            operation=base_profile.operation if base_profile else None,
        )

    def _save_current_profile(self) -> None:
        profile = self._build_profile_from_ui()
        if not profile:
            return
        save_path = save_profile(profile)
        self.loaded_profile_id = profile.id
        self._refresh_profile_options()
        self._update_profile_summary(profile)
        QMessageBox.information(self, "Modelo", f"Modelo salvo em:\n{save_path}")

    def _connection_payload(self) -> Dict[str, Any] | None:
        host = self.host_edit.text().strip()
        database = self.db_edit.text().strip()
        user = self.user_edit.text().strip()
        password = self.pwd_edit.text()
        if not all([host, database, user]):
            QMessageBox.warning(
                self,
                "Banco",
                "Preencha host, database, usuario e senha na conexao antes de importar para o Xerife.",
            )
            return None
        return {
            "host": host,
            "port": int(self.port_edit.text() or "5432"),
            "database": database,
            "user": user,
            "password": password,
            "adminUser": user,
            "adminPassword": password,
            "mode": "online",
        }

    def _validate_profile(self, *, show_blocking_dialog: bool = False) -> XerifeValidationResult | None:
        if not self.excel_reader:
            QMessageBox.warning(self, "Importacao", "Escolha a planilha antes de validar o modelo.")
            return None
        if not self._applied_excel_selection:
            QMessageBox.warning(
                self,
                "Importacao",
                "Confirme a aba e o recorte na janela do Excel antes de validar o modelo.",
            )
            return None
        profile = self._build_profile_from_ui()
        if not profile:
            return None
        if profile.target_type != "xerife_stock":
            QMessageBox.warning(self, "Importacao", f"Alvo rapido nao suportado: {profile.target_type}")
            return None

        try:
            validation = XerifeStockImporter(self.excel_reader, profile).validate()
            self._last_validation_result = validation
            self._last_import_result = None
            self._last_profile_preview = validation.preview_text()
            self.preview_text.setPlainText(validation.preview_text())
            self._update_profile_summary(profile)
            self._sync_quick_workflow_state()
            if validation.issues and show_blocking_dialog:
                QMessageBox.warning(
                    self,
                    "Validacao",
                    f"Foram encontrados {len(validation.issues)} erro(s) bloqueantes. Corrija-os antes de importar.",
                )
            return validation
        except Exception as exc:  # noqa: BLE001
            self._show_error("Erro ao validar o perfil do Xerife", exc)
            return None

    def _validate_profile_only(self) -> None:
        validation = self._validate_profile(show_blocking_dialog=False)
        if validation is None:
            return
        self.quick_page.set_step(2)

    def _import_validated_profile(self, *, show_summary_dialog: bool = False) -> None:
        validation = self._last_validation_result
        if validation is None:
            validation = self._validate_profile(show_blocking_dialog=False)
        if validation is None:
            return
        if validation.issues:
            self.quick_page.set_step(2)
            return

        connection = self._connection_payload()
        if connection is None:
            self.quick_page.set_step(3)
            return

        confirmation_lines = [f"Enviar {validation.importable_rows} item(ns) validados para o Xerife?"]
        if validation.skipped_rows:
            confirmation_lines.append(
                f"{validation.skipped_rows} linha(s) serao descartadas pelas regras do perfil e nao serao importadas."
            )
        confirmation = QMessageBox.question(self, "Importar para o Xerife", "\n".join(confirmation_lines))
        if confirmation != QMessageBox.Yes:
            return

        try:
            result = run_xerife_stock_batch(
                connection=connection,
                items=validation.prepared_items,
                run_id=uuid4().hex[:12],
            )
            self._last_import_result = result
            result_lines = [
                validation.preview_text(),
                "",
                "Resultado do lote Xerife:",
                f"Criados: {result.get('created', 0)}",
                f"Atualizados: {result.get('updated', 0)}",
                f"Ajustados: {result.get('adjusted', 0)}",
            ]
            result_items = result.get("items", [])
            if result_items:
                result_lines.extend(["", pd.DataFrame(result_items[:20]).to_string(index=False)])
            result_text = "\n".join(result_lines)
            self.preview_text.setPlainText(result_text)
            self.quick_page.validation_details.setPlainText(result_text)
            self.quick_page.set_step(3)
            self._sync_quick_workflow_state()
            if show_summary_dialog:
                QMessageBox.information(
                    self,
                    "Importacao",
                    (
                        f"Importacao concluida.\nCriados: {result.get('created', 0)}\n"
                        f"Atualizados: {result.get('updated', 0)}\n"
                        f"Descartados pelo perfil: {validation.skipped_rows}"
                    ),
                )
        except Exception as exc:  # noqa: BLE001
            self._show_error("Erro ao importar para o Xerife", exc)

    def _validate_and_import_profile(self) -> None:
        validation = self._validate_profile(show_blocking_dialog=True)
        if not validation or not validation.can_import:
            return
        self._import_validated_profile(show_summary_dialog=True)

    def _update_profile_summary(self, profile: ImportProfile | None = None) -> None:
        current_profile = profile
        if current_profile is None:
            selected_profile_id = self._selected_profile_id()
            if selected_profile_id:
                try:
                    current_profile = load_profile(selected_profile_id)
                except FileNotFoundError:
                    current_profile = None

        file_text = str(self.excel_file_path) if self.excel_file_path else "sem planilha"
        if current_profile:
            summary = f"{current_profile.summary} | arquivo={file_text}"
        else:
            summary = f"Sem modelo aplicado | arquivo={file_text}"
        self.profile_summary_label.setText(summary)
        self._sync_quick_workflow_state()

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
            self._set_connection_status(connection_text, connected=True)
            if hasattr(self, "database_status_copy_label"):
                self.database_status_copy_label.setText(connection_text)
            self._sync_quick_workflow_state()
            QMessageBox.information(self, "Banco", "Conexão realizada com sucesso")
            return True
        except Exception as exc:  # noqa: BLE001
            self._show_error("Erro ao conectar", exc)
            self._set_connection_status("Erro ao conectar", connected=False, tone="danger")
            if hasattr(self, "database_status_copy_label"):
                self.database_status_copy_label.setText("Erro ao conectar")
            self._set_db_step_ready(False)
            self._sync_quick_workflow_state()
            return False

    def _load_tables(self) -> None:
        self.table_list.clear()
        self._set_db_step_ready(False)
        for table in self.database.list_tables():
            self.table_list.addItem(table)
        self._refresh_fk_table_options()

    def _on_sheet_selected(self) -> None:
        items = self.sheet_list.selectedItems()
        if not items:
            self._set_excel_step_ready(False)
            self._update_profile_summary()
            return
        self._clear_pre_validation_state()
        self._refresh_sheet_preview()
        self._update_profile_summary()

    def _refresh_sheet_preview(self) -> None:
        self._set_excel_step_ready(False)
        if not self.excel_reader:
            self._sync_quick_workflow_state()
            return
        items = self.sheet_list.selectedItems()
        if not items:
            self._sync_quick_workflow_state()
            return
        sheet_name = items[0].text()
        col_start = self.col_start_spin.value()
        col_end = self.col_end_spin.value() or None
        if col_end is not None and col_end < col_start:
            col_end = None
            self.col_end_spin.setValue(0)
        header_excel_row = self._current_header_excel_row()
        data_start_row = self._current_data_start_excel_row(header_excel_row)
        data_end_row = self.row_end_spin.value() or None
        if data_end_row is not None and data_end_row < data_start_row:
            data_end_row = None
            self.row_end_spin.setValue(0)
        self._start_excel_grid_load(
            sheet_name=sheet_name,
            header_row=header_excel_row,
            data_start_row=data_start_row,
            data_end_row=data_end_row,
            col_start=col_start,
            col_end=col_end,
        )

    def _start_excel_grid_load(
        self,
        *,
        sheet_name: str,
        header_row: int,
        data_start_row: int,
        data_end_row: int | None,
        col_start: int | None,
        col_end: int | None,
    ) -> None:
        if not self.excel_reader:
            return
        self._excel_load_job_id += 1
        job_id = self._excel_load_job_id
        self._excel_grid_model.clear()
        self._excel_dataframe = None
        self._excel_grid_headers = []
        self._excel_grid_selection_model.clearSelection()
        for workspace in self._all_excel_workspaces():
            workspace.set_loading("Carregando planilha...", f"Aba {sheet_name}: preparando grade.")

        worker = ExcelLoadWorker(
            reader=self.excel_reader,
            sheet_name=sheet_name,
            header_row=header_row,
            data_start_row=data_start_row,
            data_end_row=data_end_row,
            col_start=col_start,
            col_end=col_end,
        )
        thread = QThread(self)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.started.connect(lambda message, current_job=job_id: self._on_excel_load_started(current_job, message))
        worker.first_chunk_ready.connect(
            lambda headers, rows, first_row, total_rows, current_job=job_id: self._on_excel_first_chunk(
                current_job,
                headers,
                rows,
                first_row,
                total_rows,
            )
        )
        worker.rows_appended.connect(
            lambda rows, current_job=job_id: self._on_excel_rows_appended(current_job, rows)
        )
        worker.finished.connect(
            lambda dataframe, headers, loaded_header_row, first_row, total_rows, total_cols, current_job=job_id:
            self._on_excel_load_finished(
                current_job,
                dataframe,
                headers,
                loaded_header_row,
                first_row,
                total_rows,
                total_cols,
            )
        )
        worker.failed.connect(lambda message, current_job=job_id: self._on_excel_load_failed(current_job, message))
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.finished.connect(
            lambda current_thread=thread, current_worker=worker: self._clear_excel_load_worker_refs(
                current_thread,
                current_worker,
            )
        )
        self._excel_load_worker = worker
        self._excel_load_thread = thread
        thread.start()

    def _clear_excel_load_worker_refs(self, thread: QThread, worker: ExcelLoadWorker) -> None:
        if self._excel_load_thread is thread:
            self._excel_load_thread = None
        if self._excel_load_worker is worker:
            self._excel_load_worker = None

    def _on_excel_load_started(self, job_id: int, message: str) -> None:
        if job_id != self._excel_load_job_id:
            return
        for workspace in self._all_excel_workspaces():
            workspace.set_loading("Carregando planilha...", message)

    def _on_excel_first_chunk(
        self,
        job_id: int,
        headers: list[str],
        rows: list[list[str]],
        first_data_row: int,
        total_rows: int,
    ) -> None:
        if job_id != self._excel_load_job_id:
            return
        self._excel_grid_headers = list(headers)
        self._excel_grid_first_data_row = first_data_row
        self._current_first_data_row = first_data_row
        self._excel_grid_model.reset_content(headers, rows, first_data_row)
        self.sheet_columns_list.clear()
        for column_name in headers:
            self.sheet_columns_list.addItem(column_name)
        self._refresh_fk_excel_options()
        if total_rows:
            status_text = f"Montando grade... {len(rows)}/{total_rows} linhas"
        else:
            status_text = "Intervalo vazio. Ajuste o recorte."
        for workspace in self._all_excel_workspaces():
            workspace.show_table(status_text=status_text, tone="warning" if total_rows else "danger")
        self._update_selection_info()
        self._sync_excel_workspaces()

    def _on_excel_rows_appended(self, job_id: int, rows: list[list[str]]) -> None:
        if job_id != self._excel_load_job_id:
            return
        self._excel_grid_model.append_rows(rows)
        current_rows = self._excel_grid_model.rowCount()
        for workspace in self._all_excel_workspaces():
            workspace.set_status(f"Montando grade... {current_rows} linhas", tone="warning")

    def _on_excel_load_finished(
        self,
        job_id: int,
        dataframe: pd.DataFrame,
        headers: list[str],
        header_row: int,
        first_data_row: int,
        total_rows: int,
        total_cols: int,
    ) -> None:
        if job_id != self._excel_load_job_id:
            return
        self._excel_dataframe = dataframe
        self._excel_grid_headers = list(headers)
        self._excel_grid_header_row = header_row
        self._excel_grid_first_data_row = first_data_row
        self._current_header_excel_row_value = header_row
        self._current_first_data_row = first_data_row
        self._set_excel_step_ready(True)
        self._sync_quick_workflow_state()
        if total_rows <= 0 or total_cols <= 0:
            self.selection_info_label.setText("Nenhum dado no intervalo atual. Ajuste cabecalho, linhas ou colunas.")
            for workspace in self._all_excel_workspaces():
                workspace.show_table(status_text="Sem dados no intervalo atual.", tone="danger")
            return
        ready_text = f"Grade pronta: {total_rows} linhas x {total_cols} colunas."
        for workspace in self._all_excel_workspaces():
            workspace.show_table(status_text=ready_text, tone="success")
        self._update_selection_info()
        self._sync_excel_workspaces()

    def _on_excel_load_failed(self, job_id: int, message: str) -> None:
        if job_id != self._excel_load_job_id:
            return
        for workspace in self._all_excel_workspaces():
            workspace.set_loading("Falha ao abrir a grade", message)
            workspace.set_status("Falha ao carregar planilha", tone="danger")
        QMessageBox.warning(self, "Excel", message)
        self._sync_quick_workflow_state()

    def _current_excel_selection(self) -> ExcelSelectionResult | None:
        selection = self._excel_grid_selection_model.selection()
        if selection.isEmpty():
            return None
        min_row = min(item.top() for item in selection)
        max_row = max(item.bottom() for item in selection)
        min_col = min(item.left() for item in selection)
        max_col = max(item.right() for item in selection)
        current_col_start = self.col_start_spin.value() or 1
        selected_columns = self._excel_grid_headers[min_col : max_col + 1]
        return ExcelSelectionResult(
            header_row=self._current_header_excel_row(),
            data_start_row=self._excel_row_from_table_row(min_row),
            data_end_row=self._excel_row_from_table_row(max_row),
            col_start=current_col_start + min_col,
            col_end=current_col_start + max_col,
            selected_columns=selected_columns,
        )

    def _apply_selection_to_range(self) -> None:
        selection = self._current_excel_selection()
        if not selection:
            return
        self._applied_excel_selection = selection
        self.row_start_spin.setValue(selection.data_start_row)
        self.row_end_spin.setValue(selection.data_end_row)
        self.col_start_spin.setValue(selection.col_start)
        self.col_end_spin.setValue(selection.col_end)
        self._last_validation_result = None
        self._last_import_result = None
        self._sync_excel_workspaces()
        self._update_profile_summary()

    def _apply_selection_to_header(self) -> None:
        selection = self._current_excel_selection()
        if not selection:
            return
        target_header = max(1, selection.data_start_row)
        self.header_row_spin.setValue(target_header)
        self.row_start_spin.setValue(target_header + 1)
        self._applied_excel_selection = None
        self._refresh_sheet_preview()

    def _clear_excel_selection(self) -> None:
        self._excel_grid_selection_model.clearSelection()
        self._update_selection_info()

    def _update_selection_info(self) -> None:
        selection = self._current_excel_selection()
        if not selection:
            self.selection_info_label.setText(self._selection_hint_text())
            self.selection_info_label.setToolTip("")
            self._sync_excel_workspaces()
            return
        col_labels = [
            f"{column_name} (col {selection.col_start + index})"
            for index, column_name in enumerate(selection.selected_columns)
        ]
        col_brief = self._compact_columns_text(col_labels, limit=6)
        text = (
            f"Linhas Excel: {selection.data_start_row} - {selection.data_end_row} | "
            f"Colunas ({len(col_labels)}): {col_brief}"
        )
        self.selection_info_label.setText(text)
        self.selection_info_label.setToolTip(", ".join(col_labels))
        self._sync_excel_workspaces()

    def _current_header_excel_row(self) -> int:
        """Return header row in Excel (1-based), clamped to at least 1."""
        if not hasattr(self, "header_row_spin"):
            return 1
        header_excel_row = self.header_row_spin.value()
        if header_excel_row < 1:
            header_excel_row = 1
            self.header_row_spin.setValue(header_excel_row)
        return header_excel_row

    def _current_data_start_excel_row(self, header_excel_row: Optional[int] = None) -> int:
        if not hasattr(self, "row_start_spin"):
            return (header_excel_row or 1) + 1
        header_value = header_excel_row or self._current_header_excel_row()
        explicit_start = self.row_start_spin.value()
        if explicit_start <= 0:
            return header_value + 1
        if explicit_start <= header_value:
            explicit_start = header_value + 1
            self.row_start_spin.setValue(explicit_start)
        return explicit_start

    def _selection_hint_text(self) -> str:
        return "Selecione celulas na grade com mouse ou Shift + setas. Confirme o recorte na janela do Excel."

    def _excel_row_from_table_row(self, row_idx: int) -> int:
        return self._excel_grid_first_data_row + row_idx

    def _normalize_lookup_key(self, value: object) -> str:
        if value is None:
            return ""
        try:
            if pd.isna(value):
                return ""
        except Exception:
            # pd.isna may not support the value type; ignore and continue.
            pass
        text = str(value)
        if self._fk_trim_whitespace:
            text = text.strip()
        return text.casefold()

    def _apply_fk_conversion(self, column: str, value: object) -> object:
        if value is None:
            return value
        mapping = self._relation_conversions.get(column)
        if not mapping:
            return value
        text = str(value)
        if self._fk_trim_whitespace:
            text = text.strip()
        return mapping.get(text, value)

    def _normalize_for_duplicates(self, value: object) -> str:
        if value is None:
            return ""
        try:
            if pd.isna(value):
                return ""
        except Exception:
            pass
        text = str(value)
        text = " ".join(text.split())
        normalized = unicodedata.normalize("NFKD", text)
        normalized = "".join(ch for ch in normalized if not unicodedata.combining(ch))
        return normalized.casefold()

    def _is_nullish(self, value: object) -> bool:
        if value is None:
            return True
        try:
            if pd.isna(value):
                return True
        except Exception:
            pass
        text = str(value)
        if not text.strip():
            return True
        return False

    def _normalize_similarity_text(self, text: str) -> str:
        cleaned = " ".join(text.split())
        cleaned = cleaned.rstrip("sS")
        normalized = unicodedata.normalize("NFKD", cleaned)
        normalized = "".join(ch for ch in normalized if not unicodedata.combining(ch))
        return normalized.casefold()

    def _are_values_similar(self, left: str, right: str) -> bool:
        norm_left = self._normalize_similarity_text(left)
        norm_right = self._normalize_similarity_text(right)
        if not norm_left or not norm_right:
            return False
        if norm_left == norm_right:
            return True
        ratio = SequenceMatcher(None, norm_left, norm_right).ratio()
        length_gap = abs(len(norm_left) - len(norm_right))
        if ratio >= 0.9:
            return True
        if length_gap <= 1 and ratio >= 0.82:
            return True
        return False

    def _trim_cell_whitespace(self, value: object) -> object:
        if isinstance(value, str):
            return value.strip()
        return value

    def _trim_dataframe_whitespace(self, df: pd.DataFrame, columns: Optional[List[str]] = None) -> pd.DataFrame:
        target_columns = columns or list(df.columns)
        for col in target_columns:
            if col not in df.columns:
                continue
            df[col] = df[col].map(self._trim_cell_whitespace)
        return df

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
        self._refresh_fk_conversion_hint()

    def _remove_virtual_extra_columns(self) -> None:
        if not self._virtual_extra_columns:
            return
        # Remove mappings that reference virtual columns
        rows_to_remove = []
        for row in range(self.mapping_table.rowCount()):
            sheet_item = self.mapping_table.item(row, 0)
            if sheet_item and sheet_item.text() in self._virtual_extra_columns:
                rows_to_remove.append(row)
        for row in sorted(rows_to_remove, reverse=True):
            self.mapping_table.removeRow(row)
        # Remove from sheet columns list
        keep_items: List[QListWidgetItem] = []
        for idx in range(self.sheet_columns_list.count()):
            item = self.sheet_columns_list.item(idx)
            if item and item.text() not in self._virtual_extra_columns:
                keep_items.append(item)
        self.sheet_columns_list.clear()
        for item in keep_items:
            self.sheet_columns_list.addItem(item.text())
        self._virtual_extra_columns.clear()
        self._refresh_fk_excel_options()

    def _next_extra_column_name(self) -> str:
        existing = {self.sheet_columns_list.item(i).text() for i in range(self.sheet_columns_list.count())}
        existing |= set(self._virtual_extra_columns)
        idx = 1
        while True:
            candidate = f"extra_{idx}"
            if candidate not in existing:
                return candidate
            idx += 1

    def _ensure_virtual_extra_column(self) -> str:
        name = self._next_extra_column_name()
        if name not in self._virtual_extra_columns:
            self._virtual_extra_columns.add(name)
            self.sheet_columns_list.addItem(name)
            self._refresh_fk_excel_options()
        return name

    def _refresh_fk_conversion_hint(self) -> None:
        if not getattr(self, "fk_conversion_status", None):
            return
        if not self._relation_conversions:
            self.fk_conversion_status.setText("Sem conversões carregadas")
            return
        total_cols = len(self._relation_conversions)
        total_maps = sum(len(m) for m in self._relation_conversions.values())
        parts = [f"{col}: {len(mapping)}" for col, mapping in list(self._relation_conversions.items())[:3]]
        if len(self._relation_conversions) > 3:
            parts.append("...")
        self.fk_conversion_status.setText(f"Conversões carregadas ({total_cols} col): " + ", ".join(parts))

    def _load_fk_conversion_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Importar arquivo de conversões (FK)", str(self.excel_file_path or Path.home()), "CSV (*.csv)"
        )
        if not path:
            return
        try:
            df = pd.read_csv(path)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "Conversões", f"Erro ao ler o CSV: {exc}")
            return
        required_cols = {"coluna_excel", "valor_original", "valor_corrigido"}
        if not required_cols.issubset(set(df.columns)):
            QMessageBox.warning(
                self,
                "Conversões",
                "CSV inválido. Esperado colunas: coluna_excel, valor_original, valor_corrigido.",
            )
            return
        conversions: Dict[str, Dict[str, str]] = {}
        for _, row in df.iterrows():
            col = str(row["coluna_excel"])
            src = row["valor_original"]
            dst = row["valor_corrigido"]
            if pd.isna(src) or pd.isna(dst):
                continue
            src_text = str(src)
            dst_text = str(dst)
            if self._fk_trim_whitespace:
                src_text = src_text.strip()
                dst_text = dst_text.strip()
            if not src_text.strip():
                continue
            bucket = conversions.setdefault(col, {})
            bucket[src_text] = dst_text
        self._relation_conversions = conversions
        self._last_conversion_file = Path(path)
        self._refresh_fk_conversion_hint()
        QMessageBox.information(
            self,
            "Conversões",
            f"Arquivo de conversões carregado.\nColunas: {len(conversions)} | Substituições: {sum(len(m) for m in conversions.values())}",
        )

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
        self._pre_validation_trim_whitespace = False
        self._pre_validation_column = None
        self._pre_validation_last_result = None
        self._split_enabled = False
        self._split_column = None
        self._split_operator = "gt"
        self._split_length = None
        self._split_extra_name = None
        self._remove_virtual_extra_columns()
        self._refresh_pre_validation_hint()
        self._clear_similarity_state()

    def _clear_similarity_state(self) -> None:
        self._similarity_replacements = {}
        self._refresh_similarity_hint()
        self._refresh_fk_conversion_hint()

    def _refresh_pre_validation_hint(self) -> None:
        if not getattr(self, "pre_validation_status", None):
            return
        rules: List[str] = []
        tooltips: List[str] = []
        if self._pre_validation_trim_whitespace:
            rules.append("remover espacos extras")
            tooltips.append("Remove espacos em branco no inicio/fim de valores textuais.")
        if self._pre_validation_remove_duplicates and self._pre_validation_column:
            rule = f"remover duplicados em '{self._pre_validation_column}'"
            if self._pre_validation_trim_whitespace:
                rule += " (remove espacos e ignora maiusculas/minusculas)"
            else:
                rule += " (ignora espacos extras e maiusculas/minusculas)"
            rules.append(rule)
            if self._pre_validation_last_result:
                total, unique = self._pre_validation_last_result
                removed = max(total - unique, 0)
                rules[-1] += f" (previsto remover {removed}/{total})"
                tooltips.append(f"Duplicados: Total {total} | Removidos {removed} | Restantes {unique}")
        if self._split_enabled and self._split_column and self._split_length is not None:
            op_text = ">" if self._split_operator == "gt" else "<"
            rule = f"mover valores de '{self._split_column}' cujo tamanho {op_text} {self._split_length} p/ coluna extra"
            rules.append(rule)
            extra = next(iter(self._virtual_extra_columns), "extra_1") if self._virtual_extra_columns else "extra_1"
            tooltips.append(f"Cria coluna {extra} com os valores deslocados; coluna original fica vazia nessas linhas.")
        if not rules:
            self.pre_validation_status.setText("Pré-validação: sem regras")
            self.pre_validation_status.setToolTip("")
            return
        text = "Pré-validação: " + " | ".join(rules)
        tooltip = "\n".join(tooltips) if tooltips else text
        self.pre_validation_status.setText(text)
        self.pre_validation_status.setToolTip(tooltip)

    def _refresh_similarity_hint(self) -> None:
        if not getattr(self, "similarity_status", None):
            return
        # Reset text/tooltip before recomputing to avoid stale content.
        self.similarity_status.setText("Padronização: inativa")
        self.similarity_status.setToolTip("")
        if getattr(self, "export_conversion_btn", None):
            self.export_conversion_btn.setEnabled(False)
        if not self._similarity_replacements:
            self._refresh_fk_conversion_hint()
            return
        parts: List[str] = []
        tooltips: List[str] = []
        total = 0
        for col, mapping in self._similarity_replacements.items():
            count = len(mapping)
            total += count
            parts.append(f"{col}: {count} trocas")
            preview_items = [f"{src} -> {dst}" for src, dst in list(mapping.items())[:3]]
            if len(mapping) > 3:
                preview_items.append(f"... (+{len(mapping) - 3})")
            tooltips.append(f"{col}: " + ", ".join(preview_items))
        self.similarity_status.setText(f"Padronização: {total} substituições")
        self.similarity_status.setToolTip("\n".join(tooltips))
        if getattr(self, "export_conversion_btn", None):
            self.export_conversion_btn.setEnabled(True)
        self._refresh_fk_conversion_hint()

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
        self._clear_similarity_state()
        self._last_skipped_null_rows = 0
        self._last_conversion_file = None
        self._relation_conversions = {}
        self._refresh_fk_conversion_hint()

    def _on_table_selected(self) -> None:
        items = self.table_list.selectedItems()
        if not items:
            self._set_db_step_ready(False)
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
        self._set_db_step_ready(True)

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
        if not self.update_radio.isChecked():
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
            start_row=self.row_start_spin.value() or None,
            end_row=self.row_end_spin.value() or None,
            start_column=self.col_start_spin.value(),
            end_column=self.col_end_spin.value() or None,
            column_mapping=mapping,
            default_values=defaults,
            fk_lookups=fk_lookups,
            operation="UPDATE" if self.update_radio.isChecked() else "INSERT",
            join_column=join_column,
            primary_key=self.primary_key_column,
            autogenerate_pk=autogenerate_pk,
            trim_whitespace=self._pre_validation_trim_whitespace,
            remove_duplicate_rows=remove_duplicates,
            duplicate_check_column=duplicate_column,
            similarity_replacements={col: mapping.copy() for col, mapping in self._similarity_replacements.items()},
            split_column=self._split_column if self._split_enabled else None,
            split_operator=self._split_operator if self._split_enabled else None,
            split_length=self._split_length if self._split_enabled else None,
            split_extra_column=self._split_extra_name if self._split_enabled else None,
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
            data_start_row=selection.start_row,
            data_end_row=selection.end_row,
            col_start=selection.start_column,
            col_end=selection.end_column,
        )
        if column not in df.columns:
            raise ValueError(f"Coluna '{column}' não encontrada na seleção atual")
        # Sempre ignora espaços extras no cálculo de duplicados para não contar valores iguais como distintos.
        df = self._trim_dataframe_whitespace(df, [column])
        if self._similarity_replacements:
            df = self._apply_similarity_replacements(df, self._similarity_replacements)
        normalized_series = df[column].map(self._normalize_for_duplicates)
        total_rows = len(normalized_series.index)
        unique_rows = len(normalized_series.drop_duplicates().index)
        return total_rows, unique_rows

    def _calculate_similarity_suggestions(
        self, selection: MappingSelection, column: str
    ) -> tuple[List[tuple[str, str, int]], int]:
        if not self.excel_reader:
            raise ValueError("Nenhuma planilha carregada")
        df = self.excel_reader._read_dataframe(
            selection.sheet_name,
            selection.header_row,
            data_start_row=selection.start_row,
            data_end_row=selection.end_row,
            col_start=selection.start_column,
            col_end=selection.end_column,
        )
        if column not in df.columns:
            raise ValueError(f"Coluna '{column}' não encontrada na seleção atual")
        if self._pre_validation_trim_whitespace:
            df = self._trim_dataframe_whitespace(df, [column])

        values: List[str] = []
        for raw in df[column]:
            if raw is None:
                continue
            try:
                if pd.isna(raw):
                    continue
            except Exception:
                pass
            text = str(raw).strip()
            if not text:
                continue
            values.append(text)
        counts = Counter(values)
        if len(counts) <= 1:
            return [], len(values)
        if len(counts) > 800:
            raise ValueError("Coluna com muitas variações para análise (limite: 800 valores distintos).")
        suggestions = self._build_similarity_suggestions(counts)
        return suggestions, len(values)

    def _build_similarity_suggestions(self, counts: Counter[str]) -> List[tuple[str, str, int]]:
        unique_values = list(counts.keys())
        parent: Dict[str, str] = {value: value for value in unique_values}

        def find(value: str) -> str:
            while parent[value] != value:
                parent[value] = parent[parent[value]]
                value = parent[value]
            return value

        def union(a: str, b: str) -> None:
            root_a = find(a)
            root_b = find(b)
            if root_a == root_b:
                return
            if counts[root_a] < counts[root_b]:
                root_a, root_b = root_b, root_a
            parent[root_b] = root_a

        for idx, left in enumerate(unique_values):
            for right in unique_values[idx + 1 :]:
                if self._are_values_similar(left, right):
                    union(left, right)

        groups: Dict[str, List[str]] = {}
        for value in unique_values:
            root = find(value)
            groups.setdefault(root, []).append(value)

        suggestions: List[tuple[str, str, int]] = []
        for members in groups.values():
            if len(members) <= 1:
                continue
            canonical = sorted(members, key=lambda v: (-counts[v], -len(v), v))[0]
            for member in sorted(members):
                if member == canonical:
                    continue
                suggestions.append((member, canonical, counts[member]))

        suggestions.sort(key=lambda item: (item[1].casefold(), item[0].casefold()))
        return suggestions

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
            trim_whitespace=self._pre_validation_trim_whitespace,
            selected_column=self._pre_validation_column,
            last_result=self._pre_validation_last_result,
            run_check=lambda col: self._calculate_duplicate_stats(selection, col),
            split_enabled=self._split_enabled,
            split_column=self._split_column,
            split_operator=self._split_operator,
            split_length=self._split_length,
        )
        if dialog.exec():
            self._pre_validation_remove_duplicates = dialog.remove_duplicates
            self._pre_validation_trim_whitespace = dialog.trim_whitespace
            self._pre_validation_column = dialog.selected_column if dialog.remove_duplicates else None
            self._pre_validation_last_result = dialog.last_result if dialog.remove_duplicates else None
            self._split_enabled = dialog.split_enabled
            self._split_column = dialog.split_column if dialog.split_enabled else None
            self._split_operator = dialog.split_operator if dialog.split_enabled else "gt"
            self._split_length = dialog.split_length if dialog.split_enabled else None
            if self._split_enabled and self._split_column and self._split_length is not None:
                self._split_extra_name = self._ensure_virtual_extra_column()
            else:
                self._split_extra_name = None
                self._remove_virtual_extra_columns()
            self._refresh_pre_validation_hint()

    def _open_similarity_validation(self) -> None:
        selection = self._collect_mapping()
        if not selection or not self.excel_reader:
            return
        sheet_items = self.sheet_columns_list.selectedItems()
        if not sheet_items:
            QMessageBox.warning(self, "Padronização", "Selecione uma coluna do Excel para validar.")
            return
        column = sheet_items[0].text()
        try:
            suggestions, total_checked = self._calculate_similarity_suggestions(selection, column)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "Padronização", str(exc))
            return
        if not suggestions:
            QMessageBox.information(self, "Padronização", "Nenhuma variação encontrada nessa coluna.")
            self._refresh_similarity_hint()
            return
        dialog = SimilarValuesDialog(self, column, suggestions, total_checked)
        if dialog.exec():
            self._similarity_replacements[column] = {src: dst for src, dst, _ in suggestions}
            self._refresh_similarity_hint()

    def _export_similarity_file_action(self) -> None:
        if not self._similarity_replacements:
            QMessageBox.information(self, "Conversões", "Nenhuma conversão de palavras encontrada para exportar.")
            return
        try:
            path = self._write_similarity_conversion_file(self._similarity_replacements)
            if path:
                QMessageBox.information(self, "Conversões", f"Arquivo de conversões salvo em:\n{path}")
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "Conversões", str(exc))

    def _generate_preview(self) -> None:
        selection = self._collect_mapping()
        if not selection or not self.excel_reader:
            return
        try:
            preview = self.excel_reader.load_sheet_preview(
                selection.sheet_name,
                header_row=selection.header_row,
                data_start_row=selection.start_row,
                data_end_row=selection.end_row,
                col_start=selection.start_column,
                col_end=selection.end_column,
            )
            sql_example = self._build_sql_example(selection)
            text = ["Pré-visualização de dados:", preview.sample.head().to_string()]
            if selection.trim_whitespace:
                text.append("")
                text.append("Limpeza: remover espacos em branco no inicio/fim das celulas de texto.")
            if selection.default_values:
                text.append("")
                text.append("Valores padrão aplicados:")
                for col, value in selection.default_values.items():
                    text.append(f"- {col}: {value}")
            if selection.similarity_replacements:
                text.append("")
                text.append("Padronização de texto:")
                for col, mapping in selection.similarity_replacements.items():
                    example = next(iter(mapping.items()), None)
                    suffix = f" (ex.: {example[0]} -> {example[1]})" if example else ""
                    text.append(f"- {col}: {len(mapping)} substituições{suffix}")
                conv_path = self._similarity_conversion_path()
                if conv_path:
                    text.append(f"Arquivo de conversões sugerido: {conv_path}")
            if selection.split_column and selection.split_length is not None and selection.split_extra_column:
                text.append("")
                op = ">" if selection.split_operator == "gt" else "<"
                text.append(
                    f"Separação por tamanho: mover '{selection.split_column}' com len {op} {selection.split_length} para '{selection.split_extra_column}'"
                )
            if selection.fk_lookups:
                text.append("")
                text.append("Relacionamentos (descrição -> ID):")
                for fk in selection.fk_lookups:
                    text.append(
                        f"- {fk.target_column} <= {fk.foreign_table}.{fk.foreign_id_column} via {fk.foreign_label_column} = Excel[{fk.excel_column}]"
                    )
            if selection.remove_duplicate_rows and selection.duplicate_check_column:
                text.append("")
                duplicate_summary = (
                    "Remover duplicados em "
                    + selection.duplicate_check_column
                    + " (ignora espaços no início/fim)"
                )
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
        first_excel_row = selection.start_row or (selection.header_row + 1)
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
        progress: QProgressDialog | None = None
        try:
            self._cancel_requested = False
            progress = self._create_progress_dialog("Importação", "Processando dados e enviando para o banco...")
            records = self._build_records_for_selection(
                selection, cancel_checker=lambda: bool(self._cancel_requested)
            )
            if self._cancel_requested:
                QMessageBox.information(self, "Importação", "Operação cancelada.")
                return
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
            msg = f"Registros processados: {affected}"
            if self._last_skipped_null_rows:
                msg += f"\nLinhas ignoradas por estarem vazias: {self._last_skipped_null_rows}"
            if selection.similarity_replacements:
                try:
                    path = self._write_similarity_conversion_file(selection.similarity_replacements)
                    if path:
                        msg += f"\nArquivo de conversões: {path}"
                except Exception as conv_exc:  # noqa: BLE001
                    QMessageBox.warning(
                        self,
                        "Conversões",
                        f"Não foi possível salvar o arquivo de conversões das palavras: {conv_exc}",
                    )
            QMessageBox.information(self, "Importação", msg)
            self._reset_after_execute()
        except RuntimeError as exc:
            if "cancelada" in str(exc).lower():
                QMessageBox.information(self, "Importação", "Operação cancelada.")
            else:
                self._show_error("Erro na importação", exc)
        except Exception as exc:  # noqa: BLE001
            self._show_error("Erro na importação", exc)
        finally:
            if progress:
                progress.close()
            self._cancel_requested = False

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

    def _apply_similarity_replacements(
        self, df: pd.DataFrame, replacements: Dict[str, Dict[str, str]]
    ) -> pd.DataFrame:
        if not replacements:
            return df
        for column, mapping in replacements.items():
            if column not in df.columns or not mapping:
                continue
            df[column] = df[column].map(lambda value: mapping.get(value, value) if isinstance(value, str) else value)
        return df

    def _apply_split_rule(self, df: pd.DataFrame, selection: MappingSelection) -> pd.DataFrame:
        if not selection.split_column or selection.split_length is None or not selection.split_extra_column:
            return df
        column = selection.split_column
        extra_col = selection.split_extra_column
        if column not in df.columns:
            raise ValueError(f"Coluna '{column}' não encontrada para separação por tamanho")
        # Avoid collision
        if extra_col in df.columns:
            extra_col = self._next_extra_column_name()
            selection.split_extra_column = extra_col
        df[extra_col] = None
        for idx, val in df[column].items():
            if val is None:
                continue
            try:
                if pd.isna(val):
                    continue
            except Exception:
                pass
            text = str(val)
            if selection.trim_whitespace:
                text = text.strip()
            length = len(text)
            if (selection.split_operator == "gt" and length > selection.split_length) or (
                selection.split_operator == "lt" and length < selection.split_length
            ):
                df.at[idx, extra_col] = val
                df.at[idx, column] = None
        # Register virtual column so UI knows
        if extra_col not in self._virtual_extra_columns:
            self._virtual_extra_columns.add(extra_col)
            self.sheet_columns_list.addItem(extra_col)
            self._refresh_fk_excel_options()
        return df

    def _similarity_conversion_path(self) -> Path | None:
        if not self.excel_file_path:
            return None
        base = Path(self.excel_file_path)
        if not self._similarity_replacements:
            return base.with_name(f"{base.stem}_conversoes.csv")
        # Usa a primeira coluna com conversões para compor o nome.
        first_col = next(iter(self._similarity_replacements.keys()), "coluna")
        col_slug = str(first_col).replace("/", "_").replace("\\", "_").replace(" ", "_")
        return base.with_name(f"{col_slug}_{base.stem}_conversoes.csv")

    def _write_similarity_conversion_file(
        self, replacements: Dict[str, Dict[str, str]], *, warn_if_missing_excel: bool = True
    ) -> Path | None:
        if not replacements:
            return None
        rows: List[Dict[str, str]] = []
        for column, mapping in replacements.items():
            for src, dst in mapping.items():
                rows.append(
                    {
                        "coluna_excel": column,
                        "valor_original": src,
                        "valor_corrigido": dst,
                    }
                )
        if not rows:
            return None
        target_path = self._similarity_conversion_path()
        if target_path is None:
            if warn_if_missing_excel:
                raise ValueError(
                    "Não foi possível determinar onde salvar o arquivo de conversões (precisa do Excel carregado)."
                )
            return None
        df = pd.DataFrame(rows)
        df.to_csv(target_path, index=False)
        self._last_conversion_file = target_path
        return target_path

    def _create_progress_dialog(self, title: str, label: str) -> QProgressDialog:
        dialog = QProgressDialog(label, "Cancelar", 0, 0, self)
        dialog.setWindowTitle(title)
        dialog.setWindowModality(Qt.WindowModal)
        dialog.setMinimumDuration(0)
        dialog.canceled.connect(lambda: self._request_cancel(dialog))
        dialog.show()
        QApplication.processEvents()
        return dialog

    def _request_cancel(self, dialog: QProgressDialog) -> None:
        self._cancel_requested = True
        dialog.setLabelText("Cancelando... aguarde")

    def _build_records_for_selection(
        self,
        selection: MappingSelection,
        cancel_checker: Optional[Callable[[], bool]] = None,
    ) -> List[Dict[str, object]]:
        # Carrega todas as colunas necessárias (mapeamento + lookups de FK)
        df = self.excel_reader._read_dataframe(
            selection.sheet_name,
            selection.header_row,
            data_start_row=selection.start_row,
            data_end_row=selection.end_row,
            col_start=selection.start_column,
            col_end=selection.end_column,
        )
        if selection.trim_whitespace:
            df = self._trim_dataframe_whitespace(df)
        if selection.similarity_replacements:
            df = self._apply_similarity_replacements(df, selection.similarity_replacements)
        df = self._apply_split_rule(df, selection)
        if selection.remove_duplicate_rows and selection.duplicate_check_column:
            if selection.duplicate_check_column not in df.columns:
                raise ValueError(
                    f"Coluna '{selection.duplicate_check_column}' não encontrada para remover duplicados"
                )
            dedup_key = f"__dupkey_{selection.duplicate_check_column}"
            df[dedup_key] = df[selection.duplicate_check_column].map(self._normalize_for_duplicates)
            total_rows = len(df.index)
            df = df.drop_duplicates(subset=[dedup_key], keep="first")
            df = df.reset_index(drop=True)
            df = df.drop(columns=[dedup_key])
            self._pre_validation_last_result = (total_rows, len(df.index))
        column_mapping = selection.column_mapping
        if selection.autogenerate_pk and selection.primary_key:
            column_mapping = [(s, t) for s, t in selection.column_mapping if t != selection.primary_key]
        needed_excel_cols = {s for s, _ in column_mapping} | {fk.excel_column for fk in selection.fk_lookups}
        missing_excel = [col for col in needed_excel_cols if col not in df.columns]
        if missing_excel:
            raise ValueError(f"Colunas da planilha não encontradas: {', '.join(missing_excel)}")

        records: List[Dict[str, object]] = []
        record_source_rows: List[tuple[int, pd.Series]] = []
        source_columns_for_null_check = [s for s, _ in column_mapping]
        skipped_null_rows = 0
        first_excel_row = selection.start_row or (selection.header_row + 1)
        for idx, (_, row) in enumerate(df.iterrows()):
            if source_columns_for_null_check:
                if all(self._is_nullish(row.get(col)) for col in source_columns_for_null_check):
                    skipped_null_rows += 1
                    continue
            if cancel_checker and cancel_checker():
                raise RuntimeError("Operação cancelada pelo usuário")
            record: Dict[str, object] = {}
            for sheet_col, table_col in column_mapping:
                if sheet_col in row:
                    record[table_col] = self.excel_reader._normalize_cell(row[sheet_col])
            records.append(record)
            record_source_rows.append((first_excel_row + idx, row))
        self._last_skipped_null_rows = skipped_null_rows

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
            for record_idx, (excel_row, row) in enumerate(record_source_rows):
                for fk in selection.fk_lookups:
                    raw_value = row.get(fk.excel_column)
                    raw_value = self._apply_fk_conversion(fk.excel_column, raw_value)
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
                        records[record_idx][fk.target_column] = mapped
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
                default_path = self._default_export_path("csv", selection)
                path, _ = QFileDialog.getSaveFileName(
                    self, "Salvar CSV mapeado", default_path, "CSV (*.csv)"
                )
                if not path:
                    return
                df.to_csv(path, index=False)
            else:
                default_path = self._default_export_path("xlsx", selection)
                path, _ = QFileDialog.getSaveFileName(
                    self, "Salvar Excel mapeado", default_path, "Excel (*.xlsx)"
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

    def _default_export_path(self, kind: str, selection: MappingSelection) -> str:
        sheet_slug = selection.sheet_name.replace("/", "_").replace("\\", "_").replace(" ", "_")
        if self.excel_file_path:
            base = self.excel_file_path
            name = f"{base.stem}_{sheet_slug}_mapeado.{kind}"
            return str(base.with_name(name))
        return str(Path.home() / f"planilha_{sheet_slug}_mapeado.{kind}")


class SimilarValuesDialog(QDialog):
    def __init__(
        self,
        parent: QWidget,
        column: str,
        suggestions: List[tuple[str, str, int]],
        total_values: int,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Padronizar palavras parecidas")
        self.resize(900, 600)
        self.setMinimumSize(720, 480)

        layout = QVBoxLayout(self)
        header = QLabel(
            f"Encontramos {len(suggestions)} sugestões na coluna '{column}' (avaliadas {total_values} células)."
            " Deseja aplicar essas substituições?"
        )
        header.setWordWrap(True)
        layout.addWidget(header)

        table = QTableWidget(len(suggestions), 3)
        table.setMinimumHeight(360)
        table.setHorizontalHeaderLabels(["Valor encontrado", "Sugerido", "Ocorrências"])
        for row, (source, target, count) in enumerate(suggestions):
            table.setItem(row, 0, QTableWidgetItem(source))
            table.setItem(row, 1, QTableWidgetItem(target))
            table.setItem(row, 2, QTableWidgetItem(str(count)))
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        layout.addWidget(table)

        buttons = QHBoxLayout()
        apply_btn = QPushButton("Aplicar correções")
        cancel_btn = QPushButton("Cancelar")
        apply_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)
        buttons.addWidget(apply_btn)
        buttons.addStretch()
        buttons.addWidget(cancel_btn)
        layout.addLayout(buttons)


class PreValidationDialog(QDialog):
    def __init__(
        self,
        parent: QWidget,
        *,
        columns: List[str],
        remove_duplicates: bool,
        trim_whitespace: bool,
        selected_column: Optional[str],
        last_result: Optional[Tuple[int, int]],
        run_check: Callable[[str], tuple[int, int]],
        split_enabled: bool,
        split_column: Optional[str],
        split_operator: str,
        split_length: Optional[int],
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
        self.trim_whitespace_checkbox = QCheckBox("Remover espacos em branco no inicio/fim das celulas")
        self.trim_whitespace_checkbox.setChecked(trim_whitespace)
        layout.addWidget(self.trim_whitespace_checkbox)

        column_layout = QHBoxLayout()
        column_layout.addWidget(QLabel("Coluna para checar:"))
        self.duplicate_column_combo = QComboBox()
        for col in columns:
            self.duplicate_column_combo.addItem(col)
        if selected_column and selected_column in columns:
            self.duplicate_column_combo.setCurrentText(selected_column)
        column_layout.addWidget(self.duplicate_column_combo)
        layout.addLayout(column_layout)

        split_group = QGroupBox("Mover valores por tamanho para coluna extra")
        split_layout = QHBoxLayout(split_group)
        self.split_checkbox = QCheckBox("Ativar")
        self.split_checkbox.setChecked(split_enabled)
        split_layout.addWidget(self.split_checkbox)
        self.split_column_combo = QComboBox()
        for col in columns:
            self.split_column_combo.addItem(col)
        if split_column and split_column in columns:
            self.split_column_combo.setCurrentText(split_column)
        split_layout.addWidget(QLabel("Coluna"))
        split_layout.addWidget(self.split_column_combo)
        self.split_operator_combo = QComboBox()
        self.split_operator_combo.addItem(">", "gt")
        self.split_operator_combo.addItem("<", "lt")
        if split_operator in ("gt", "lt"):
            self.split_operator_combo.setCurrentIndex(0 if split_operator == "gt" else 1)
        split_layout.addWidget(self.split_operator_combo)
        self.split_length_spin = QSpinBox()
        self.split_length_spin.setMinimum(1)
        self.split_length_spin.setMaximum(10_000)
        if split_length:
            self.split_length_spin.setValue(split_length)
        split_layout.addWidget(QLabel("Qtde caracteres"))
        split_layout.addWidget(self.split_length_spin)
        layout.addWidget(split_group)

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
    def trim_whitespace(self) -> bool:
        return self.trim_whitespace_checkbox.isChecked()

    @property
    def selected_column(self) -> Optional[str]:
        if self.duplicate_column_combo.count() == 0:
            return None
        return self.duplicate_column_combo.currentText()

    @property
    def last_result(self) -> Optional[Tuple[int, int]]:
        return self._last_result if self.remove_duplicates else None

    @property
    def split_enabled(self) -> bool:
        return self.split_checkbox.isChecked()

    @property
    def split_column(self) -> Optional[str]:
        if not self.split_enabled or self.split_column_combo.count() == 0:
            return None
        return self.split_column_combo.currentText()

    @property
    def split_operator(self) -> str:
        return self.split_operator_combo.currentData() or "gt"

    @property
    def split_length(self) -> Optional[int]:
        return self.split_length_spin.value() if self.split_enabled else None

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
        if self.split_enabled:
            if not self.split_column:
                QMessageBox.warning(self, "Pré-validação", "Selecione a coluna para mover os valores.")
                return
            if not self.split_length or self.split_length <= 0:
                QMessageBox.warning(self, "Pré-validação", "Informe um valor mínimo de caracteres maior que zero.")
                return
        self.accept()


def main() -> int:
    app = QApplication([])
    window = MainWindow()
    window.showMaximized()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
