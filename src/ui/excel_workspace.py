from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
from PySide6.QtCore import QAbstractTableModel, QModelIndex, QObject, Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QStackedWidget,
    QTableView,
    QVBoxLayout,
    QWidget,
)


def _refresh_widget_style(widget: QWidget) -> None:
    widget.style().unpolish(widget)
    widget.style().polish(widget)
    widget.update()


@dataclass
class ExcelSelectionResult:
    header_row: int
    data_start_row: int
    data_end_row: int
    col_start: int
    col_end: int
    selected_columns: list[str]


class FileDropArea(QFrame):
    fileDropped = Signal(str)
    fileChosen = Signal()

    def __init__(self) -> None:
        super().__init__()
        self.setAcceptDrops(True)
        self.setProperty("card", True)
        self.setProperty("dropZone", True)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(10)

        title = QLabel("Importe a planilha")
        title.setProperty("role", "card-title")
        layout.addWidget(title, alignment=Qt.AlignCenter)

        subtitle = QLabel("Arraste o arquivo para esta area ou escolha manualmente.")
        subtitle.setWordWrap(True)
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setProperty("role", "muted")
        layout.addWidget(subtitle)

        self.select_button = QPushButton("Escolher arquivo Excel")
        self.select_button.setProperty("variant", "primary")
        self.select_button.clicked.connect(self.fileChosen.emit)
        layout.addWidget(self.select_button, alignment=Qt.AlignCenter)

    def dragEnterEvent(self, event) -> None:  # type: ignore[override]
        urls = event.mimeData().urls()
        if urls and urls[0].isLocalFile():
            event.acceptProposedAction()
            return
        event.ignore()

    def dropEvent(self, event) -> None:  # type: ignore[override]
        urls = event.mimeData().urls()
        if not urls:
            event.ignore()
            return
        local_file = urls[0].toLocalFile()
        if local_file:
            self.fileDropped.emit(local_file)
            event.acceptProposedAction()
            return
        event.ignore()


class ExcelGridModel(QAbstractTableModel):
    def __init__(self) -> None:
        super().__init__()
        self._headers: list[str] = []
        self._rows: list[list[str]] = []
        self._first_excel_row = 2

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: N802
        if parent.isValid():
            return 0
        return len(self._rows)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: N802
        if parent.isValid():
            return 0
        return len(self._headers)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) -> Any:  # noqa: D401
        if not index.isValid():
            return None
        if role in (Qt.DisplayRole, Qt.EditRole):
            try:
                return self._rows[index.row()][index.column()]
            except IndexError:
                return ""
        if role == Qt.TextAlignmentRole:
            return int(Qt.AlignLeft | Qt.AlignVCenter)
        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole) -> Any:  # noqa: N802
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            if 0 <= section < len(self._headers):
                return self._headers[section]
            return ""
        return str(self._first_excel_row + section)

    def clear(self) -> None:
        self.beginResetModel()
        self._headers = []
        self._rows = []
        self._first_excel_row = 2
        self.endResetModel()

    def reset_content(self, headers: list[str], rows: list[list[str]], first_excel_row: int) -> None:
        self.beginResetModel()
        self._headers = list(headers)
        self._rows = list(rows)
        self._first_excel_row = first_excel_row
        self.endResetModel()

    def append_rows(self, rows: list[list[str]]) -> None:
        if not rows:
            return
        start_row = len(self._rows)
        end_row = start_row + len(rows) - 1
        self.beginInsertRows(QModelIndex(), start_row, end_row)
        self._rows.extend(rows)
        self.endInsertRows()

    @property
    def headers(self) -> list[str]:
        return list(self._headers)

    @property
    def first_excel_row(self) -> int:
        return self._first_excel_row


class ExcelLoadWorker(QObject):
    started = Signal(str)
    first_chunk_ready = Signal(list, list, int, int)
    rows_appended = Signal(list)
    finished = Signal(object, list, int, int, int, int)
    failed = Signal(str)

    def __init__(
        self,
        *,
        reader,
        sheet_name: str,
        header_row: int,
        data_start_row: int,
        data_end_row: int | None,
        col_start: int | None,
        col_end: int | None,
        chunk_size: int = 200,
    ) -> None:
        super().__init__()
        self.reader = reader
        self.sheet_name = sheet_name
        self.header_row = header_row
        self.data_start_row = data_start_row
        self.data_end_row = data_end_row
        self.col_start = col_start
        self.col_end = col_end
        self.chunk_size = chunk_size

    def run(self) -> None:
        try:
            self.started.emit("Carregando planilha...")
            dataframe = self.reader.load_sheet_dataframe(
                self.sheet_name,
                header_row=self.header_row,
                data_start_row=self.data_start_row,
                data_end_row=self.data_end_row,
                col_start=self.col_start,
                col_end=self.col_end,
            )
            headers = list(dataframe.columns)
            all_rows = [self._row_to_display(row) for _, row in dataframe.iterrows()]

            first_chunk = all_rows[: self.chunk_size]
            self.first_chunk_ready.emit(headers, first_chunk, self.data_start_row, len(all_rows))

            for start_index in range(self.chunk_size, len(all_rows), self.chunk_size):
                chunk = all_rows[start_index : start_index + self.chunk_size]
                self.rows_appended.emit(chunk)

            self.finished.emit(
                dataframe,
                headers,
                self.header_row,
                self.data_start_row,
                len(all_rows),
                len(headers),
            )
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(str(exc) or exc.__class__.__name__)

    def _row_to_display(self, row: pd.Series) -> list[str]:
        values: list[str] = []
        for value in row.tolist():
            if value is None:
                values.append("")
                continue
            try:
                if pd.isna(value):
                    values.append("")
                    continue
            except Exception:  # noqa: BLE001
                pass
            values.append(str(value))
        return values


class ExcelSelectionWorkspace(QWidget):
    chooseFileRequested = Signal()
    fileDropped = Signal(str)
    sheetChanged = Signal(str)
    reloadRequested = Signal()
    applySelectionRequested = Signal()
    useSelectionAsHeaderRequested = Signal()
    clearSelectionRequested = Signal()

    def __init__(self, *, subtitle: str) -> None:
        super().__init__()

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(16)

        self.file_drop = FileDropArea()
        self.file_drop.fileChosen.connect(self.chooseFileRequested.emit)
        self.file_drop.fileDropped.connect(self.fileDropped.emit)
        root.addWidget(self.file_drop)

        info_card = QFrame()
        info_card.setProperty("card", True)
        info_layout = QVBoxLayout(info_card)
        info_layout.setContentsMargins(18, 18, 18, 18)
        info_layout.setSpacing(12)

        title = QLabel("Selecione o recorte util da planilha")
        title.setProperty("role", "section-title")
        info_layout.addWidget(title)

        subtitle_label = QLabel(subtitle)
        subtitle_label.setWordWrap(True)
        subtitle_label.setProperty("role", "muted")
        info_layout.addWidget(subtitle_label)

        meta_row = QHBoxLayout()
        meta_row.setSpacing(12)

        file_block = QVBoxLayout()
        file_block.setSpacing(4)
        file_label = QLabel("Arquivo")
        file_label.setProperty("role", "eyebrow")
        file_block.addWidget(file_label)
        self.file_name_label = QLabel("Nenhum arquivo selecionado.")
        self.file_name_label.setWordWrap(True)
        file_block.addWidget(self.file_name_label)
        meta_row.addLayout(file_block, 2)

        sheet_block = QVBoxLayout()
        sheet_block.setSpacing(4)
        sheet_label = QLabel("Aba")
        sheet_label.setProperty("role", "eyebrow")
        sheet_block.addWidget(sheet_label)
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self._emit_sheet_change)
        sheet_block.addWidget(self.sheet_combo)
        meta_row.addLayout(sheet_block, 1)

        status_block = QVBoxLayout()
        status_block.setSpacing(4)
        status_label = QLabel("Status")
        status_label.setProperty("role", "eyebrow")
        status_block.addWidget(status_label)
        self.grid_status_label = QLabel("Aguardando planilha.")
        self.grid_status_label.setProperty("badgeTone", "neutral")
        status_block.addWidget(self.grid_status_label, alignment=Qt.AlignLeft)
        meta_row.addLayout(status_block, 1)

        info_layout.addLayout(meta_row)

        self.selection_summary_label = QLabel(
            "Selecione um bloco com o mouse ou com Shift + setas para definir o recorte."
        )
        self.selection_summary_label.setWordWrap(True)
        info_layout.addWidget(self.selection_summary_label)

        self.applied_summary_label = QLabel("Nenhum recorte aplicado ainda.")
        self.applied_summary_label.setWordWrap(True)
        self.applied_summary_label.setProperty("role", "muted")
        info_layout.addWidget(self.applied_summary_label)

        actions = QHBoxLayout()
        actions.setSpacing(10)
        self.reload_button = QPushButton("Recarregar grade")
        self.reload_button.clicked.connect(self.reloadRequested.emit)
        actions.addWidget(self.reload_button)

        self.apply_button = QPushButton("Aplicar selecao")
        self.apply_button.setProperty("variant", "primary")
        self.apply_button.clicked.connect(self.applySelectionRequested.emit)
        actions.addWidget(self.apply_button)

        self.use_header_button = QPushButton("Usar 1a linha como cabecalho")
        self.use_header_button.clicked.connect(self.useSelectionAsHeaderRequested.emit)
        actions.addWidget(self.use_header_button)

        self.clear_selection_button = QPushButton("Limpar selecao")
        self.clear_selection_button.clicked.connect(self.clearSelectionRequested.emit)
        actions.addWidget(self.clear_selection_button)
        actions.addStretch()
        info_layout.addLayout(actions)

        root.addWidget(info_card)

        table_shell = QFrame()
        table_shell.setProperty("card", True)
        table_shell.setProperty("workspaceShell", True)
        table_layout = QVBoxLayout(table_shell)
        table_layout.setContentsMargins(18, 18, 18, 18)
        table_layout.setSpacing(0)

        self.table_stack = QStackedWidget()
        self.table_stack.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        loading_page = QWidget()
        loading_layout = QVBoxLayout(loading_page)
        loading_layout.setContentsMargins(24, 40, 24, 40)
        loading_layout.setSpacing(8)
        loading_layout.addStretch()
        self.loading_title_label = QLabel("Aguardando planilha")
        self.loading_title_label.setProperty("role", "section-title")
        loading_layout.addWidget(self.loading_title_label, alignment=Qt.AlignCenter)
        self.loading_hint_label = QLabel(
            "O sistema vai abrir a grade assim que o arquivo e a aba estiverem prontos."
        )
        self.loading_hint_label.setProperty("role", "muted")
        self.loading_hint_label.setWordWrap(True)
        self.loading_hint_label.setAlignment(Qt.AlignCenter)
        loading_layout.addWidget(self.loading_hint_label, alignment=Qt.AlignCenter)
        loading_layout.addStretch()

        table_page = QWidget()
        table_page_layout = QVBoxLayout(table_page)
        table_page_layout.setContentsMargins(0, 0, 0, 0)
        table_page_layout.setSpacing(0)

        self.table_view = QTableView()
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table_view.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.table_view.setCornerButtonEnabled(True)
        self.table_view.setWordWrap(False)
        self.table_view.setShowGrid(True)
        self.table_view.setSortingEnabled(False)
        self.table_view.horizontalHeader().setStretchLastSection(False)
        self.table_view.verticalHeader().setDefaultSectionSize(24)
        self.table_view.setMinimumHeight(360)
        table_page_layout.addWidget(self.table_view)

        self.table_stack.addWidget(loading_page)
        self.table_stack.addWidget(table_page)
        self.table_stack.setCurrentWidget(loading_page)
        table_layout.addWidget(self.table_stack, 1)

        root.addWidget(table_shell, 1)

    def _emit_sheet_change(self, sheet_name: str) -> None:
        if sheet_name:
            self.sheetChanged.emit(sheet_name)

    def bind_grid(self, model: ExcelGridModel, selection_model) -> None:
        self.table_view.setModel(model)
        self.table_view.setSelectionModel(selection_model)
        self.table_view.horizontalHeader().setDefaultSectionSize(160)
        self.table_view.verticalHeader().setMinimumWidth(64)

    def set_file_path(self, file_path: Path | None) -> None:
        self.file_name_label.setText(str(file_path) if file_path else "Nenhum arquivo selecionado.")

    def set_sheet_options(self, sheet_names: Iterable[str], current_sheet: str | None = None) -> None:
        current_text = current_sheet or self.sheet_combo.currentText()
        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        for name in sheet_names:
            self.sheet_combo.addItem(name)
        if current_text:
            self.sheet_combo.setCurrentText(current_text)
        self.sheet_combo.blockSignals(False)

    def set_status(self, text: str, *, tone: str = "neutral") -> None:
        self.grid_status_label.setProperty("badgeTone", tone)
        self.grid_status_label.setText(text)
        _refresh_widget_style(self.grid_status_label)

    def set_loading(self, title: str, hint: str) -> None:
        self.loading_title_label.setText(title)
        self.loading_hint_label.setText(hint)
        self.table_stack.setCurrentIndex(0)
        self.set_status(title, tone="warning")

    def show_table(self, *, status_text: str, tone: str = "success") -> None:
        self.table_stack.setCurrentIndex(1)
        self.set_status(status_text, tone=tone)

    def set_selection_summary(self, text: str) -> None:
        self.selection_summary_label.setText(text)

    def set_applied_summary(self, text: str) -> None:
        self.applied_summary_label.setText(text)
