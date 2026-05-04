from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

import pandas as pd
from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
from PySide6.QtGui import QCursor, QGuiApplication
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableView,
    QVBoxLayout,
)

from src.excel.reader import ExcelReader


@dataclass
class ExcelSelectionResult:
    sheet_name: str
    header_row: int
    data_start_row: int
    data_end_row: Optional[int]
    col_start: int
    col_end: Optional[int]


class RawExcelTableModel(QAbstractTableModel):
    def __init__(self, frame: pd.DataFrame | None = None) -> None:
        super().__init__()
        self._frame = frame if frame is not None else pd.DataFrame()

    def set_frame(self, frame: pd.DataFrame) -> None:
        self.beginResetModel()
        self._frame = frame
        self.endResetModel()

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: N802
        if parent.isValid():
            return 0
        return int(self._frame.shape[0])

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: N802
        if parent.isValid():
            return 0
        return int(self._frame.shape[1])

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) -> object:
        if not index.isValid() or role not in (Qt.DisplayRole, Qt.EditRole):
            return None
        value = self._frame.iat[index.row(), index.column()]
        if pd.isna(value):
            return ""
        return str(value)

    def headerData(  # noqa: N802
        self,
        section: int,
        orientation: Qt.Orientation,
        role: int = Qt.DisplayRole,
    ) -> object:
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Vertical:
            return str(section + 1)
        return self._excel_column_name(section + 1)

    @staticmethod
    def _excel_column_name(number: int) -> str:
        label = ""
        while number:
            number, rem = divmod(number - 1, 26)
            label = chr(65 + rem) + label
        return label


class ExcelSelectionDialog(QDialog):
    """Full-sheet selection dialog used by the legacy import flow."""

    def __init__(
        self,
        parent: object,
        reader: ExcelReader,
        *,
        initial_sheet: str | None = None,
        header_row: int = 1,
        data_end_row: int | None = None,
        col_start: int = 1,
        col_end: int | None = None,
    ) -> None:
        super().__init__(parent)
        self.reader = reader
        self.result: ExcelSelectionResult | None = None
        self._initial_header_row = max(1, header_row)
        self._initial_data_end_row = data_end_row
        self._initial_col_start = max(1, col_start)
        self._initial_col_end = col_end

        self.setWindowTitle("Selecionar dados da planilha")
        self.resize(1280, 820)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        top_layout = QHBoxLayout()
        top_layout.addWidget(QLabel("Aba"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.addItems(reader.sheet_names())
        if initial_sheet:
            index = self.sheet_combo.findText(initial_sheet)
            if index >= 0:
                self.sheet_combo.setCurrentIndex(index)
        self.sheet_combo.currentTextChanged.connect(self._load_current_sheet)
        top_layout.addWidget(self.sheet_combo, 1)

        self.summary_label = QLabel("")
        self.summary_label.setWordWrap(True)
        top_layout.addWidget(self.summary_label, 3)
        layout.addLayout(top_layout)

        self.hint_label = QLabel(
            "Selecione da linha do cabecalho ate a ultima linha desejada. "
            "A primeira linha selecionada sera o cabecalho; as colunas selecionadas serao importadas."
        )
        self.hint_label.setWordWrap(True)
        layout.addWidget(self.hint_label)

        self.table_model = RawExcelTableModel()
        self.table = QTableView()
        self.table.setModel(self.table_model)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(False)
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.verticalHeader().setDefaultSectionSize(24)
        self.table.setStyleSheet("QTableView { gridline-color: #d7dde7; }")
        layout.addWidget(self.table, 1)

        footer = QHBoxLayout()
        self.select_all_btn = QPushButton("Selecionar area inteira")
        self.select_all_btn.clicked.connect(self._select_full_used_range)
        footer.addWidget(self.select_all_btn)
        footer.addStretch()

        buttons = QDialogButtonBox(QDialogButtonBox.Cancel | QDialogButtonBox.Ok)
        buttons.button(QDialogButtonBox.Ok).setText("Confirmar selecao")
        buttons.button(QDialogButtonBox.Cancel).setText("Cancelar")
        buttons.accepted.connect(self._confirm)
        buttons.rejected.connect(self.reject)
        footer.addWidget(buttons)
        layout.addLayout(footer)

        self.table.selectionModel().selectionChanged.connect(lambda *_: self._update_summary())
        self._load_current_sheet()
        self.showMaximized()

    def _load_current_sheet(self) -> None:
        sheet_name = self.sheet_combo.currentText()
        if not sheet_name:
            return
        QGuiApplication.setOverrideCursor(QCursor(Qt.WaitCursor))
        self.summary_label.setText("Carregando planilha...")
        try:
            frame = self.reader.load_sheet_raw(sheet_name)
            self.table_model.set_frame(frame)
            self.table.resizeColumnsToContents()
            self._select_initial_range()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "Excel", f"Nao foi possivel carregar a aba:\n{exc}")
            self.table_model.set_frame(pd.DataFrame())
            self._update_summary()
        finally:
            QGuiApplication.restoreOverrideCursor()

    def _select_initial_range(self) -> None:
        rows = self.table_model.rowCount()
        cols = self.table_model.columnCount()
        if rows <= 0 or cols <= 0:
            self._update_summary()
            return

        top = min(max(self._initial_header_row, 1), rows) - 1
        bottom_excel = self._initial_data_end_row if self._initial_data_end_row else rows
        bottom = min(max(bottom_excel, top + 1), rows) - 1
        left = min(max(self._initial_col_start, 1), cols) - 1
        right_excel = self._initial_col_end if self._initial_col_end else cols
        right = min(max(right_excel, left + 1), cols) - 1
        self._select_range(top, left, bottom, right)

    def _select_full_used_range(self) -> None:
        rows = self.table_model.rowCount()
        cols = self.table_model.columnCount()
        if rows <= 0 or cols <= 0:
            return
        self._select_range(0, 0, rows - 1, cols - 1)

    def _select_range(self, top: int, left: int, bottom: int, right: int) -> None:
        from PySide6.QtCore import QItemSelection, QItemSelectionModel

        top_left = self.table_model.index(top, left)
        bottom_right = self.table_model.index(bottom, right)
        selection = QItemSelection(top_left, bottom_right)
        self.table.selectionModel().select(selection, QItemSelectionModel.ClearAndSelect)
        self.table.scrollTo(top_left)
        self._update_summary()

    def _selection_bounds(self) -> tuple[int, int, int, int] | None:
        selection = self.table.selectionModel().selection()
        if selection.isEmpty():
            return None
        ranges = [selection.at(idx) for idx in range(selection.count())]
        top = min(r.top() for r in ranges)
        bottom = max(r.bottom() for r in ranges)
        left = min(r.left() for r in ranges)
        right = max(r.right() for r in ranges)
        return top, bottom, left, right

    def _current_result(self) -> ExcelSelectionResult | None:
        rows = self.table_model.rowCount()
        cols = self.table_model.columnCount()
        if rows <= 0 or cols <= 0:
            return None

        bounds = self._selection_bounds()
        if bounds is None:
            top, bottom, left, right = 0, rows - 1, 0, cols - 1
        else:
            top, bottom, left, right = bounds

        header_row = top + 1
        data_start_row = header_row + 1
        data_end_row = bottom + 1 if bottom + 1 > header_row else rows
        col_start = left + 1
        col_end = right + 1 if right + 1 < cols else None

        return ExcelSelectionResult(
            sheet_name=self.sheet_combo.currentText(),
            header_row=header_row,
            data_start_row=data_start_row,
            data_end_row=data_end_row,
            col_start=col_start,
            col_end=col_end,
        )

    def _update_summary(self) -> None:
        result = self._current_result()
        if result is None:
            self.summary_label.setText("Nenhum dado encontrado na aba atual.")
            return
        col_end = result.col_end if result.col_end is not None else self.table_model.columnCount()
        self.summary_label.setText(
            "Cabecalho: linha {header} | Dados: linhas {start}-{end} | Colunas: {col_start}-{col_end}".format(
                header=result.header_row,
                start=result.data_start_row,
                end=result.data_end_row or "fim",
                col_start=result.col_start,
                col_end=col_end,
            )
        )

    def _confirm(self) -> None:
        result = self._current_result()
        if result is None:
            QMessageBox.warning(self, "Selecao", "Nenhum dado encontrado para confirmar.")
            return
        if result.data_end_row is not None and result.data_end_row < result.data_start_row:
            QMessageBox.warning(self, "Selecao", "Selecione pelo menos uma linha de dados abaixo do cabecalho.")
            return
        self.result = result
        self.accept()
