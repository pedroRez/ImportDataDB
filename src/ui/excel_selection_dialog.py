from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from src.ui.excel_workspace import ExcelSelectionWorkspace


class ExcelSelectionDialog(QDialog):
    def __init__(self, *, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Selecionar dados do Excel")
        self.setModal(True)
        self._fit_to_available_screen()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(14)

        body_scroll = QScrollArea()
        body_scroll.setWidgetResizable(True)
        body_scroll.setFrameShape(QFrame.NoFrame)
        body_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        body_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(0, 0, 0, 0)
        body_layout.setSpacing(14)

        header = QFrame()
        header.setProperty("card", True)
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(18, 18, 18, 18)
        header_layout.setSpacing(6)

        eyebrow = QLabel("Seletor de planilha")
        eyebrow.setProperty("role", "eyebrow")
        header_layout.addWidget(eyebrow)

        title = QLabel("Confirme a aba, o cabeçalho e o bloco que será importado")
        title.setProperty("role", "section-title")
        header_layout.addWidget(title)

        subtitle = QLabel(
            "A seleção é atualizada ao vivo. Quando terminar, confirme para salvar o recorte na importação."
        )
        subtitle.setWordWrap(True)
        subtitle.setProperty("role", "muted")
        header_layout.addWidget(subtitle)
        body_layout.addWidget(header)

        self.workspace = ExcelSelectionWorkspace(
            subtitle=(
                "Escolha a aba e marque somente as linhas e colunas úteis. "
                "O resumo abaixo acompanha a seleção em tempo real."
            )
        )
        self.workspace.file_drop.hide()
        self.workspace.apply_button.hide()
        self.workspace.use_header_button.hide()
        self.workspace.clear_selection_button.setText("Limpar bloco")
        self.workspace.reload_button.setText("Recarregar aba")
        body_layout.addWidget(self.workspace, 1)
        body_scroll.setWidget(body)
        layout.addWidget(body_scroll, 1)

        footer = QFrame()
        footer.setProperty("card", True)
        footer_layout = QVBoxLayout(footer)
        footer_layout.setContentsMargins(18, 16, 18, 16)
        footer_layout.setSpacing(12)

        self.header_checkbox = QCheckBox("Usar a primeira linha selecionada como cabeçalho ao confirmar")
        footer_layout.addWidget(self.header_checkbox)

        action_row = QHBoxLayout()
        action_row.setSpacing(10)
        self.selection_warning_label = QLabel("Selecione um bloco ou confirme o intervalo atual.")
        self.selection_warning_label.setProperty("role", "muted")
        self.selection_warning_label.setWordWrap(True)
        action_row.addWidget(self.selection_warning_label, 1)

        self.cancel_btn = QPushButton("Cancelar")
        self.cancel_btn.clicked.connect(self.reject)
        action_row.addWidget(self.cancel_btn)

        self.confirm_btn = QPushButton("Confirmar seleção")
        self.confirm_btn.setProperty("variant", "primary")
        self.confirm_btn.clicked.connect(self.accept)
        action_row.addWidget(self.confirm_btn)
        footer_layout.addLayout(action_row)

        layout.addWidget(footer)

    def _fit_to_available_screen(self) -> None:
        screen = self.screen() or QApplication.primaryScreen()
        if not screen:
            self.resize(1200, 760)
            self.setMinimumSize(560, 420)
            return
        available = screen.availableGeometry()
        width = max(560, min(1400, int(available.width() * 0.94)))
        height = max(420, min(900, int(available.height() * 0.90)))
        self.resize(width, height)
        self.setMinimumSize(min(560, width), min(420, height))

    def show_maximized(self) -> None:
        self.showMaximized()

    def set_file_hint(self, file_path: Path | None) -> None:
        self.workspace.set_file_path(file_path)

    def set_footer_hint(self, text: str) -> None:
        self.selection_warning_label.setText(text)
