from __future__ import annotations

from typing import Iterable

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QComboBox,
    QSizePolicy,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QToolButton,
    QVBoxLayout,
    QWidget,
)


def _refresh_widget_style(widget: QWidget) -> None:
    widget.style().unpolish(widget)
    widget.style().polish(widget)
    widget.update()


class ConnectionStatusBadge(QLabel):
    def __init__(self) -> None:
        super().__init__("Sem conexão")
        self.setWordWrap(True)
        self.set_status(False, "Sem conexão")

    def set_status(self, connected: bool, text: str, *, tone: str | None = None) -> None:
        badge_tone = tone or ("success" if connected else "warning")
        self.setProperty("badgeTone", badge_tone)
        self.setText(text)
        _refresh_widget_style(self)


class ValidationSummaryCard(QFrame):
    def __init__(self, title: str, value: str = "--", caption: str = "") -> None:
        super().__init__()
        self.setProperty("card", True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(6)

        self.title_label = QLabel(title)
        self.title_label.setProperty("role", "muted")
        layout.addWidget(self.title_label)

        self.value_label = QLabel(value)
        self.value_label.setProperty("role", "section-title")
        layout.addWidget(self.value_label)

        self.caption_label = QLabel(caption)
        self.caption_label.setWordWrap(True)
        self.caption_label.setProperty("role", "muted")
        layout.addWidget(self.caption_label)
        layout.addStretch()

    def set_value(self, value: str, caption: str = "", *, tone: str = "neutral") -> None:
        self.setProperty("tone", tone)
        self.value_label.setText(value)
        self.caption_label.setText(caption)
        _refresh_widget_style(self)


class ImportResultCard(QFrame):
    def __init__(self) -> None:
        super().__init__()
        self.setProperty("card", True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(16)

        title = QLabel("Resultado da importação")
        title.setProperty("role", "card-title")
        layout.addWidget(title)

        stats_layout = QGridLayout()
        stats_layout.setHorizontalSpacing(14)
        stats_layout.setVerticalSpacing(14)

        self.cards = {
            "created": ValidationSummaryCard("Criados", "0", "Novos itens incluídos."),
            "updated": ValidationSummaryCard("Atualizados", "0", "Itens já existentes atualizados."),
            "adjusted": ValidationSummaryCard("Ajustados", "0", "Registros com saldo ajustado."),
            "discarded": ValidationSummaryCard("Descartados", "0", "Linhas ignoradas pelas regras."),
        }
        stats_layout.addWidget(self.cards["created"], 0, 0)
        stats_layout.addWidget(self.cards["updated"], 0, 1)
        stats_layout.addWidget(self.cards["adjusted"], 1, 0)
        stats_layout.addWidget(self.cards["discarded"], 1, 1)
        layout.addLayout(stats_layout)

        self.summary_label = QLabel("Nenhuma importação executada nesta sessão.")
        self.summary_label.setWordWrap(True)
        self.summary_label.setProperty("role", "muted")
        layout.addWidget(self.summary_label)

    def set_result(self, result: dict[str, int] | None, *, skipped_rows: int = 0) -> None:
        if not result:
            for key, card in self.cards.items():
                card.set_value("0", tone="neutral")
            self.summary_label.setText("Nenhuma importação executada nesta sessão.")
            return

        self.cards["created"].set_value(str(result.get("created", 0)), tone="success")
        self.cards["updated"].set_value(str(result.get("updated", 0)), tone="neutral")
        self.cards["adjusted"].set_value(str(result.get("adjusted", 0)), tone="warning")
        self.cards["discarded"].set_value(str(skipped_rows), tone="warning")
        self.summary_label.setText(
            "Importação concluída com lote processado no Xerife. "
            f"Descartes do perfil nesta execução: {skipped_rows}."
        )


class FileDropArea(QFrame):
    fileDropped = Signal(str)

    def __init__(self) -> None:
        super().__init__()
        self.setAcceptDrops(True)
        self.setProperty("card", True)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(10)

        title = QLabel("Arraste a planilha aqui")
        title.setProperty("role", "card-title")
        layout.addWidget(title, alignment=Qt.AlignCenter)

        subtitle = QLabel("Ou escolha o arquivo manualmente. São aceitos .xls, .xlsx e .xlsm.")
        subtitle.setWordWrap(True)
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setProperty("role", "muted")
        layout.addWidget(subtitle)

        self.select_button = QPushButton("Escolher planilha")
        self.select_button.setProperty("variant", "primary")
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


class QuickImportStepper(QWidget):
    stepRequested = Signal(int)

    def __init__(self, steps: Iterable[str]) -> None:
        super().__init__()
        self._frames: list[QFrame] = []
        self._markers: list[QLabel] = []
        self._titles: list[QLabel] = []
        self._details: list[QLabel] = []

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        for index, title in enumerate(steps):
            frame = QFrame()
            frame.setProperty("stepItem", True)
            frame.setProperty("stepState", "pending")

            row = QHBoxLayout(frame)
            row.setContentsMargins(16, 16, 16, 16)
            row.setSpacing(12)

            marker = QLabel(str(index + 1))
            marker.setProperty("stepMarker", True)
            row.addWidget(marker, alignment=Qt.AlignTop)

            text_column = QVBoxLayout()
            text_column.setSpacing(4)

            button = QToolButton()
            button.setText(title)
            button.setToolButtonStyle(Qt.ToolButtonTextOnly)
            button.setAutoRaise(True)
            button.clicked.connect(lambda _checked=False, step=index: self.stepRequested.emit(step))
            button.setStyleSheet("text-align: left; font-weight: 600; padding: 0; border: none; background: transparent;")
            text_column.addWidget(button, alignment=Qt.AlignLeft)

            detail = QLabel("Aguardando.")
            detail.setWordWrap(True)
            detail.setProperty("role", "muted")
            text_column.addWidget(detail)

            row.addLayout(text_column, 1)
            layout.addWidget(frame)

            self._frames.append(frame)
            self._markers.append(marker)
            self._titles.append(button)
            self._details.append(detail)

        layout.addStretch()

    def set_step_state(self, step: int, *, title: str, detail: str, state: str) -> None:
        self._titles[step].setText(title)
        self._details[step].setText(detail)
        self._frames[step].setProperty("stepState", state)
        marker_text = str(step + 1)
        if state == "ready":
            marker_text = "✓"
        elif state == "warning":
            marker_text = "!"
        elif state == "error":
            marker_text = "×"
        self._markers[step].setText(marker_text)
        _refresh_widget_style(self._frames[step])
        _refresh_widget_style(self._markers[step])

    def set_current_step(self, step: int) -> None:
        for index, frame in enumerate(self._frames):
            state = frame.property("stepState") or "pending"
            if index == step and state == "pending":
                frame.setProperty("stepState", "active")
                self._markers[index].setText(str(index + 1))
            elif index == step and state not in {"warning", "error", "ready"}:
                frame.setProperty("stepState", "active")
            elif state == "active":
                frame.setProperty("stepState", "pending")
                self._markers[index].setText(str(index + 1))
            _refresh_widget_style(frame)
            _refresh_widget_style(self._markers[index])


class QuickImportPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(24, 24, 24, 24)
        main_layout.setSpacing(24)

        left_column = QFrame()
        left_column.setProperty("card", True)
        left_layout = QVBoxLayout(left_column)
        left_layout.setContentsMargins(20, 20, 20, 20)
        left_layout.setSpacing(16)

        stepper_title = QLabel("Fluxo rápido")
        stepper_title.setProperty("role", "card-title")
        left_layout.addWidget(stepper_title)

        self.stepper = QuickImportStepper(["Modelo", "Planilha", "Revisão", "Importar"])
        left_layout.addWidget(self.stepper, 1)

        self.connection_hint = QLabel("Conexão com o Xerife pendente.")
        self.connection_hint.setWordWrap(True)
        self.connection_hint.setProperty("role", "muted")
        left_layout.addWidget(self.connection_hint)

        left_column.setMaximumWidth(320)
        main_layout.addWidget(left_column, 0)

        self.cards_stack = QStackedWidget()
        self.cards_stack.addWidget(self._build_model_card())
        self.cards_stack.addWidget(self._build_spreadsheet_card())
        self.cards_stack.addWidget(self._build_validation_card())
        self.cards_stack.addWidget(self._build_import_card())
        main_layout.addWidget(self.cards_stack, 1)

    def _build_model_card(self) -> QWidget:
        card = QFrame()
        card.setProperty("card", True)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(28, 28, 28, 28)
        layout.setSpacing(18)

        eyebrow = QLabel("Etapa 1")
        eyebrow.setProperty("role", "eyebrow")
        layout.addWidget(eyebrow)

        title = QLabel("Escolha um modelo salvo")
        title.setProperty("role", "title")
        layout.addWidget(title)

        subtitle = QLabel(
            "Selecione o perfil que já conhece a estrutura da planilha e o alvo do Xerife. "
            "O modo avançado continua disponível para ajustes técnicos."
        )
        subtitle.setWordWrap(True)
        subtitle.setProperty("role", "muted")
        layout.addWidget(subtitle)

        self.profile_list = QListWidget()
        self.profile_list.setMinimumHeight(240)
        layout.addWidget(self.profile_list, 1)

        self.profile_detail_label = QLabel("Nenhum modelo aplicado.")
        self.profile_detail_label.setWordWrap(True)
        self.profile_detail_label.setProperty("role", "muted")
        layout.addWidget(self.profile_detail_label)

        actions = QHBoxLayout()
        self.use_profile_btn = QPushButton("Usar este modelo")
        self.use_profile_btn.setProperty("variant", "primary")
        actions.addWidget(self.use_profile_btn)

        self.new_profile_btn = QPushButton("Novo modelo")
        self.new_profile_btn.setProperty("variant", "ghost")
        actions.addWidget(self.new_profile_btn)
        actions.addStretch()
        layout.addLayout(actions)

        self.model_next_btn = QPushButton("Continuar para planilha")
        self.model_next_btn.setProperty("variant", "primary")
        layout.addWidget(self.model_next_btn, alignment=Qt.AlignRight)

        return card

    def _build_spreadsheet_card(self) -> QWidget:
        card = QFrame()
        card.setProperty("card", True)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(28, 28, 28, 28)
        layout.setSpacing(18)

        eyebrow = QLabel("Etapa 2")
        eyebrow.setProperty("role", "eyebrow")
        layout.addWidget(eyebrow)

        title = QLabel("Selecione a planilha")
        title.setProperty("role", "title")
        layout.addWidget(title)

        subtitle = QLabel(
            "Use a faixa configurada no perfil e confirme a aba. "
            "Se precisar mexer em cabeçalho ou colunas, faça isso no modo avançado."
        )
        subtitle.setWordWrap(True)
        subtitle.setProperty("role", "muted")
        layout.addWidget(subtitle)

        self.file_drop = FileDropArea()
        layout.addWidget(self.file_drop)

        info_grid = QGridLayout()
        info_grid.setHorizontalSpacing(16)
        info_grid.setVerticalSpacing(12)

        info_grid.addWidget(QLabel("Arquivo"), 0, 0)
        self.file_name_label = QLabel("Nenhum arquivo selecionado.")
        self.file_name_label.setWordWrap(True)
        info_grid.addWidget(self.file_name_label, 0, 1)

        info_grid.addWidget(QLabel("Aba"), 1, 0)
        self.quick_sheet_combo = QComboBox()
        info_grid.addWidget(self.quick_sheet_combo, 1, 1)

        info_grid.addWidget(QLabel("Faixa aplicada"), 2, 0)
        self.range_summary_label = QLabel("Aguardando modelo e planilha.")
        self.range_summary_label.setWordWrap(True)
        info_grid.addWidget(self.range_summary_label, 2, 1)

        layout.addLayout(info_grid)

        self.quick_preview_table = QTableWidget()
        self.quick_preview_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.quick_preview_table.setAlternatingRowColors(True)
        self.quick_preview_table.setMinimumHeight(220)
        layout.addWidget(self.quick_preview_table, 1)

        actions = QHBoxLayout()
        self.spreadsheet_back_btn = QPushButton("Voltar")
        actions.addWidget(self.spreadsheet_back_btn)
        self.refresh_preview_btn = QPushButton("Atualizar prévia")
        actions.addWidget(self.refresh_preview_btn)
        actions.addStretch()
        self.spreadsheet_next_btn = QPushButton("Continuar para revisão")
        self.spreadsheet_next_btn.setProperty("variant", "primary")
        actions.addWidget(self.spreadsheet_next_btn)
        layout.addLayout(actions)

        return card

    def _build_validation_card(self) -> QWidget:
        card = QFrame()
        card.setProperty("card", True)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(28, 28, 28, 28)
        layout.setSpacing(18)

        eyebrow = QLabel("Etapa 3")
        eyebrow.setProperty("role", "eyebrow")
        layout.addWidget(eyebrow)

        title = QLabel("Revise e valide")
        title.setProperty("role", "title")
        layout.addWidget(title)

        subtitle = QLabel(
            "Confira o resumo antes do envio. Linhas descartadas pelo perfil aparecem aqui sem bloquear o lote."
        )
        subtitle.setWordWrap(True)
        subtitle.setProperty("role", "muted")
        layout.addWidget(subtitle)

        cards_grid = QGridLayout()
        cards_grid.setHorizontalSpacing(14)
        cards_grid.setVerticalSpacing(14)
        self.validation_cards = {
            "total": ValidationSummaryCard("Linhas lidas", "--", "Total bruto da faixa selecionada."),
            "importable": ValidationSummaryCard("Importáveis", "--", "Itens prontos para envio."),
            "skipped": ValidationSummaryCard("Descartadas", "--", "Ignoradas pelas regras do perfil."),
            "blocking": ValidationSummaryCard("Erros bloqueantes", "--", "Impedem o envio do lote."),
        }
        cards_grid.addWidget(self.validation_cards["total"], 0, 0)
        cards_grid.addWidget(self.validation_cards["importable"], 0, 1)
        cards_grid.addWidget(self.validation_cards["skipped"], 1, 0)
        cards_grid.addWidget(self.validation_cards["blocking"], 1, 1)
        layout.addLayout(cards_grid)

        self.validation_status_label = QLabel("Validação ainda não executada.")
        self.validation_status_label.setWordWrap(True)
        self.validation_status_label.setProperty("badgeTone", "neutral")
        layout.addWidget(self.validation_status_label, alignment=Qt.AlignLeft)

        self.validation_table = QTableWidget(0, 4)
        self.validation_table.setHorizontalHeaderLabels(["Linha", "Status", "Código", "Mensagem"])
        self.validation_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.validation_table.setAlternatingRowColors(True)
        self.validation_table.verticalHeader().setVisible(False)
        self.validation_table.setMinimumHeight(220)
        layout.addWidget(self.validation_table, 1)

        detail_row = QHBoxLayout()
        self.validation_details_btn = QPushButton("Ver detalhes")
        detail_row.addWidget(self.validation_details_btn)
        self.validate_btn = QPushButton("Validar modelo")
        self.validate_btn.setProperty("variant", "primary")
        detail_row.addWidget(self.validate_btn)
        detail_row.addStretch()
        layout.addLayout(detail_row)

        self.validation_details = QTextEdit()
        self.validation_details.setReadOnly(True)
        self.validation_details.setVisible(False)
        self.validation_details.setMinimumHeight(160)
        layout.addWidget(self.validation_details)

        actions = QHBoxLayout()
        self.validation_back_btn = QPushButton("Voltar")
        actions.addWidget(self.validation_back_btn)
        actions.addStretch()
        self.validation_next_btn = QPushButton("Ir para importar")
        self.validation_next_btn.setProperty("variant", "primary")
        actions.addWidget(self.validation_next_btn)
        layout.addLayout(actions)

        return card

    def _build_import_card(self) -> QWidget:
        card = QFrame()
        card.setProperty("card", True)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(28, 28, 28, 28)
        layout.setSpacing(18)

        eyebrow = QLabel("Etapa 4")
        eyebrow.setProperty("role", "eyebrow")
        layout.addWidget(eyebrow)

        title = QLabel("Importe para o Xerife")
        title.setProperty("role", "title")
        layout.addWidget(title)

        subtitle = QLabel(
            "O lote só é enviado quando houver conexão ativa e uma validação sem erros bloqueantes."
        )
        subtitle.setWordWrap(True)
        subtitle.setProperty("role", "muted")
        layout.addWidget(subtitle)

        self.import_checklist_label = QLabel("Conecte-se ao Xerife e valide a planilha antes de importar.")
        self.import_checklist_label.setWordWrap(True)
        layout.addWidget(self.import_checklist_label)

        self.import_status_label = QLabel("Aguardando validação.")
        self.import_status_label.setProperty("badgeTone", "neutral")
        layout.addWidget(self.import_status_label, alignment=Qt.AlignLeft)

        self.import_btn = QPushButton("Importar para o Xerife")
        self.import_btn.setProperty("variant", "primary")
        layout.addWidget(self.import_btn, alignment=Qt.AlignLeft)

        self.result_card = ImportResultCard()
        layout.addWidget(self.result_card, 1)

        actions = QHBoxLayout()
        self.import_back_btn = QPushButton("Voltar")
        actions.addWidget(self.import_back_btn)
        actions.addStretch()
        layout.addLayout(actions)

        return card

    def set_step(self, step: int) -> None:
        self.cards_stack.setCurrentIndex(step)
        self.stepper.set_current_step(step)

    def set_profile_items(self, items: list[tuple[str, str, str]]) -> None:
        current_id = self.selected_profile_id()
        self.profile_list.blockSignals(True)
        self.profile_list.clear()
        for profile_id, name, summary in items:
            item = QListWidgetItem(name)
            item.setData(Qt.UserRole, profile_id)
            item.setToolTip(summary)
            self.profile_list.addItem(item)
        if current_id:
            self.select_profile(current_id)
        elif self.profile_list.count() > 0:
            self.profile_list.setCurrentRow(0)
        self.profile_list.blockSignals(False)

    def selected_profile_id(self) -> str | None:
        item = self.profile_list.currentItem()
        if not item:
            return None
        return str(item.data(Qt.UserRole))

    def select_profile(self, profile_id: str | None) -> None:
        if not profile_id:
            return
        for index in range(self.profile_list.count()):
            item = self.profile_list.item(index)
            if str(item.data(Qt.UserRole)) == profile_id:
                self.profile_list.setCurrentRow(index)
                return

    def populate_validation_rows(self, rows: list[tuple[str, str, str, str]]) -> None:
        self.validation_table.setRowCount(len(rows))
        for row_index, values in enumerate(rows):
            for column_index, value in enumerate(values):
                self.validation_table.setItem(row_index, column_index, QTableWidgetItem(value))
        self.validation_table.resizeColumnsToContents()
        self.validation_table.horizontalHeader().setStretchLastSection(True)
