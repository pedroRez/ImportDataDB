from __future__ import annotations


def build_app_stylesheet() -> str:
    return """
    QWidget {
        background: #f4f7fb;
        color: #1f2937;
        font-size: 13px;
    }

    QMainWindow {
        background: #f4f7fb;
    }

    QLabel[role="eyebrow"] {
        color: #5b6b80;
        font-size: 11px;
        font-weight: 600;
        letter-spacing: 0.08em;
        text-transform: uppercase;
    }

    QLabel[role="title"] {
        color: #10233b;
        font-size: 26px;
        font-weight: 700;
    }

    QLabel[role="section-title"] {
        color: #10233b;
        font-size: 19px;
        font-weight: 700;
    }

    QLabel[role="card-title"] {
        color: #10233b;
        font-size: 14px;
        font-weight: 600;
    }

    QLabel[role="muted"] {
        color: #5b6b80;
    }

    QFrame[card="true"],
    QGroupBox {
        background: #ffffff;
        border: 1px solid #d8e0ea;
        border-radius: 16px;
    }

    QGroupBox {
        margin-top: 12px;
        padding-top: 14px;
        font-weight: 600;
        color: #10233b;
    }

    QGroupBox::title {
        subcontrol-origin: margin;
        left: 14px;
        padding: 0 6px;
    }

    QFrame[modeShell="true"] {
        background: transparent;
        border: none;
    }

    QPushButton,
    QToolButton {
        background: #ffffff;
        border: 1px solid #c8d3df;
        border-radius: 10px;
        padding: 10px 14px;
        min-height: 18px;
    }

    QPushButton:hover,
    QToolButton:hover {
        border-color: #8fa7c2;
        background: #f8fbff;
    }

    QPushButton:disabled,
    QToolButton:disabled {
        background: #eef3f8;
        color: #8ba0b8;
        border-color: #d7e0ea;
    }

    QPushButton[variant="primary"],
    QToolButton[variant="primary"] {
        background: #0f4c81;
        color: #ffffff;
        border-color: #0f4c81;
        font-weight: 600;
        min-height: 22px;
    }

    QPushButton[variant="primary"]:hover,
    QToolButton[variant="primary"]:hover {
        background: #155b99;
        border-color: #155b99;
    }

    QPushButton[variant="ghost"],
    QToolButton[variant="ghost"] {
        background: transparent;
        border-color: transparent;
        color: #0f4c81;
    }

    QPushButton[variant="ghost"]:hover,
    QToolButton[variant="ghost"]:hover {
        background: #edf4fb;
        border-color: #d2e2f0;
    }

    QToolButton[modeButton="true"] {
        min-width: 170px;
        font-weight: 600;
    }

    QToolButton[modeButton="true"]:checked {
        background: #0f4c81;
        color: #ffffff;
        border-color: #0f4c81;
    }

    QLineEdit,
    QComboBox,
    QSpinBox,
    QDateEdit,
    QListWidget,
    QTableView,
    QTableWidget,
    QTextEdit {
        background: #ffffff;
        border: 1px solid #cdd8e4;
        border-radius: 10px;
        padding: 6px 8px;
        selection-background-color: #d9ebfb;
        selection-color: #10233b;
    }

    QLineEdit:focus,
    QComboBox:focus,
    QSpinBox:focus,
    QDateEdit:focus,
    QListWidget:focus,
    QTableView:focus,
    QTableWidget:focus,
    QTextEdit:focus {
        border-color: #0f4c81;
    }

    QListWidget::item {
        padding: 8px 10px;
        border-radius: 8px;
    }

    QListWidget::item:selected {
        background: #e8f2fb;
        color: #10233b;
    }

    QHeaderView::section {
        background: #eef3f8;
        color: #44566c;
        border: none;
        border-bottom: 1px solid #d8e0ea;
        padding: 8px;
        font-weight: 600;
    }

    QTableView,
    QTableWidget {
        gridline-color: #e8edf3;
    }

    QScrollArea {
        border: none;
        background: transparent;
    }

    QScrollBar:vertical {
        background: #edf2f7;
        width: 12px;
        margin: 0;
        border-radius: 6px;
    }

    QScrollBar::handle:vertical {
        background: #bfd0e2;
        min-height: 36px;
        border-radius: 6px;
    }

    QScrollBar:horizontal {
        background: #edf2f7;
        height: 12px;
        margin: 0;
        border-radius: 6px;
    }

    QScrollBar::handle:horizontal {
        background: #bfd0e2;
        min-width: 36px;
        border-radius: 6px;
    }

    QScrollBar::add-line,
    QScrollBar::sub-line,
    QScrollBar::add-page,
    QScrollBar::sub-page {
        background: transparent;
        border: none;
    }

    QLabel[badgeTone="neutral"],
    QLabel[badgeTone="success"],
    QLabel[badgeTone="warning"],
    QLabel[badgeTone="danger"] {
        border-radius: 999px;
        padding: 6px 10px;
        font-weight: 600;
    }

    QLabel[badgeTone="neutral"] {
        background: #eaf0f6;
        color: #40576f;
    }

    QLabel[badgeTone="success"] {
        background: #e4f5ea;
        color: #1d6b3b;
    }

    QLabel[badgeTone="warning"] {
        background: #fff4de;
        color: #8a5a00;
    }

    QLabel[badgeTone="danger"] {
        background: #fde9e7;
        color: #993126;
    }

    QFrame[stepItem="true"] {
        background: #ffffff;
        border: 1px solid #d8e0ea;
        border-radius: 14px;
    }

    QFrame[workspaceShell="true"] {
        background: #ffffff;
    }

    QFrame[dropZone="true"] {
        border-style: dashed;
        border-width: 2px;
        border-color: #b7c7d8;
        background: #fbfdff;
    }

    QFrame[stageSlot="true"] {
        background: transparent;
        border: none;
    }

    QFrame[stepState="active"] {
        border-color: #0f4c81;
        background: #f7fbff;
    }

    QFrame[stepState="ready"] {
        border-color: #8cc4a0;
        background: #f7fcf8;
    }

    QFrame[stepState="warning"] {
        border-color: #e7c576;
        background: #fffdf7;
    }

    QFrame[stepState="error"] {
        border-color: #e7a39b;
        background: #fff8f7;
    }

    QLabel[stepMarker="true"] {
        min-width: 28px;
        max-width: 28px;
        min-height: 28px;
        max-height: 28px;
        border-radius: 14px;
        background: #eaf0f6;
        color: #40576f;
        font-weight: 700;
        qproperty-alignment: AlignCenter;
    }

    QFrame[stepState="active"] QLabel[stepMarker="true"] {
        background: #0f4c81;
        color: #ffffff;
    }

    QFrame[stepState="ready"] QLabel[stepMarker="true"] {
        background: #1d6b3b;
        color: #ffffff;
    }

    QFrame[stepState="warning"] QLabel[stepMarker="true"] {
        background: #d98d00;
        color: #ffffff;
    }

    QFrame[stepState="error"] QLabel[stepMarker="true"] {
        background: #b84235;
        color: #ffffff;
    }
    """
