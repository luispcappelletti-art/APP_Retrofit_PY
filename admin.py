import sys
import json
import os
from datetime import datetime
import re
from PyQt6 import QtWidgets, QtGui, QtCore
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QColor
from PyQt6.QtWidgets import QGraphicsDropShadowEffect, QStyle
import firebase_admin
from firebase_admin import credentials, firestore, auth

# --- CONFIGURAÇÃO ---
try:
    cred = credentials.Certificate("serviceAccountKey.json")
    if not firebase_admin._apps:
        firebase_admin.initialize_app(cred)
except FileNotFoundError:
    print("ERRO: O arquivo 'serviceAccountKey.json' não foi encontrado.")
    sys.exit(1)
except Exception as e:
    print(f"Ocorreu um erro ao inicializar o Firebase: {e}")
    sys.exit(1)

db = firestore.client()

# --- ARQUIVOS LOCAIS ---
LIMITE_FILE = "limites.json"
RELATORIOS_FILE = "relatorios.json"
PRECOS_FILE = "precos.json"
USUARIOS_FILE = "usuarios.json"
PERGUNTAS_INICIAIS_FILE = "perguntas_iniciais.json"


# --- CORES E ESTILOS ---
class Colors:
    PRIMARY = "#4a90e2"
    PRIMARY_DARK = "#357abd"
    BACKGROUND = "#1e1e1e"
    CARD = "#252526"
    TEXT = "#ffffff"
    TEXT_SECONDARY = "#b7c0cd"
    BORDER = "#4a90e2"
    HOVER = "#2f3338"
    SUCCESS = "#28a745"
    DANGER = "#dc3545"
    WARNING = "#ff9800"
    INFO = "#17a2b8"


def format_currency(value):
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def parse_decimal(value):
    text = str(value or "").strip()
    if not text:
        return None
    text = text.replace(" ", "").replace(".", "").replace(",", ".")
    return float(text)


def safe_date_from_iso(iso_text):
    try:
        return datetime.fromisoformat(iso_text)
    except (TypeError, ValueError):
        return None

def converter_firestore_para_json(data):
    from datetime import datetime

    if isinstance(data, dict):
        return {k: converter_firestore_para_json(v) for k, v in data.items()}

    elif isinstance(data, list):
        return [converter_firestore_para_json(v) for v in data]

    elif isinstance(data, datetime):
        return data.isoformat()

    else:
        return data

# ================== COMPONENTES CUSTOMIZADOS ==================
class ModernButton(QtWidgets.QPushButton):
    """Botão com estilo moderno."""

    def __init__(self, text="", icon_char="", variant="primary", parent=None):
        full_text = f"{icon_char} {text}" if icon_char else text
        super().__init__(full_text, parent)
        self.variant = variant
        self.setCursor(QtGui.QCursor(Qt.CursorShape.PointingHandCursor))
        self.setup_style()

    def setup_style(self):
        styles = {
            "primary": f"background-color: {Colors.PRIMARY}; color: white;",
            "success": f"background-color: {Colors.SUCCESS}; color: white;",
            "danger": f"background-color: {Colors.DANGER}; color: white;",
            "secondary": f"background-color: {Colors.CARD}; color: {Colors.TEXT}; border: 1px solid {Colors.BORDER};"
        }
        hover_styles = {
            "primary": f"background-color: {Colors.PRIMARY_DARK};",
            "success": f"background-color: #218838;",
            "danger": f"background-color: #c82333;",
            "secondary": f"background-color: #f1f5f9;"
        }
        base_style = """
            QPushButton {{
                padding: 8px 16px;
                border-radius: 8px;
                font-weight: 600;
                font-size: 14px;
                border: none;
                {variant_style}
            }}
            QPushButton:hover {{
                {hover_style}
            }}
            QPushButton:disabled {{
                background-color: #e0e0e0;
                color: #aaaaaa;
            }}
        """
        self.setStyleSheet(base_style.format(
            variant_style=styles.get(self.variant, styles["primary"]),
            hover_style=hover_styles.get(self.variant, styles["primary"])
        ))


class ModernCard(QtWidgets.QFrame):
    """Cartão base para agrupar conteúdo, com sombra e bordas arredondadas."""

    def __init__(self, title="", parent=None):
        super().__init__(parent)
        self.setStyleSheet(f"""
            QFrame {{
                background-color: {Colors.CARD};
                border-radius: 12px;
                border: 1px solid {Colors.BORDER};
            }}
        """)
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(20)
        shadow.setXOffset(0)
        shadow.setYOffset(4)
        shadow.setColor(QColor(0, 0, 0, 25))
        self.setGraphicsEffect(shadow)

        self.main_layout = QtWidgets.QVBoxLayout(self)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.setSpacing(15)

        if title:
            title_label = QtWidgets.QLabel(title)
            title_label.setStyleSheet(f"""
                font-size: 18px;
                font-weight: 700;
                color: {Colors.TEXT};
                margin-bottom: 5px;
                border: none;
            """)
            self.main_layout.addWidget(title_label)


class StatKpiCard(QtWidgets.QFrame):
    """Card enxuto para destacar métricas-chave da aba de estatísticas."""

    def __init__(self, title, accent=Colors.PRIMARY, parent=None):
        super().__init__(parent)
        self.setStyleSheet(f"""
            QFrame {{
                background-color: {Colors.CARD};
                border-radius: 12px;
                border: 1px solid {Colors.BORDER};
            }}
        """)
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(6)

        title_label = QtWidgets.QLabel(title)
        title_label.setStyleSheet(f"""
            color: {Colors.TEXT_SECONDARY};
            font-size: 12px;
            font-weight: 600;
            border: none;
        """)

        self.value_label = QtWidgets.QLabel("-")
        self.value_label.setStyleSheet(f"""
            color: {accent};
            font-size: 22px;
            font-weight: 800;
            border: none;
        """)

        self.subtitle_label = QtWidgets.QLabel("")
        self.subtitle_label.setStyleSheet(f"""
            color: {Colors.TEXT_SECONDARY};
            font-size: 11px;
            border: none;
        """)
        self.subtitle_label.setWordWrap(True)

        layout.addWidget(title_label)
        layout.addWidget(self.value_label)
        layout.addWidget(self.subtitle_label)
        layout.addStretch()

    def set_data(self, value, subtitle=""):
        self.value_label.setText(value)
        self.subtitle_label.setText(subtitle)


# --- DIÁLOGOS ---
class ModernDialog(QtWidgets.QDialog):
    """Diálogo base com estilo moderno."""

    def __init__(self, title, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setMinimumWidth(450)
        self.setStyleSheet(f"""
            QDialog {{ background-color: {Colors.BACKGROUND}; }}
            QGroupBox {{ font-weight: bold; border: 1px solid {Colors.BORDER}; border-radius: 8px; margin-top: 10px; }}
            QGroupBox::title {{ subcontrol-origin: margin; left: 10px; padding: 0 5px; }}
            QLineEdit, QComboBox, QListWidget {{
                background-color: {Colors.CARD};
                border: 1px solid {Colors.BORDER};
                border-radius: 6px;
                padding: 8px;
                font-size: 14px;
                color: {Colors.TEXT};
            }}
            QLineEdit:focus, QComboBox:focus, QListWidget:focus {{
                border-color: {Colors.PRIMARY};
            }}
        """)


class PerguntaDialog(ModernDialog):
    def __init__(self, pergunta_data=None, parent=None):
        title = "Editar Pergunta" if pergunta_data else "Adicionar Pergunta"
        super().__init__(title, parent)

        self.layout = QtWidgets.QVBoxLayout(self)
        form_layout = QtWidgets.QFormLayout()
        self.titulo_edit = QtWidgets.QLineEdit()
        self.tipo_combo = QtWidgets.QComboBox()
        self.tipo_combo.addItems(["texto_livre", "opcoes"])
        form_layout.addRow("Pergunta:", self.titulo_edit)
        form_layout.addRow("Tipo:", self.tipo_combo)
        self.layout.addLayout(form_layout)

        self.opcoes_widget = QtWidgets.QGroupBox("Opções de Resposta")
        opcoes_layout = QtWidgets.QVBoxLayout()
        self.opcoes_list = QtWidgets.QListWidget()
        opcoes_btn_layout = QtWidgets.QHBoxLayout()
        self.add_opcao_btn = ModernButton("Adicionar", icon_char="➕", variant="success")
        self.remove_opcao_btn = ModernButton("Remover", icon_char="🗑️", variant="danger")
        opcoes_btn_layout.addWidget(self.add_opcao_btn)
        opcoes_btn_layout.addWidget(self.remove_opcao_btn)
        opcoes_layout.addWidget(self.opcoes_list)
        opcoes_layout.addLayout(opcoes_btn_layout)
        self.opcoes_widget.setLayout(opcoes_layout)
        self.layout.addWidget(self.opcoes_widget)

        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addStretch()
        cancel_btn = ModernButton("Cancelar", variant="secondary")
        ok_btn = ModernButton("OK", variant="primary")
        button_layout.addWidget(cancel_btn)
        button_layout.addWidget(ok_btn)
        self.layout.addLayout(button_layout)

        self.tipo_combo.currentTextChanged.connect(self.toggle_opcoes_widget)
        self.add_opcao_btn.clicked.connect(self.add_opcao)
        self.remove_opcao_btn.clicked.connect(self.remove_opcao)
        ok_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)

        if pergunta_data:
            self.titulo_edit.setText(pergunta_data.get("titulo", ""))
            self.tipo_combo.setCurrentText(pergunta_data.get("tipo", "texto_livre"))
            if pergunta_data.get("tipo") == "opcoes":
                self.opcoes_list.addItems(pergunta_data.get("opcoes", []))
        self.toggle_opcoes_widget(self.tipo_combo.currentText())

    def toggle_opcoes_widget(self, tipo):
        self.opcoes_widget.setVisible(tipo == "opcoes")

    def add_opcao(self):
        dialog = GetTextDialog("Nova Opção", "Texto da opção:", self)
        if dialog.exec():
            text = dialog.get_text()
            if text:
                self.opcoes_list.addItem(text)

    def remove_opcao(self):
        selected_items = self.opcoes_list.selectedItems()
        if not selected_items: return
        for item in selected_items:
            self.opcoes_list.takeItem(self.opcoes_list.row(item))

    def get_data(self):
        data = {
            "titulo": self.titulo_edit.text().strip(),
            "tipo": self.tipo_combo.currentText()
        }
        if data["tipo"] == "opcoes":
            data["opcoes"] = [self.opcoes_list.item(i).text() for i in range(self.opcoes_list.count())]
        if not data["titulo"]:
            ModernMessageBox.warning(self, "Erro", "O título da pergunta não pode estar vazio.")
            return None
        return data


class ReportDetailsDialog(ModernDialog):

    def __init__(self, report_data, parent=None):

        super().__init__("Relatório de Orçamento", parent)

        self.setMinimumSize(980, 680)
        layout = QtWidgets.QVBoxLayout(self)
        scroll_area = QtWidgets.QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll_area.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        scroll_content = QtWidgets.QWidget()
        main_layout = QtWidgets.QVBoxLayout(scroll_content)
        main_layout.setSpacing(14)

        # =====================
        # TÍTULO
        # =====================

        titulo = QtWidgets.QLabel("RELATÓRIO DE ORÇAMENTO")
        titulo.setStyleSheet("""
        font-size:26px;
        font-weight:700;
        color:white;
        """)
        main_layout.addWidget(titulo)

        # =====================
        # INFORMAÇÕES PRINCIPAIS
        # =====================

        data_str = report_data.get("criadoEm", "")

        try:
            data_formatada = datetime.fromisoformat(data_str).strftime("%d/%m/%Y %H:%M")
        except:
            data_formatada = "N/A"

        info = QtWidgets.QLabel(
            f"<b>Data:</b> {data_formatada}<br>"
            f"<b>Orçamentista:</b> {report_data.get('orcamentistaEmail','N/A')}<br>"
            f"<b>Estimativa:</b> {report_data.get('estimativaFormatada','N/A')}"
        )

        info.setTextFormat(Qt.TextFormat.RichText)

        resumo_box = QtWidgets.QGroupBox("Resumo")
        resumo_layout = QtWidgets.QVBoxLayout()
        resumo_layout.addWidget(info)
        resumo_box.setLayout(resumo_layout)
        main_layout.addWidget(resumo_box)

        # =====================
        # PERGUNTAS INICIAIS
        # =====================

        respostas_iniciais = report_data.get("respostasIniciais", {})

        if respostas_iniciais:

            box = QtWidgets.QGroupBox("Perguntas iniciais")
            box_layout = QtWidgets.QVBoxLayout()

            for pergunta, resposta in respostas_iniciais.items():

                label = QtWidgets.QLabel(f"<b>{pergunta}</b>: {resposta}")
                label.setWordWrap(True)

                box_layout.addWidget(label)

            box.setLayout(box_layout)
            main_layout.addWidget(box)

        # =====================
        # PERGUNTAS TÉCNICAS
        # =====================

        respostas_tecnicas = report_data.get("respostasQuestionario", {})

        if respostas_tecnicas:

            box = QtWidgets.QGroupBox("Perguntas técnicas")
            box_layout = QtWidgets.QVBoxLayout()

            for pergunta, resposta in respostas_tecnicas.items():

                label = QtWidgets.QLabel(f"<b>{pergunta}</b>: {resposta}")
                label.setWordWrap(True)

                box_layout.addWidget(label)

            box.setLayout(box_layout)
            main_layout.addWidget(box)

        # =====================
        # ITENS DO ORÇAMENTO
        # =====================

        itens = report_data.get("itensOrcamento", [])

        # remover itens com valor 0
        itens_filtrados = [i for i in itens if i.get("valor", 0) > 0]

        if itens_filtrados:

            box = QtWidgets.QGroupBox("Itens do orçamento")
            box_layout = QtWidgets.QVBoxLayout()

            total = 0

            for item in itens_filtrados:

                descricao = item.get("descricao", "")
                valor = item.get("valor", 0)

                total += valor

                valor_formatado = format_currency(valor)

                linha = QtWidgets.QLabel(
                    f"• <b>{descricao}</b> — {valor_formatado}"
                )

                linha.setTextFormat(Qt.TextFormat.RichText)
                linha.setWordWrap(True)

                box_layout.addWidget(linha)

            # =====================
            # TOTAL
            # =====================

            box.setLayout(box_layout)
            main_layout.addWidget(box)

            total_box = QtWidgets.QGroupBox("Total")
            total_layout = QtWidgets.QVBoxLayout()
            total_label = QtWidgets.QLabel(f"<b>{format_currency(total)}</b>")
            total_label.setStyleSheet("font-size: 20px; color: #00ff9c;")
            total_layout.addWidget(total_label)
            total_box.setLayout(total_layout)
            main_layout.addWidget(total_box)

        # =====================
        # BOTÃO FECHAR
        # =====================

        copiar = ModernButton("Copiar Relatório", icon_char="📋", variant="primary")
        copiar.clicked.connect(lambda: self._copiar_relatorio(report_data, itens_filtrados))
        fechar = ModernButton("Fechar", variant="secondary")
        fechar.clicked.connect(self.accept)
        btns = QtWidgets.QHBoxLayout()
        btns.addWidget(copiar)
        btns.addStretch()
        btns.addWidget(fechar)
        main_layout.addLayout(btns)
        scroll_area.setWidget(scroll_content)
        layout.addWidget(scroll_area)

    def _copiar_relatorio(self, report_data, itens_filtrados):
        linhas = [
            "RELATÓRIO DE ORÇAMENTO",
            f"Data: {report_data.get('criadoEm', 'N/A')}",
            f"Orçamentista: {report_data.get('orcamentistaEmail', 'N/A')}",
            "",
            "Itens do orçamento:",
        ]
        total = 0
        for item in itens_filtrados:
            valor = item.get("valor", 0) or 0
            total += valor
            linhas.append(f"- {item.get('descricao', 'Sem descrição')}: {format_currency(valor)}")
        linhas.append(f"\nTotal: {format_currency(total)}")
        QtWidgets.QApplication.clipboard().setText("\n".join(linhas))
        ModernMessageBox.information(self, "Copiado", "Relatório copiado para a área de transferência.")


class AddUserDialog(ModernDialog):
    def __init__(self, parent=None):
        super().__init__("Adicionar Novo Usuário", parent)

        layout = QtWidgets.QVBoxLayout(self)
        form_layout = QtWidgets.QFormLayout()

        self.email_edit = QtWidgets.QLineEdit()
        self.email_edit.setPlaceholderText("exemplo@email.com")
        self.password_edit = QtWidgets.QLineEdit()
        self.password_edit.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)

        form_layout.addRow("Email:", self.email_edit)
        form_layout.addRow("Senha:", self.password_edit)
        layout.addLayout(form_layout)

        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addStretch()
        cancel_btn = ModernButton("Cancelar", variant="secondary")
        ok_btn = ModernButton("Salvar", variant="primary")
        button_layout.addWidget(cancel_btn)
        button_layout.addWidget(ok_btn)
        layout.addLayout(button_layout)

        ok_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)

    def get_data(self):
        email = self.email_edit.text().strip()
        password = self.password_edit.text()
        if not email or not password:
            ModernMessageBox.warning(self, "Campos Vazios", "Email e senha são obrigatórios.")
            return None, None
        return email, password


class EditUserStatusDialog(ModernDialog):
    def __init__(self, email, current_status, parent=None):
        super().__init__(f"Editar Status de {email}", parent)
        self.setMinimumWidth(400)

        layout = QtWidgets.QVBoxLayout(self)
        form_layout = QtWidgets.QFormLayout()

        self.status_combo = QtWidgets.QComboBox()
        self.status_combo.addItems(["Ativo", "Desativado"])
        current_index = 0 if current_status == "Ativo" else 1
        self.status_combo.setCurrentIndex(current_index)

        form_layout.addRow("Status:", self.status_combo)
        layout.addLayout(form_layout)

        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addStretch()
        cancel_btn = ModernButton("Cancelar", variant="secondary")
        ok_btn = ModernButton("Salvar", variant="primary")
        button_layout.addWidget(cancel_btn)
        button_layout.addWidget(ok_btn)
        layout.addLayout(button_layout)

        ok_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)

    def get_status(self):
        return self.status_combo.currentText()


class GetTextDialog(ModernDialog):
    def __init__(self, title, label, parent=None):
        super().__init__(title, parent)
        self.setMinimumWidth(400)

        layout = QtWidgets.QVBoxLayout(self)
        form_layout = QtWidgets.QFormLayout()

        self.text_edit = QtWidgets.QLineEdit()
        form_layout.addRow(label, self.text_edit)
        layout.addLayout(form_layout)

        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addStretch()
        cancel_btn = ModernButton("Cancelar", variant="secondary")
        ok_btn = ModernButton("OK", variant="primary")
        button_layout.addWidget(cancel_btn)
        button_layout.addWidget(ok_btn)
        layout.addLayout(button_layout)

        ok_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)

    def get_text(self):
        return self.text_edit.text().strip()


class ModernMessageBox(ModernDialog):
    def __init__(self, parent, icon, title, text, buttons):
        super().__init__(title, parent)
        self.setMinimumWidth(400)
        self.result = None

        layout = QtWidgets.QVBoxLayout(self)
        layout.setSpacing(20)

        content_layout = QtWidgets.QHBoxLayout()
        content_layout.setSpacing(15)

        if icon:
            icon_label = QtWidgets.QLabel()
            pixmap = self.style().standardPixmap(icon)
            icon_label.setPixmap(
                pixmap.scaled(48, 48, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
            content_layout.addWidget(icon_label)

        text_label = QtWidgets.QLabel(text)
        text_label.setWordWrap(True)
        content_layout.addWidget(text_label, 1)
        layout.addLayout(content_layout)

        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addStretch()

        if buttons & QtWidgets.QDialogButtonBox.StandardButton.Ok:
            ok_button = ModernButton("OK", variant="primary")
            ok_button.clicked.connect(self.accept)
            button_layout.addWidget(ok_button)
            self.result = QtWidgets.QMessageBox.StandardButton.Ok

        if buttons & QtWidgets.QDialogButtonBox.StandardButton.Yes:
            yes_button = ModernButton("Sim", variant="success")
            yes_button.clicked.connect(self.on_yes)
            button_layout.addWidget(yes_button)

        if buttons & QtWidgets.QDialogButtonBox.StandardButton.No:
            no_button = ModernButton("Não", variant="danger")
            no_button.clicked.connect(self.on_no)
            button_layout.addWidget(no_button)

        if buttons & QtWidgets.QDialogButtonBox.StandardButton.Cancel:
            cancel_button = ModernButton("Cancelar", variant="secondary")
            cancel_button.clicked.connect(self.reject)
            button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)

    def on_yes(self):
        self.result = QtWidgets.QMessageBox.StandardButton.Yes
        self.accept()

    def on_no(self):
        self.result = QtWidgets.QMessageBox.StandardButton.No
        self.accept()

    @staticmethod
    def show_message(parent, icon_enum, title, text, buttons_enum):
        icon_map = {
            QtWidgets.QMessageBox.Icon.Information: QStyle.StandardPixmap.SP_MessageBoxInformation,
            QtWidgets.QMessageBox.Icon.Warning: QStyle.StandardPixmap.SP_MessageBoxWarning,
            QtWidgets.QMessageBox.Icon.Critical: QStyle.StandardPixmap.SP_MessageBoxCritical,
            QtWidgets.QMessageBox.Icon.Question: QStyle.StandardPixmap.SP_MessageBoxQuestion
        }
        dialog = ModernMessageBox(parent, icon_map.get(icon_enum), title, text, buttons_enum)
        dialog.exec()
        return dialog.result

    @staticmethod
    def information(parent, title, text):
        return ModernMessageBox.show_message(parent, QtWidgets.QMessageBox.Icon.Information, title, text,
                                             QtWidgets.QDialogButtonBox.StandardButton.Ok)

    @staticmethod
    def warning(parent, title, text):
        return ModernMessageBox.show_message(parent, QtWidgets.QMessageBox.Icon.Warning, title, text,
                                             QtWidgets.QDialogButtonBox.StandardButton.Ok)

    @staticmethod
    def critical(parent, title, text):
        return ModernMessageBox.show_message(parent, QtWidgets.QMessageBox.Icon.Critical, title, text,
                                             QtWidgets.QDialogButtonBox.StandardButton.Ok)

    @staticmethod
    def question(parent, title, text):
        buttons = QtWidgets.QDialogButtonBox.StandardButton.Yes | QtWidgets.QDialogButtonBox.StandardButton.No
        return ModernMessageBox.show_message(parent, QtWidgets.QMessageBox.Icon.Question, title, text, buttons)


# ================== JANELA PRINCIPAL ==================
class FirebaseManager(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.max_acessos_pecas, self.max_acessos_usuarios = 5, 5
        self.max_acessos_perg_sync, self.max_acessos_perg_save, self.max_acessos_rel_sync = 1, 2, 1
        self.limites = self.carregar_limites()
        self.novos_usuarios_cache = []  # NOVO: Cache para novos usuários

        self.reset_timer = QtCore.QTimer(self)
        self.reset_timer.timeout.connect(self.verificar_reset_diario)
        self.reset_timer.start(60000)

        self.setWindowTitle("Gerenciador Firebase")
        self.setMinimumSize(1024, 768)
        self.doc_ids, self.perguntas_doc_id = {}, None
        self.perguntas_data = {"ordem": [], "perguntas": {}}
        self.relatorios_filtrados_cache, self.local_reports_data = [], []

        self.setup_ui()
        self.load_all_data()

    def setup_ui(self):
        self.setStyleSheet(f"""
            QMainWindow {{ background-color: {Colors.BACKGROUND}; }}
            QWidget {{ font-size: 14px; color: {Colors.TEXT}; }}
            QTabWidget::pane {{ border-top: 1px solid {Colors.BORDER}; }}
            QTabBar::tab {{
                background-color: transparent; color: {Colors.TEXT_SECONDARY};
                padding: 10px 20px; border: 1px solid transparent; 
                border-bottom: none;
                border-top-left-radius: 8px; border-top-right-radius: 8px;
                font-weight: 600;
            }}
            QTabBar::tab:selected {{ background-color: {Colors.CARD}; color: {Colors.TEXT}; border-color: {Colors.BORDER}; }}
            QTabBar::tab:hover {{ color: {Colors.PRIMARY}; }}
            QTableWidget, QTreeWidget {{
                background-color: {Colors.CARD}; alternate-background-color: #2b2d30;
                border: 1px solid {Colors.BORDER}; border-radius: 8px;
                gridline-color: {Colors.BORDER};
            }}
            QTableWidget::item:hover, QTreeWidget::item:hover {{ background-color: {Colors.HOVER}; }}
            QHeaderView::section {{
                background-color: #2a2a2a; padding: 10px;
                border: none; border-bottom: 1px solid {Colors.BORDER};
                font-weight: 600;
            }}
            QTableWidget::item:selected, QTreeWidget::item:selected {{
                background-color: {Colors.PRIMARY}; color: white;
            }}
            QStatusBar {{ background-color: {Colors.CARD}; border-top: 1px solid {Colors.BORDER}; }}
            QGroupBox {{ font-weight: bold; border: 1px solid {Colors.BORDER}; border-radius: 8px; margin-top: 10px; }}
            QGroupBox::title {{ subcontrol-origin: margin; left: 10px; padding: 0 5px; }}
            QDateEdit, QComboBox, QLineEdit {{
                background-color: {Colors.CARD}; border: 1px solid {Colors.BORDER};
                border-radius: 6px; padding: 6px;
            }}
            QComboBox QAbstractItemView {{
                background-color: {Colors.CARD};
                border: 1px solid {Colors.BORDER};
                selection-background-color: {Colors.PRIMARY};
                color: {Colors.TEXT};
            }}
            QCalendarWidget QWidget#qt_calendar_navigationbar {{ 
                background-color: {Colors.BACKGROUND}; 
            }}
            QCalendarWidget QAbstractItemView {{
                background-color: {Colors.CARD};
                color: {Colors.TEXT};
                selection-background-color: {Colors.PRIMARY};
                selection-color: white;
            }}
        """)

        self.tabs = QtWidgets.QTabWidget()
        self.setCentralWidget(self.tabs)
        self.statusBar = QtWidgets.QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("Pronto.")
        self.init_precos_tab()
        self.init_perguntas_tab()
        self.init_usuarios_tab()
        self.init_relatorios_tab()
        self.init_estatisticas_tab()
        self.tabs.currentChanged.connect(self.on_tab_changed)

    def on_tab_changed(self, index):
        if self.tabs.tabText(index) == "Estatísticas":
            self.gerar_estatisticas()

    def load_all_data(self):
        self.atualizar_labels()
        self.load_local_prices()
        self.load_local_perguntas()
        self.load_local_reports()
        self.load_local_users()

    # --- Abas ---
    def init_precos_tab(self):
        tab = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(tab)
        card = ModernCard("Gerenciamento de Preços")
        layout.addWidget(card)

        filtro_layout = QtWidgets.QHBoxLayout()
        filtro_layout.addWidget(QtWidgets.QLabel("Buscar item:"))
        self.prices_search_edit = QtWidgets.QLineEdit()
        self.prices_search_edit.setPlaceholderText("Digite item/serviço...")
        self.prices_search_edit.textChanged.connect(self.apply_prices_filter)
        filtro_layout.addWidget(self.prices_search_edit)
        card.main_layout.addLayout(filtro_layout)

        self.table = QtWidgets.QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Item/Serviço", "Preço"])
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        card.main_layout.addWidget(self.table)

        self.label_pecas = QtWidgets.QLabel()
        card.main_layout.addWidget(self.label_pecas)

        btn_layout = QtWidgets.QHBoxLayout()
        sync_btn = ModernButton("Sincronizar", icon_char="🔄", variant="secondary")
        save_btn = ModernButton("Salvar Alterações", icon_char="💾", variant="primary")
        add_btn = ModernButton("Adicionar Linha", icon_char="➕", variant="success")
        delete_btn = ModernButton("Remover Linha", icon_char="🗑️", variant="danger")
        sync_btn.clicked.connect(self.sync_prices_from_firebase)
        save_btn.clicked.connect(self.save_data)
        add_btn.clicked.connect(self.add_row)
        delete_btn.clicked.connect(self.delete_row)
        btn_layout.addWidget(sync_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.addWidget(save_btn)
        card.main_layout.addLayout(btn_layout)
        self.tabs.addTab(tab, "Preços")

    def init_perguntas_tab(self):
        tab = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(tab)
        card = ModernCard("Configuração de Perguntas Iniciais")
        layout.addWidget(card)

        self.perguntas_tree = QtWidgets.QTreeWidget()
        self.perguntas_tree.setHeaderLabels(["Pergunta", "Tipo / Opção"])
        self.perguntas_tree.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        header = self.perguntas_tree.header()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        card.main_layout.addWidget(self.perguntas_tree)

        self.label_perguntas = QtWidgets.QLabel()
        card.main_layout.addWidget(self.label_perguntas)

        btn_layout = QtWidgets.QGridLayout()
        sync_btn = ModernButton("Sincronizar", icon_char="🔄", variant="secondary")
        save_btn = ModernButton("Salvar no Firebase", icon_char="💾", variant="primary")
        add_btn = ModernButton("Adicionar", icon_char="➕", variant="success")
        edit_btn = ModernButton("Editar", icon_char="✏️", variant="secondary")
        delete_btn = ModernButton("Remover", icon_char="🗑️", variant="danger")
        move_up_btn = ModernButton("Mover para Cima", icon_char="⬆️", variant="secondary")
        move_down_btn = ModernButton("Mover para Baixo", icon_char="⬇️", variant="secondary")

        btn_layout.addWidget(sync_btn, 0, 0, 1, 2)
        btn_layout.addWidget(save_btn, 0, 2, 1, 2)
        btn_layout.addWidget(add_btn, 1, 0)
        btn_layout.addWidget(edit_btn, 1, 1)
        btn_layout.addWidget(delete_btn, 1, 2)
        btn_layout.addWidget(move_up_btn, 2, 0)
        btn_layout.addWidget(move_down_btn, 2, 1)

        sync_btn.clicked.connect(self.sync_perguntas_from_firebase)
        save_btn.clicked.connect(self.save_perguntas_to_firebase)
        add_btn.clicked.connect(self.add_pergunta)
        edit_btn.clicked.connect(self.edit_pergunta)
        delete_btn.clicked.connect(self.delete_pergunta)
        move_up_btn.clicked.connect(self.move_pergunta_up)
        move_down_btn.clicked.connect(self.move_pergunta_down)

        card.main_layout.addLayout(btn_layout)
        self.tabs.addTab(tab, "Perguntas")

    def init_usuarios_tab(self):
        tab = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(tab)
        card = ModernCard("Gerenciamento de Usuários")
        layout.addWidget(card)

        search_layout = QtWidgets.QHBoxLayout()
        search_layout.addWidget(QtWidgets.QLabel("Buscar email:"))
        self.user_search_edit = QtWidgets.QLineEdit()
        self.user_search_edit.setPlaceholderText("Filtrar usuários por email...")
        self.user_search_edit.textChanged.connect(self.apply_users_filter)
        search_layout.addWidget(self.user_search_edit)
        card.main_layout.addLayout(search_layout)

        self.user_table = QtWidgets.QTableWidget()
        self.user_table.setColumnCount(3)
        self.user_table.setHorizontalHeaderLabels(["UID", "Email", "Status"])
        self.user_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.user_table.setSelectionMode(
            QtWidgets.QAbstractItemView.SelectionMode.MultiSelection)  # Permitir multisseleção
        header = self.user_table.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        card.main_layout.addWidget(self.user_table)

        self.label_usuarios = QtWidgets.QLabel()
        card.main_layout.addWidget(self.label_usuarios)

        btn_layout = QtWidgets.QHBoxLayout()
        sync_btn = ModernButton("Sincronizar", icon_char="🔄", variant="secondary")
        add_btn = ModernButton("Adicionar Usuário", icon_char="➕", variant="secondary")
        self.save_users_btn = ModernButton("Salvar Novos Usuários", icon_char="💾", variant="primary")
        self.save_users_btn.setEnabled(False)
        discard_btn = ModernButton("Descartar", icon_char="🗑️", variant="danger")
        edit_btn = ModernButton("Editar Status", icon_char="✏️", variant="secondary")
        # *** CORREÇÃO: Adicionar o botão de apagar ***
        delete_btn = ModernButton("Apagar Usuário", icon_char="🗑️", variant="danger")

        sync_btn.clicked.connect(self.sync_users_from_firebase)
        add_btn.clicked.connect(self.add_user_local)
        self.save_users_btn.clicked.connect(self.save_new_users_to_firebase)
        discard_btn.clicked.connect(self.discard_new_users)
        edit_btn.clicked.connect(self.edit_user)
        # *** CORREÇÃO: Conectar o botão ***
        delete_btn.clicked.connect(self.delete_user)

        btn_layout.addWidget(sync_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(discard_btn)
        btn_layout.addWidget(self.save_users_btn)
        btn_layout.addWidget(edit_btn)
        # *** CORREÇÃO: Adicionar o botão ao layout ***
        btn_layout.addWidget(delete_btn)

        card.main_layout.addLayout(btn_layout)
        self.tabs.addTab(tab, "Usuários")

    def init_relatorios_tab(self):
        tab = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(tab)
        card = ModernCard("Relatórios de Orçamentos")
        layout.addWidget(card)

        filter_layout = QtWidgets.QHBoxLayout()
        filter_layout.addWidget(QtWidgets.QLabel("Buscar:"))
        self.report_search_edit = QtWidgets.QLineEdit()
        self.report_search_edit.setPlaceholderText("Cliente, vendedor ou item do orçamento...")
        self.report_search_edit.textChanged.connect(self.apply_reports_filter)
        filter_layout.addWidget(self.report_search_edit)
        card.main_layout.addLayout(filter_layout)

        self.reports_table = QtWidgets.QTableWidget()
        self.reports_table.setColumnCount(4)
        self.reports_table.setHorizontalHeaderLabels(["Data", "Cliente", "Orçamentista", "Valor"])
        self.reports_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.reports_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.reports_table.setAlternatingRowColors(True)
        self.reports_table.setSortingEnabled(True)
        self.reports_table.cellDoubleClicked.connect(self.show_report_details)
        header = self.reports_table.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        card.main_layout.addWidget(self.reports_table)

        self.label_relatorios = QtWidgets.QLabel("Nenhum relatório carregado.")
        card.main_layout.addWidget(self.label_relatorios)

        resumo_layout = QtWidgets.QHBoxLayout()
        self.reports_kpi_filtrados = QtWidgets.QLabel("Exibidos: 0")
        self.reports_kpi_valor = QtWidgets.QLabel("Valor filtrado: R$ 0,00")
        self.reports_kpi_ticket = QtWidgets.QLabel("Ticket filtrado: R$ 0,00")
        for lbl in (self.reports_kpi_filtrados, self.reports_kpi_valor, self.reports_kpi_ticket):
            lbl.setStyleSheet(f"color: {Colors.TEXT_SECONDARY}; font-weight: 600;")
            resumo_layout.addWidget(lbl)
        resumo_layout.addStretch()
        card.main_layout.addLayout(resumo_layout)

        btn_layout = QtWidgets.QHBoxLayout()
        sync_btn = ModernButton("Importar Novos Relatórios", icon_char="📥", variant="primary")
        sync_btn.clicked.connect(self.sync_reports_from_firebase)
        btn_layout.addStretch()
        btn_layout.addWidget(sync_btn)
        card.main_layout.addLayout(btn_layout)
        self.tabs.addTab(tab, "Relatórios")

    def init_estatisticas_tab(self):
        tab = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(tab)
        layout.setSpacing(14)

        filter_card = ModernCard("Visão Estratégica")
        filter_layout = QtWidgets.QHBoxLayout()

        filter_layout.addWidget(QtWidgets.QLabel("Período:"))
        self.start_date_edit = QtWidgets.QDateEdit(calendarPopup=True)
        self.start_date_edit.setDate(datetime.now().date().replace(day=1))
        self.start_date_edit.setMaximumWidth(130)
        filter_layout.addWidget(self.start_date_edit)

        filter_layout.addWidget(QtWidgets.QLabel("até"))
        self.end_date_edit = QtWidgets.QDateEdit(calendarPopup=True)
        self.end_date_edit.setDate(datetime.now().date())
        self.end_date_edit.setMaximumWidth(130)
        filter_layout.addWidget(self.end_date_edit)

        filter_layout.addSpacing(12)
        filter_layout.addWidget(QtWidgets.QLabel("Vendedor:"))
        self.vendedor_combo = QtWidgets.QComboBox()
        self.vendedor_combo.setMinimumWidth(280)
        self.vendedor_combo.setSizePolicy(QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Fixed)
        self.vendedor_combo.setToolTip("Selecione um vendedor para filtrar")
        filter_layout.addWidget(self.vendedor_combo)

        refresh_btn = ModernButton("Atualizar", icon_char="🔎", variant="secondary")
        refresh_btn.clicked.connect(self.gerar_estatisticas)
        filter_layout.addWidget(refresh_btn)

        filter_card.main_layout.addLayout(filter_layout)

        self.periodo_aplicado_label = QtWidgets.QLabel("Nenhum filtro aplicado.")
        self.periodo_aplicado_label.setStyleSheet(f"color: {Colors.TEXT_SECONDARY}; font-size: 12px; border: none;")
        filter_card.main_layout.addWidget(self.periodo_aplicado_label)
        layout.addWidget(filter_card)

        kpi_layout = QtWidgets.QGridLayout()
        kpi_layout.setHorizontalSpacing(12)
        kpi_layout.setVerticalSpacing(12)

        self.kpi_total_orcamentos = StatKpiCard("Total de orçamentos", Colors.PRIMARY)
        self.kpi_ticket = StatKpiCard("Ticket médio", Colors.INFO)
        self.kpi_mediana = StatKpiCard("Mediana", Colors.SUCCESS)
        self.kpi_cobertura = StatKpiCard("Cobertura do questionário", Colors.WARNING)
        self.kpi_maior = StatKpiCard("Maior orçamento", Colors.PRIMARY_DARK)
        self.kpi_menor = StatKpiCard("Menor orçamento", Colors.DANGER)

        kpi_layout.addWidget(self.kpi_total_orcamentos, 0, 0)
        kpi_layout.addWidget(self.kpi_ticket, 0, 1)
        kpi_layout.addWidget(self.kpi_mediana, 0, 2)
        kpi_layout.addWidget(self.kpi_cobertura, 1, 0)
        kpi_layout.addWidget(self.kpi_maior, 1, 1)
        kpi_layout.addWidget(self.kpi_menor, 1, 2)
        layout.addLayout(kpi_layout)

        rankings_layout = QtWidgets.QHBoxLayout()

        self.vendedor_ranking_tree = QtWidgets.QTreeWidget()
        self.vendedor_ranking_tree.setHeaderLabels(["Vendedor", "Qtd", "%", "Ticket médio"])
        self.vendedor_ranking_tree.header().setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeMode.Stretch)
        self.vendedor_ranking_tree.header().setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        self.vendedor_ranking_tree.header().setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        self.vendedor_ranking_tree.header().setSectionResizeMode(3, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        vendedores_card = ModernCard("Performance por vendedor")
        vendedores_card.main_layout.addWidget(self.vendedor_ranking_tree)
        rankings_layout.addWidget(vendedores_card)

        self.faixa_valor_tree = QtWidgets.QTreeWidget()
        self.faixa_valor_tree.setHeaderLabels(["Faixa", "Qtd", "%"])
        self.faixa_valor_tree.header().setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeMode.Stretch)
        self.faixa_valor_tree.header().setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        self.faixa_valor_tree.header().setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        faixa_card = ModernCard("Distribuição por faixa de orçamento")
        faixa_card.main_layout.addWidget(self.faixa_valor_tree)
        rankings_layout.addWidget(faixa_card)

        layout.addLayout(rankings_layout, 2)

        perguntas_card = ModernCard("Insights do questionário")
        pergunta_filter_layout = QtWidgets.QHBoxLayout()
        pergunta_filter_layout.addWidget(QtWidgets.QLabel("Pergunta:"))
        self.perguntas_iniciais_combo = QtWidgets.QComboBox()
        self.perguntas_iniciais_combo.setSizePolicy(QtWidgets.QSizePolicy.Policy.Expanding,
                                                    QtWidgets.QSizePolicy.Policy.Fixed)
        self.perguntas_iniciais_combo.setToolTip("Selecione uma pergunta para ver as estatísticas")
        self.perguntas_iniciais_combo.currentIndexChanged.connect(self.atualizar_estatisticas_pergunta_selecionada)
        pergunta_filter_layout.addWidget(self.perguntas_iniciais_combo)
        perguntas_card.main_layout.addLayout(pergunta_filter_layout)

        self.perguntas_iniciais_stats_tree = QtWidgets.QTreeWidget()
        self.perguntas_iniciais_stats_tree.setHeaderLabels(["Resposta", "Qtd", "%", "Preço médio"])
        self.perguntas_iniciais_stats_tree.header().setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeMode.Stretch)
        self.perguntas_iniciais_stats_tree.header().setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        self.perguntas_iniciais_stats_tree.header().setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        self.perguntas_iniciais_stats_tree.header().setSectionResizeMode(3, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        self.perguntas_iniciais_stats_tree.setMinimumHeight(220)
        perguntas_card.main_layout.addWidget(self.perguntas_iniciais_stats_tree)
        layout.addWidget(perguntas_card, 1)

        self.start_date_edit.dateChanged.connect(self.gerar_estatisticas)
        self.end_date_edit.dateChanged.connect(self.gerar_estatisticas)
        self.vendedor_combo.currentTextChanged.connect(self.gerar_estatisticas)
        self.tabs.addTab(tab, "Estatísticas")

    def atualizar_filtro_vendedores(self):
        self.vendedor_combo.clear()
        self.vendedor_combo.addItem("Todos")
        emails = sorted(
            {r.get("orcamentistaEmail", "N/A") for r in self.local_reports_data if r.get("orcamentistaEmail")})
        self.vendedor_combo.addItems(emails)
        try:
            self.vendedor_combo.currentTextChanged.disconnect(self._update_vendedor_tooltip)
        except TypeError:
            pass
        self.vendedor_combo.currentTextChanged.connect(self._update_vendedor_tooltip)
        self._update_vendedor_tooltip(self.vendedor_combo.currentText())

    def _update_vendedor_tooltip(self, text):
        if text and text != "Todos":
            self.vendedor_combo.setToolTip(f"Vendedor selecionado: {text}")
        else:
            self.vendedor_combo.setToolTip("Selecione um vendedor para filtrar")

    # --- Métodos de Limite e Labels ---
    def carregar_limites(self):
        if os.path.exists(LIMITE_FILE):
            with open(LIMITE_FILE, "r") as f:
                return json.load(f)
        return {"data": str(datetime.now().date()), "pecas": 0, "usuarios": 0,
                "perguntas_sync": 0, "perguntas_save": 0, "relatorios_sync": 0}

    def salvar_limites(self):
        with open(LIMITE_FILE, "w") as f:
            json.dump(self.limites, f)

    def verificar_reset_diario(self):
        hoje = str(datetime.now().date())
        if self.limites.get("data") != hoje:
            self.limites = {"data": hoje, "pecas": 0, "usuarios": 0,
                            "perguntas_sync": 0, "perguntas_save": 0, "relatorios_sync": 0}
            self.salvar_limites()
            self.atualizar_labels()

    def atualizar_labels(self):
        self.label_pecas.setText(f"Acessos de Peças (hoje): {self.limites.get('pecas', 0)}/{self.max_acessos_pecas}")
        self.label_usuarios.setText(
            f"Acessos de Usuários (hoje): {self.limites.get('usuarios', 0)}/{self.max_acessos_usuarios}")
        psync, psave = self.limites.get('perguntas_sync', 0), self.limites.get('perguntas_save', 0)
        self.label_perguntas.setText(
            f"Sincronizações (hoje): {psync}/{self.max_acessos_perg_sync} | Envios (hoje): {psave}/{self.max_acessos_perg_save}")
        rsync = self.limites.get('relatorios_sync', 0)
        self.label_relatorios.setText(
            f"{len(self.local_reports_data)} relatórios carregados. | Sincronizações (hoje): {rsync}/{self.max_acessos_rel_sync}")

    # --- MÉTODOS FUNCIONAIS ---
    def load_local_perguntas(self):
        if os.path.exists(PERGUNTAS_INICIAIS_FILE):
            try:
                with open(PERGUNTAS_INICIAIS_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self.perguntas_doc_id = data.get("_id")
                self.perguntas_data = {"ordem": data.get("ordem", []), "perguntas": data.get("perguntas", {})}
                self.statusBar.showMessage("Perguntas iniciais carregadas do arquivo local.")
            except (json.JSONDecodeError, KeyError) as e:
                self.statusBar.showMessage(f"Erro ao ler arquivo de perguntas: {e}. Inicie uma sincronização.")
                self.perguntas_data = {"ordem": [], "perguntas": {}}
        else:
            self.statusBar.showMessage("Arquivo de perguntas não encontrado. Sincronize do Firebase.")
        self.populate_perguntas_tree()
        self._stats_cache_key = None

    def sync_perguntas_from_firebase(self):
        if self.limites.get("perguntas_sync", 0) >= self.max_acessos_perg_sync:
            ModernMessageBox.critical(self, "Erro",
                                      f"Limite diário de sincronização atingido ({self.max_acessos_perg_sync}/{self.max_acessos_perg_sync})!")
            return
        self.statusBar.showMessage("Sincronizando perguntas iniciais do Firebase...")
        try:
            docs = list(db.collection("perguntas_iniciais").limit(1).stream())
            if not docs:
                ModernMessageBox.information(self, "Aviso", "Nenhuma configuração de perguntas encontrada no Firebase.")
                self.perguntas_doc_id, self.perguntas_data = None, {"ordem": [], "perguntas": {}}
                if os.path.exists(PERGUNTAS_INICIAIS_FILE): os.remove(PERGUNTAS_INICIAIS_FILE)
            else:
                doc = docs[0]
                self.perguntas_doc_id = doc.id
                self.perguntas_data = doc.to_dict()
                if "ordem" not in self.perguntas_data: self.perguntas_data["ordem"] = []
                if "perguntas" not in self.perguntas_data: self.perguntas_data["perguntas"] = {}
                data_to_save = {**self.perguntas_data, "_id": self.perguntas_doc_id}
                with open(PERGUNTAS_INICIAIS_FILE, 'w', encoding='utf-8') as f:
                    json.dump(data_to_save, f, indent=4, ensure_ascii=False)
                ModernMessageBox.information(self, "Sucesso", "Perguntas iniciais sincronizadas com sucesso!")
            self.limites["perguntas_sync"] += 1
            self.salvar_limites()
            self.atualizar_labels()
            self.populate_perguntas_tree()
            self.statusBar.showMessage("Sincronização de perguntas concluída.")
        except Exception as e:
            ModernMessageBox.critical(self, "Erro", f"Falha ao sincronizar perguntas:\n{e}")
            self.statusBar.showMessage("Erro de sincronização.")

    def populate_perguntas_tree(self):
        self.perguntas_tree.clear()
        for titulo in self.perguntas_data.get("ordem", []):
            config = self.perguntas_data.get("perguntas", {}).get(titulo, {})
            item = QtWidgets.QTreeWidgetItem([titulo, config.get("tipo", "desconhecido")])
            self.perguntas_tree.addTopLevelItem(item)
            if config.get("tipo") == "opcoes":
                for opcao in config.get("opcoes", []):
                    item.addChild(QtWidgets.QTreeWidgetItem(["", opcao]))
        self.perguntas_tree.expandAll()

    def add_pergunta(self):
        dialog = PerguntaDialog(parent=self)
        if dialog.exec():
            data = dialog.get_data()
            if data:
                titulo = data.pop("titulo")
                if titulo in self.perguntas_data["perguntas"]:
                    ModernMessageBox.warning(self, "Erro", "Uma pergunta com este título já existe.")
                    return
                self.perguntas_data["perguntas"][titulo] = data
                self.perguntas_data["ordem"].append(titulo)
                self.populate_perguntas_tree()

    def edit_pergunta(self):
        selected = self.perguntas_tree.currentItem()
        if not selected or selected.parent():
            ModernMessageBox.warning(self, "Aviso", "Selecione uma pergunta para editar.")
            return
        titulo_original = selected.text(0)
        config_original = self.perguntas_data["perguntas"][titulo_original]
        pergunta_data = {"titulo": titulo_original, **config_original}
        dialog = PerguntaDialog(pergunta_data=pergunta_data, parent=self)
        if dialog.exec():
            novo_data = dialog.get_data()
            if novo_data:
                novo_titulo = novo_data.pop("titulo")
                self.perguntas_data["perguntas"].pop(titulo_original)
                self.perguntas_data["perguntas"][novo_titulo] = novo_data
                index = self.perguntas_data["ordem"].index(titulo_original)
                self.perguntas_data["ordem"][index] = novo_titulo
                self.populate_perguntas_tree()

    def delete_pergunta(self):
        selected = self.perguntas_tree.currentItem()
        if not selected or selected.parent():
            ModernMessageBox.warning(self, "Aviso", "Selecione uma pergunta para remover.")
            return
        titulo = selected.text(0)
        reply = ModernMessageBox.question(self, "Confirmar Remoção",
                                          f"Tem certeza que deseja remover a pergunta '{titulo}'?")
        if reply == QtWidgets.QMessageBox.StandardButton.Yes:
            self.perguntas_data["perguntas"].pop(titulo, None)
            if titulo in self.perguntas_data["ordem"]: self.perguntas_data["ordem"].remove(titulo)
            self.populate_perguntas_tree()

    def _move_pergunta(self, direction):
        selected = self.perguntas_tree.currentItem()
        if not selected or selected.parent(): return
        titulo = selected.text(0)
        index = self.perguntas_data["ordem"].index(titulo)
        new_index = index + direction
        if 0 <= new_index < len(self.perguntas_data["ordem"]):
            self.perguntas_data["ordem"].insert(new_index, self.perguntas_data["ordem"].pop(index))
            self.populate_perguntas_tree()
            for i in range(self.perguntas_tree.topLevelItemCount()):
                item = self.perguntas_tree.topLevelItem(i)
                if item.text(0) == titulo:
                    self.perguntas_tree.setCurrentItem(item)
                    break

    def move_pergunta_up(self):
        self._move_pergunta(-1)

    def move_pergunta_down(self):
        self._move_pergunta(1)

    def save_perguntas_to_firebase(self):
        if self.limites.get("perguntas_save", 0) >= self.max_acessos_perg_save:
            ModernMessageBox.critical(self, "Erro",
                                      f"Limite diário para salvar perguntas atingido ({self.max_acessos_perg_save}/{self.max_acessos_perg_save})!")
            return
        reply = ModernMessageBox.question(self, "Confirmar",
                                          "Isso irá sobrescrever a configuração no Firebase. Deseja continuar?")
        if reply != QtWidgets.QMessageBox.StandardButton.Yes: return
        self.statusBar.showMessage("Salvando perguntas no Firebase...")
        try:
            if self.perguntas_doc_id:
                db.collection("perguntas_iniciais").document(self.perguntas_doc_id).set(self.perguntas_data)
            else:
                _, doc_ref = db.collection("perguntas_iniciais").add(self.perguntas_data)
                self.perguntas_doc_id = doc_ref.id
            data_to_save = {**self.perguntas_data, "_id": self.perguntas_doc_id}
            with open(PERGUNTAS_INICIAIS_FILE, 'w', encoding='utf-8') as f:
                json.dump(data_to_save, f, indent=4, ensure_ascii=False)
            self.limites["perguntas_save"] += 1
            self.salvar_limites()
            self.atualizar_labels()
            ModernMessageBox.information(self, "Sucesso", "Perguntas salvas com sucesso!")
            self.statusBar.showMessage("Pronto.")
        except Exception as e:
            ModernMessageBox.critical(self, "Erro", f"Falha ao salvar perguntas:\n{e}")
            self.statusBar.showMessage("Erro ao salvar.")

    def load_local_prices(self):
        self.table.setRowCount(0)
        self.doc_ids.clear()
        if not os.path.exists(PRECOS_FILE):
            self.statusBar.showMessage("Arquivo de preços local não encontrado.")
            return
        with open(PRECOS_FILE, 'r', encoding='utf-8') as f:
            local_prices = json.load(f)
        for i, item in enumerate(local_prices):
            self.table.insertRow(i)
            self.table.setItem(i, 0, QtWidgets.QTableWidgetItem(item.get("selecionado", "")))
            self.table.setItem(i, 1, QtWidgets.QTableWidgetItem(str(item.get("precos", ""))))
            self.doc_ids[i] = item.get("_id")
        self.apply_prices_filter()
        self.statusBar.showMessage(f"{len(local_prices)} preços carregados do arquivo local.")

    def apply_prices_filter(self):
        search = self.prices_search_edit.text().strip().lower() if hasattr(self, "prices_search_edit") else ""
        for row in range(self.table.rowCount()):
            item_text = (self.table.item(row, 0).text() if self.table.item(row, 0) else "").lower()
            self.table.setRowHidden(row, bool(search and search not in item_text))

    def sync_prices_from_firebase(self):
        if self.limites.get("pecas", 0) >= self.max_acessos_pecas:
            ModernMessageBox.critical(self, "Erro", "Limite diário de acessos às peças atingido!")
            return
        self.statusBar.showMessage("Sincronizando preços do Firebase...")
        try:
            docs = db.collection("precos").stream()
            prices_to_save = [{**doc.to_dict(), "_id": doc.id} for doc in docs]
            with open(PRECOS_FILE, 'w', encoding='utf-8') as f:
                json.dump(prices_to_save, f, indent=4, ensure_ascii=False)
            self.limites["pecas"] += 1
            self.salvar_limites()
            self.atualizar_labels()
            self.load_local_prices()
            ModernMessageBox.information(self, "Sucesso", f"{len(prices_to_save)} preços sincronizados!")
        except Exception as e:
            ModernMessageBox.critical(self, "Erro", f"Falha ao sincronizar preços:\n{e}")
        finally:
            self.statusBar.showMessage("Pronto.")

    def save_data(self):
        try:
            for row in range(self.table.rowCount()):
                selecionado = self.table.item(row, 0).text() if self.table.item(row, 0) else ""
                preco = self.table.item(row, 1).text() if self.table.item(row, 1) else ""
                if preco:
                    try:
                        parse_decimal(preco)
                    except ValueError:
                        ModernMessageBox.warning(self, "Erro",
                                                 f"O valor '{preco}' na linha {row + 1} não é um número válido.")
                        return
                if preco and parse_decimal(preco) is None:
                    ModernMessageBox.warning(self, "Erro",
                                             f"O valor '{preco}' na linha {row + 1} não é válido.")
                    return
                doc_id = self.doc_ids.get(row)
                data_to_save = {"selecionado": selecionado, "precos": preco}
                if doc_id:
                    db.collection("precos").document(doc_id).update(data_to_save)
                else:
                    _, doc_ref = db.collection("precos").add(data_to_save)
                    self.doc_ids[row] = doc_ref.id
            self._save_prices_to_local_file()
            ModernMessageBox.information(self, "Sucesso", "Alterações salvas com sucesso!")
        except Exception as e:
            ModernMessageBox.critical(self, "Erro", f"Falha ao salvar dados:\n{e}")

    def _save_prices_to_local_file(self):
        prices_to_save = []
        for row in range(self.table.rowCount()):
            selecionado = self.table.item(row, 0).text() if self.table.item(row, 0) else ""
            preco = self.table.item(row, 1).text() if self.table.item(row, 1) else ""
            prices_to_save.append({"_id": self.doc_ids.get(row), "selecionado": selecionado, "precos": preco})
        with open(PRECOS_FILE, 'w', encoding='utf-8') as f:
            json.dump(prices_to_save, f, indent=4, ensure_ascii=False)

    def add_row(self):
        self.table.insertRow(self.table.rowCount())

    def delete_row(self):
        selected_rows = sorted({index.row() for index in self.table.selectedIndexes()}, reverse=True)
        if not selected_rows: return
        if ModernMessageBox.question(self, "Confirmação",
                                     f"Deletar {len(selected_rows)} linha(s)?") == QtWidgets.QMessageBox.StandardButton.Yes:
            for row in selected_rows:
                doc_id = self.doc_ids.get(row)
                if doc_id: db.collection("precos").document(doc_id).delete()
                self.table.removeRow(row)
            self._save_prices_to_local_file()
            ModernMessageBox.information(self, "Sucesso", "Linha(s) deletada(s)!")

    def load_local_users(self):
        self.user_table.setRowCount(0)
        if not os.path.exists(USUARIOS_FILE):
            self.statusBar.showMessage("Arquivo de usuários local não encontrado.")
        else:
            with open(USUARIOS_FILE, 'r', encoding='utf-8') as f:
                local_users = json.load(f)
            for i, user in enumerate(local_users):
                self.user_table.insertRow(i)
                uid_item = QtWidgets.QTableWidgetItem(user.get("uid", ""))
                email_item = QtWidgets.QTableWidgetItem(user.get("email", ""))
                status_item = QtWidgets.QTableWidgetItem("Desativado" if user.get("disabled") else "Ativo")
                if user.get("disabled"):
                    for item in (uid_item, email_item, status_item):
                        item.setForeground(QColor("#ff7b7b"))
                self.user_table.setItem(i, 0, uid_item)
                self.user_table.setItem(i, 1, email_item)
                self.user_table.setItem(i, 2, status_item)
            self.statusBar.showMessage(f"{len(local_users)} usuários carregados.")

        # Adiciona usuários do cache à tabela
        for user_data in self.novos_usuarios_cache:
            row_position = self.user_table.rowCount()
            self.user_table.insertRow(row_position)

            uid_item = QtWidgets.QTableWidgetItem("(Pendente)")
            uid_item.setForeground(QColor(Colors.TEXT_SECONDARY))

            email_item = QtWidgets.QTableWidgetItem(user_data['email'])

            status_item = QtWidgets.QTableWidgetItem("Pendente")
            status_item.setForeground(QColor(Colors.PRIMARY))
            status_item.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))

            for col in range(3):
                item = QtWidgets.QTableWidgetItem()
                item.setBackground(QColor("#e7f3fe"))
                self.user_table.setItem(row_position, col, item)

            self.user_table.setItem(row_position, 0, uid_item)
            self.user_table.setItem(row_position, 1, email_item)
            self.user_table.setItem(row_position, 2, status_item)

        self.save_users_btn.setEnabled(bool(self.novos_usuarios_cache))
        self.apply_users_filter()

    def apply_users_filter(self):
        search = self.user_search_edit.text().strip().lower() if hasattr(self, "user_search_edit") else ""
        for row in range(self.user_table.rowCount()):
            email = (self.user_table.item(row, 1).text() if self.user_table.item(row, 1) else "").lower()
            self.user_table.setRowHidden(row, bool(search and search not in email))

    def sync_users_from_firebase(self):
        if self.novos_usuarios_cache:
            reply = ModernMessageBox.question(self, "Descartar Alterações?",
                                              "Você tem usuários pendentes não salvos. Sincronizar irá descartá-los. Deseja continuar?")
            if reply != QtWidgets.QMessageBox.StandardButton.Yes:
                return

        self.novos_usuarios_cache.clear()

        if self.limites.get("usuarios", 0) >= self.max_acessos_usuarios:
            ModernMessageBox.critical(self, "Erro", "Limite diário de acessos aos usuários atingido!")
            return
        self.statusBar.showMessage("Sincronizando usuários...")
        try:
            users_to_save = [{"uid": u.uid, "email": u.email or "", "disabled": u.disabled} for u in
                             auth.list_users().users]
            with open(USUARIOS_FILE, 'w', encoding='utf-8') as f:
                json.dump(users_to_save, f, indent=4, ensure_ascii=False)
            self.limites["usuarios"] += 1
            self.salvar_limites()
            self.atualizar_labels()
            self.load_local_users()
            ModernMessageBox.information(self, "Sucesso", f"{len(users_to_save)} usuários sincronizados!")
        except Exception as e:
            ModernMessageBox.critical(self, "Erro", f"Erro ao sincronizar usuários: {e}")
        finally:
            self.statusBar.showMessage("Pronto.")

    def add_user_local(self):
        dialog = AddUserDialog(self)
        if dialog.exec():
            email, senha = dialog.get_data()
            if email and senha:
                self.novos_usuarios_cache.append({"email": email, "password": senha})
                self.load_local_users()  # Recarrega a tabela para mostrar o novo usuário pendente

    def save_new_users_to_firebase(self):
        if not self.novos_usuarios_cache:
            ModernMessageBox.information(self, "Aviso", "Nenhum novo usuário para salvar.")
            return

        total_novos = len(self.novos_usuarios_cache)
        sucesso, falha = 0, 0
        erros_msg = []

        self.statusBar.showMessage(f"Salvando {total_novos} novo(s) usuário(s)...")

        for user_data in self.novos_usuarios_cache:
            try:
                auth.create_user(email=user_data['email'], password=user_data['password'])
                sucesso += 1
            except Exception as e:
                falha += 1
                erros_msg.append(f"- {user_data['email']}: {e}")

        self.novos_usuarios_cache.clear()

        resultado_msg = f"{sucesso} de {total_novos} usuários salvos com sucesso."
        if falha > 0:
            resultado_msg += f"\n\n{falha} falharam:\n" + "\n".join(erros_msg)
            ModernMessageBox.warning(self, "Resultado do Salvamento", resultado_msg)
        else:
            ModernMessageBox.information(self, "Resultado do Salvamento", resultado_msg)

        self.sync_users_from_firebase()

    def discard_new_users(self):
        if not self.novos_usuarios_cache:
            return
        reply = ModernMessageBox.question(self, "Descartar Novos Usuários?",
                                          "Tem certeza que deseja descartar todos os usuários pendentes?")
        if reply == QtWidgets.QMessageBox.StandardButton.Yes:
            self.novos_usuarios_cache.clear()
            self.load_local_users()
            self.statusBar.showMessage("Alterações descartadas.")

    def edit_user(self):
        selected_rows = self.user_table.selectionModel().selectedRows()
        if not selected_rows:
            ModernMessageBox.warning(self, "Aviso", "Selecione um usuário para editar.")
            return
        if len(selected_rows) > 1:
            ModernMessageBox.warning(self, "Aviso", "Selecione apenas um usuário para editar.")
            return

        row = selected_rows[0].row()
        uid = self.user_table.item(row, 0).text()
        current_email = self.user_table.item(row, 1).text()
        current_status = self.user_table.item(row, 2).text()

        if not uid or uid == "(Pendente)":
            ModernMessageBox.warning(self, "Aviso", "Não é possível editar um usuário pendente. Salve primeiro.")
            return

        dialog = EditUserStatusDialog(current_email, current_status, self)
        if dialog.exec():
            status = dialog.get_status()
            try:
                auth.update_user(uid, disabled=(status == "Desativado"))
                ModernMessageBox.information(self, "Sucesso", f"Status do usuário atualizado!")
                self.sync_users_from_firebase()
            except Exception as e:
                ModernMessageBox.critical(self, "Erro", f"Falha ao atualizar usuário:\n{e}")

    # *** CORREÇÃO: Função delete_user atualizada ***
    def delete_user(self):
        selected_row_items = self.user_table.selectionModel().selectedRows()
        if not selected_row_items:
            ModernMessageBox.warning(self, "Aviso", "Selecione pelo menos um usuário para apagar.")
            return

        users_to_delete_firebase = []
        rows_to_discard_local = []

        for row_item in selected_row_items:
            row = row_item.row()
            uid_item = self.user_table.item(row, 0)
            email_item = self.user_table.item(row, 1)

            if not uid_item or not email_item:
                continue

            uid = uid_item.text()
            email = email_item.text()

            if uid == "(Pendente)":
                rows_to_discard_local.append({"row": row, "email": email})
            else:
                users_to_delete_firebase.append({"uid": uid, "email": email})

        # --- Processar usuários pendentes ---
        if rows_to_discard_local:
            count = len(rows_to_discard_local)
            reply = ModernMessageBox.question(self, "Descartar Usuários Pendentes?",
                                              f"{count} usuário(s) selecionado(s) são pendentes (não salvos).\n"
                                              "Eles serão descartados localmente. Deseja continuar?")

            if reply == QtWidgets.QMessageBox.StandardButton.Yes:
                # É mais seguro iterar ao contrário ao remover da GUI
                for user_data in sorted(rows_to_discard_local, key=lambda x: x['row'], reverse=True):
                    # Remove do cache
                    self.novos_usuarios_cache = [u for u in self.novos_usuarios_cache if
                                                 u['email'] != user_data['email']]
                    # Remove da tabela
                    self.user_table.removeRow(user_data['row'])

                self.load_local_users()  # Recarrega para garantir consistência
                self.statusBar.showMessage(f"{count} usuário(s) pendente(s) descartado(s).")

            if not users_to_delete_firebase:  # Se SÓ havia pendentes, paramos aqui
                return

        # --- Processar usuários do Firebase ---
        if not users_to_delete_firebase:
            return

        count = len(users_to_delete_firebase)
        email_list = "\n".join([f"- {u['email']}" for u in users_to_delete_firebase[:5]])  # Mostra até 5 emails
        if count > 5:
            email_list += f"\n... e mais {count - 5}."

        reply = ModernMessageBox.question(self, "Confirmar Deleção Permanente",
                                          f"Tem certeza que deseja apagar {count} usuário(s) PERMANENTEMENTE do Firebase?\n\n{email_list}\n\nEsta ação não pode ser desfeita.")

        if reply != QtWidgets.QMessageBox.StandardButton.Yes:
            return

        sucesso = 0
        falha_msgs = []
        self.statusBar.showMessage(f"Apagando {count} usuário(s)...")
        try:
            for user in users_to_delete_firebase:
                try:
                    auth.delete_user(user['uid'])
                    sucesso += 1
                except Exception as e:
                    falha_msgs.append(f"- {user['email']}: {e}")

            if falha_msgs:
                ModernMessageBox.warning(self, "Resultado da Deleção",
                                         f"{sucesso} usuário(s) apagado(s) com sucesso.\n\n"
                                         f"Falha ao apagar {len(falha_msgs)}:\n" + "\n".join(falha_msgs))
            else:
                ModernMessageBox.information(self, "Sucesso", f"{sucesso} usuário(s) apagado(s) com sucesso.")

            # Sincroniza para atualizar a lista após apagar
            self.sync_users_from_firebase()

        except Exception as e:
            ModernMessageBox.critical(self, "Erro Crítico", f"Falha ao processar deleções:\n{e}")
            self.statusBar.showMessage("Erro ao apagar usuários.")

    def load_local_reports(self):
        self.local_reports_data = []
        if not os.path.exists(RELATORIOS_FILE):
            self.atualizar_labels()
            return
        with open(RELATORIOS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            self.local_reports_data = data.get("reports", [])
        self._stats_cache_key = None
        self.apply_reports_filter()
        self.atualizar_labels()
        self.atualizar_filtro_vendedores()
        self.gerar_estatisticas()

    def _extract_report_value(self, report):
        total = 0.0
        for item in report.get("itensOrcamento", []):
            valor = item.get("valor", 0) or 0
            if valor > 0:
                total += valor
        if total > 0:
            return total
        valor_str = report.get("estimativaFormatada", "")
        numbers = re.findall(r'[\d\.,]+', valor_str)
        vals = []
        for n in numbers:
            try:
                vals.append(float(n.replace('.', '').replace(',', '.')))
            except ValueError:
                pass
        return (sum(vals) / len(vals)) if vals else 0.0

    def apply_reports_filter(self):
        query = self.report_search_edit.text().strip().lower() if hasattr(self, "report_search_edit") else ""
        self.reports_table.setSortingEnabled(False)
        self.reports_table.setRowCount(0)
        self.relatorios_filtrados_cache = []
        total_filtrado = 0.0
        primeira_pergunta = next(iter(self.perguntas_data.get("ordem", [])), "N/A")
        for report in self.local_reports_data:
            cliente = (report.get("respostasIniciais", {}).get(primeira_pergunta, "N/A") or "N/A")
            vendedor = report.get("orcamentistaEmail", "N/A")
            itens_texto = " ".join((i.get("descricao", "") for i in report.get("itensOrcamento", []))).lower()
            lookup = f"{cliente} {vendedor} {itens_texto}".lower()
            if query and query not in lookup:
                continue
            row = self.reports_table.rowCount()
            self.reports_table.insertRow(row)
            data_obj = safe_date_from_iso(report.get("criadoEm"))
            data_texto = data_obj.strftime("%d/%m/%Y %H:%M") if data_obj else "Data Inválida"
            item_data = QtWidgets.QTableWidgetItem(data_texto)
            if data_obj:
                item_data.setData(Qt.ItemDataRole.UserRole, data_obj.timestamp())
            item_data.setData(Qt.ItemDataRole.UserRole + 1, report.get("_id", str(id(report))))
            self.reports_table.setItem(row, 0, item_data)
            self.reports_table.setItem(row, 1, QtWidgets.QTableWidgetItem(cliente))
            self.reports_table.setItem(row, 2, QtWidgets.QTableWidgetItem(vendedor))
            valor = self._extract_report_value(report)
            total_filtrado += valor
            valor_item = QtWidgets.QTableWidgetItem(format_currency(valor))
            valor_item.setData(Qt.ItemDataRole.UserRole, valor)
            self.reports_table.setItem(row, 3, valor_item)
            self.relatorios_filtrados_cache.append(report)
        self.reports_table.setSortingEnabled(True)

        qtd = len(self.relatorios_filtrados_cache)
        ticket = total_filtrado / qtd if qtd else 0
        self.reports_kpi_filtrados.setText(f"Exibidos: {qtd}")
        self.reports_kpi_valor.setText(f"Valor filtrado: {format_currency(total_filtrado)}")
        self.reports_kpi_ticket.setText(f"Ticket filtrado: {format_currency(ticket)}")

    def sync_reports_from_firebase(self):

        self.statusBar.showMessage("Sincronizando relatórios...")

        local_data = {"_metadata": {}, "reports": []}

        if os.path.exists(RELATORIOS_FILE):
            with open(RELATORIOS_FILE, "r", encoding="utf-8") as f:
                local_data = json.load(f)

        last_sync = local_data.get("_metadata", {}).get("last_sync")

        if last_sync:
            last_sync = datetime.fromisoformat(last_sync)

        try:

            query = db.collection("relatorios").order_by("criadoEm")

            if last_sync:
                query = query.where("criadoEm", ">", last_sync)

            docs = list(query.stream())

            if not docs:
                ModernMessageBox.information(self, "Sincronização", "Nenhum relatório novo encontrado.")
                self.statusBar.showMessage("Pronto.")
                return

            newest_timestamp = last_sync
            novos = 0

            for doc in docs:

                report_data = converter_firestore_para_json(doc.to_dict())
                report_data["_id"] = doc.id

                report_data.pop("escopoEmailTexto", None)

                criado_em = doc.to_dict().get("criadoEm")

                if criado_em and isinstance(criado_em, datetime):

                    if not newest_timestamp or criado_em > newest_timestamp:
                        newest_timestamp = criado_em

                    report_data["criadoEm"] = criado_em.isoformat()

                local_data["reports"].append(report_data)

                novos += 1

            if newest_timestamp:
                local_data.setdefault("_metadata", {})
                local_data["_metadata"]["last_sync"] = newest_timestamp.isoformat()

            with open(RELATORIOS_FILE, "w", encoding="utf-8") as f:
                json.dump(local_data, f, indent=4, ensure_ascii=False)

            self.limites["relatorios_sync"] += 1
            self.salvar_limites()

            self.load_local_reports()

            ModernMessageBox.information(
                self,
                "Sucesso",
                f"{novos} novo(s) relatório(s) importado(s)!"
            )

            self.statusBar.showMessage("Pronto.")

        except Exception as e:

            ModernMessageBox.critical(
                self,
                "Erro",
                f"Falha ao sincronizar relatórios:\n{e}"
            )

            self.statusBar.showMessage("Erro de sincronização.")



    def show_report_details(self, row, column):
        item = self.reports_table.item(row, 0)
        if not item:
            return
        report_id = item.data(Qt.ItemDataRole.UserRole + 1)
        for report in self.relatorios_filtrados_cache:
            if report.get("_id", str(id(report))) == report_id:
                dialog = ReportDetailsDialog(report, self)
                dialog.exec()
                return

    def gerar_estatisticas(self):
        cache_key = (
            self.start_date_edit.date().toString(Qt.DateFormat.ISODate),
            self.end_date_edit.date().toString(Qt.DateFormat.ISODate),
            self.vendedor_combo.currentText(),
            len(self.local_reports_data)
        )
        if getattr(self, "_stats_cache_key", None) == cache_key:
            return
        self._stats_cache_key = cache_key

        start_date = self.start_date_edit.date().toPyDate()
        end_date = self.end_date_edit.date().toPyDate()
        vendedor_selecionado = self.vendedor_combo.currentText()

        relatorios_filtrados = []
        for report in self.local_reports_data:
            criado_em = safe_date_from_iso(report.get("criadoEm"))
            if not criado_em:
                continue
            if not (start_date <= criado_em.date() <= end_date):
                continue
            if vendedor_selecionado != "Todos" and report.get("orcamentistaEmail") != vendedor_selecionado:
                continue
            relatorios_filtrados.append(report)

        self.relatorios_filtrados_cache = relatorios_filtrados
        total_relatorios = len(relatorios_filtrados)

        soma_valores = 0
        valores = []
        vendedor_stats = {}
        faixa_counts = {
            "Até R$ 25 mil": 0,
            "R$ 25 mil a R$ 50 mil": 0,
            "R$ 50 mil a R$ 100 mil": 0,
            "Acima de R$ 100 mil": 0
        }
        respostas_respondidas = 0
        respostas_totais = 0

        todas_as_perguntas = set(self.perguntas_data.get("ordem", [])) if self.perguntas_data else set()

        for report in relatorios_filtrados:
            report_value = self._extract_report_value(report)
            soma_valores += report_value
            valores.append(report_value)

            vendedor = report.get("orcamentistaEmail") or "N/A"
            if vendedor not in vendedor_stats:
                vendedor_stats[vendedor] = {"count": 0, "total_valor": 0.0}
            vendedor_stats[vendedor]["count"] += 1
            vendedor_stats[vendedor]["total_valor"] += report_value

            if report_value <= 25000:
                faixa_counts["Até R$ 25 mil"] += 1
            elif report_value <= 50000:
                faixa_counts["R$ 25 mil a R$ 50 mil"] += 1
            elif report_value <= 100000:
                faixa_counts["R$ 50 mil a R$ 100 mil"] += 1
            else:
                faixa_counts["Acima de R$ 100 mil"] += 1

            respostas_questionario = report.get("respostasQuestionario", {})
            respostas_iniciais = report.get("respostasIniciais", {})
            if respostas_questionario:
                todas_as_perguntas.update(respostas_questionario.keys())
            if respostas_iniciais:
                todas_as_perguntas.update(respostas_iniciais.keys())

            respostas_unificadas = {}
            respostas_unificadas.update(respostas_questionario)
            respostas_unificadas.update(respostas_iniciais)
            respostas_totais += len(respostas_unificadas)
            respostas_respondidas += sum(1 for v in respostas_unificadas.values() if str(v or "").strip())

        valor_medio = soma_valores / total_relatorios if total_relatorios else 0
        maior_valor = max(valores) if valores else 0
        menor_valor = min(valores) if valores else 0
        if valores:
            valores_ordenados = sorted(valores)
            meio = len(valores_ordenados) // 2
            if len(valores_ordenados) % 2 == 0:
                mediana_valor = (valores_ordenados[meio - 1] + valores_ordenados[meio]) / 2
            else:
                mediana_valor = valores_ordenados[meio]
        else:
            mediana_valor = 0
        cobertura_questionario = (respostas_respondidas / respostas_totais * 100) if respostas_totais else 0

        periodo = f"{start_date.strftime('%d/%m/%Y')} até {end_date.strftime('%d/%m/%Y')}"
        filtro_vendedor = vendedor_selecionado if vendedor_selecionado != "Todos" else "todos os vendedores"
        self.periodo_aplicado_label.setText(
            f"Análise de {periodo} • Filtro: {filtro_vendedor} • {total_relatorios} orçamento(s)"
        )

        self.kpi_total_orcamentos.set_data(str(total_relatorios), "Quantidade de propostas no período")
        self.kpi_ticket.set_data(format_currency(valor_medio), "Média dos orçamentos gerados")
        self.kpi_mediana.set_data(format_currency(mediana_valor), "Valor central para reduzir efeito de outliers")
        self.kpi_cobertura.set_data(f"{cobertura_questionario:.1f}%", "Percentual de respostas preenchidas")
        self.kpi_maior.set_data(format_currency(maior_valor), "Maior orçamento registrado")
        self.kpi_menor.set_data(format_currency(menor_valor), "Menor orçamento registrado")

        self.vendedor_ranking_tree.clear()
        for vendedor, data in sorted(vendedor_stats.items(), key=lambda item: item[1]["count"], reverse=True):
            qtd = data["count"]
            pct = (qtd / total_relatorios * 100) if total_relatorios else 0
            ticket_medio_vendedor = data["total_valor"] / qtd if qtd else 0
            self.vendedor_ranking_tree.addTopLevelItem(
                QtWidgets.QTreeWidgetItem([
                    vendedor,
                    str(qtd),
                    f"{pct:.1f}%",
                    format_currency(ticket_medio_vendedor)
                ])
            )

        self.faixa_valor_tree.clear()
        for faixa, qtd in faixa_counts.items():
            pct = (qtd / total_relatorios * 100) if total_relatorios else 0
            self.faixa_valor_tree.addTopLevelItem(
                QtWidgets.QTreeWidgetItem([faixa, str(qtd), f"{pct:.1f}%"])
            )

        self.perguntas_iniciais_combo.blockSignals(True)
        self.perguntas_iniciais_combo.clear()
        self.perguntas_iniciais_combo.addItem("Selecione uma pergunta...")
        self.perguntas_iniciais_combo.addItems(sorted(todas_as_perguntas))
        self.perguntas_iniciais_combo.blockSignals(False)

        self.atualizar_estatisticas_pergunta_selecionada()

    def atualizar_estatisticas_pergunta_selecionada(self):
        self.perguntas_iniciais_stats_tree.clear()
        pergunta_selecionada = self.perguntas_iniciais_combo.currentText()
        if not pergunta_selecionada or pergunta_selecionada == "Selecione uma pergunta...":
            return

        respostas_stats = {}
        for report in self.relatorios_filtrados_cache:
            resposta = report.get("respostasIniciais", {}).get(pergunta_selecionada)
            if resposta is None:
                resposta = report.get("respostasQuestionario", {}).get(pergunta_selecionada)
            if not resposta:
                resposta = "Não respondido"

            resposta = str(resposta)
            if resposta not in respostas_stats:
                respostas_stats[resposta] = {"count": 0, "total_valor": 0.0}
            respostas_stats[resposta]["count"] += 1
            respostas_stats[resposta]["total_valor"] += self._extract_report_value(report)

        total = sum(item["count"] for item in respostas_stats.values()) or 1
        ordenado = sorted(respostas_stats.items(), key=lambda item: item[1]["count"], reverse=True)
        for resposta, dados in ordenado:
            qtd = dados["count"]
            pct = (qtd / total) * 100
            media_preco = dados["total_valor"] / qtd if qtd else 0
            self.perguntas_iniciais_stats_tree.addTopLevelItem(
                QtWidgets.QTreeWidgetItem([resposta, str(qtd), f"{pct:.1f}%", format_currency(media_preco)])
            )



if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 10))
    window = FirebaseManager()
    window.show()
    sys.exit(app.exec())
