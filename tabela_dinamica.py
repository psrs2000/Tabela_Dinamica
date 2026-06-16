import sys
import os
import sqlite3
import datetime
import re

try:
    import pandas as pd
    PANDAS_OK = True
except ImportError:
    PANDAS_OK = False

try:
    import openpyxl
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget,
    QVBoxLayout, QHBoxLayout, QGridLayout, QFormLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QCheckBox,
    QRadioButton, QButtonGroup, QGroupBox, QFileDialog,
    QMessageBox, QTableWidget, QTableWidgetItem,
    QTreeWidget, QTreeWidgetItem, QHeaderView, QSplitter,
    QAbstractItemView, QStatusBar, QFrame, QCompleter,
    QMenu, QWidgetAction,
)
from PyQt5.QtCore import Qt, QSize, QSortFilterProxyModel, QStringListModel
from PyQt5.QtGui import QColor, QBrush, QFont, QPalette

def _app_dir() -> str:
    """Retorna a pasta do .exe (quando compilado) ou do .py (em desenvolvimento)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

DB_PATH  = os.path.join(_app_dir(), "dados.db")
CFG_PATH = os.path.join(_app_dir(), "config.json")


def cfg_load() -> dict:
    try:
        import json
        with open(CFG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def cfg_save(data: dict):
    import json
    try:
        existing = cfg_load()
        existing.update(data)
        with open(CFG_PATH, "w", encoding="utf-8") as f:
            json.dump(existing, f, indent=2)
    except Exception:
        pass

NOMES_MESES = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio",    6: "Junho",     7: "Julho", 8: "Agosto",
    9: "Setembro",10: "Outubro",  11: "Novembro", 12: "Dezembro",
}

COLOR_NEG   = QColor("#c62828")   # vermelho escuro
COLOR_POS   = QColor("#1b5e20")   # verde escuro
COLOR_ZERO  = QColor("#424242")   # cinza
BG_GRUPO    = QColor("#dce8f5")   # azul claro
BG_TOTAL    = QColor("#c8e6c9")   # verde claro
BG_ALT      = QColor("#f9f9f9")   # alternado


# ═══════════════════════════════════════════════════════════
#  BANCO DE DADOS
# ═══════════════════════════════════════════════════════════

def init_db():
    con = sqlite3.connect(DB_PATH)
    con.execute("""
        CREATE TABLE IF NOT EXISTS registros (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            Data          TEXT,
            Mes           INTEGER,
            Ano           INTEGER,
            Categoria     TEXT,
            Sub_Categoria TEXT,
            Transacao     TEXT,
            Descricao     TEXT,
            Valor         REAL
        )
    """)
    con.commit()
    # garante que o banco nunca fica vazio (evita crash no pandas/Qt)
    cur = con.execute("SELECT COUNT(*) FROM registros")
    if cur.fetchone()[0] == 0:
        con.execute(
            "INSERT INTO registros (Data,Mes,Ano,Categoria,Sub_Categoria,Transacao,Descricao,Valor)"
            " VALUES (?,?,?,?,?,?,?,?)",
            ("01/01/1900", 1, 1900, "", "", "", "", 0.0)
        )
        con.commit()
    con.close()


def _mes_ano(data_str: str):
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y",
                "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            d = datetime.datetime.strptime(data_str.strip(), fmt)
            return d.month, d.year
        except ValueError:
            pass
    return None, None


def _parse_valor(raw) -> float:
    s = re.sub(r"[R$\s]", "", str(raw).strip())
    if "," in s and "." in s:
        if s.rindex(",") > s.rindex("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def inserir(row: dict):
    mes, ano = _mes_ano(row.get("Data", ""))
    con = sqlite3.connect(DB_PATH)
    con.execute(
        "INSERT INTO registros (Data,Mes,Ano,Categoria,Sub_Categoria,Transacao,Descricao,Valor)"
        " VALUES (?,?,?,?,?,?,?,?)",
        (row["Data"], mes, ano, row["Categoria"], row["Sub_Categoria"],
         row["Transacao"], row["Descricao"], float(row["Valor"] or 0))
    )
    con.commit()
    con.close()


def atualizar(rid, row: dict):
    mes, ano = _mes_ano(row.get("Data", ""))
    con = sqlite3.connect(DB_PATH)
    con.execute(
        "UPDATE registros SET Data=?,Mes=?,Ano=?,Categoria=?,Sub_Categoria=?,"
        "Transacao=?,Descricao=?,Valor=? WHERE id=?",
        (row["Data"], mes, ano, row["Categoria"], row["Sub_Categoria"],
         row["Transacao"], row["Descricao"], float(row["Valor"] or 0), rid)
    )
    con.commit()
    con.close()


def _inserir_dummy(con):
    con.execute(
        "INSERT INTO registros (Data,Mes,Ano,Categoria,Sub_Categoria,Transacao,Descricao,Valor)"
        " VALUES (?,?,?,?,?,?,?,?)",
        ("01/01/1900", 1, 1900, "", "", "", "", 0.0)
    )


def deletar(rid):
    con = sqlite3.connect(DB_PATH)
    con.execute("DELETE FROM registros WHERE id=?", (rid,))
    con.commit()
    cur = con.execute("SELECT COUNT(*) FROM registros")
    if cur.fetchone()[0] == 0:
        _inserir_dummy(con)
        con.commit()
    con.close()


def buscar_todos():
    con = sqlite3.connect(DB_PATH)
    cur = con.execute("SELECT * FROM registros ORDER BY Data")
    rows = cur.fetchall()
    cols = [d[0] for d in cur.description]
    con.close()
    return cols, rows


def apagar_banco():
    con = sqlite3.connect(DB_PATH)
    con.execute("DELETE FROM registros")
    _inserir_dummy(con)
    con.commit()
    con.close()


def importar_df(df, modo: str):
    col_map = {c.lower().strip(): c for c in df.columns}

    def get(*names):
        for name in names:
            key = name.lower().strip()
            if key in col_map:
                return col_map[key]
        prefix = names[0].lower().strip()
        for k, v in col_map.items():
            if k.startswith(prefix):
                return v
        return None

    if modo == "sobrescrever":
        apagar_banco()

    col_data  = get("data e hora", "data")
    col_cat   = get("categoria")
    col_sub   = get("sub_categoria", "sub-categoria", "subcategoria")
    col_tran  = get("transação", "transacao", "transação")
    col_desc  = get("descrição", "descricao", "descrição")
    col_valor = get("valor")

    con = sqlite3.connect(DB_PATH)
    for _, r in df.iterrows():
        data_val = str(r[col_data]).strip() if col_data else ""
        mes, ano = _mes_ano(data_val)
        valor    = _parse_valor(r[col_valor]) if col_valor else 0.0
        con.execute(
            "INSERT INTO registros (Data,Mes,Ano,Categoria,Sub_Categoria,Transacao,Descricao,Valor)"
            " VALUES (?,?,?,?,?,?,?,?)",
            (data_val, mes, ano,
             str(r[col_cat])  if col_cat  else "",
             str(r[col_sub])  if col_sub  else "",
             str(r[col_tran]) if col_tran else "",
             str(r[col_desc]) if col_desc else "",
             valor)
        )
    con.commit()
    con.close()


# ═══════════════════════════════════════════════════════════
#  HELPERS VISUAIS
# ═══════════════════════════════════════════════════════════

def fmt_valor(v) -> str:
    if not isinstance(v, (int, float)):
        return str(v)
    s = f"R$ {abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"-{s}" if v < 0 else s


def cor_valor(v) -> QColor:
    try:
        f = float(v) if not isinstance(v, float) else v
        if f < 0:   return COLOR_NEG
        if f > 0:   return COLOR_POS
        return COLOR_ZERO
    except (ValueError, TypeError):
        return COLOR_ZERO


class NumericItem(QTableWidgetItem):
    """Item que ordena pelo valor numérico armazenado em UserRole."""
    def __lt__(self, other):
        a = self.data(Qt.UserRole)
        b = other.data(Qt.UserRole)
        try:
            return float(a) < float(b)
        except (TypeError, ValueError):
            return super().__lt__(other)


def item_valor(v) -> QTableWidgetItem:
    it = NumericItem(fmt_valor(v))
    it.setData(Qt.UserRole, float(v) if v is not None else 0.0)
    it.setForeground(QBrush(cor_valor(v)))
    it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
    return it


def buscar_distintos(campo: str) -> list:
    con = sqlite3.connect(DB_PATH)
    cur = con.execute(
        f"SELECT DISTINCT {campo} FROM registros WHERE {campo} IS NOT NULL AND {campo}!='' ORDER BY {campo}")
    vals = [r[0] for r in cur.fetchall()]
    con.close()
    return vals


def _btn(texto, cor_bg, slot, min_w=100):
    """Cria um QPushButton padronizado."""
    b = QPushButton(texto)
    b.setStyleSheet(
        f"QPushButton{{background:{cor_bg};color:white;border-radius:5px;"
        f"padding:6px 18px;font-weight:bold;font-size:12px;}}"
        f"QPushButton:hover{{background:{cor_bg};border:1px solid rgba(0,0,0,0.2);}}"
    )
    b.setMinimumWidth(min_w)
    b.clicked.connect(slot)
    return b


# ═══════════════════════════════════════════════════════════
#  ABA DADOS
# ═══════════════════════════════════════════════════════════

COLS_DADOS = ["id", "Data", "Mês", "Ano", "Categoria", "Sub-Categoria",
              "Transação", "Descrição", "Valor"]

class AbaForm(QWidget):
    def __init__(self):
        super().__init__()
        self._edit_id  = None
        self._all_rows = []   # cache completo para filtro local
        self._dirty    = set()  # campos modificados pelo usuário desde última seleção
        self._build()

    # ── construção ────────────────────────────────────────
    def _build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(10, 8, 10, 6)
        root.setSpacing(6)

        # ── formulário ────────────────────────────────────
        grp = QGroupBox("Registro")
        form = QGridLayout(grp)
        form.setSpacing(5)
        form.setContentsMargins(8, 6, 8, 6)

        self._campos = {}
        specs = [
            ("Data",          "Data (dd/mm/aaaa hh:mm:ss)", False),
            ("Categoria",     "Categoria",                   True),
            ("Sub_Categoria", "Sub-Categoria",               True),
            ("Transacao",     "Transação",                   True),
            ("Descricao",     "Descrição",                   False),
            ("Valor",         "Valor",                       False),
        ]
        for row_idx, (key, lbl_txt, is_combo) in enumerate(specs):
            form.addWidget(QLabel(lbl_txt + ":"), row_idx, 0, Qt.AlignRight)
            if is_combo:
                w = QComboBox()
                w.setEditable(True)
                w.setInsertPolicy(QComboBox.NoInsert)
                w.setMinimumWidth(260)
                comp = QCompleter([], w)
                comp.setCaseSensitivity(Qt.CaseInsensitive)
                comp.setFilterMode(Qt.MatchContains)
                w.setCompleter(comp)
            else:
                w = QLineEdit()
                if key == "Descricao":
                    w.setMinimumWidth(450)
            form.addWidget(w, row_idx, 1, Qt.AlignLeft)
            self._campos[key] = w

        # botões do formulário
        btn_row = QHBoxLayout()
        self._btn_salvar  = _btn("Salvar",             "#4CAF50", self._salvar,  140)
        self._btn_limpar  = _btn("Limpar",             "#2196F3", self._limpar,  140)
        self._btn_excluir = _btn("Excluir Selecionado","#f44336", self._excluir, 180)
        self._btn_lote    = _btn("Aplicar a Selecionados", "#E65100", self._aplicar_lote, 200)
        for b in (self._btn_salvar, self._btn_limpar, self._btn_excluir, self._btn_lote):
            b.setStyleSheet(b.styleSheet().replace("padding:6px 18px", "padding:10px 28px")
                                          .replace("font-size:12px", "font-size:14px"))
        self._btn_lote.setVisible(False)
        btn_row.addWidget(self._btn_salvar)
        btn_row.addWidget(self._btn_limpar)
        btn_row.addWidget(self._btn_excluir)
        btn_row.addWidget(self._btn_lote)
        btn_row.addStretch()
        form.addLayout(btn_row, len(specs), 0, 1, 2)
        root.addWidget(grp)

        # conectar dirty tracking em todos os campos
        for key, w in self._campos.items():
            if isinstance(w, QComboBox):
                w.currentTextChanged.connect(lambda _, k=key: self._marcar_dirty(k))
            else:
                w.textChanged.connect(lambda _, k=key: self._marcar_dirty(k))

        # ── barra de filtros ──────────────────────────────
        flt_grp = QGroupBox("Filtros / Pesquisa")
        flt_lay = QHBoxLayout(flt_grp)
        flt_lay.setSpacing(6)
        flt_lay.setContentsMargins(8, 4, 8, 4)

        self._flt_widgets = {}
        flt_specs = [
            ("cat",  "Categoria",     120),
            ("sub",  "Sub-Categoria", 120),
            ("tran", "Transação",     150),
            ("desc", "Descrição",     160),
        ]
        for key, placeholder, w_px in flt_specs:
            ed = QLineEdit()
            ed.setPlaceholderText(placeholder)
            ed.setFixedWidth(w_px)
            ed.textChanged.connect(self._aplicar_filtro)
            flt_lay.addWidget(ed)
            self._flt_widgets[key] = ed

        flt_lay.addWidget(QLabel("Ano:"))
        self._flt_ano = QComboBox(); self._flt_ano.setFixedWidth(90)
        self._flt_ano.currentTextChanged.connect(self._aplicar_filtro)
        flt_lay.addWidget(self._flt_ano)

        flt_lay.addWidget(QLabel("Mês:"))
        self._flt_mes = QComboBox(); self._flt_mes.setFixedWidth(130)
        self._flt_mes.currentTextChanged.connect(self._aplicar_filtro)
        flt_lay.addWidget(self._flt_mes)

        btn_limpar_flt = _btn("✕ Limpar", "#757575", self._limpar_filtros, 80)
        flt_lay.addWidget(btn_limpar_flt)
        flt_lay.addStretch()
        root.addWidget(flt_grp)

        # ── tabela ────────────────────────────────────────
        self._table = QTableWidget(0, len(COLS_DADOS))
        self._table.setHorizontalHeaderLabels(COLS_DADOS)
        self._table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._table.setAlternatingRowColors(True)
        self._table.verticalHeader().setVisible(False)
        # manter destaque da seleção mesmo quando a tabela não tem foco
        self._table.setStyleSheet("""
            QTableWidget::item:selected {
                background-color: #1976D2;
                color: white;
            }
            QTableWidget::item:selected:!active {
                background-color: #1976D2;
                color: white;
            }
        """)
        hdr = self._table.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.Interactive)
        hdr.setStretchLastSection(False)
        hdr.setSortIndicatorShown(True)
        self._table.setSortingEnabled(True)
        self._table.itemSelectionChanged.connect(self._on_select)
        # salva larguras automaticamente quando o usuário arrastar uma coluna
        hdr.sectionResized.connect(self._salvar_layout)
        root.addWidget(self._table, 1)

        # ── rodapé: status ────────────────────────────────
        self._status = QLabel("")
        self._status.setStyleSheet("color:#555; font-size:10px")
        root.addWidget(self._status)

        # ── exportação da aba Dados ───────────────────────
        exp_row = QHBoxLayout()
        exp_row.addStretch()
        exp_row.addWidget(_btn("Exportar XLSX", "#1565C0", self._exportar_xlsx, 130))
        exp_row.addWidget(_btn("Exportar CSV",  "#6A1B9A", self._exportar_csv,  120))
        root.addLayout(exp_row)

        self._carregar()

    # ── campos helpers ────────────────────────────────────
    def _get_text(self, key) -> str:
        w = self._campos[key]
        if isinstance(w, QComboBox):
            return w.currentText().strip()
        return w.text().strip()

    def _set_text(self, key, val):
        w = self._campos[key]
        if isinstance(w, QComboBox):
            idx = w.findText(val)
            if idx >= 0:
                w.setCurrentIndex(idx)
            else:
                w.setCurrentText(val)
        else:
            w.setText(val)

    def _clear_field(self, key):
        w = self._campos[key]
        if isinstance(w, QComboBox):
            w.setCurrentIndex(-1)
            w.clearEditText()
        else:
            w.clear()

    def _atualizar_combos(self):
        """Recarrega listas de Categoria, Sub_Categoria e Transacao a partir do banco."""
        for key, campo in [("Categoria", "Categoria"), ("Sub_Categoria", "Sub_Categoria"),
                           ("Transacao", "Transacao")]:
            vals = buscar_distintos(campo)
            w = self._campos[key]
            cur = w.currentText()
            w.blockSignals(True)
            w.clear()
            w.addItems(vals)
            w.setCurrentText(cur)
            w.blockSignals(False)
            if w.completer():
                w.completer().setModel(QStringListModel(vals, w.completer()))

    def _atualizar_filtros_combo(self):
        anos  = ["(todos)"] + sorted(set(
            str(r[3]) for r in self._all_rows if r[3] is not None))
        meses = ["(todos)"] + [f"{i} – {NOMES_MESES[i]}" for i in range(1, 13)]
        for cb, vals in ((self._flt_ano, anos), (self._flt_mes, meses)):
            cur = cb.currentText()
            cb.blockSignals(True)
            cb.clear(); cb.addItems(vals)
            idx = cb.findText(cur)
            cb.setCurrentIndex(idx if idx >= 0 else 0)
            cb.blockSignals(False)

    # ── ações formulário ──────────────────────────────────
    def _salvar(self):
        row = {k: self._get_text(k) for k in self._campos}
        if not row["Data"]:
            QMessageBox.warning(self, "Atenção", "Data é obrigatória.")
            return
        mes, ano = _mes_ano(row["Data"])
        if mes is None:
            QMessageBox.warning(self, "Atenção",
                "Formato de data inválido.\nUse dd/mm/aaaa ou dd/mm/aaaa hh:mm:ss")
            return
        try:
            float(row["Valor"] or 0)
        except ValueError:
            QMessageBox.warning(self, "Atenção", "Valor deve ser numérico.")
            return
        if self._edit_id:
            atualizar(self._edit_id, row)
        else:
            inserir(row)
        self._limpar()
        self._carregar()

    def _limpar(self):
        self._carregando_selecao = True
        for key in self._campos:
            self._clear_field(key)
        self._carregando_selecao = False
        self._edit_id = None
        self._dirty.clear()
        self._btn_salvar.setText("Salvar")
        self._btn_salvar.setVisible(True)
        self._btn_lote.setVisible(False)

    def _excluir(self):
        sel = self._table.selectionModel().selectedRows()
        if not sel:
            QMessageBox.information(self, "Info", "Selecione um ou mais registros para excluir.")
            return
        count = len(sel)
        msg = (f"Excluir {count} registros selecionados?"
               if count > 1 else "Excluir registro selecionado?")
        if QMessageBox.question(self, "Confirmar", msg,
                                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            rids = [int(self._table.item(idx.row(), 0).text()) for idx in sel]
            for rid in rids:
                deletar(rid)
            self._limpar()
            self._carregar()

    def _marcar_dirty(self, key):
        # ignora mudanças causadas programaticamente durante _on_select
        if getattr(self, "_carregando_selecao", False):
            return
        self._dirty.add(key)

    def _on_select(self):
        sel = self._table.selectionModel().selectedRows()
        if not sel:
            self._btn_lote.setVisible(False)
            return
        multi = len(sel) > 1
        self._btn_lote.setVisible(multi)
        self._btn_salvar.setVisible(not multi)

        # carrega dados da primeira linha selecionada no formulário
        self._carregando_selecao = True
        self._dirty.clear()
        r = sel[0].row()
        self._edit_id = int(self._table.item(r, 0).text())
        self._set_text("Data",          self._table.item(r, 1).text())
        self._set_text("Categoria",     self._table.item(r, 4).text())
        self._set_text("Sub_Categoria", self._table.item(r, 5).text())
        self._set_text("Transacao",     self._table.item(r, 6).text())
        self._set_text("Descricao",     self._table.item(r, 7).text())
        raw = self._table.item(r, 8).data(Qt.UserRole)
        self._set_text("Valor", str(raw) if raw is not None else "")
        self._carregando_selecao = False
        self._btn_salvar.setText("Atualizar" if not multi else "Salvar")

    def _aplicar_lote(self):
        sel = self._table.selectionModel().selectedRows()
        if not sel or not self._dirty:
            QMessageBox.information(self, "Info",
                "Altere pelo menos um campo antes de aplicar.")
            return
        rids   = [int(self._table.item(idx.row(), 0).text()) for idx in sel]
        campos = sorted(self._dirty)
        nomes  = {"Data": "Data", "Categoria": "Categoria",
                  "Sub_Categoria": "Sub-Categoria", "Transacao": "Transação",
                  "Descricao": "Descrição", "Valor": "Valor"}
        lista  = ", ".join(nomes.get(c, c) for c in campos)
        msg    = (f"Atualizar o(s) campo(s)  【{lista}】\n"
                  f"em {len(rids)} registros selecionados?")
        if QMessageBox.question(self, "Confirmar edição em lote", msg,
                                QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return
        # busca registro completo e substitui apenas os campos dirty
        con = sqlite3.connect(DB_PATH)
        for rid in rids:
            cur = con.execute(
                "SELECT Data,Categoria,Sub_Categoria,Transacao,Descricao,Valor"
                " FROM registros WHERE id=?", (rid,))
            row_db = cur.fetchone()
            if not row_db:
                continue
            row = dict(zip(
                ["Data","Categoria","Sub_Categoria","Transacao","Descricao","Valor"],
                row_db))
            for campo in campos:
                row[campo] = self._get_text(campo)
            mes, ano = _mes_ano(row["Data"])
            try:
                valor = float(row["Valor"] or 0)
            except ValueError:
                valor = 0.0
            con.execute(
                "UPDATE registros SET Data=?,Mes=?,Ano=?,Categoria=?,Sub_Categoria=?,"
                "Transacao=?,Descricao=?,Valor=? WHERE id=?",
                (row["Data"], mes, ano, row["Categoria"], row["Sub_Categoria"],
                 row["Transacao"], row["Descricao"], valor, rid))
        con.commit()
        con.close()
        self._dirty.clear()
        self._carregar()

    # ── filtros ───────────────────────────────────────────
    def _limpar_filtros(self):
        for ed in self._flt_widgets.values():
            ed.blockSignals(True); ed.clear(); ed.blockSignals(False)
        self._flt_ano.blockSignals(True)
        self._flt_ano.setCurrentIndex(0)
        self._flt_ano.blockSignals(False)
        self._flt_mes.blockSignals(True)
        self._flt_mes.setCurrentIndex(0)
        self._flt_mes.blockSignals(False)
        self._aplicar_filtro()

    def _aplicar_filtro(self):
        f_cat  = self._flt_widgets["cat"].text().lower()
        f_sub  = self._flt_widgets["sub"].text().lower()
        f_tran = self._flt_widgets["tran"].text().lower()
        f_desc = self._flt_widgets["desc"].text().lower()
        f_ano  = self._flt_ano.currentText()
        f_mes  = self._flt_mes.currentText()
        num_mes = int(f_mes.split(" – ")[0]) if f_mes != "(todos)" and " – " in f_mes else None
        num_ano = int(f_ano) if f_ano not in ("", "(todos)") else None

        self._table.setSortingEnabled(False)
        self._table.setRowCount(0)
        soma = 0.0
        vis = 0
        for r in self._all_rows:
            # r = (id, Data, Mes, Ano, Categoria, Sub_Categoria, Transacao, Descricao, Valor)
            if f_cat  and f_cat  not in str(r[4]).lower(): continue
            if f_sub  and f_sub  not in str(r[5]).lower(): continue
            if f_tran and f_tran not in str(r[6]).lower(): continue
            if f_desc and f_desc not in str(r[7]).lower(): continue
            if num_ano is not None and r[3] != num_ano:    continue
            if num_mes is not None and r[2] != num_mes:    continue
            i = self._table.rowCount()
            self._table.insertRow(i)
            for j, v in enumerate(r):
                if j == 8:                          # Valor
                    it = item_valor(v)
                elif j in (0, 2, 3):               # id, Mês, Ano
                    it = NumericItem(str(v) if v is not None else "")
                    it.setData(Qt.UserRole, float(v) if v is not None else 0.0)
                    it.setTextAlignment(Qt.AlignCenter)
                elif j == 1:                        # Data — ordena por ISO
                    it = NumericItem(str(v) if v is not None else "")
                    mes, ano = _mes_ano(str(v)) if v else (0, 0)
                    # converte para AAAAMMDD para ordenação correta
                    try:
                        import datetime as _dt
                        for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M",
                                    "%d/%m/%Y", "%Y-%m-%d %H:%M:%S",
                                    "%Y-%m-%d %H:%M", "%Y-%m-%d"):
                            try:
                                sort_key = float(
                                    _dt.datetime.strptime(str(v).strip(), fmt)
                                    .strftime("%Y%m%d%H%M%S"))
                                break
                            except ValueError:
                                sort_key = 0.0
                    except Exception:
                        sort_key = 0.0
                    it.setData(Qt.UserRole, sort_key)
                    it.setTextAlignment(Qt.AlignCenter)
                else:
                    it = QTableWidgetItem(str(v) if v is not None else "")
                    it.setTextAlignment(Qt.AlignCenter)
                self._table.setItem(i, j, it)
            soma += float(r[8] or 0)
            vis += 1
        self._table.setSortingEnabled(True)
        # atualiza cabeçalho da coluna Valor com a soma ANTES do resize
        cor_hex = "#c62828" if soma < 0 else "#1b5e20"
        hdr_item = self._table.horizontalHeaderItem(8)
        hdr_item.setText(f"Valor  |  {fmt_valor(soma)}")
        hdr_item.setForeground(QBrush(QColor(cor_hex)))
        hdr = self._table.horizontalHeader()
        # bloqueia sinal para que o auto-ajuste NÃO sobrescreva o config salvo
        hdr.sectionResized.disconnect(self._salvar_layout)
        saved = cfg_load().get("dados_col_widths")
        if saved and len(saved) == self._table.columnCount():
            # layout salvo pelo usuário: restaura direto, ignora auto-ajuste
            for c, w in enumerate(saved):
                self._table.setColumnWidth(c, w)
        else:
            # sem layout salvo: aplica auto-ajuste por conteúdo + cabeçalho
            self._table.resizeColumnsToContents()
            fm = self._table.fontMetrics()
            for col in range(self._table.columnCount()):
                hdr_txt = self._table.horizontalHeaderItem(col)
                hdr_w = fm.horizontalAdvance(hdr_txt.text() if hdr_txt else "") + 24
                if hdr_w > hdr.sectionSize(col):
                    hdr.resizeSection(col, hdr_w)
        hdr.sectionResized.connect(self._salvar_layout)
        total = len(self._all_rows)
        self._status.setText(
            f"Exibindo {vis} de {total} registros" +
            (f"  |  filtro ativo" if vis < total else ""))

    # ── carga ─────────────────────────────────────────────
    def _carregar(self):
        _, rows = buscar_todos()
        self._all_rows = rows
        self._atualizar_combos()
        self._atualizar_filtros_combo()
        self._aplicar_filtro()

    # ── layout de colunas ─────────────────────────────────
    def _salvar_layout(self):
        """Chamado automaticamente ao arrastar qualquer coluna."""
        widths = [self._table.columnWidth(c)
                  for c in range(self._table.columnCount())]
        cfg_save({"dados_col_widths": widths})


    # ── exportação ────────────────────────────────────────
    def _linhas_visiveis(self):
        """Retorna linhas visíveis; coluna Valor como número puro (sem R$)."""
        rows = []
        for i in range(self._table.rowCount()):
            row = []
            for j in range(self._table.columnCount()):
                it = self._table.item(i, j)
                if j == 8 and it:                          # coluna Valor
                    raw = it.data(Qt.UserRole)
                    row.append(raw if raw is not None else "")
                else:
                    row.append(it.text() if it else "")
            rows.append(row)
        return rows

    def _exportar_xlsx(self):
        if not OPENPYXL_OK:
            QMessageBox.critical(self, "Erro",
                "openpyxl não instalado.\nExecute: pip install openpyxl")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Exportar como XLSX", "", "Excel (*.xlsx)")
        if not path:
            return
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Dados"
        # cabeçalho (sem a soma no texto)
        hdrs = [COLS_DADOS[i] if i != 8 else "Valor"
                for i in range(len(COLS_DADOS))]
        for c, h in enumerate(hdrs, 1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4472C4")
            cell.alignment = Alignment(horizontal="center")
        for r, row in enumerate(self._linhas_visiveis(), 2):
            for c, val in enumerate(row, 1):
                cell = ws.cell(row=r, column=c, value=val)
                cell.alignment = Alignment(
                    horizontal="right" if c == len(hdrs) else "center")
        for col in ws.columns:
            w = max((len(str(c.value or "")) for c in col), default=8)
            ws.column_dimensions[col[0].column_letter].width = min(w + 2, 40)
        wb.save(path)
        QMessageBox.information(self, "Sucesso",
            f"{self._table.rowCount()} registros exportados para:\n{path}")

    def _exportar_csv(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Exportar como CSV", "", "CSV (*.csv)")
        if not path:
            return
        import csv
        hdrs = [COLS_DADOS[i] if i != 8 else "Valor"
                for i in range(len(COLS_DADOS))]
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(hdrs)
            writer.writerows(self._linhas_visiveis())
        QMessageBox.information(self, "Sucesso",
            f"{self._table.rowCount()} registros exportados para:\n{path}")

    def refresh(self):
        self._carregar()


# ═══════════════════════════════════════════════════════════
#  ABA IMPORTAÇÃO
# ═══════════════════════════════════════════════════════════

class AbaImport(QWidget):
    def __init__(self, aba_form: AbaForm):
        super().__init__()
        self._aba_form = aba_form
        self._build()

    def _build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 16, 12, 8)
        root.setSpacing(10)

        if not PANDAS_OK:
            root.addWidget(QLabel(
                "pandas não instalado.\nExecute: pip install pandas openpyxl xlrd"))
            return

        # ── título centralizado ───────────────────────────
        title = QLabel("Importar Arquivo")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size:16px; font-weight:bold; margin-bottom:6px;")
        root.addWidget(title)

        # ── container centralizado com formulário ─────────
        center_container = QWidget()
        center_layout = QHBoxLayout(center_container)
        center_layout.setContentsMargins(0, 0, 0, 0)

        form_widget = QWidget()
        form_widget.setMaximumWidth(500)
        form_layout = QFormLayout(form_widget)
        form_layout.setSpacing(10)
        form_layout.setContentsMargins(0, 0, 0, 0)
        form_layout.setLabelAlignment(Qt.AlignRight)

        # Arquivo
        arq_lbl = QLabel("Arquivo:")
        arq_row = QHBoxLayout()
        self._path = QLineEdit()
        self._path.setReadOnly(True)
        self._path.setMinimumWidth(260)
        btn_browse = QPushButton("...")
        btn_browse.setFixedWidth(36)
        btn_browse.clicked.connect(self._browse)
        arq_row.addWidget(self._path)
        arq_row.addWidget(btn_browse)
        form_layout.addRow(arq_lbl, arq_row)

        # Separador CSV
        sep_lbl = QLabel("Separador CSV:")
        self._sep = QLineEdit(";")
        self._sep.setFixedWidth(40)
        form_layout.addRow(sep_lbl, self._sep)

        center_layout.addStretch()
        center_layout.addWidget(form_widget)
        center_layout.addStretch()
        root.addWidget(center_container)

        # ── modo de importação (label + radio buttons centrados) ──
        modo_lbl = QLabel("Modo de importação:")
        modo_lbl.setAlignment(Qt.AlignCenter)
        modo_lbl.setStyleSheet("font-weight:bold; font-size:15px; margin-top:4px;")
        root.addWidget(modo_lbl)

        radio_container = QWidget()
        radio_layout = QHBoxLayout(radio_container)
        radio_layout.setContentsMargins(0, 0, 0, 0)
        self._rb_add = QRadioButton("Acrescentar ao banco existente")
        self._rb_ow  = QRadioButton("Sobrescrever banco (apaga tudo antes)")
        self._rb_ow.setStyleSheet("color:#c62828; font-size:15px; font-weight:bold")
        self._rb_add.setChecked(True)
        radio_layout.addStretch()
        radio_layout.addWidget(self._rb_add)
        radio_layout.addSpacing(20)
        radio_layout.addWidget(self._rb_ow)
        radio_layout.addStretch()
        root.addWidget(radio_container)

        # ── botão Importar centralizado ───────────────────
        btn_container = QWidget()
        btn_layout = QHBoxLayout(btn_container)
        btn_layout.setContentsMargins(0, 4, 0, 4)
        btn_imp = QPushButton("Importar")
        btn_imp.setStyleSheet(
            "QPushButton{background:#4CAF50;color:white;border-radius:6px;"
            "padding:10px 28px;font-weight:bold;font-size:14px;}"
            "QPushButton:hover{background:#43A047;border:1px solid rgba(0,0,0,0.2);}"
        )
        btn_imp.setFixedWidth(180)
        btn_imp.clicked.connect(self._importar)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_imp)
        btn_layout.addStretch()
        root.addWidget(btn_container)

        # ── status centralizado ───────────────────────────
        self._status = QLabel("")
        self._status.setAlignment(Qt.AlignCenter)
        root.addWidget(self._status)

        # ── preview (preenche espaço restante) ────────────
        grp_prev = QGroupBox("Pré-visualização (20 primeiras linhas)")
        lay_prev = QVBoxLayout(grp_prev)
        self._preview = QTableWidget(0, 0)
        self._preview.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._preview.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self._preview.verticalHeader().setVisible(False)
        lay_prev.addWidget(self._preview)
        root.addWidget(grp_prev, 1)

    def _browse(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Selecionar arquivo", "",
            "Planilhas (*.csv *.xls *.xlsx);;Todos (*.*)")
        if path:
            self._path.setText(path)
            self._load_preview(path)

    def _ler(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext == ".csv":
            sep = self._sep.text() or ";"
            try:
                return pd.read_csv(path, sep=sep, encoding="utf-8-sig", dtype=str)
            except UnicodeDecodeError:
                return pd.read_csv(path, sep=sep, encoding="latin1", dtype=str)
        elif ext == ".xlsx":
            return pd.read_excel(path, dtype=str, engine="openpyxl")
        elif ext == ".xls":
            return pd.read_excel(path, dtype=str, engine="xlrd")
        raise ValueError(f"Formato não suportado: {ext}")

    def _load_preview(self, path):
        try:
            df = self._ler(path)
            self._preview.clear()
            self._preview.setColumnCount(len(df.columns))
            self._preview.setHorizontalHeaderLabels(list(df.columns))
            self._preview.setRowCount(min(20, len(df)))
            for i, (_, row) in enumerate(df.head(20).iterrows()):
                for j, v in enumerate(row):
                    it = QTableWidgetItem(str(v))
                    self._preview.setItem(i, j, it)
            self._status.setText(f"{len(df)} linhas encontradas.")
            self._status.setStyleSheet("color:#1b5e20")
        except Exception as e:
            self._status.setText(f"Erro: {e}")
            self._status.setStyleSheet("color:#c62828")

    def _importar(self):
        path = self._path.text().strip()
        if not path:
            QMessageBox.warning(self, "Atenção", "Selecione um arquivo primeiro.")
            return
        modo = "sobrescrever" if self._rb_ow.isChecked() else "acrescentar"
        if modo == "sobrescrever":
            if QMessageBox.question(self, "Confirmar",
                    "Isso apagará TODOS os registros existentes. Continuar?",
                    QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
                return
        try:
            df = self._ler(path)
            importar_df(df, modo)
            self._status.setText(f"{len(df)} registros importados com sucesso.")
            self._status.setStyleSheet("color:#1b5e20")
            self._aba_form.refresh()
        except Exception as e:
            self._status.setText(f"Erro: {e}")
            self._status.setStyleSheet("color:#c62828")


# ═══════════════════════════════════════════════════════════
#  ABA TABELA DINÂMICA
# ═══════════════════════════════════════════════════════════

ALL_DIM = ["Ano", "Mes", "Categoria", "Sub_Categoria", "Transacao", "Descricao"]

class AbaPivot(QWidget):
    def __init__(self):
        super().__init__()
        self._export_rows = []
        self._build()

    def _build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 8, 12, 8)
        root.setSpacing(6)

        # ── estrutura ─────────────────────────────────────
        grp_str = QGroupBox("Estrutura")
        lay_str = QGridLayout(grp_str)
        lay_str.setSpacing(6)

        def lbl(t): return QLabel(t)
        def combo(vals, default):
            cb = QComboBox(); cb.addItems(vals)
            cb.setCurrentText(default); return cb

        lay_str.addWidget(lbl("Linha 1 (grupo):"),    0, 0, Qt.AlignRight)
        self._row1 = combo(ALL_DIM, "Categoria")
        lay_str.addWidget(self._row1, 0, 1)
        self._btn_excl1 = QPushButton("Excluir itens ▼")
        self._btn_excl1.setFixedWidth(130)
        self._btn_excl1.clicked.connect(lambda: self._abrir_menu_exclusao(1))
        lay_str.addWidget(self._btn_excl1, 0, 2)

        lay_str.addWidget(lbl("Linha 2 (subgrupo):"), 0, 3, Qt.AlignRight)
        self._row2 = combo(["(nenhuma)"] + ALL_DIM, "Descricao")
        lay_str.addWidget(self._row2, 0, 4)
        self._btn_excl2 = QPushButton("Excluir itens ▼")
        self._btn_excl2.setFixedWidth(130)
        self._btn_excl2.clicked.connect(lambda: self._abrir_menu_exclusao(2))
        lay_str.addWidget(self._btn_excl2, 0, 5)

        lay_str.addWidget(lbl("Colunas:"),            0, 6, Qt.AlignRight)
        self._cols = combo(["(nenhuma)"] + ALL_DIM, "Mes")
        lay_str.addWidget(self._cols, 0, 7)

        lay_str.addWidget(lbl("Agregar:"),            1, 0, Qt.AlignRight)
        self._agg = combo(["sum", "count", "mean", "min", "max"], "sum")
        self._agg.setFixedWidth(90)
        lay_str.addWidget(self._agg, 1, 1)

        self._chk_sub = QCheckBox("Subtotais"); self._chk_sub.setChecked(True)
        lay_str.addWidget(self._chk_sub, 1, 3, 1, 2)
        self._chk_total = QCheckBox("Total Geral"); self._chk_total.setChecked(True)
        lay_str.addWidget(self._chk_total, 1, 5, 1, 2)
        self._chk_pct = QCheckBox("Mostrar como %"); self._chk_pct.setChecked(False)
        lay_str.addWidget(self._chk_pct, 1, 7)
        root.addWidget(grp_str)

        # sets de exclusão e estado de expansão (populados dinamicamente)
        self._excluidos1 = set()
        self._excluidos2 = set()
        self._expandidos = set()  # nomes dos grupos expandidos

        # ── filtros ───────────────────────────────────────
        grp_flt = QGroupBox("Filtros de Relatório")
        lay_flt = QGridLayout(grp_flt)
        lay_flt.setSpacing(6)

        lay_flt.addWidget(lbl("Ano:"),           0, 0, Qt.AlignRight)
        self._f_ano  = QComboBox(); self._f_ano.setFixedWidth(80)
        lay_flt.addWidget(self._f_ano, 0, 1)

        lay_flt.addWidget(lbl("Mês:"),           0, 2, Qt.AlignRight)
        self._f_mes  = QComboBox(); self._f_mes.setMinimumWidth(150)
        lay_flt.addWidget(self._f_mes, 0, 3)

        lay_flt.addWidget(lbl("Categoria:"),     0, 4, Qt.AlignRight)
        self._f_cat  = QComboBox(); self._f_cat.setMinimumWidth(160)
        lay_flt.addWidget(self._f_cat, 0, 5)

        lay_flt.addWidget(lbl("Transação:"),     1, 0, Qt.AlignRight)
        self._f_tran = QComboBox(); self._f_tran.setMinimumWidth(200)
        lay_flt.addWidget(self._f_tran, 1, 1, 1, 2)

        lay_flt.addWidget(lbl("Sub-Categoria:"), 1, 4, Qt.AlignRight)
        self._f_sub  = QComboBox(); self._f_sub.setMinimumWidth(160)
        lay_flt.addWidget(self._f_sub, 1, 5)

        # radio buttons: todos / somente positivos / somente negativos
        lay_flt.addWidget(lbl("Valores:"), 2, 0, Qt.AlignRight)
        self._rb_todos = QRadioButton("Todos");          self._rb_todos.setChecked(True)
        self._rb_pos   = QRadioButton("Somente positivos")
        self._rb_neg   = QRadioButton("Somente negativos")
        self._rb_grp   = QButtonGroup(self)
        for rb in (self._rb_todos, self._rb_pos, self._rb_neg):
            self._rb_grp.addButton(rb)
        rb_row = QHBoxLayout()
        rb_row.setSpacing(16)
        for rb in (self._rb_todos, self._rb_pos, self._rb_neg):
            rb_row.addWidget(rb)
        rb_row.addStretch()
        lay_flt.addLayout(rb_row, 2, 1, 1, 5)
        root.addWidget(grp_flt)

        # ── botões expandir/recolher ──────────────────────
        btn_tree_row = QHBoxLayout()
        btn_exp = QPushButton("▼  Expandir Tudo")
        btn_rec = QPushButton("▶  Recolher Tudo")
        for b in (btn_exp, btn_rec):
            b.setFixedHeight(28)
            b.setFixedWidth(150)
        btn_exp.clicked.connect(lambda: self._tree.expandAll())
        btn_rec.clicked.connect(lambda: self._tree.collapseAll())
        btn_tree_row.addWidget(btn_exp)
        btn_tree_row.addWidget(btn_rec)
        btn_tree_row.addStretch()
        root.addLayout(btn_tree_row)

        # ── resultado ─────────────────────────────────────
        self._tree = QTreeWidget()
        self._tree.setAlternatingRowColors(False)
        self._tree.setUniformRowHeights(True)
        self._tree.setRootIsDecorated(True)
        self._tree.setSortingEnabled(False)
        self._tree.header().setSectionResizeMode(QHeaderView.Interactive)
        self._tree.setFont(QFont("Segoe UI", 11))
        root.addWidget(self._tree, 1)

        self._status = QLabel("")
        self._status.setStyleSheet("color:#555")
        root.addWidget(self._status)

        # ── conectar sinais para auto-gerar ───────────────
        for cb in (self._row1, self._row2, self._cols, self._agg,
                   self._f_ano, self._f_mes, self._f_cat, self._f_tran, self._f_sub):
            cb.currentIndexChanged.connect(self._gerar)
        for chk in (self._chk_sub, self._chk_total, self._chk_pct):
            chk.stateChanged.connect(self._gerar)
        for rb in (self._rb_todos, self._rb_pos, self._rb_neg):
            rb.toggled.connect(self._gerar)

        self._tree.itemExpanded.connect(
            lambda item: self._on_expansao(item, True))
        self._tree.itemCollapsed.connect(
            lambda item: self._on_expansao(item, False))

        self._atualizar_filtros()
        self._restaurar_config_pivot()

    # ── dados ─────────────────────────────────────────────
    def _carregar_df(self):
        if not PANDAS_OK:
            return None
        cols, rows = buscar_todos()
        if not rows:
            return pd.DataFrame(columns=["id","Data","Mes","Ano","Categoria",
                                         "Sub_Categoria","Transacao","Descricao","Valor"])
        df = pd.DataFrame(rows, columns=cols)
        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
        df["Mes"]   = pd.to_numeric(df["Mes"],   errors="coerce")
        df["Ano"]   = pd.to_numeric(df["Ano"],   errors="coerce")
        return df

    def _atualizar_filtros(self):
        df = self._carregar_df()
        todos = ["(todos)"]
        if df is None or df.empty:
            for cb in (self._f_ano, self._f_cat, self._f_tran, self._f_sub, self._f_mes):
                cb.blockSignals(True)
                cb.clear()
                cb.addItems(todos)
                cb.blockSignals(False)
            self._tree.clear()
            return
        anos  = todos + sorted(df["Ano"].dropna().unique().astype(int).astype(str).tolist())
        cats  = todos + sorted(df["Categoria"].dropna().unique().tolist())
        trans = todos + sorted(df["Transacao"].dropna().unique().tolist())
        subs  = todos + sorted(df["Sub_Categoria"].dropna().unique().tolist())
        meses = todos + [f"{i} – {NOMES_MESES[i]}" for i in range(1, 13)]
        for cb, vals in ((self._f_ano, anos), (self._f_cat, cats),
                         (self._f_tran, trans), (self._f_sub, subs),
                         (self._f_mes, meses)):
            cur = cb.currentText()
            cb.blockSignals(True)
            cb.clear(); cb.addItems(vals)
            idx = cb.findText(cur)
            if idx >= 0: cb.setCurrentIndex(idx)
            cb.blockSignals(False)

    # ── restaurar configuração salva ──────────────────────
    def _restaurar_config_pivot(self):
        cfg = cfg_load().get("pivot_config")
        if not cfg:
            return
        # block signals during restore to avoid multiple _gerar calls
        widgets = [self._row1, self._row2, self._cols, self._agg,
                   self._f_ano, self._f_mes, self._f_cat, self._f_tran, self._f_sub,
                   self._chk_sub, self._chk_total, self._chk_pct,
                   self._rb_todos, self._rb_pos, self._rb_neg]
        for w in widgets:
            w.blockSignals(True)
        try:
            for attr, key in [("_row1", "row1"), ("_row2", "row2"),
                               ("_cols", "cols"), ("_agg", "agg"),
                               ("_f_ano", "f_ano"), ("_f_mes", "f_mes"),
                               ("_f_cat", "f_cat"), ("_f_tran", "f_tran"),
                               ("_f_sub", "f_sub")]:
                val = cfg.get(key)
                if val is not None:
                    cb = getattr(self, attr)
                    idx = cb.findText(val)
                    if idx >= 0:
                        cb.setCurrentIndex(idx)
            if "subtotais" in cfg:
                self._chk_sub.setChecked(bool(cfg["subtotais"]))
            if "total_geral" in cfg:
                self._chk_total.setChecked(bool(cfg["total_geral"]))
            if "mostrar_pct" in cfg:
                self._chk_pct.setChecked(bool(cfg["mostrar_pct"]))
            fv = cfg.get("filtro_valor", "todos")
            if fv == "pos":  self._rb_pos.setChecked(True)
            elif fv == "neg": self._rb_neg.setChecked(True)
            else:             self._rb_todos.setChecked(True)
            if "excluidos1" in cfg:
                self._excluidos1 = set(cfg["excluidos1"])
            if "excluidos2" in cfg:
                self._excluidos2 = set(cfg["excluidos2"])
            if "expandidos" in cfg:
                self._expandidos = set(cfg["expandidos"])
        finally:
            for w in widgets:
                w.blockSignals(False)

    # ── estado de expansão dos grupos ────────────────────
    def _on_expansao(self, item: QTreeWidgetItem, expandido: bool):
        nome = item.text(0)
        if expandido:
            self._expandidos.add(nome)
        else:
            self._expandidos.discard(nome)
        cfg_save({"pivot_config": {
            **cfg_load().get("pivot_config", {}),
            "expandidos": list(self._expandidos),
        }})

    # ── menu de exclusão de itens ─────────────────────────
    def _abrir_menu_exclusao(self, linha: int):
        df = self._carregar_df()
        if df is None or df.empty:
            return
        campo = self._row1.currentText() if linha == 1 else self._row2.currentText()
        if campo == "(nenhuma)":
            return
        excluidos = self._excluidos1 if linha == 1 else self._excluidos2
        btn      = self._btn_excl1  if linha == 1 else self._btn_excl2

        itens = sorted(df[campo].dropna().unique().tolist(), key=str)

        menu = QMenu(self)
        menu.setStyleSheet("QMenu { padding:4px; } QCheckBox { padding:4px 8px; }")

        for item in itens:
            wa = QWidgetAction(menu)
            chk = QCheckBox(str(item), menu)
            chk.setChecked(str(item) in excluidos)
            item_str = str(item)
            def make_toggle(s, ex):
                def toggle(checked):
                    if checked: ex.add(s)
                    else:        ex.discard(s)
                    self._gerar()
                return toggle
            chk.toggled.connect(make_toggle(item_str, excluidos))
            wa.setDefaultWidget(chk)
            menu.addAction(wa)

        menu.addSeparator()
        act_clear = menu.addAction("Limpar exclusões")
        act_clear.triggered.connect(lambda: self._limpar_exclusoes(linha))

        menu.exec_(btn.mapToGlobal(btn.rect().bottomLeft()))

    def _limpar_exclusoes(self, linha: int):
        if linha == 1: self._excluidos1.clear()
        else:          self._excluidos2.clear()
        self._gerar()

    # ── gerar ─────────────────────────────────────────────
    def _gerar(self):
        self._atualizar_filtros()
        df = self._carregar_df()
        if df is None:
            QMessageBox.warning(self, "Erro", "pandas não instalado.")
            return
        if df.empty:
            return

        # filtros
        if self._f_ano.currentText()  != "(todos)":
            df = df[df["Ano"] == int(self._f_ano.currentText())]
        mes_sel = self._f_mes.currentText()
        if mes_sel != "(todos)":
            df = df[df["Mes"] == int(mes_sel.split(" – ")[0])]
        if self._f_cat.currentText()  != "(todos)":
            df = df[df["Categoria"]     == self._f_cat.currentText()]
        if self._f_tran.currentText() != "(todos)":
            df = df[df["Transacao"]     == self._f_tran.currentText()]
        if self._f_sub.currentText()  != "(todos)":
            df = df[df["Sub_Categoria"] == self._f_sub.currentText()]

        # exclusões de itens de linha 1 e linha 2
        row1_campo = self._row1.currentText()
        row2_campo = self._row2.currentText()
        if self._excluidos1 and row1_campo in df.columns:
            df = df[~df[row1_campo].astype(str).isin(self._excluidos1)]
        if self._excluidos2 and row2_campo in df.columns:
            df = df[~df[row2_campo].astype(str).isin(self._excluidos2)]

        # filtro positivos / negativos
        if self._rb_pos.isChecked():
            df = df[df["Valor"] > 0]
        elif self._rb_neg.isChecked():
            df = df[df["Valor"] < 0]

        if df.empty:
            self._tree.clear()
            self._status.setText("Nenhum dado para os filtros selecionados.")
            return

        row1     = self._row1.currentText()
        row2     = self._row2.currentText()
        col_fld  = self._cols.currentText()
        agg      = self._agg.currentText()
        use_row2 = row2 != "(nenhuma)"
        use_cols = col_fld != "(nenhuma)"

        col_vals = (
            sorted(df[col_fld].dropna().unique().tolist(),
                   key=lambda x: (int(x) if str(x).lstrip("-").isdigit() else 0, str(x)))
            if use_cols else ["Valor"]
        )

        def agregar(sub):
            if sub.empty: return 0.0
            if agg == "sum":   return float(sub["Valor"].sum())
            if agg == "count": return float(sub["Valor"].count())
            v = getattr(sub["Valor"], agg)()
            return 0.0 if pd.isna(v) else float(v)

        def vals_por_col(sub):
            return ({str(cv): agregar(sub[sub[col_fld] == cv]) for cv in col_vals}
                    if use_cols else {"Valor": agregar(sub)})

        def soma_d(d):
            return sum(v for v in d.values() if isinstance(v, (int, float)) and not pd.isna(v))

        # ── configurar QTreeWidget ────────────────────────
        hdrs = [f"{row1}" + (f" / {row2}" if use_row2 else "")] + \
               [str(cv) for cv in col_vals] + ["Total Geral"]
        self._tree.clear()
        self._tree.setColumnCount(len(hdrs))
        self._tree.setHeaderLabels(hdrs)
        self._tree.header().setSectionResizeMode(0, QHeaderView.Interactive)
        self._tree.setColumnWidth(0, 210)
        for c in range(1, len(hdrs)):
            self._tree.setColumnWidth(c, 105)
            self._tree.header().setSectionResizeMode(c, QHeaderView.Interactive)

        font_bold = QFont("Segoe UI", 11, QFont.Bold)
        font_reg  = QFont("Segoe UI", 11)

        usar_pct = self._chk_pct.isChecked()

        # ── 1ª passagem: calcular todos os valores e o total geral ──
        grand = {str(cv): 0.0 for cv in col_vals}
        grupos = sorted(df[row1].dropna().unique().tolist())
        dados_grupos = []
        for g in grupos:
            g_df  = df[df[row1] == g]
            g_cv  = vals_por_col(g_df)
            g_tot = soma_d(g_cv)
            for k in grand: grand[k] += g_cv.get(k, 0.0)
            subgrupos_dados = []
            if use_row2:
                for sg in sorted(g_df[row2].dropna().unique().tolist()):
                    sg_df = g_df[g_df[row2] == sg]
                    sg_cv = vals_por_col(sg_df)
                    sg_tot = soma_d(sg_cv)
                    subgrupos_dados.append((sg, sg_cv, sg_tot))
            dados_grupos.append((g, g_cv, g_tot, subgrupos_dados))

        gt_tot = soma_d(grand)

        def fmt_cel(v, is_total_geral_cell=False):
            """Formata célula em valor ou %."""
            if not usar_pct:
                return fmt_valor(v)
            if is_total_geral_cell:
                return "100,00%"
            base = abs(gt_tot) if gt_tot != 0 else 1.0
            return f"{v / base * 100:.2f}%".replace(".", ",")

        def cor_cel(v):
            return cor_valor(v)

        # ── 2ª passagem: montar a árvore ─────────────────────
        export_rows = [hdrs]

        for g, g_cv, g_tot, subgrupos_dados in dados_grupos:
            g_strs = [fmt_cel(g_cv.get(str(cv), 0)) for cv in col_vals] + \
                     [fmt_cel(g_tot, usar_pct and g_tot == gt_tot)]
            parent = QTreeWidgetItem([str(g)] + g_strs)
            parent.setFont(0, font_bold)
            parent.setBackground(0, QBrush(BG_GRUPO))
            for c in range(1, len(hdrs)):
                parent.setFont(c, font_bold)
                parent.setBackground(c, QBrush(BG_GRUPO))
                v = g_cv.get(str(col_vals[c-1]), g_tot) if c < len(hdrs)-1 else g_tot
                parent.setForeground(c, QBrush(cor_cel(v)))
                parent.setTextAlignment(c, Qt.AlignRight | Qt.AlignVCenter)
            export_rows.append([str(g)] + g_strs)

            for sg, sg_cv, sg_tot in subgrupos_dados:
                sg_strs = [fmt_cel(sg_cv.get(str(cv), 0)) for cv in col_vals] + \
                          [fmt_cel(sg_tot)]
                child = QTreeWidgetItem([str(sg)] + sg_strs)
                child.setFont(0, font_reg)
                for c in range(1, len(hdrs)):
                    child.setFont(c, font_reg)
                    v = sg_cv.get(str(col_vals[c-1]), sg_tot) if c < len(hdrs)-1 else sg_tot
                    child.setForeground(c, QBrush(cor_cel(v)))
                    child.setTextAlignment(c, Qt.AlignRight | Qt.AlignVCenter)
                parent.addChild(child)
                export_rows.append(["  " + str(sg)] + sg_strs)

            self._tree.addTopLevelItem(parent)
            parent.setExpanded(str(g) in self._expandidos)

        # Total Geral
        if self._chk_total.isChecked():
            gt_strs = [fmt_cel(grand.get(str(cv), 0)) for cv in col_vals] + \
                      [fmt_cel(gt_tot, usar_pct)]
            total_item = QTreeWidgetItem(["Total Geral"] + gt_strs)
            total_item.setFont(0, font_bold)
            total_item.setBackground(0, QBrush(BG_TOTAL))
            for c in range(1, len(hdrs)):
                total_item.setFont(c, font_bold)
                total_item.setBackground(c, QBrush(BG_TOTAL))
                v = grand.get(str(col_vals[c-1]), gt_tot) if c < len(hdrs)-1 else gt_tot
                total_item.setForeground(c, QBrush(cor_cel(v)))
                total_item.setTextAlignment(c, Qt.AlignRight | Qt.AlignVCenter)
            self._tree.addTopLevelItem(total_item)
            export_rows.append(["Total Geral"] + gt_strs)

        self._export_rows = export_rows
        # auto-ajuste da coluna de rótulo; demais via ResizeToContents
        self._tree.header().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        for c in range(1, len(hdrs)):
            self._tree.header().setSectionResizeMode(c, QHeaderView.ResizeToContents)
        self._status.setText(
            f"{len(grupos)} grupos  |  {len(df)} registros  |  agreg: {agg}"
            + ("  —  clique no ▶ para expandir" if use_row2 else ""))

        # ── salvar configuração em config.json ────────────
        cfg_save({"pivot_config": {
            "row1":        self._row1.currentText(),
            "row2":        self._row2.currentText(),
            "cols":        self._cols.currentText(),
            "agg":         self._agg.currentText(),
            "subtotais":   self._chk_sub.isChecked(),
            "total_geral": self._chk_total.isChecked(),
            "mostrar_pct": self._chk_pct.isChecked(),
            "filtro_valor": "pos" if self._rb_pos.isChecked() else "neg" if self._rb_neg.isChecked() else "todos",
            "f_ano":       self._f_ano.currentText(),
            "f_mes":       self._f_mes.currentText(),
            "f_cat":       self._f_cat.currentText(),
            "f_tran":      self._f_tran.currentText(),
            "f_sub":       self._f_sub.currentText(),
            "excluidos1":  list(self._excluidos1),
            "excluidos2":  list(self._excluidos2),
            "expandidos":  list(self._expandidos),
        }})

    # ── exportar ──────────────────────────────────────────
    def _exportar(self):
        if not self._export_rows:
            QMessageBox.information(self, "Info", "Gere a tabela antes de exportar.")
            return
        if not OPENPYXL_OK:
            QMessageBox.critical(self, "Erro",
                "openpyxl não instalado.\nExecute: pip install openpyxl")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Salvar como", "", "Excel (*.xlsx)")
        if not path:
            return
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Tabela Dinâmica"
        for r_idx, row in enumerate(self._export_rows, 1):
            for c_idx, val in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.alignment = Alignment(horizontal="right" if c_idx > 1 else "left")
                if r_idx == 1:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="4472C4")
                elif "Total" in str(row[0]):
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill("solid", fgColor="E2EFDA")
                elif c_idx > 1:
                    try:
                        v = float(str(val).replace("R$","").replace(".","")
                                         .replace(",",".").strip())
                        cell.font = Font(color="C62828" if v < 0 else "1B5E20")
                    except (ValueError, AttributeError):
                        pass
        for col in ws.columns:
            w = max((len(str(c.value or "")) for c in col), default=8)
            ws.column_dimensions[col[0].column_letter].width = min(w + 2, 40)
        wb.save(path)
        QMessageBox.information(self, "Sucesso", f"Exportado para:\n{path}")


# ═══════════════════════════════════════════════════════════
#  JANELA PRINCIPAL
# ═══════════════════════════════════════════════════════════

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Tabela Dinâmica")
        self.setMinimumSize(800, 560)

        # restaurar geometria salva
        cfg = cfg_load().get("janela", {})
        if cfg.get("maximizado"):
            self.resize(cfg.get("largura", 1100), cfg.get("altura", 720))
            self.showMaximized()
        else:
            self.resize(cfg.get("largura", 1100), cfg.get("altura", 720))
            if "x" in cfg and "y" in cfg:
                self.move(cfg["x"], cfg["y"])

        tabs = QTabWidget()
        tabs.setDocumentMode(True)
        tabs.setFont(QFont("Segoe UI", 10))

        self._aba_form   = AbaForm()
        self._aba_import = AbaImport(self._aba_form)
        self._aba_pivot  = AbaPivot()

        tabs.addTab(self._aba_form,   "  Dados  ")
        tabs.addTab(self._aba_import, "  Importar  ")
        tabs.addTab(self._aba_pivot,  "  Tabela Dinâmica  ")
        tabs.currentChanged.connect(self._on_tab)

        self.setCentralWidget(tabs)
        self.statusBar().showMessage("Pronto")

    def closeEvent(self, event):
        maximizado = self.isMaximized()
        if maximizado:
            self.showNormal()  # para capturar tamanho restaurado
        cfg_save({"janela": {
            "maximizado": maximizado,
            "largura":    self.width(),
            "altura":     self.height(),
            "x":          self.x(),
            "y":          self.y(),
        }})
        event.accept()

    def _on_tab(self, idx):
        if idx == 2:
            self._aba_pivot._gerar()


# ═══════════════════════════════════════════════════════════

if __name__ == "__main__":
    init_db()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())
