import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import os
import datetime

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

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dados.db")

CAMPOS = ["Data", "Categoria", "Sub_Categoria", "Transacao", "Descricao", "Valor"]


# ─────────────────────────── DB ────────────────────────────

def init_db():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS registros (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            Data        TEXT,
            Mes         INTEGER,
            Ano         INTEGER,
            Categoria   TEXT,
            Sub_Categoria TEXT,
            Transacao   TEXT,
            Descricao   TEXT,
            Valor       REAL
        )
    """)
    con.commit()
    con.close()


def _mes_ano(data_str):
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y",
                "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            d = datetime.datetime.strptime(data_str.strip(), fmt)
            return d.month, d.year
        except ValueError:
            pass
    return None, None


def inserir(row: dict):
    mes, ano = _mes_ano(row.get("Data", ""))
    con = sqlite3.connect(DB_PATH)
    con.execute("""
        INSERT INTO registros (Data, Mes, Ano, Categoria, Sub_Categoria, Transacao, Descricao, Valor)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (row["Data"], mes, ano,
          row["Categoria"], row["Sub_Categoria"],
          row["Transacao"], row["Descricao"],
          float(row["Valor"] or 0)))
    con.commit()
    con.close()


def atualizar(rid, row: dict):
    mes, ano = _mes_ano(row.get("Data", ""))
    con = sqlite3.connect(DB_PATH)
    con.execute("""
        UPDATE registros SET Data=?, Mes=?, Ano=?, Categoria=?, Sub_Categoria=?,
            Transacao=?, Descricao=?, Valor=? WHERE id=?
    """, (row["Data"], mes, ano,
          row["Categoria"], row["Sub_Categoria"],
          row["Transacao"], row["Descricao"],
          float(row["Valor"] or 0), rid))
    con.commit()
    con.close()


def deletar(rid):
    con = sqlite3.connect(DB_PATH)
    con.execute("DELETE FROM registros WHERE id=?", (rid,))
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
    con.commit()
    con.close()


def _parse_valor(raw) -> float:
    """Converte strings como '-R$ 55,48' ou '1.234,56' para float."""
    import re
    s = str(raw).strip()
    s = re.sub(r"[R$\s]", "", s)   # remove R$, espaços
    # detecta formato BR (último separador é vírgula): 1.234,56
    if "," in s and "." in s:
        if s.rindex(",") > s.rindex("."):   # BR: 1.234,56
            s = s.replace(".", "").replace(",", ".")
        else:                               # EN: 1,234.56
            s = s.replace(",", "")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def importar_df(df: "pd.DataFrame", modo: str):
    col_map = {c.lower().strip(): c for c in df.columns}

    def get(*names):
        """Retorna o nome original da coluna para qualquer alias."""
        for name in names:
            key = name.lower().strip()
            if key in col_map:
                return col_map[key]
        # busca parcial: coluna que começa com o primeiro nome
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
        con.execute("""
            INSERT INTO registros (Data, Mes, Ano, Categoria, Sub_Categoria, Transacao, Descricao, Valor)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            data_val, mes, ano,
            str(r[col_cat])  if col_cat  else "",
            str(r[col_sub])  if col_sub  else "",
            str(r[col_tran]) if col_tran else "",
            str(r[col_desc]) if col_desc else "",
            valor,
        ))
    con.commit()
    con.close()


# ─────────────────────────── FORMULÁRIO ────────────────────

class FormularioFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg="#f5f5f5")
        self.app = app
        self._edit_id = None
        self._build()

    def _build(self):
        tk.Label(self, text="Entrada de Dados", font=("Segoe UI", 14, "bold"),
                 bg="#f5f5f5").grid(row=0, column=0, columnspan=2, pady=(12, 6))

        labels = ["Data (dd/mm/aaaa hh:mm:ss)", "Categoria", "Sub-Categoria",
                  "Transação", "Descrição", "Valor"]
        self._vars = []
        for i, lbl in enumerate(labels):
            tk.Label(self, text=lbl, bg="#f5f5f5",
                     font=("Segoe UI", 10)).grid(row=i+1, column=0, sticky="e", padx=8, pady=4)
            v = tk.StringVar()
            entry = tk.Entry(self, textvariable=v, width=36, font=("Segoe UI", 10))
            entry.grid(row=i+1, column=1, padx=8, pady=4, sticky="w")
            self._vars.append(v)

        btn_frame = tk.Frame(self, bg="#f5f5f5")
        btn_frame.grid(row=8, column=0, columnspan=2, pady=10)
        tk.Button(btn_frame, text="Salvar", width=12, command=self._salvar,
                  bg="#4CAF50", fg="white", font=("Segoe UI", 10, "bold")).pack(side="left", padx=4)
        tk.Button(btn_frame, text="Limpar", width=12, command=self._limpar,
                  bg="#2196F3", fg="white", font=("Segoe UI", 10)).pack(side="left", padx=4)
        tk.Button(btn_frame, text="Excluir Selecionado", width=16, command=self._excluir,
                  bg="#f44336", fg="white", font=("Segoe UI", 10)).pack(side="left", padx=4)

        # tabela de registros
        frame_tree = tk.Frame(self)
        frame_tree.grid(row=9, column=0, columnspan=2, sticky="nsew", padx=8, pady=4)
        self.rowconfigure(9, weight=1)
        self.columnconfigure(1, weight=1)

        cols = ("id", "Data", "Mes", "Ano", "Categoria", "Sub_Categoria",
                "Transacao", "Descricao", "Valor")
        self.tree = ttk.Treeview(frame_tree, columns=cols, show="headings", height=14)
        widths = [40, 140, 40, 50, 100, 110, 110, 150, 80]
        for col, w in zip(cols, widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor="center")
        vsb = ttk.Scrollbar(frame_tree, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(frame_tree, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame_tree.rowconfigure(0, weight=1)
        frame_tree.columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self._carregar_tree()

    def _campos(self):
        return {
            "Data": self._vars[0].get(),
            "Categoria": self._vars[1].get(),
            "Sub_Categoria": self._vars[2].get(),
            "Transacao": self._vars[3].get(),
            "Descricao": self._vars[4].get(),
            "Valor": self._vars[5].get(),
        }

    def _salvar(self):
        row = self._campos()
        if not row["Data"]:
            messagebox.showwarning("Atenção", "Data é obrigatória.")
            return
        mes, ano = _mes_ano(row["Data"])
        if mes is None:
            messagebox.showwarning("Atenção", "Formato de data inválido.\nUse dd/mm/aaaa ou dd/mm/aaaa hh:mm:ss")
            return
        try:
            float(row["Valor"] or 0)
        except ValueError:
            messagebox.showwarning("Atenção", "Valor deve ser numérico.")
            return
        if self._edit_id:
            atualizar(self._edit_id, row)
        else:
            inserir(row)
        self._limpar()
        self._carregar_tree()

    def _limpar(self):
        for v in self._vars:
            v.set("")
        self._edit_id = None

    def _excluir(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Info", "Selecione um registro para excluir.")
            return
        if messagebox.askyesno("Confirmar", "Excluir registro selecionado?"):
            rid = self.tree.item(sel[0])["values"][0]
            deletar(rid)
            self._limpar()
            self._carregar_tree()

    def _on_select(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0])["values"]
        # id, Data, Mes, Ano, Categoria, Sub_Categoria, Transacao, Descricao, Valor
        self._edit_id = vals[0]
        self._vars[0].set(vals[1])
        self._vars[1].set(vals[4])
        self._vars[2].set(vals[5])
        self._vars[3].set(vals[6])
        self._vars[4].set(vals[7])
        self._vars[5].set(vals[8])

    def _carregar_tree(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        _, rows = buscar_todos()
        for r in rows:
            self.tree.insert("", "end", values=r)

    def refresh(self):
        self._carregar_tree()


# ─────────────────────────── IMPORTAÇÃO ────────────────────

class ImportacaoFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg="#f5f5f5")
        self.app = app
        self._build()

    def _build(self):
        tk.Label(self, text="Importar Arquivo", font=("Segoe UI", 14, "bold"),
                 bg="#f5f5f5").pack(pady=12)

        if not PANDAS_OK:
            tk.Label(self, text="pandas não instalado.\nExecute: pip install pandas openpyxl xlrd",
                     fg="red", bg="#f5f5f5", font=("Segoe UI", 11)).pack(pady=20)
            return

        frm = tk.Frame(self, bg="#f5f5f5")
        frm.pack(pady=6)
        tk.Label(frm, text="Arquivo:", bg="#f5f5f5",
                 font=("Segoe UI", 10)).grid(row=0, column=0, sticky="e", padx=6)
        self._path_var = tk.StringVar()
        tk.Entry(frm, textvariable=self._path_var, width=48,
                 font=("Segoe UI", 10)).grid(row=0, column=1, padx=4)
        tk.Button(frm, text="...", width=3, command=self._browse).grid(row=0, column=2, padx=4)

        tk.Label(frm, text="Separador CSV:", bg="#f5f5f5",
                 font=("Segoe UI", 10)).grid(row=1, column=0, sticky="e", padx=6, pady=6)
        self._sep_var = tk.StringVar(value=";")
        tk.Entry(frm, textvariable=self._sep_var, width=4,
                 font=("Segoe UI", 10)).grid(row=1, column=1, sticky="w", padx=4)

        self._modo = tk.StringVar(value="acrescentar")
        tk.Label(self, text="Modo de importação:", bg="#f5f5f5",
                 font=("Segoe UI", 10, "bold")).pack()
        rb_frame = tk.Frame(self, bg="#f5f5f5")
        rb_frame.pack()
        tk.Radiobutton(rb_frame, text="Acrescentar ao banco existente",
                       variable=self._modo, value="acrescentar",
                       bg="#f5f5f5", font=("Segoe UI", 10)).pack(anchor="w")
        tk.Radiobutton(rb_frame, text="Sobrescrever banco (apaga tudo antes)",
                       variable=self._modo, value="sobrescrever",
                       bg="#f5f5f5", font=("Segoe UI", 10),
                       fg="#c62828").pack(anchor="w")

        tk.Button(self, text="Importar", width=16, command=self._importar,
                  bg="#4CAF50", fg="white", font=("Segoe UI", 11, "bold")).pack(pady=14)

        self._status = tk.Label(self, text="", bg="#f5f5f5", font=("Segoe UI", 10))
        self._status.pack()

        # preview
        prev_frame = tk.Frame(self)
        prev_frame.pack(fill="both", expand=True, padx=8, pady=4)
        self._preview = ttk.Treeview(prev_frame, show="headings", height=10)
        vsb = ttk.Scrollbar(prev_frame, orient="vertical", command=self._preview.yview)
        hsb = ttk.Scrollbar(prev_frame, orient="horizontal", command=self._preview.xview)
        self._preview.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._preview.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        prev_frame.rowconfigure(0, weight=1)
        prev_frame.columnconfigure(0, weight=1)

    def _browse(self):
        path = filedialog.askopenfilename(
            filetypes=[("Planilhas", "*.csv *.xls *.xlsx"), ("Todos", "*.*")])
        if path:
            self._path_var.set(path)
            self._load_preview(path)

    def _load_preview(self, path):
        try:
            df = self._ler(path)
            self._preview["columns"] = list(df.columns)
            for col in df.columns:
                self._preview.heading(col, text=col)
                self._preview.column(col, width=100)
            for i in self._preview.get_children():
                self._preview.delete(i)
            for _, row in df.head(20).iterrows():
                self._preview.insert("", "end", values=list(row))
        except Exception as e:
            self._status.config(text=f"Erro ao pré-visualizar: {e}", fg="red")

    def _ler(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext == ".csv":
            sep = self._sep_var.get() or ";"
            try:
                return pd.read_csv(path, sep=sep, encoding="utf-8-sig", dtype=str)
            except UnicodeDecodeError:
                return pd.read_csv(path, sep=sep, encoding="latin1", dtype=str)
        elif ext == ".xlsx":
            return pd.read_excel(path, dtype=str, engine="openpyxl")
        elif ext == ".xls":
            return pd.read_excel(path, dtype=str, engine="xlrd")
        else:
            raise ValueError(f"Formato não suportado: {ext}")

    def _importar(self):
        path = self._path_var.get().strip()
        if not path:
            messagebox.showwarning("Atenção", "Selecione um arquivo primeiro.")
            return
        modo = self._modo.get()
        if modo == "sobrescrever":
            if not messagebox.askyesno("Confirmar", "Isso apagará TODOS os registros existentes. Continuar?"):
                return
        try:
            df = self._ler(path)
            importar_df(df, modo)
            self._status.config(text=f"{len(df)} registros importados com sucesso.", fg="#2e7d32")
            self.app.formulario.refresh()
        except Exception as e:
            self._status.config(text=f"Erro: {e}", fg="red")


# ─────────────────────────── TABELA DINÂMICA ───────────────

NOMES_MESES = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",
               7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}

class TabelaDinamicaFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg="#f5f5f5")
        self.app = app
        self._df = None
        self._export_data = None   # lista de listas para exportar
        self._export_cols = []
        self._build()

    # ── construção da UI ──────────────────────────────────────
    def _build(self):
        tk.Label(self, text="Tabela Dinâmica", font=("Segoe UI", 14, "bold"),
                 bg="#f5f5f5").grid(row=0, column=0, columnspan=6, pady=(10, 4))

        if not PANDAS_OK:
            tk.Label(self, text="pandas não instalado.", fg="red",
                     bg="#f5f5f5").grid(row=1, column=0)
            return

        def lbl(parent, text, row, col, **kw):
            tk.Label(parent, text=text, bg="#f5f5f5",
                     font=("Segoe UI", 10, "bold"), **kw).grid(
                row=row, column=col, padx=6, pady=3, sticky="w")

        all_dim = ["Ano", "Mes", "Categoria", "Sub_Categoria", "Transacao", "Descricao"]

        # ── bloco ESTRUTURA ───────────────────────────────────
        grp_struct = tk.LabelFrame(self, text="  Estrutura  ", bg="#f5f5f5",
                                   font=("Segoe UI", 9, "bold"))
        grp_struct.grid(row=1, column=0, columnspan=6, padx=10, pady=4, sticky="ew")

        lbl(grp_struct, "Linha 1 (grupo):", 0, 0)
        self._row1_var = tk.StringVar(value="Categoria")
        ttk.Combobox(grp_struct, textvariable=self._row1_var, values=all_dim,
                     width=16, state="readonly").grid(row=0, column=1, padx=4, pady=3)

        lbl(grp_struct, "Linha 2 (subgrupo):", 0, 2)
        self._row2_var = tk.StringVar(value="Descricao")
        ttk.Combobox(grp_struct, textvariable=self._row2_var,
                     values=["(nenhuma)"] + all_dim,
                     width=16, state="readonly").grid(row=0, column=3, padx=4, pady=3)

        lbl(grp_struct, "Colunas:", 0, 4)
        self._cols_var = tk.StringVar(value="Mes")
        ttk.Combobox(grp_struct, textvariable=self._cols_var,
                     values=["(nenhuma)"] + all_dim,
                     width=14, state="readonly").grid(row=0, column=5, padx=4, pady=3)

        lbl(grp_struct, "Agregar Valor:", 1, 0)
        self._agg_var = tk.StringVar(value="sum")
        ttk.Combobox(grp_struct, textvariable=self._agg_var,
                     values=["sum", "count", "mean", "min", "max"],
                     width=10, state="readonly").grid(row=1, column=1, padx=4, pady=3)

        lbl(grp_struct, "Subtotais:", 1, 2)
        self._subtotal_var = tk.BooleanVar(value=True)
        tk.Checkbutton(grp_struct, variable=self._subtotal_var,
                       bg="#f5f5f5").grid(row=1, column=3, sticky="w")

        lbl(grp_struct, "Total Geral:", 1, 4)
        self._total_var = tk.BooleanVar(value=True)
        tk.Checkbutton(grp_struct, variable=self._total_var,
                       bg="#f5f5f5").grid(row=1, column=5, sticky="w")

        # ── bloco FILTROS ─────────────────────────────────────
        grp_filt = tk.LabelFrame(self, text="  Filtros de Relatório  ", bg="#f5f5f5",
                                  font=("Segoe UI", 9, "bold"))
        grp_filt.grid(row=2, column=0, columnspan=6, padx=10, pady=4, sticky="ew")

        lbl(grp_filt, "Ano:", 0, 0)
        self._f_ano = tk.StringVar(value="(todos)")
        self._cb_ano = ttk.Combobox(grp_filt, textvariable=self._f_ano, width=8, state="readonly")
        self._cb_ano.grid(row=0, column=1, padx=4, pady=3)

        lbl(grp_filt, "Mês:", 0, 2)
        self._f_mes = tk.StringVar(value="(todos)")
        meses_vals = ["(todos)"] + [f"{i} – {NOMES_MESES[i]}" for i in range(1, 13)]
        self._cb_mes = ttk.Combobox(grp_filt, textvariable=self._f_mes,
                                     values=meses_vals, width=14, state="readonly")
        self._cb_mes.grid(row=0, column=3, padx=4, pady=3)

        lbl(grp_filt, "Categoria:", 0, 4)
        self._f_cat = tk.StringVar(value="(todos)")
        self._cb_cat = ttk.Combobox(grp_filt, textvariable=self._f_cat, width=18, state="readonly")
        self._cb_cat.grid(row=0, column=5, padx=4, pady=3)

        lbl(grp_filt, "Transação:", 1, 0)
        self._f_tran = tk.StringVar(value="(todos)")
        self._cb_tran = ttk.Combobox(grp_filt, textvariable=self._f_tran, width=20, state="readonly")
        self._cb_tran.grid(row=1, column=1, columnspan=2, padx=4, pady=3, sticky="w")

        lbl(grp_filt, "Sub-Categoria:", 1, 4)
        self._f_sub = tk.StringVar(value="(todos)")
        self._cb_sub = ttk.Combobox(grp_filt, textvariable=self._f_sub, width=18, state="readonly")
        self._cb_sub.grid(row=1, column=5, padx=4, pady=3)

        # ── botões ────────────────────────────────────────────
        btn_frame = tk.Frame(self, bg="#f5f5f5")
        btn_frame.grid(row=3, column=0, columnspan=6, pady=6)
        tk.Button(btn_frame, text="▶  Gerar Tabela", command=self._gerar,
                  bg="#4CAF50", fg="white", width=16,
                  font=("Segoe UI", 10, "bold")).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Exportar XLSX", command=self._exportar,
                  bg="#1565C0", fg="white", width=14,
                  font=("Segoe UI", 10)).pack(side="left", padx=6)
        tk.Button(btn_frame, text="↺  Atualizar Filtros", command=self._atualizar_filtros,
                  bg="#757575", fg="white", width=16,
                  font=("Segoe UI", 10)).pack(side="left", padx=6)

        # ── resultado ─────────────────────────────────────────
        res_frame = tk.Frame(self)
        res_frame.grid(row=4, column=0, columnspan=6, sticky="nsew", padx=8, pady=4)
        self.rowconfigure(4, weight=1)
        self.columnconfigure(5, weight=1)

        style = ttk.Style()
        style.configure("Pivot.Treeview", font=("Segoe UI", 9), rowheight=22)
        style.configure("Pivot.Treeview.Heading", font=("Segoe UI", 9, "bold"))

        self.result_tree = ttk.Treeview(res_frame, show="tree headings",
                                        style="Pivot.Treeview", selectmode="browse")
        self.result_tree.tag_configure("grupo",   font=("Segoe UI", 9, "bold"), background="#dce8f5")
        self.result_tree.tag_configure("total",   font=("Segoe UI", 9, "bold"), background="#c8e6c9")
        self.result_tree.tag_configure("subitem", font=("Segoe UI", 9))

        vsb = ttk.Scrollbar(res_frame, orient="vertical",   command=self.result_tree.yview)
        hsb = ttk.Scrollbar(res_frame, orient="horizontal", command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.result_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        res_frame.rowconfigure(0, weight=1)
        res_frame.columnconfigure(0, weight=1)

        self._status_lbl = tk.Label(self, text="", bg="#f5f5f5", font=("Segoe UI", 9))
        self._status_lbl.grid(row=5, column=0, columnspan=6)
        self._atualizar_filtros()

    # ── dados ─────────────────────────────────────────────────
    def _carregar_df(self):
        cols, rows = buscar_todos()
        if not rows:
            return pd.DataFrame()
        df = pd.DataFrame(rows, columns=cols)
        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
        df["Mes"]   = pd.to_numeric(df["Mes"],   errors="coerce")
        df["Ano"]   = pd.to_numeric(df["Ano"],   errors="coerce")
        return df

    def _atualizar_filtros(self):
        df = self._carregar_df()
        self._df = df
        vazio = ["(todos)"]
        if df.empty:
            for cb in (self._cb_ano, self._cb_cat, self._cb_tran, self._cb_sub):
                cb["values"] = vazio
            return
        anos  = vazio + sorted(df["Ano"].dropna().unique().astype(int).astype(str).tolist())
        cats  = vazio + sorted(df["Categoria"].dropna().unique().tolist())
        trans = vazio + sorted(df["Transacao"].dropna().unique().tolist())
        subs  = vazio + sorted(df["Sub_Categoria"].dropna().unique().tolist())
        self._cb_ano["values"]  = anos
        self._cb_cat["values"]  = cats
        self._cb_tran["values"] = trans
        self._cb_sub["values"]  = subs

    # ── formatar valor ────────────────────────────────────────
    @staticmethod
    def _fmt(v):
        if isinstance(v, (int, float)):
            return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return str(v) if v is not None else ""

    # ── gerar tabela ──────────────────────────────────────────
    def _gerar(self):
        self._atualizar_filtros()
        df = self._df.copy() if self._df is not None else self._carregar_df()
        if df.empty:
            messagebox.showinfo("Info", "Banco de dados vazio.")
            return

        # aplicar filtros
        if self._f_ano.get() != "(todos)":
            df = df[df["Ano"] == int(self._f_ano.get())]
        mes_sel = self._f_mes.get()
        if mes_sel != "(todos)":
            num_mes = int(mes_sel.split(" – ")[0])
            df = df[df["Mes"] == num_mes]
        if self._f_cat.get()  != "(todos)": df = df[df["Categoria"]    == self._f_cat.get()]
        if self._f_tran.get() != "(todos)": df = df[df["Transacao"]    == self._f_tran.get()]
        if self._f_sub.get()  != "(todos)": df = df[df["Sub_Categoria"]== self._f_sub.get()]

        if df.empty:
            messagebox.showinfo("Info", "Nenhum dado para os filtros selecionados.")
            return

        row1     = self._row1_var.get()
        row2     = self._row2_var.get()
        col_fld  = self._cols_var.get()
        agg      = self._agg_var.get()
        subtotal = self._subtotal_var.get()
        total_g  = self._total_var.get()
        use_row2 = (row2 != "(nenhuma)")
        use_cols = (col_fld != "(nenhuma)")

        # determinar colunas dinâmicas
        if use_cols:
            col_vals = sorted(df[col_fld].dropna().unique().tolist(),
                              key=lambda x: (int(x) if str(x).lstrip("-").isdigit() else 0, str(x)))
        else:
            col_vals = ["Valor"]

        # helper de agregação — retorna 0 em vez de NaN para subconjuntos vazios
        def agregar(sub):
            if sub.empty:
                return 0.0
            if agg == "sum":   return float(sub["Valor"].sum())
            if agg == "count": return float(sub["Valor"].count())
            if agg == "mean":  v = sub["Valor"].mean();  return 0.0 if pd.isna(v) else float(v)
            if agg == "min":   v = sub["Valor"].min();   return 0.0 if pd.isna(v) else float(v)
            if agg == "max":   v = sub["Valor"].max();   return 0.0 if pd.isna(v) else float(v)
            return 0.0

        def vals_por_col(sub):
            if use_cols:
                return {str(cv): agregar(sub[sub[col_fld] == cv]) for cv in col_vals}
            else:
                return {"Valor": agregar(sub)}

        def soma_dict(d):
            return sum(v for v in d.values() if isinstance(v, (int, float)) and not pd.isna(v))

        # ── configurar colunas do treeview ────────────────────
        # #0 = coluna de árvore (rótulo da linha)
        # colunas nomeadas = valores dinâmicos + Total Geral
        val_col_labels = [str(cv) for cv in col_vals] + ["Total Geral"]
        val_col_ids    = [f"v{i}" for i in range(len(val_col_labels))]

        self.result_tree["columns"] = val_col_ids
        for i in self.result_tree.get_children():
            self.result_tree.delete(i)

        # coluna da árvore (#0)
        tree_lbl = f"{row1}" + (f"  /  {row2}" if use_row2 else "")
        self.result_tree.heading("#0", text=tree_lbl, anchor="w")
        self.result_tree.column("#0", width=200, stretch=False, anchor="w")

        for vid, vlbl in zip(val_col_ids, val_col_labels):
            self.result_tree.heading(vid, text=str(vlbl), anchor="e")
            w = 105 if vid != val_col_ids[-1] else 115
            self.result_tree.column(vid, width=w, stretch=False, anchor="e")

        # ── preencher árvore ──────────────────────────────────
        export_rows = [[tree_lbl] + val_col_labels]
        grand_totals = {str(cv): 0.0 for cv in col_vals}

        grupos = sorted(df[row1].dropna().unique().tolist())
        for g in grupos:
            g_df       = df[df[row1] == g]
            g_col_vals = vals_por_col(g_df)
            g_total    = soma_dict(g_col_vals)

            # acumular grand total (somente sum/count fazem sentido; mean acumula separado)
            for k in grand_totals:
                grand_totals[k] += g_col_vals.get(k, 0.0)

            # valores formatados para o nó de grupo (= subtotal da categoria)
            g_vals_fmt = [self._fmt(g_col_vals.get(str(cv), 0)) for cv in col_vals] + \
                         [self._fmt(g_total)]

            safe_iid = f"g_{g}"
            if use_row2:
                # nó pai: fechado por padrão, mostra subtotais
                self.result_tree.insert("", "end", iid=safe_iid, text=g,
                                        values=g_vals_fmt, open=False, tags=("grupo",))
                export_rows.append([g] + g_vals_fmt)

                subgrupos = sorted(g_df[row2].dropna().unique().tolist())
                for sg in subgrupos:
                    sg_df  = g_df[g_df[row2] == sg]
                    sg_cv  = vals_por_col(sg_df)
                    sg_tot = soma_dict(sg_cv)
                    sg_fmt = [self._fmt(sg_cv.get(str(cv), 0)) for cv in col_vals] + \
                             [self._fmt(sg_tot)]
                    self.result_tree.insert(safe_iid, "end", text=sg,
                                            values=sg_fmt, tags=("subitem",))
                    export_rows.append(["  " + sg] + sg_fmt)
            else:
                self.result_tree.insert("", "end", iid=safe_iid, text=g,
                                        values=g_vals_fmt, tags=("grupo",))
                export_rows.append([g] + g_vals_fmt)

        if total_g:
            gt_total    = soma_dict(grand_totals)
            gt_vals_fmt = [self._fmt(grand_totals.get(str(cv), 0)) for cv in col_vals] + \
                          [self._fmt(gt_total)]
            self.result_tree.insert("", "end", text="Total Geral",
                                    values=gt_vals_fmt, tags=("total",))
            export_rows.append(["Total Geral"] + gt_vals_fmt)

        self._export_data = export_rows
        self._export_cols = [tree_lbl] + val_col_labels
        self._status_lbl.config(
            text=f"{len(grupos)} grupos | {len(df)} registros | agreg: {agg}  "
                 f"{'(clique no ▶ para expandir grupos)' if use_row2 else ''}")

    # ── exportar ──────────────────────────────────────────────
    def _exportar(self):
        if not self._export_data:
            messagebox.showinfo("Info", "Gere a tabela antes de exportar.")
            return
        if not OPENPYXL_OK:
            messagebox.showerror("Erro", "openpyxl não instalado.\nExecute: pip install openpyxl")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Tabela Dinâmica"
        for r_idx, row in enumerate(self._export_data, start=1):
            for c_idx, val in enumerate(row, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                if r_idx == 1:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill("solid", fgColor="4472C4")
                    cell.font = Font(bold=True, color="FFFFFF")
                elif "Total" in str(row[0]):
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill("solid", fgColor="E2EFDA")
                cell.alignment = Alignment(horizontal="right" if c_idx > 2 else "left")
        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=8)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)
        wb.save(path)
        messagebox.showinfo("Sucesso", f"Exportado para:\n{path}")


# ─────────────────────────── APLICATIVO PRINCIPAL ──────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tabela Dinâmica")
        self.geometry("980x700")
        self.minsize(800, 560)
        self._build()

    def _build(self):
        # barra de abas
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)

        self.formulario = FormularioFrame(nb, self)
        self.importacao = ImportacaoFrame(nb, self)
        self.pivot = TabelaDinamicaFrame(nb, self)

        nb.add(self.formulario, text="  Dados  ")
        nb.add(self.importacao, text="  Importar  ")
        nb.add(self.pivot, text="  Tabela Dinâmica  ")

        nb.bind("<<NotebookTabChanged>>", self._on_tab)

    def _on_tab(self, event):
        tab = event.widget.tab(event.widget.select(), "text").strip()
        if tab == "Tabela Dinâmica":
            self.pivot._atualizar_filtros()


if __name__ == "__main__":
    init_db()
    app = App()
    app.mainloop()
