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


def importar_df(df: "pd.DataFrame", modo: str):
    col_map = {c.lower(): c for c in df.columns}

    def get(name):
        return col_map.get(name.lower(), col_map.get(name.lower().replace(" ", "_"), None))

    if modo == "sobrescrever":
        apagar_banco()

    con = sqlite3.connect(DB_PATH)
    for _, r in df.iterrows():
        data_val = str(r.get(get("Data") or "", "")).strip()
        mes, ano = _mes_ano(data_val)
        try:
            valor = float(str(r.get(get("Valor") or "", 0)).replace(",", ".") or 0)
        except (ValueError, TypeError):
            valor = 0.0
        con.execute("""
            INSERT INTO registros (Data, Mes, Ano, Categoria, Sub_Categoria, Transacao, Descricao, Valor)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            data_val, mes, ano,
            str(r.get(get("Categoria") or "", "") or ""),
            str(r.get(get("Sub_Categoria") or get("Sub-Categoria") or "", "") or ""),
            str(r.get(get("Transacao") or get("Transação") or "", "") or ""),
            str(r.get(get("Descricao") or get("Descrição") or "", "") or ""),
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

class TabelaDinamicaFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg="#f5f5f5")
        self.app = app
        self._df = None
        self._build()

    def _build(self):
        tk.Label(self, text="Tabela Dinâmica", font=("Segoe UI", 14, "bold"),
                 bg="#f5f5f5").grid(row=0, column=0, columnspan=4, pady=10)

        if not PANDAS_OK:
            tk.Label(self, text="pandas não instalado.", fg="red",
                     bg="#f5f5f5").grid(row=1, column=0)
            return

        cfg = tk.Frame(self, bg="#f5f5f5", bd=1, relief="groove")
        cfg.grid(row=1, column=0, columnspan=4, padx=10, pady=6, sticky="ew")

        all_cols = ["Ano", "Mes", "Categoria", "Sub_Categoria",
                    "Transacao", "Descricao", "Valor", "Data"]

        def lbl(parent, text, row, col):
            tk.Label(parent, text=text, bg="#f5f5f5",
                     font=("Segoe UI", 10, "bold")).grid(row=row, column=col, padx=6, pady=3, sticky="w")

        lbl(cfg, "Linhas:", 0, 0)
        self._rows_var = tk.StringVar(value="Categoria")
        ttk.Combobox(cfg, textvariable=self._rows_var, values=all_cols, width=16,
                     state="readonly").grid(row=0, column=1, padx=4)

        lbl(cfg, "Colunas:", 0, 2)
        self._cols_var = tk.StringVar(value="Ano")
        ttk.Combobox(cfg, textvariable=self._cols_var, values=["(nenhuma)"] + all_cols, width=16,
                     state="readonly").grid(row=0, column=3, padx=4)

        lbl(cfg, "Valores:", 1, 0)
        self._vals_var = tk.StringVar(value="Valor")
        ttk.Combobox(cfg, textvariable=self._vals_var, values=["Valor"], width=16,
                     state="readonly").grid(row=1, column=1, padx=4)

        lbl(cfg, "Agregar:", 1, 2)
        self._agg_var = tk.StringVar(value="sum")
        ttk.Combobox(cfg, textvariable=self._agg_var,
                     values=["sum", "count", "mean", "min", "max"], width=16,
                     state="readonly").grid(row=1, column=3, padx=4)

        # filtros
        filter_frame = tk.Frame(self, bg="#f5f5f5", bd=1, relief="groove")
        filter_frame.grid(row=2, column=0, columnspan=4, padx=10, pady=4, sticky="ew")
        lbl(filter_frame, "Filtros:", 0, 0)

        lbl(filter_frame, "Ano:", 0, 1)
        self._f_ano = tk.StringVar(value="(todos)")
        self._cb_ano = ttk.Combobox(filter_frame, textvariable=self._f_ano, width=8, state="readonly")
        self._cb_ano.grid(row=0, column=2, padx=4)

        lbl(filter_frame, "Mês:", 0, 3)
        self._f_mes = tk.StringVar(value="(todos)")
        meses = ["(todos)", "1", "2", "3", "4", "5", "6",
                 "7", "8", "9", "10", "11", "12"]
        ttk.Combobox(filter_frame, textvariable=self._f_mes, values=meses,
                     width=6, state="readonly").grid(row=0, column=4, padx=4)

        lbl(filter_frame, "Categoria:", 0, 5)
        self._f_cat = tk.StringVar(value="(todos)")
        self._cb_cat = ttk.Combobox(filter_frame, textvariable=self._f_cat, width=14, state="readonly")
        self._cb_cat.grid(row=0, column=6, padx=4)

        btn_frame = tk.Frame(self, bg="#f5f5f5")
        btn_frame.grid(row=3, column=0, columnspan=4, pady=8)
        tk.Button(btn_frame, text="Gerar Tabela", command=self._gerar,
                  bg="#4CAF50", fg="white", width=14,
                  font=("Segoe UI", 10, "bold")).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Exportar XLSX", command=self._exportar,
                  bg="#1565C0", fg="white", width=14,
                  font=("Segoe UI", 10)).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Atualizar Filtros", command=self._atualizar_filtros,
                  bg="#757575", fg="white", width=14,
                  font=("Segoe UI", 10)).pack(side="left", padx=6)

        # resultado
        res_frame = tk.Frame(self)
        res_frame.grid(row=4, column=0, columnspan=4, sticky="nsew", padx=8, pady=4)
        self.rowconfigure(4, weight=1)
        self.columnconfigure(3, weight=1)

        self.result_tree = ttk.Treeview(res_frame, show="headings")
        vsb = ttk.Scrollbar(res_frame, orient="vertical", command=self.result_tree.yview)
        hsb = ttk.Scrollbar(res_frame, orient="horizontal", command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.result_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        res_frame.rowconfigure(0, weight=1)
        res_frame.columnconfigure(0, weight=1)

        self._status_lbl = tk.Label(self, text="", bg="#f5f5f5", font=("Segoe UI", 9))
        self._status_lbl.grid(row=5, column=0, columnspan=4)
        self._atualizar_filtros()

    def _carregar_df(self):
        cols, rows = buscar_todos()
        if not rows:
            return pd.DataFrame()
        df = pd.DataFrame(rows, columns=cols)
        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
        df["Mes"] = pd.to_numeric(df["Mes"], errors="coerce")
        df["Ano"] = pd.to_numeric(df["Ano"], errors="coerce")
        return df

    def _atualizar_filtros(self):
        df = self._carregar_df()
        self._df = df
        if df.empty:
            self._cb_ano["values"] = ["(todos)"]
            self._cb_cat["values"] = ["(todos)"]
            return
        anos = ["(todos)"] + sorted(df["Ano"].dropna().unique().astype(int).astype(str).tolist())
        cats = ["(todos)"] + sorted(df["Categoria"].dropna().unique().tolist())
        self._cb_ano["values"] = anos
        self._cb_cat["values"] = cats

    def _gerar(self):
        self._atualizar_filtros()
        df = self._df.copy() if self._df is not None else self._carregar_df()
        if df.empty:
            messagebox.showinfo("Info", "Banco de dados vazio.")
            return

        # filtros
        if self._f_ano.get() != "(todos)":
            df = df[df["Ano"] == int(self._f_ano.get())]
        if self._f_mes.get() != "(todos)":
            df = df[df["Mes"] == int(self._f_mes.get())]
        if self._f_cat.get() != "(todos)":
            df = df[df["Categoria"] == self._f_cat.get()]

        if df.empty:
            messagebox.showinfo("Info", "Nenhum dado para os filtros selecionados.")
            return

        row_field = self._rows_var.get()
        col_field = self._cols_var.get()
        agg = self._agg_var.get()

        try:
            if col_field == "(nenhuma)":
                pivot = df.groupby(row_field)["Valor"].agg(agg).reset_index()
                pivot.columns = [row_field, agg.upper()]
                pivot["TOTAL"] = pivot[agg.upper()]
            else:
                pivot = pd.pivot_table(
                    df,
                    values="Valor",
                    index=row_field,
                    columns=col_field,
                    aggfunc=agg,
                    fill_value=0,
                    margins=True,
                    margins_name="TOTAL",
                )
                pivot = pivot.reset_index()
                pivot.columns = [str(c) for c in pivot.columns]
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            return

        self._pivot_df = pivot
        cols = list(pivot.columns)
        self.result_tree["columns"] = cols
        for i in self.result_tree.get_children():
            self.result_tree.delete(i)
        for col in cols:
            self.result_tree.heading(col, text=str(col))
            self.result_tree.column(col, width=max(80, len(str(col)) * 9), anchor="center")

        for _, row in pivot.iterrows():
            vals = []
            for v in row:
                if isinstance(v, float):
                    vals.append(f"{v:,.2f}")
                else:
                    vals.append(str(v))
            self.result_tree.insert("", "end", values=vals)

        self._status_lbl.config(
            text=f"{len(pivot)-1 if col_field != '(nenhuma)' else len(pivot)} linhas | agreg: {agg} | filtros aplicados")

    def _exportar(self):
        if not hasattr(self, "_pivot_df") or self._pivot_df is None:
            messagebox.showinfo("Info", "Gere a tabela antes de exportar.")
            return
        if not OPENPYXL_OK:
            messagebox.showerror("Erro", "openpyxl não instalado.\nExecute: pip install openpyxl")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if path:
            self._pivot_df.to_excel(path, index=False)
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
