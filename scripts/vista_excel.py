from __future__ import annotations
import os
import sys
import math
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional, List

# Dependencias externas
try:
    import pandas as pd
except Exception as e:
    raise SystemExit("Se requiere 'pandas'. Instala con: pip install pandas openpyxl")

SUPPORTED_EXT = (".xlsx", ".xlsm", ".xltx", ".xltm", ".csv")


class ExcelTableViewer(tk.Toplevel):
    def __init__(self, master: tk.Misc | None = None, filepath: Optional[str] = None, page_size: int = 500):
        super().__init__(master)
        self.title("Visor de Excel – FacturasPower")
        self.geometry("1100x650")
        self.minsize(900, 500)
        self.page_size = page_size
        self.current_page = 0
        self.sort_state = {}   
        self.source_path: Optional[str] = None
        self._df_full = None        # DataFrame completo (hoja actual / csv)
        self._df_filtered = None    # DF con filtro aplicado
        self._sheet_names: List[str] = []
        self.after(0, lambda: self.bind("<Escape>", lambda e: self.destroy()))
        self.protocol("WM_DELETE_WINDOW", self.destroy)

        self._build_ui()
        if filepath:
            self.load_file(filepath)

    # ---------------- UI -----------------
    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=10, pady=8)

        # Hoja
        ttk.Label(top, text="Hoja:").pack(side=tk.LEFT, padx=(12, 4))
        self.cb_sheet = ttk.Combobox(top, state="readonly", width=32)
        self.cb_sheet.pack(side=tk.LEFT)
        self.cb_sheet.bind("<<ComboboxSelected>>", lambda e: self._on_sheet_selected())

        # Búsqueda
        ttk.Label(top, text="Buscar:").pack(side=tk.LEFT, padx=(12, 4))
        self.var_query = tk.StringVar()
        self.entry_query = ttk.Entry(top, textvariable=self.var_query, width=28)
        self.entry_query.pack(side=tk.LEFT)
        self.entry_query.bind("<KeyRelease>", lambda e: self.apply_filter())

        # Info
        self.lbl_info = ttk.Label(top, text="Sin archivo")
        self.lbl_info.pack(side=tk.RIGHT)

        # Tree + scrolls
        mid = ttk.Frame(self)
        mid.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 8))

        self.tree = ttk.Treeview(mid, columns=(), show="headings")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        vsb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(mid, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

        # página inferior
        bottom = ttk.Frame(self)
        bottom.pack(fill=tk.X, padx=10, pady=(0,10))
        self.btn_prev = ttk.Button(bottom, text="◀ Anterior", command=self.prev_page)
        self.btn_next = ttk.Button(bottom, text="Siguiente ▶", command=self.next_page)
        self.lbl_page = ttk.Label(bottom, text="Página 0/0")
        self.btn_prev.pack(side=tk.LEFT)
        self.btn_next.pack(side=tk.LEFT, padx=6)
        self.lbl_page.pack(side=tk.LEFT, padx=12)

        # double click: ver celda completa
        self.tree.bind("<Double-1>", self._on_double_click_cell)

    #----------------- Cargar archivo -----------------
    def load_file(self, path: str):
        ext = os.path.splitext(path)[1].lower()
        if ext not in SUPPORTED_EXT:
            messagebox.showerror("Formato no soportado", f"Extensión no soportada: {ext}")
            return
        self.source_path = path
        self.title(f"Visor de Excel – {os.path.basename(path)}")

        try:
            if ext == ".csv":
                df = pd.read_csv(path, dtype=str, keep_default_na=False)
                self._sheet_names = ["(CSV)"]
                self.cb_sheet["values"] = self._sheet_names
                self.cb_sheet.current(0)
                self._df_full = df
            else:
                # Excel con múltiples hojas
                xls = pd.ExcelFile(path, engine="openpyxl")
                self._sheet_names = xls.sheet_names
                self.cb_sheet["values"] = self._sheet_names
                self.cb_sheet.current(0)
                self._df_full = xls.parse(self._sheet_names[0], dtype=str, keep_default_na=False)
            self.var_query.set("")
            self.current_page = 0
            self.sort_state.clear()
            self.apply_filter()
            self._update_info()
        except Exception as e:
            messagebox.showerror("Error al abrir", str(e))

    def _on_sheet_selected(self):
        if not self.source_path:
            return
        try:
            xls = pd.ExcelFile(self.source_path, engine="openpyxl")
            sheet = self.cb_sheet.get()
            self._df_full = xls.parse(sheet, dtype=str, keep_default_na=False)
            self.var_query.set("")
            self.current_page = 0
            self.sort_state.clear()
            self.apply_filter()
            self._update_info()
        except Exception as e:
            messagebox.showerror("Error al cargar hoja", str(e))
    
    def _update_info(self):
        try:
            if self._df_full is None:
                txt = "Sin archivo"
            else:
                filas = len(self._df_full)
                cols = len(self._df_full.columns)
                txt = f"Filas: {filas} • Columnas: {cols}"
            self.lbl_info.config(text=txt)
        except Exception:
            pass

    # -------------- Filtro / orden / paginación --------------
    def apply_filter(self):
        q = self.var_query.get().strip().lower()
        df = self._df_full.copy() if self._df_full is not None else None
        if df is None:
            return
        if q:
            # filtro contiene en cualquier columna (string contains casefold)
            def row_match(series):
                return any((str(v).lower().find(q) >= 0) for v in series.values)
            mask = df.apply(row_match, axis=1)
            df = df[mask]
        self._df_filtered = df.reset_index(drop=True)
        self.current_page = 0
        self.render_page()

    def _sort_by(self, col_name: str):
        if self._df_filtered is None:
            return
        asc = not self.sort_state.get(col_name, True)
        # intentar convertir a número donde se pueda
        try:
            col_series = pd.to_numeric(self._df_filtered[col_name].str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce")
            # si hay suficientes numéricos, ordenar por esa serie; sino por texto
            if col_series.notna().sum() >= max(3, int(0.6 * len(col_series))):
                self._df_filtered = self._df_filtered.assign(__num=col_series).sort_values(["__num", col_name], ascending=[asc, asc]).drop(columns=["__num"]).reset_index(drop=True)
            else:
                self._df_filtered = self._df_filtered.sort_values(col_name, ascending=asc, key=lambda s: s.str.lower()).reset_index(drop=True)
        except Exception:
            self._df_filtered = self._df_filtered.sort_values(col_name, ascending=asc, key=lambda s: s.str.lower()).reset_index(drop=True)
        self.sort_state[col_name] = asc
        self.current_page = 0
        self.render_page()

    def render_page(self):
        df = self._df_filtered
        if df is None:
            return
        # configurar columnas
        cols = list(df.columns)
        self.tree["columns"] = cols
        for c in self.tree["columns"]:
            self.tree.heading(c, text=c, command=lambda cn=c: self._sort_by(cn))
            self.tree.column(c, width=120, stretch=True, anchor=tk.W)
        # limpiar filas
        self.tree.delete(*self.tree.get_children())

        # paginación
        total_rows = len(df)
        if total_rows == 0:
            self.lbl_page.config(text="Página 0/0 • 0 filas")
            return
        pages = math.ceil(total_rows / self.page_size)
        start = self.current_page * self.page_size
        end = min(start + self.page_size, total_rows)

        # insertar filas
        view_df = df.iloc[start:end]
        for _, row in view_df.iterrows():
            values = ["" if pd.isna(v) else str(v) for v in row.tolist()]
            self.tree.insert("", tk.END, values=values)

        self.lbl_page.config(text=f"Página {self.current_page+1}/{pages} • filas {start+1}-{end} de {total_rows}")
        # ajustar ancho inicial (una sola vez por tamaño de tabla)
        self.after(50, self._autosize_once)

    def _autosize_once(self):
        # intenta ajustar un poco los anchos según headers
        for c in self.tree["columns"]:
            text = c if len(c) < 30 else c[:30]
            px = 8 * max(10, len(text))  # heurística simple
            cur = self.tree.column(c, option="width")
            if cur < px:
                self.tree.column(c, width=px)

    def next_page(self):
        if self._df_filtered is None:
            return
        total_rows = len(self._df_filtered)
        pages = max(1, math.ceil(total_rows / self.page_size))
        if self.current_page + 1 < pages:
            self.current_page += 1
            self.render_page()

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.render_page()

    # -------------- UX: ver celda completa --------------
    def _on_double_click_cell(self, event):
        item_id = self.tree.identify_row(event.y)
        colid = self.tree.identify_column(event.x)
        if not item_id or not colid:
            return
        col_index = int(colid.replace('#', '')) - 1
        cols = self.tree["columns"]
        if col_index < 0 or col_index >= len(cols):
            return
        values = self.tree.item(item_id, "values")
        contenido = values[col_index] if col_index < len(values) else ""

        top = tk.Toplevel(self)
        top.title(f"{cols[col_index]}")
        top.geometry("600x300")
        txt = tk.Text(top, wrap="word")
        txt.pack(fill=tk.BOTH, expand=True)
        txt.insert("1.0", contenido)
        txt.config(state=tk.DISABLED)
        ttk.Button(top, text="Cerrar", command=top.destroy).pack(pady=6)


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  
    ExcelTableViewer(root)
    root.mainloop()
