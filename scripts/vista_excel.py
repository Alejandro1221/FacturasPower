from __future__ import annotations
import os
import math
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, List

import json, hashlib, time
from pathlib import Path

# Directorio de caché en disco y caché de sesión
CACHE_DIR = Path(__file__).resolve().parent / ".cache_resultados"
CACHE_DIR.mkdir(exist_ok=True)
SESSION_CACHE: dict[str, list[dict]] = {}
SESSION_CACHE.clear()

for f in CACHE_DIR.glob("*.json"):
    f.unlink(missing_ok=True)

# Dependencias externas
try:
    import pandas as pd
except Exception as e:
    raise SystemExit("Se requiere 'pandas'. Instala con: pip install pandas openpyxl")

import threading
from comparador_facturas import comparar_desde_excel
from extraer_TotalFactura import DEFAULT_DIR as DEFAULT_PDF_DIR

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
        self._detail_win = None
        self._detail_text = None
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

        # Botón comparar totales
        self.btn_compare = ttk.Button(
            top,
            text="Comparar totales",
            command=self._comparar_totales_ui
        )
        self.btn_compare.pack(side=tk.RIGHT, padx=(8, 0))

        # Botón resetear caché
        self.btn_reset_cache = ttk.Button(
            top,
            text="Resetear caché",
            command=self._reset_cache_ui
        )
        self.btn_reset_cache.pack(side=tk.RIGHT, padx=(8, 8))

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

        # doble clic: ver celda completa
        self.tree.bind("<Double-1>", self._on_double_click_cell)
    
    # ---------------- Helpers internos -----------------
    def _ensure_resultado_column(self):
        if self._df_full is None:
            return

        # normalizamos columnas
        cols_norm = {c.strip().lower(): c for c in self._df_full.columns}

        if "observaciones" in cols_norm:
            real_col = cols_norm["observaciones"]
            self._df_full.drop(columns=[real_col], inplace=True)

        if "Resultado" not in self._df_full.columns:
            self._df_full["Resultado"] = ""

        cols = [c for c in self._df_full.columns if c != "Resultado"] + ["Resultado"]
        self._df_full = self._df_full[cols]

    def _drop_empty_factura_rows(self):
        """Elimina filas donde la columna 'Factura' esté vacía o en blanco."""
        if self._df_full is None:
            return
        if "Factura" in self._df_full.columns:
            self._df_full["Factura"] = self._df_full["Factura"].astype(str).str.strip()
            self._df_full = self._df_full[self._df_full["Factura"] != ""].reset_index(drop=True)

    #----------------- Cargar archivo -----------------
    def load_file(self, path: str):
        ext = os.path.splitext(path)[1].lower()
        if ext not in SUPPORTED_EXT:
            messagebox.showerror("Formato no soportado", f"Extensión no soportada: {ext}")
            return

        # --- LIMPIAR CACHÉ ANTERIOR SI CAMBIA EL ARCHIVO ---
        old_path = getattr(self, "_last_source_path", None)
        if old_path and os.path.abspath(old_path) != os.path.abspath(path):
            # limpiar caché de sesión
            SESSION_CACHE.clear()
            # eliminar cachés de disco viejos
            try:
                for f in CACHE_DIR.glob("*.json"):
                    f.unlink(missing_ok=True)
            except Exception:
                pass
        self._last_source_path = path

        self.source_path = path
        self.title(f"Visor de Excel – {os.path.basename(path)}")

        try:
            if ext == ".csv":
                # Detección de separador simple
                with open(path, "rb") as fh:
                    head = fh.read(4096)
                sample = head.decode("latin-1", errors="ignore")
                sep_guess = ";" if sample.count(";") > sample.count(",") else ","

                # Intentar abrir CSV con varios encodings
                encodings = ["utf-8-sig", "latin-1", "cp1252"]
                df = None
                for enc in encodings:
                    try:
                        df = pd.read_csv(path, sep=sep_guess, dtype=str, keep_default_na=False, encoding=enc)
                        break
                    except UnicodeDecodeError:
                        df = None
                if df is None:
                    raise Exception("No se pudo leer el CSV con los encodings comunes (utf-8, latin-1, cp1252).")

                self._sheet_names = ["(CSV)"]
                self.cb_sheet["values"] = self._sheet_names
                self.cb_sheet.current(0)
                self._df_full = df

                self._ensure_resultado_column()
                self._drop_empty_factura_rows()

            else:
                # Excel con múltiples hojas
                xls = pd.ExcelFile(path, engine="openpyxl")
                self._sheet_names = xls.sheet_names
                self.cb_sheet["values"] = self._sheet_names
                self.cb_sheet.current(0)
                self._df_full = xls.parse(self._sheet_names[0], dtype=str, keep_default_na=False)

                self._ensure_resultado_column()
                self._drop_empty_factura_rows()

            # Aplicar cachés (primero disco, luego sesión)
            self._apply_cache_to_df()
            self._apply_session_cache()

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

            self._ensure_resultado_column()
            self._drop_empty_factura_rows()

            # Al cambiar de hoja, aplicar caché de disco y luego caché de sesión
            self._apply_cache_to_df()
            self._apply_session_cache()

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

    #----------------- Comparar totales -----------------
    def _comparar_totales_ui(self):
        if not self.source_path:
            messagebox.showinfo("Archivo", "Abre primero un Excel/CSV.")
            return

        self.btn_compare.config(state="disabled")
        self.lbl_info.config(text="Comparando…")

        def worker():
            try:
                res = list(comparar_desde_excel(
                    path_excel=self.source_path,
                    carpeta_pdfs=str(DEFAULT_PDF_DIR),
                    col_factura="Factura",   # fijo
                    col_total=None,          # autodetección por 'total'
                    limite=None
                ))

                # Actualiza tabla en UI (escribe en columna Resultado)
                self.after(0, lambda: self._insertar_resultados_en_tabla(res))
                # Guardar cachés
                self.after(0, lambda: self._save_session_cache(list(res)))
                self.after(0, lambda: self._save_cache(res))

                # Mini resumen en la barra de info
                ok  = sum(1 for r in res if r.get("estado") == "OK")
                noc = sum(1 for r in res if r.get("estado") == "NO_COINCIDE")
                nf  = sum(1 for r in res if r.get("estado") == "pdf_no_encontrado")
                self.after(0, lambda: self.lbl_info.config(
                    text=f"Filas: {len(self._df_full)} • Resultado → Correcto:{ok}  Verificar:{noc}  No encontrado:{nf}"
                ))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error", str(e)))
            finally:
                self.after(0, lambda: self.btn_compare.config(state="normal"))

        threading.Thread(target=worker, daemon=True).start()
    
    def _reset_cache_ui(self):
        """Limpia la caché de disco y de sesión, y borra la columna Resultado."""
        from tkinter import messagebox

        msg = (
            "¿Seguro que deseas borrar toda la caché?\n\n"
            "• Se eliminarán resultados guardados.\n"
            "• Se borrará la caché de disco.\n"
            "• La columna 'Resultado' se limpiará.\n\n"
            "Tendrás que volver a comparar las facturas."
        )

        resp = messagebox.askyesno("Resetear caché", msg)
        if not resp:
            return

        # 1) Limpiar caché de sesión
        SESSION_CACHE.clear()

        # 2) Limpiar caché en disco
        try:
            for f in CACHE_DIR.glob("*.json"):
                f.unlink(missing_ok=True)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo limpiar caché:\n{e}")
            return

        # 3) Limpiar resultados de la tabla
        if self._df_full is not None and "Resultado" in self._df_full.columns:
            self._df_full["Resultado"] = ""

        # Refrescar UI
        self.apply_filter()
        self._update_info()

        messagebox.showinfo("Listo", "La caché ha sido eliminada correctamente.")

    #--Inserción de resultados en tabla y caché --
    def _insertar_resultados_en_tabla(self, resultados):
        if self._df_full is None:
            return

        # Aseguramos que Observaciones ya no exista y Resultado esté lista
        self._ensure_resultado_column()
        df = self._df_full

        # Map factura -> estado
        estado_map = {r.get("factura"): r.get("estado") for r in resultados}

        # Escribir Resultado sin usar nunca Observaciones
        for i, row in df.iterrows():
            factura = str(row.get("Factura", "")).strip()
            if not factura:
                continue

            estado = estado_map.get(factura)
            if not estado:
                # si no hay resultado, dejamos lo que tenga (vacío o manual)
                continue

            df.at[i, "Resultado"] = (
                "Correcto"      if estado == "OK" else
                "Verificar"     if estado == "NO_COINCIDE" else
                "No encontrado" if estado == "pdf_no_encontrado" else
                "Verificar"
            )

        # Mover Resultado al final
        if "Resultado" in df.columns:
            cols = [c for c in df.columns if c != "Resultado"] + ["Resultado"]
            self._df_full = df[cols]
        else:
            self._df_full = df

        # refrescar vista
        self.apply_filter()

    # -------------- Filtro / orden / paginación --------------
    def apply_filter(self):
        q = self.var_query.get().strip().lower()
        df = self._df_full.copy() if self._df_full is not None else None
        if df is None:
            return
        if q:
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
        try:
            col_series = pd.to_numeric(
                self._df_filtered[col_name].str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
                errors="coerce"
            )
            if col_series.notna().sum() >= max(3, int(0.6 * len(col_series))):
                self._df_filtered = (
                    self._df_filtered
                    .assign(__num=col_series)
                    .sort_values(["__num", col_name], ascending=[asc, asc])
                    .drop(columns=["__num"])
                    .reset_index(drop=True)
                )
            else:
                self._df_filtered = self._df_filtered.sort_values(
                    col_name, ascending=asc, key=lambda s: s.str.lower()
                ).reset_index(drop=True)
        except Exception:
            self._df_filtered = self._df_filtered.sort_values(
                col_name, ascending=asc, key=lambda s: s.str.lower()
            ).reset_index(drop=True)
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
        self.tree.tag_configure("ok", background="#eaffea")           # verde muy suave
        self.tree.tag_configure("verificar", background="#fff3b0")    # amarillo claro
        self.tree.tag_configure("no_encontrado", background="#ffc9c9") # rojo claro
        self.tree.tag_configure("vacio", background="#ffffff")         # blanco normal
        for _, row in view_df.iterrows():
            values = ["" if pd.isna(v) else str(v) for v in row.tolist()]
            resultado_idx = None
            if "Resultado" in df.columns:
                resultado_idx = df.columns.get_loc("Resultado")

            tag = "vacio"
            if resultado_idx is not None:
                valor_resultado = str(row.iloc[resultado_idx]).strip().lower()
                if "correcto" in valor_resultado:
                    tag = "ok"
                elif "verificar" in valor_resultado:
                    tag = "verificar"
                elif "no encontrado" in valor_resultado:
                    tag = "no_encontrado"

            self.tree.insert("", tk.END, values=values, tags=(tag,))

        self.lbl_page.config(text=f"Página {self.current_page+1}/{pages} • filas {start+1}-{end} de {total_rows}")
        self.after(50, self._autosize_once)

    def _autosize_once(self):
        for c in self.tree["columns"]:
            text = c if len(c) < 30 else c[:30]
            px = 8 * max(10, len(text))
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

    # -------------- Caché (disco) --------------
    def _cache_key(self) -> str:
        """Incluye ruta, hoja, tamaño y mtime para invalidar cuando el archivo cambie."""
        sheet = self.cb_sheet.get() or "(CSV)"
        abspath = os.path.abspath(self.source_path) if self.source_path else ""
        try:
            st = os.stat(abspath)
            sig = f"{abspath}|{sheet}|{st.st_size}|{int(st.st_mtime)}"
        except Exception:
            sig = f"{abspath}|{sheet}"
        return hashlib.sha1(sig.encode("utf-8")).hexdigest()

    def _cache_path(self) -> Path:
        return CACHE_DIR / f"{self._cache_key()}.json"

    def _save_cache(self, resultados: list[dict]):
        try:
            data = {
                "ts": time.time(),
                "resultados": resultados,
            }
            with open(self._cache_path(), "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2, default=str)
        except Exception:
            pass  # best-effort

    def _load_cache(self) -> list[dict] | None:
        try:
            p = self._cache_path()
            if not p.exists():
                return None
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data.get("resultados")
        except Exception:
            return None

    def _apply_cache_to_df(self):
        """Rellena 'Resultado' usando caché en disco (si existe).
           Solo escribe donde 'Resultado' esté vacío para no pisar ediciones manuales."""
        if self._df_full is None:
            return

        cached = self._load_cache()
        if not cached:
            return

        # asegurar columna Resultado
        self._ensure_resultado_column()

        estado_map = {r.get("factura"): r.get("estado") for r in cached}

        for i, row in self._df_full.iterrows():
            factura = str(row.get("Factura", "")).strip()
            if not factura:
                continue
            if str(self._df_full.at[i, "Resultado"]).strip() != "":
                continue
            est = estado_map.get(factura)
            if not est:
                continue
            self._df_full.at[i, "Resultado"] = (
                "Correcto" if est == "OK" else
                "Verificar" if est == "NO_COINCIDE" else
                "No encontrado" if est == "pdf_no_encontrado" else
                "Verificar"
            )

        # mantener Resultado al final
        if "Resultado" in self._df_full.columns:
            cols = [c for c in self._df_full.columns if c != "Resultado"] + ["Resultado"]
            self._df_full = self._df_full[cols]

    # -------------- Caché de sesión (temporal) --------------
    def _source_key(self) -> str | None:
        return os.path.abspath(self.source_path) if self.source_path else None

    def _save_session_cache(self, resultados: list[dict]):
        key = self._source_key()
        if key:
            SESSION_CACHE[key] = resultados

    def _apply_session_cache(self):
        """Rellena 'Resultado' desde caché de sesión (si existe para este archivo).
           No pisa celdas que ya tengan contenido."""
        if self._df_full is None:
            return
        key = self._source_key()
        cached = SESSION_CACHE.get(key) if key else None
        if not cached:
            return

        self._ensure_resultado_column()

        estado_map = {r.get("factura"): r.get("estado") for r in cached}

        for i, row in self._df_full.iterrows():
            factura = str(row.get("Factura", "")).strip()
            if not factura:
                continue
            if str(self._df_full.at[i, "Resultado"]).strip():
                continue
            est = estado_map.get(factura)
            if not est:
                continue
            self._df_full.at[i, "Resultado"] = (
                "Correcto" if est == "OK" else
                "Verificar" if est == "NO_COINCIDE" else
                "No encontrado" if est == "pdf_no_encontrado" else
                "Verificar"
            )

        # mantener Resultado al final
        if "Resultado" in self._df_full.columns:
            cols = [c for c in self._df_full.columns if c != "Resultado"] + ["Resultado"]
            self._df_full = self._df_full[cols]

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
        titulo = f"{cols[col_index]}"

        # --- Reusar ventana si ya existe ---
        if self._detail_win and self._detail_win.winfo_exists():
            self._detail_win.title(titulo)
            self._detail_text.config(state=tk.NORMAL)
            self._detail_text.delete("1.0", tk.END)
            self._detail_text.insert("1.0", contenido)
            self._detail_text.config(state=tk.DISABLED)
            self._detail_win.lift()
            self._detail_win.focus_force()
            return

        # --- Crear ventana nueva si no existe ---
        top = tk.Toplevel(self)
        top.title(titulo)
        top.geometry("600x300")

        txt = tk.Text(top, wrap="word")
        txt.pack(fill=tk.BOTH, expand=True)
        txt.insert("1.0", contenido)
        txt.config(state=tk.DISABLED)

        btn = ttk.Button(top, text="Cerrar", command=top.destroy)
        btn.pack(pady=6)

        # Guardar referencias para reutilizar
        self._detail_win = top
        self._detail_text = txt

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    ExcelTableViewer(root)
    root.mainloop()
