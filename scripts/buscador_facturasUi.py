import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import threading
from pathlib import Path
from buscar_facturas import buscar as buscar_en_sharepoint
from buscar_facturas import set_graph_token 
from vista_excel import ExcelTableViewer

BASE_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent


# === PALETA DE COLORES ===
PALETTE = {
    "bg": "#f4f7fb",        # fondo app
    "card": "#ffffff",      # tarjetas
    "muted": "#eef3fb",     # encabezados de panel
    "text": "#2b2f36",
    "subtext": "#5f6b7a",
    "accent": "#3b82f6",  
    "accent_hover": "#2563eb",
    "success": "#22c55e",
    "border": "#e5eaf2",
}


# === APLICAR TEMA ===
def apply_theme(root: tk.Tk):
    root.configure(bg=PALETTE["bg"])
    style = ttk.Style(root)
    style.theme_use("clam")

    # App global
    style.configure(".", background=PALETTE["bg"], foreground=PALETTE["text"], font=("Segoe UI", 10))

    # Cards
    style.configure("Card.TFrame", background=PALETTE["card"], relief="flat")
    style.configure("Muted.TFrame", background=PALETTE["muted"], relief="flat")
    style.configure("Card.TLabelframe", background=PALETTE["card"], relief="flat", bordercolor=PALETTE["border"])
    style.configure("Card.TLabelframe.Label", background=PALETTE["muted"], foreground=PALETTE["text"],
                    padding=(10,6), font=("Segoe UI Semibold", 10))
    style.map("Card.TLabelframe", bordercolor=[("focus", PALETTE["accent"])])

    # Entradas
    style.configure("Rounded.TEntry", fieldbackground="#ffffff", background="#ffffff",
                    bordercolor=PALETTE["border"], lightcolor=PALETTE["accent"], darkcolor=PALETTE["border"],
                    relief="flat", padding=6)

    # Labels
    style.configure("Title.TLabel", background=PALETTE["card"], font=("Segoe UI Semibold", 11))
    style.configure("Kicker.TLabel", foreground=PALETTE["subtext"], background=PALETTE["card"])

    # Botones
    style.configure("TButton", padding=(12,6), background=PALETTE["card"], bordercolor=PALETTE["border"])
    style.map("TButton", background=[("active", PALETTE["muted"])])

    style.configure("Success.TButton", foreground="white", background=PALETTE["success"], bordercolor=PALETTE["success"])
    style.map("Success.TButton", background=[("active", "#16a34a")]) 

    style.configure(
        "Accent.TButton",
        foreground="white",
        background=PALETTE["accent"],
        bordercolor=PALETTE["accent"],
    )
    style.map(
        "Accent.TButton",
        background=[
            ("disabled", "#cbd5e1"),               # gris cuando está deshabilitado
            ("pressed", PALETTE["accent_hover"]),  # clic
            ("active", PALETTE["accent_hover"]),   # hover
        ],
        foreground=[
            ("disabled", "#888"),
            ("!disabled", "white"),
        ],
        bordercolor=[
            ("disabled", "#cbd5e1"),
            ("!disabled", PALETTE["accent"]),
        ],
    )

        # Scrollbars (vertical)
    style.configure(
        "Custom.Vertical.TScrollbar",
        gripcount=0,
        background="#8cdceb",   
        darkcolor="#dbe7ff",
        lightcolor="#dbe7ff",
        troughcolor=PALETTE["muted"],  
        bordercolor=PALETTE["border"],
        relief="flat",
        arrowsize=12
    )
    style.map(
        "Custom.Vertical.TScrollbar",
        background=[("active", "#c7d8ff"), ("pressed", "#b5ccff")],
        arrowcolor=[("!disabled", PALETTE["subtext"])]
    )


    # Progressbar 
    style.configure(
        "Loading.Horizontal.TProgressbar",
        troughcolor=PALETTE["muted"],  
        background=PALETTE["accent"],  
        bordercolor=PALETTE["border"],
        lightcolor=PALETTE["accent"],  
        darkcolor=PALETTE["accent"],
        thickness=10,              
        troughrelief="flat",
        relief="flat"
    )
    style.map(
        "Loading.Horizontal.TProgressbar",
        background=[("active", PALETTE["accent_hover"])]
    )


# === UTILIDADES ===
def detectar_columna_factura(df: pd.DataFrame):
    if df is None or df.empty:
        return None

    cols = list(df.columns)
    normalizadas = {c: str(c).strip().lower() for c in cols}

    # 1) Exactamente 'factura'
    for c, n in normalizadas.items():
        if n == "factura":
            return c

    # 2) Empiece por 'factura'
    candidatos = [c for c, n in normalizadas.items() if n.startswith("factura")]
    if candidatos:
        return candidatos[0]

    # 3) Contenga 'factura' pero NO 'fecha'
    candidatos = [
        c for c, n in normalizadas.items()
        if "factura" in n and "fecha" not in n
    ]
    if candidatos:
        return candidatos[0]

    # 4) Último recurso: cualquier columna que contenga 'factura'
    for c, n in normalizadas.items():
        if "factura" in n:
            return c

    return None


def leer_dataframe_robusto(ruta: str) -> pd.DataFrame:
    low = ruta.lower()

    # ODS
    if low.endswith(".ods"):
        # requiere odfpy
        return pd.read_excel(ruta, engine="odf", dtype=str)

    # CSV
    if low.endswith(".csv"):
        with open(ruta, "rb") as fh:
            head = fh.read(4096)
        sample = head.decode("latin-1", errors="ignore")
        sep_guess = ";" if sample.count(";") > sample.count(",") else ","
        for enc in ["utf-8", "utf-8-sig", "latin-1", "cp1252"]:
            try:
                return pd.read_csv(ruta, sep=sep_guess, engine="python", encoding=enc, dtype=str)
            except Exception:
                pass
        raise Exception("No se pudo leer el CSV con los encodings probados.")

    # XLSX (openpyxl)
    if low.endswith(".xlsx"):
        # requiere openpyxl
        return pd.read_excel(ruta, engine="openpyxl", dtype=str)

    # XLS (opcional)
    if low.endswith(".xls"):
        # xlrd>=2 ya no lee .xls; usa la 1.2.0
        try:
            return pd.read_excel(ruta, engine="xlrd", dtype=str)
        except Exception as e:
            raise Exception("Para .xls instala xlrd==1.2.0 o conviértelo a .xlsx") from e

    # Por defecto: intentar como Excel moderno
    return pd.read_excel(ruta, engine="openpyxl", dtype=str)

# APLICACIÓN PRINCIPAL 
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("FacturasPower")
        try:
            BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
            icono = BASE_DIR / "icono.ico"
            if icono.exists():
                self.iconbitmap(icono)
        except Exception as e:
            print("No se pudo cargar el icono:", e)
        self.geometry("1024x620")
        self.minsize(960, 560)

        apply_theme(self)

        # Variables de estado
        self.ruta_archivo = tk.StringVar()
        self.columna_detectada = tk.StringVar(value="(sin cargar)")
        self.facturas = []
        self.status = tk.StringVar(value="Listo.")
        self._loading_win = None

        self._loading_win = None
        self._loading_label = None
        self._excel_viewer = None
        self._build_ui()

    def _build_ui(self):
        # Contenedor principal tipo “card”
        shell = ttk.Frame(self, style="Card.TFrame", padding=14)
        shell.pack(fill="both", expand=True, padx=18, pady=18)

        # TOP: selector de archivo
        top = ttk.Frame(shell, style="Card.TFrame")
        top.pack(fill="x")
        ttk.Label(top, text="Tabla Excel", style="Kicker.TLabel").pack(side="left", padx=(2,8))
        ttk.Entry(top, textvariable=self.ruta_archivo, style="Rounded.TEntry").pack(side="left", padx=8, fill="x", expand=True)
        ttk.Button(top, text="Buscar", style="Accent.TButton", command=self.elegir_archivo).pack(side="left")

        # SUB: columna detectada
        sub = ttk.Frame(shell, style="Card.TFrame")
        sub.pack(fill="x", pady=(10,0))
        ttk.Label(sub, text="Columna detectada:", style="Kicker.TLabel").pack(side="left")
        ttk.Label(sub, textvariable=self.columna_detectada, foreground="#10b981", style="Kicker.TLabel").pack(side="left", padx=(6,0))

        # ==== Sección Configuración (botones lado a lado) ====
        cfg = ttk.Frame(shell, style="Card.TFrame")
        cfg.pack(fill="x", pady=(10, 0))

        btn_conf = ttk.Button(
            cfg, text="Configurar token", style="Accent.TButton",
            command=self._open_token_config
        )
        btn_conf.pack(side="left")

        ttk.Label(cfg, text="  ", style="Card.TFrame").pack(side="left") 

        ttk.Button(
            cfg, text="Tabla Excel", style="TButton",
            command=lambda: self._abrir_visor_excel()
        ).pack(side="left")

        # Paneles de listas
        body = ttk.Frame(shell, style="Card.TFrame")
        body.pack(fill="both", expand=True, pady=12)

        def make_panel(parent, title):
            card = ttk.Labelframe(parent, text=title, style="Card.TLabelframe", padding=8)
            wrapper = ttk.Frame(card, style="Card.TFrame")
            wrapper.pack(fill="both", expand=True)
            lb = tk.Listbox(wrapper, relief="flat", highlightthickness=1,
                            highlightbackground=PALETTE["border"],
                            bd=0, activestyle="dotbox", font=("Segoe UI", 10))
            sb = ttk.Scrollbar(wrapper, orient="vertical", command=lb.yview, style="Custom.Vertical.TScrollbar")
            lb.configure(yscrollcommand=sb.set)
            lb.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")
            return card, lb

        left_col, self.lb_lista = make_panel(body, "Lista Facturas")
        mid_col,  self.lb_ok    = make_panel(body, "Facturas encontradas")
        right_col, self.lb_nok  = make_panel(body, "Facturas NO encontradas")

        left_col.grid(row=0, column=0, sticky="nsew", padx=(0,8))
        mid_col.grid(row=0, column=1, sticky="nsew", padx=8)
        right_col.grid(row=0, column=2, sticky="nsew", padx=(8,0))
        for c in range(3): body.columnconfigure(c, weight=1)
        body.rowconfigure(0, weight=1)

        # BARRA INFERIOR
        bottom = ttk.Frame(shell, style="Card.TFrame")
        bottom.pack(fill="x", pady=(4,0))

        self.btn_limpiar = ttk.Button(bottom,text="Limpiar lista",style="Success.TButton",command=self.limpiar)
        self.btn_limpiar.pack(side="left")

        self.btn_vaciar = ttk.Button(bottom, text="Limpiar carpeta",command=self.vaciar_descargas)
        self.btn_vaciar.pack(side="right", padx=(8,8))

        self.btn_exportar = ttk.Button(bottom, text="📦 Exportar carpeta", command=self.exportar_descargas, state=tk.DISABLED)
        self.btn_exportar.pack(side="right", padx=(0,8))

        self.btn_descargar = ttk.Button(bottom,text="📂 Abrir carpeta de descargas",command=self.abrir_carpeta_descargas)
        self.btn_descargar.pack(side="right", padx=(0,8))

        self.btn_buscar_sp = ttk.Button(bottom, text="🔎 Buscar en SharePoint", style="Accent.TButton", command=self._buscar_sharepoint, state=tk.DISABLED)
        self.btn_buscar_sp.pack(side="right", padx=(8,8))

        # STATUS BAR
        ttk.Label(self, textvariable=self.status, padding=(16,8)).pack(fill="x", padx=10, pady=(0,6))
        self._actualizar_boton_exportar()

    def _abrir_visor_excel(self, path: str | None = None):
            """
            Abre una sola ventana de visor Excel.
            Si ya está abierta, la trae al frente.
            """
            p = path or (self.ruta_archivo.get().strip() or None)

            # Si ya hay una ventana abierta, reúsala
            if self._excel_viewer and self._excel_viewer.winfo_exists():
                try:
                    if p:
                        self._excel_viewer.load_file(p)  # si tu ExcelTableViewer tiene este método
                    self._excel_viewer.lift()
                    self._excel_viewer.focus_force()
                    return
                except Exception:
                    pass

            # Si no hay una abierta, crea una nueva
            from vista_excel import ExcelTableViewer
            self._excel_viewer = ExcelTableViewer(self, filepath=p)

    def _open_token_config(self):
        win = tk.Toplevel(self)
        win.title("Configurar token de Microsoft Graph")
        win.transient(self)
        win.grab_set()
        win.resizable(False, False)
        frm = ttk.Frame(win, padding=16, style="Card.TFrame")
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Pega tu token (JWT):", style="Title.TLabel").grid(row=0, column=0, sticky="w")

        var_token = tk.StringVar(value="")
        ent = ttk.Entry(frm, textvariable=var_token, width=68, show="•", style="Rounded.TEntry")
        ent.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(6,10))
        ent.focus_set()

        def toggle_show():
            ent.config(show="" if ent.cget("show") else "•")
            btn_ver.config(text="Ocultar" if ent.cget("show")=="" else "Mostrar")

        def pegar():
            try:
                txt = win.clipboard_get()
                var_token.set(txt.strip())
            except Exception:
                messagebox.showwarning("Portapapeles", "No se pudo leer el portapapeles.", parent=win)

        def guardar():
            tok = var_token.get().strip()
            if len(tok) < 20:
                messagebox.showerror("Token", "Token inválido o muy corto.", parent=win)
                return
            try:
                set_graph_token(tok, persist=True)  
                messagebox.showinfo("Token", "Token guardado correctamente.", parent=win)
                win.destroy()
            except Exception as e:
                messagebox.showerror("Error", str(e), parent=win)

        btn_ver = ttk.Button(frm, text="Mostrar", command=toggle_show)
        btn_pegar = ttk.Button(frm, text="Pegar", command=pegar)
        btn_guardar = ttk.Button(frm, text="Guardar", style="Accent.TButton", command=guardar)
        btn_cancelar = ttk.Button(frm, text="Cancelar", command=win.destroy)

        btn_ver.grid(row=2, column=0, sticky="w", pady=(0,4))
        btn_pegar.grid(row=2, column=1, sticky="w", pady=(0,4), padx=(8,0))
        btn_guardar.grid(row=3, column=2, sticky="e", pady=(6,0))
        btn_cancelar.grid(row=3, column=3, sticky="e", pady=(6,0))

    
    # ------- MODAL DE CARGANDO -------
    def _show_loading(self, message="Procesando..."):
        if self._loading_win and tk.Toplevel.winfo_exists(self._loading_win):
            # ya existe: actualiza y retorna
            self._loading_label.config(text=message)
            return

        win = tk.Toplevel(self)
        win.title("Trabajando…")
        win.transient(self)
        win.grab_set()
        win.resizable(False, False)
        win.configure(bg=PALETTE["card"])
        win.protocol("WM_DELETE_WINDOW", lambda: None)  # bloquear cerrar

        # centrar
        win.update_idletasks()
        w, h = 420, 140
        x = self.winfo_rootx() + (self.winfo_width() - w) // 2
        y = self.winfo_rooty() + (self.winfo_height() - h) // 2
        win.geometry(f"{w}x{h}+{x}+{y}")

        frm = ttk.Frame(win, style="Card.TFrame", padding=16)
        frm.pack(fill="both", expand=True)

        self._loading_label = ttk.Label(frm, text=message, style="Title.TLabel")
        self._loading_label.pack(anchor="w", pady=(0,8))

        pb = ttk.Progressbar(frm, mode="indeterminate", length=360, style="Loading.Horizontal.TProgressbar")
        pb.pack(fill="x")
        pb.start(12)

        self._loading_win = win
        self.update_idletasks()

    def _update_loading_message(self, text):
        if self._loading_win and tk.Toplevel.winfo_exists(self._loading_win):
            self._loading_label.config(text=text)

    def _hide_loading(self):
        if self._loading_win and tk.Toplevel.winfo_exists(self._loading_win):
            try:
                self._loading_win.grab_release()
            except:
                pass
            self._loading_win.destroy()
        self._loading_win = None
        self._loading_label = None

    def _run_in_bg_with_progress(self, start_message, worker_func, done_callback):
        self._show_loading(start_message)

        def worker():
            res, err = None, None
            try:
                def on_progress(msg: str):
                    # actualizar desde el hilo, pero en el mainloop
                    self.after(0, lambda: self._update_loading_message(msg))
                # el worker debe aceptar ese callback
                res = worker_func(on_progress)
            except Exception as e:
                err = e
            # cerrar modal + callback en hilo principal
            self.after(0, lambda: (self._hide_loading(), done_callback(res, err)))

        threading.Thread(target=worker, daemon=True).start()


    # ===== FUNCIONALIDAD =====
    def elegir_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo",
            filetypes=[
                ("CSV", "*.csv"),
                ("ODS (LibreOffice)", "*.ods"),
                ("Excel", "*.xlsx *.xls"),
                ("Todos", "*.*"),
            ],
        )
        if not ruta:
            return
        self.ruta_archivo.set(ruta)
        self.cargar_archivo(ruta)

    def cargar_archivo(self, ruta):
        try:
            df = leer_dataframe_robusto(ruta)
            col = detectar_columna_factura(df)
            if not col:
                self.columna_detectada.set("(no encontrada)")
                self.facturas = []
                self.refrescar_listas()
                messagebox.showwarning("Columna no encontrada",
                                       "No se halló una columna que contenga 'Factura'.")
                return

            self.columna_detectada.set(col)
            serie = df[col].astype(str).str.strip()
            serie = serie[serie.notna() & (serie != "") & (serie.str.lower() != "nan")]

            # únicos preservando orden
            vistos, facturas = set(), []
            for v in serie:
                if v not in vistos:
                    vistos.add(v)
                    facturas.append(v)

            self.facturas = facturas
            self.refrescar_listas()
            self.status.set(f"Cargadas {len(self.facturas)} facturas.")
            self.btn_buscar_sp.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("Error al leer archivo", str(e))
            self.status.set("Error al procesar el archivo.")

    
    def _buscar_sharepoint(self):
        if not self.facturas:
            messagebox.showinfo("Sin facturas", "Primero carga un archivo con facturas.")
            return

        self.btn_buscar_sp.config(state=tk.DISABLED)

        def worker_func(on_progress):
            # pasa el callback al backend
            return buscar_en_sharepoint(self.facturas, on_progress=on_progress)

        def done_callback(resultado, err):
            self.btn_buscar_sp.config(state=tk.NORMAL)

            # --- NUEVO: si falta token, abrir diálogo y reintentar una vez
            if err and ("Falta GRAPH_TOKEN" in str(err) or "Acceso denegado (401" in str(err) or "Acceso denegado (403" in str(err)):
                resp = messagebox.askyesno(
                    "Token requerido",
                    "Parece que falta el token de Graph o no es válido.\n¿Deseas configurarlo ahora?"
                )
                if resp:
                    self._open_token_config()
                    messagebox.showinfo(
                        "Info",
                        "Después de guardar el token, vuelve a presionar 'Buscar en SharePoint'.",
                        parent=self
                    )
                else:
                    self.status.set("🔒 Operación cancelada por falta de token.")
                return


            if err:
                messagebox.showerror("Error al buscar", str(err))
                self.status.set("❌ Error durante la búsqueda.")
                return

            encontradas = resultado.get("encontradas", [])
            no_encontradas = resultado.get("no_encontradas", [])
            descargadas = resultado.get("descargadas", [])

            self.lb_ok.delete(0, "end")
            self.lb_nok.delete(0, "end")
            for f in encontradas:
                self.lb_ok.insert("end", f)
            for f in no_encontradas:
                self.lb_nok.insert("end", f)

            total = len(encontradas) + len(no_encontradas)
            self.status.set(
                f"Finalizado: {len(encontradas)} encontradas, "
                f"{len(descargadas)} descargadas, {len(no_encontradas)} no encontradas."
            )
            self._actualizar_boton_exportar()

        self._run_in_bg_with_progress("Buscando en SharePoint…", worker_func, done_callback)

    def refrescar_listas(self):
        self.lb_lista.delete(0, "end")
        for f in self.facturas[:1000]:
            self.lb_lista.insert("end", f)
        self.lb_ok.delete(0, "end")
        self.lb_nok.delete(0, "end")

    def abrir_carpeta_descargas(self):
        try:
            ruta = BASE_DIR / "Facturas_descargadas"
            ruta.mkdir(exist_ok=True, parents=True)
            if os.name == "nt":
                os.startfile(ruta)
            elif sys.platform == "darwin":
                import subprocess; subprocess.run(["open", str(ruta)], check=False)
            else:
                import subprocess; subprocess.run(["xdg-open", str(ruta)], check=False)
            self.status.set(f"📂 Carpeta abierta: {ruta}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta:\n{e}")
            self.status.set("Error al abrir la carpeta de descargas.")

    # Exportar carpeta de descargas
    def exportar_descargas(self):
        from tkinter import filedialog

        origen = BASE_DIR / "Facturas_descargadas"
        origen.mkdir(exist_ok=True, parents=True)

        if not any(origen.iterdir()):
            messagebox.showinfo("Exportar", "No hay archivos para exportar.")
            return

        destino = filedialog.askdirectory(
            title="Selecciona la carpeta donde guardar las facturas"
        )

        if not destino:
            return

        try:
            import shutil
            destino = Path(destino)

            # Copiar archivos uno por uno
            copiados = 0
            for file in origen.iterdir():
                if file.is_file():
                    shutil.copy2(file, destino / file.name)
                    copiados += 1

            messagebox.showinfo(
                "Exportación completa",
                f"Se copiaron {copiados} archivos correctamente a:\n\n{destino}"
            )

            self.status.set(f"Archivos exportados a {destino}")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron exportar los archivos:\n{e}")
        
    def _actualizar_boton_exportar(self):
        try:
            origen = BASE_DIR / "Facturas_descargadas"
            origen.mkdir(exist_ok=True, parents=True)

            hay_archivos = any(origen.iterdir())
            self.btn_exportar.config(state=tk.NORMAL if hay_archivos else tk.DISABLED)
        except Exception:
            # Si algo raro pasa, mejor dejarlo deshabilitado
            self.btn_exportar.config(state=tk.DISABLED)

    def vaciar_descargas(self):
        origen = BASE_DIR / "Facturas_descargadas"
        origen.mkdir(exist_ok=True, parents=True)

        # Confirmación
        resp = messagebox.askyesno(
            "Vaciar carpeta",
            "¿Seguro que deseas eliminar todos los archivos descargados?\n\n"
        )
        if not resp:
            return

        try:
            import os
            import shutil

            eliminados = 0
            for file in origen.iterdir():
                if file.is_file():
                    os.remove(file)
                    eliminados += 1
                elif file.is_dir():
                    shutil.rmtree(file)
                    eliminados += 1

            messagebox.showinfo(
                "Carpeta vacía",
                f"Se eliminaron {eliminados} elementos correctamente."
            )

            # Actualizar estado del botón Exportar
            self._actualizar_boton_exportar()

            self.status.set("Carpeta de descargas vaciada.")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo vaciar la carpeta:\n{e}")

    
    
    def limpiar(self):
        self.ruta_archivo.set("")
        self.columna_detectada.set("(sin cargar)")
        self.facturas = []
        self.refrescar_listas()
        self.status.set("Listo.")


if __name__ == "__main__":
    app = App()
    app.mainloop()
