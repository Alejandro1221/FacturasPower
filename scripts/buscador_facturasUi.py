import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

from buscar_facturas import buscar as buscar_en_sharepoint

APP_TITLE = "Buscador de Facturas Garantías"

# ---------- utilidades ----------
def detectar_columna_factura(df: pd.DataFrame):
    for c in df.columns:
        if "factura" in str(c).lower():
            return c
    return None

def leer_dataframe_robusto(ruta: str) -> pd.DataFrame:
    low = ruta.lower()
    if low.endswith(".ods"):
        return pd.read_excel(ruta, engine="odf", dtype=str)

    if low.endswith(".csv"):
        # detectar separador rápido y probar varios encodings
        with open(ruta, "rb") as fh:
            head = fh.read(4096)
        sample = head.decode("latin-1", errors="ignore")
        sep_guess = ";" if sample.count(";") > sample.count(",") else ","

        encs = ["utf-8", "utf-8-sig", "latin-1", "cp1252"]
        last = None
        for enc in encs:
            try:
                return pd.read_csv(ruta, sep=sep_guess, engine="python", encoding=enc, dtype=str)
            except Exception as e:
                last = e
        raise last or Exception("No se pudo leer el CSV con los encodings probados.")

    return pd.read_excel(ruta, dtype=str)

# ---------- app ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1000x600")
        self.minsize(900, 520)

        self.ruta_archivo = tk.StringVar()
        self.columna_detectada = tk.StringVar(value="(sin cargar)")
        self.facturas = []

        self._build_ui()

    def _build_ui(self):
        # top bar
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        ttk.Label(top, text="Tabla excel").pack(side="left")
        ttk.Entry(top, textvariable=self.ruta_archivo).pack(side="left", padx=8, fill="x", expand=True)
        ttk.Button(top, text="Buscar", command=self.elegir_archivo).pack(side="left")

        sub = ttk.Frame(self, padding=(10, 0))
        sub.pack(fill="x")
        ttk.Label(sub, text="Columna detectada: ").pack(side="left")
        ttk.Label(sub, textvariable=self.columna_detectada, foreground="#0a7").pack(side="left")

        # 3 columnas (lista / encontradas / no encontradas)
        body = ttk.Frame(self, padding=10)
        body.pack(fill="both", expand=True)

        # helpers para crear listbox con scroll
        def make_listbox(parent, title):
            col = ttk.Frame(parent)
            ttk.Label(col, text=title, font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 6))
            wrapper = ttk.Frame(col)
            wrapper.pack(fill="both", expand=True)
            lb = tk.Listbox(wrapper)
            sb = ttk.Scrollbar(wrapper, orient="vertical", command=lb.yview)
            lb.configure(yscrollcommand=sb.set)
            lb.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")
            return col, lb

        left_col, self.lb_lista = make_listbox(body, "Lista Facturas")
        mid_col,  self.lb_ok    = make_listbox(body, "Facturas encontradas")
        right_col, self.lb_nok  = make_listbox(body, "Facturas NO encontradas")

        left_col.grid(row=0, column=0, sticky="nsew", padx=(0,10))
        mid_col.grid(row=0, column=1, sticky="nsew", padx=10)
        right_col.grid(row=0, column=2, sticky="nsew", padx=(10,0))

        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.columnconfigure(2, weight=1)
        body.rowconfigure(0, weight=1)

        # barra inferior de acciones
        bottom = ttk.Frame(self, padding=10)
        bottom.pack(fill="x")

        self.btn_limpiar = ttk.Button(bottom, text="Limpiar lista", command=self.limpiar)
        self.btn_limpiar.pack(side="left")

        # placeholders (no funcionales por ahora)
        self.btn_buscar_sp = ttk.Button(bottom, text="Buscar en SharePoint", command=self._buscar_sharepoint)
        self.btn_descargar = ttk.Button(bottom, text="Facturas descargadas (próx.)", state=tk.DISABLED)

        self.btn_descargar.pack(side="right")
        self.btn_buscar_sp.pack(side="right", padx=8)

        # status
        self.status = tk.StringVar(value="Listo.")
        ttk.Label(self, textvariable=self.status, padding=(10, 6), anchor="w").pack(fill="x")

    # ------- acciones -------
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
        except Exception as e:
            messagebox.showerror("Error al leer archivo", str(e))
            self.status.set("Error al procesar el archivo.")

    def _buscar_sharepoint(self):
        if not self.facturas:
            messagebox.showinfo("Sin facturas", "Primero carga un archivo con facturas.")
            return

        self.status.set("Buscando en SharePoint...")
        self.update_idletasks()
        self.btn_buscar_sp.config(state=tk.DISABLED)

        try:
            resultado = buscar_en_sharepoint(self.facturas)
            encontradas = resultado.get("encontradas", [])
            no_encontradas = resultado.get("no_encontradas", [])

            self.lb_ok.delete(0, "end")
            self.lb_nok.delete(0, "end")

            for f in encontradas:
                self.lb_ok.insert("end", f)

            for f in no_encontradas:
                self.lb_nok.insert("end", f)

            total = len(encontradas) + len(no_encontradas)
            self.status.set(f"Resultado: {len(encontradas)} encontradas, {len(no_encontradas)} no encontradas (de {total}).")

        except Exception as e:
            messagebox.showerror("Error al buscar", str(e))
            self.status.set("Error durante la búsqueda.")
        finally:
            self.btn_buscar_sp.config(state=tk.NORMAL)

    def refrescar_listas(self):
        # Llena SOLO la primera lista. Las otras 2 quedan vacías 
        self.lb_lista.delete(0, "end")
        for f in self.facturas[:1000]:
            self.lb_lista.insert("end", f)

        self.lb_ok.delete(0, "end")
        self.lb_nok.delete(0, "end")

    def limpiar(self):
        self.ruta_archivo.set("")
        self.columna_detectada.set("(sin cargar)")
        self.facturas = []
        self.refrescar_listas()
        self.status.set("Listo.")

if __name__ == "__main__":
    app = App()
    app.mainloop()
