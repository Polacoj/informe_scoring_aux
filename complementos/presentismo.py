#!/Users/alexisjankowicz/Python/informe_scoring/.venv/bin/activate   python

import re
import tkinter as tk
import traceback
import unicodedata
from pathlib import Path as _Path
from tkinter import filedialog, messagebox, scrolledtext
from typing import Callable, Optional

import pandas as pd

# Obtener el directorio del script actual y crear carpeta DATOS si no existe
SCRIPT_DIR = _Path(__file__).parent
DATOS_DIR = SCRIPT_DIR / "DATOS"
DATOS_DIR.mkdir(exist_ok=True)

# Rutas de los archivos de salida dentro de la carpeta DATOS
output = DATOS_DIR / "filtrado_presentismo_grupos.xlsx"
salida = DATOS_DIR / "presentismo_resumen.xlsx"

ruta = []
ventana_log = None

# output = "filtrado_presentismo_grupos.xlsx"
# salida = "presentismo_resumen.xlsx"
hojas_a_procesar: list[int] | None = None  # None procesa todas las hojas del archivo
COLUMNAS_MIN = 42

# CONFIG: valores por defecto y lista de meses con 31 días
dias = 30
d = 34 if dias == 30 else 35
meses_31 = ["enero", "marzo", "mayo", "julio", "agosto", "octubre", "diciembre"]


# funccion para normalizar texto (eliminar acentos, pasar a minusculas, etc)
def normalizar_texto(s):
    if pd.isna(s):
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = s.encode("ascii", errors="ignore").decode("utf-8")
    return s.strip().lower()


def map_zona_text(v):
    if pd.isna(v):
        return ""
    s = normalizar_texto(v)
    if s == "":
        return ""
    if s in ("lxuc", "lxm", "lxp", "fac", "lxf", "art"):
        return "LM"
    if s in ("san", "a"):
        return "AUS"
    return str(v).strip().upper()


# funcion para normalizar el nombre del grupo
# Solo se usa para la hoja 12X60, donde los grupos ADM se expresan con números 1-6
def normalizar_grupo(g_raw):
    if not isinstance(g_raw, str):
        return "GRUPO DESCONOCIDO"
    g = normalizar_texto(g_raw)
    m = re.search(r"\b(?:grupo\s*|g\s*)([1-6])\b", g, re.IGNORECASE)
    if m:
        return "GRUPO ADM."
    return "GRUPO DESCONOCIDO"


# funcion para log visual y consola
def log_mensaje(msg):
    """Añade mensaje al log visual"""
    global ventana_log
    if ventana_log:
        ventana_log.config(state=tk.NORMAL)
        ventana_log.insert(tk.END, msg + "\n")
        ventana_log.see(tk.END)
        ventana_log.config(state=tk.DISABLED)
        ventana_log.update()
    print(msg)


# Regex para detectar grupos y días administrativos
# Capture grupos A-F (con o sin la palabra 'GRUPO') y grupos numéricos 1-6
patron_grupo = re.compile(
    r"\b(?:grupo\s*([A-F])|g\s*([1-6])|([A-F]))\b",
    re.IGNORECASE,
)
patron_adm = re.compile(
    r"\b(lun(\.|es)?|mar(\.|tes)?|mie(\.|rcoles)?|jue(\.|ves)?|"
    r"vier(\.|nes)?|sab(\.|ado|ados)?|dom(\.|ingo)?|feriad(o|os)?)(\s*\d{1,2})?\b",
    re.IGNORECASE,
)


# Funcion principal de procesamiento
def procesar_presentismo(
    archivo_origen,
    carpeta_salida=None,
    log_fn=None,
    mostrar_dialogos=True,
):
    """Función principal de procesamiento"""
    log = log_fn if log_fn is not None else log_mensaje
    carpeta_salida = _Path(carpeta_salida) if carpeta_salida else DATOS_DIR
    carpeta_salida.mkdir(parents=True, exist_ok=True)
    ruta_output = carpeta_salida / "filtrado_presentismo_grupos.xlsx"
    ruta_salida = carpeta_salida / "presentismo_resumen.xlsx"

    try:
        log("\n[INICIANDO PROCESO]")
        log(f"Archivo: {archivo_origen}")
        log(f"Carpeta de salida: {carpeta_salida}")

        xls = pd.ExcelFile(archivo_origen)
        df_total = pd.DataFrame()

        if hojas_a_procesar is None:
            hojas = range(len(xls.sheet_names))
        else:
            hojas = [
                i
                for i in hojas_a_procesar
                if isinstance(i, int) and i < len(xls.sheet_names)
            ]

        for i in hojas:
            nombre_hoja = xls.sheet_names[i]
            log(f"\n→ Procesando hoja: {nombre_hoja}")

            df = pd.read_excel(archivo_origen, sheet_name=nombre_hoja)

            if str(nombre_hoja).strip().upper() == "CMU VILLA LUGANO":
                df = df.iloc[:, 1:]

            if str(nombre_hoja).strip().upper() == "CMU PUERTO MADERO":
                df = df.iloc[:, 1:]

            if i != 0:
                df.insert(3, "col_extra_P", "ext.")

            if df.shape[1] < COLUMNAS_MIN:
                faltan = COLUMNAS_MIN - df.shape[1]
                for k in range(faltan):
                    df[f"extra_{k}"] = "P"

            if df.shape[1] > COLUMNAS_MIN:
                df = df.iloc[:, :COLUMNAS_MIN]

            df.columns = list(range(COLUMNAS_MIN))
            columnas_a_borrar = [3, 4, 5] + list(range(37, 42))
            df = df.drop(columns=columnas_a_borrar, errors="ignore")

            filas_a_agregar = []
            grupo_actual = None
            modo_adm = False

            for idx, fila in df.iterrows():
                txt = normalizar_texto(fila.iloc[0])

                if "nivel" in txt:
                    fila = fila.copy()
                    fila.iloc[0] = "AUX."
                    etiqueta = str(nombre_hoja) + (" ADM" if modo_adm else "")
                    etiqueta_final = (
                        f"{etiqueta} - {grupo_actual or 'GRUPO DESCONOCIDO'}"
                    )
                    nueva_fila = fila.copy()
                    nueva_fila["HOJA_ORIGEN"] = etiqueta_final
                    filas_a_agregar.append(nueva_fila)
                    continue

                if patron_adm.search(txt):
                    modo_adm = True
                    continue

                m = patron_grupo.search(txt)
                if m:
                    if str(nombre_hoja).strip().upper() == "12X60":
                        grupo_actual = normalizar_grupo(m.group(0))
                    else:
                        grupo_letra = m.group(1) or m.group(3)
                        if grupo_letra:
                            grupo_actual = f"GRUPO {grupo_letra.upper()}"
                        else:
                            grupo_actual = str(m.group(0)).strip().upper()
                    modo_adm = False
                    continue

                if "contratado" in txt:
                    etiqueta = str(nombre_hoja) + (" ADM" if modo_adm else "")
                    etiqueta_final = (
                        f"{etiqueta} - {grupo_actual or 'GRUPO DESCONOCIDO'}"
                    )
                    nueva_fila = fila.copy()
                    nueva_fila["HOJA_ORIGEN"] = etiqueta_final
                    filas_a_agregar.append(nueva_fila)

            if filas_a_agregar:
                df_filtrado = pd.DataFrame(filas_a_agregar)
                cols = ["HOJA_ORIGEN"] + [
                    c for c in df_filtrado.columns if c != "HOJA_ORIGEN"
                ]
                df_filtrado = df_filtrado[cols]

                cols_a_mapear = list(range(4, 37))
                for c in cols_a_mapear:
                    if c in df_filtrado.columns:
                        df_filtrado[c] = df_filtrado[c].apply(map_zona_text)

                zona = (
                    df_filtrado.iloc[:, 4:d]
                    .astype(str)
                    .apply(lambda x: x.str.upper().str.strip())
                )
                df_filtrado["-LM-"] = zona.apply(lambda r: r.eq("LM").sum(), axis=1)
                df_filtrado["-LLT-"] = zona.apply(lambda r: r.eq("LLT").sum(), axis=1)
                df_filtrado["-AUS-"] = zona.apply(lambda r: r.eq("AUS").sum(), axis=1)
                df_filtrado["-LA-"] = zona.apply(lambda r: r.eq("LA").sum(), axis=1)

                df_total = pd.concat([df_total, df_filtrado], ignore_index=True)
                log(f"  ✓ {len(filas_a_agregar)} filas procesadas")

        df_final = df_total.reindex(
            [2, 1, 0, "HOJA_ORIGEN", "-LM-", "-LLT-", "-AUS-", "-LA-"], axis=1
        )
        df_final.rename(
            columns={
                0: "JERARQUIA",
                1: "NOMBRE",
                2: "LP/DNI",
                "HOJA_ORIGEN": "TURNO & GRUPO",
            },
            inplace=True,
        )

        df_total.to_excel(ruta_output, index=False)
        df_final.to_excel(ruta_salida, index=False)

        log(f"\n✅ EXPORTADO: {ruta_output}")
        log(f"✅ EXPORTADO: {ruta_salida}")
        if mostrar_dialogos:
            messagebox.showinfo(
                "Éxito", f"Archivos generados:\n{ruta_output}\n{ruta_salida}"
            )
        return True, str(ruta_output), str(ruta_salida)

    except Exception as e:
        log(f"\n❌ ERROR: {str(e)}")
        traceback.print_exc()
        if mostrar_dialogos:
            messagebox.showerror("Error", f"Error al procesar: {str(e)}")
        return False, f"Error al procesar: {str(e)}", None


class PresentismoProcessor:
    """Procesa presentismo y agrega su resumen al archivo unificado."""

    def __init__(self, log_fn: Optional[Callable[[str], None]] = None):
        self.log = log_fn if log_fn is not None else print

    @staticmethod
    def _resumen_a_eventos(df_resumen):
        """Convierte el resumen en tres bloques de eventos compatibles."""
        eventos = []
        for codigo, columna in (
            ("LICENCIA MEDICA 24 HS (Reiteradas en el mes)", "-LM-"),
            ("LLEGADA TARDE", "-LLT-"),
            ("AUSENTE", "-AUS-"),
        ):
            for _, fila in df_resumen.iterrows():
                eventos.append(
                    {
                        "FECHA": "",
                        "HORA": "",
                        "OPERADOR": fila.get("NOMBRE", ""),
                        "TIPO DE EVENTO": f"{codigo}",
                        "LIC. MED/LL.T/AUS.": fila.get(columna, 0),
                        "BARRIO": "PRESENTISMO",
                        "ORIGEN": "PRESENTISMO",
                        "GAP": "",
                        "SAE": "",
                    }
                )
        return pd.DataFrame(eventos)

    def procesar_presentismo_file(self, ruta_archivo, ruta_unificado):
        carpeta_salida = _Path(ruta_unificado).parent
        ok, ruta_filtrado, ruta_resumen = procesar_presentismo(
            ruta_archivo,
            carpeta_salida=carpeta_salida,
            log_fn=self.log,
            mostrar_dialogos=False,
        )
        if not ok:
            return False, ruta_filtrado

        try:
            df_existente = pd.read_excel(ruta_unificado, engine="openpyxl")
            df_resumen = pd.read_excel(ruta_resumen, engine="openpyxl")
            df_eventos_presentismo = self._resumen_a_eventos(df_resumen)
            df_combinado = pd.concat(
                [df_existente, df_eventos_presentismo], ignore_index=True, sort=False
            )
            with pd.ExcelWriter(ruta_unificado, engine="openpyxl") as writer:
                df_combinado.to_excel(writer, sheet_name="eventos", index=False)

            self.log(
                f"✅ Presentismo agregado a {ruta_unificado} "
                f"({len(df_eventos_presentismo)} eventos)"
            )
            return (
                True,
                "Presentismo procesado y agregado al archivo unificado.\n"
                f"Filtrado: {ruta_filtrado}\n"
                f"Resumen: {ruta_resumen}",
            )
        except Exception as e:
            return False, f"Error al agregar presentismo al archivo unificado: {e}"


# funciones de la interfaz gráfica
def seleccionar_archivos():
    """Seleccionar archivo excel de presentismo"""
    global ruta
    try:
        archivos = filedialog.askopenfilenames(
            title="Seleccionar archivos EXCEL",
            filetypes=[
                ("Archivos EXCEL", "*.XLSX *.XLSM"),
                ("Todos los archivos", "*"),
            ],
        )
        if archivos:
            ruta = archivos[0]  # Tomar el primer archivo
            log_mensaje(f"Archivo seleccionado: {ruta}")
            boton_ejecutar.config(state=tk.NORMAL)
    except Exception as e:
        log_mensaje(f"Error: {str(e)}")
        messagebox.showerror("Error", f"Error al seleccionar: {str(e)}")


def ejecutar_procesamiento():
    global dias, d
    if not ruta:
        messagebox.showwarning(
            "Advertencia", "Por favor, selecciona un archivo primero"
        )
        return

    nombre = str(ruta).lower()

    # Si el nombre de archivo contiene alguno de los meses con 31 días, ajusta 'd'
    if any(m in nombre for m in meses_31):
        dias = 31
        d = 35
    elif "febrero" in nombre:
        dias = 28
        d = 32
    else:
        dias = 30
        d = 34

    procesar_presentismo(ruta)


def iniciar_interfaz():
    """Inicia la interfaz independiente del complemento."""
    global boton_ejecutar, ventana_log

    root = tk.Tk()
    root.title("Analizar y resumir PRESENTISMO")
    root.config(bg="#0E8EA4")
    root.geometry("900x600")

    frame_botones = tk.Frame(root, bg="#0E8EA4")
    frame_botones.pack(fill=tk.X, padx=10, pady=10)

    boton_seleccionar = tk.Button(
        frame_botones,
        text="📁 Seleccionar Archivos",
        width=20,
        height=2,
        command=seleccionar_archivos,
        bg="blue",
        fg="white",
        font=("Helvetica", 12, "bold"),
    )
    boton_seleccionar.pack(side=tk.LEFT, padx=5)

    boton_ejecutar = tk.Button(
        frame_botones,
        text="👍🏻 Ejecutar",
        width=20,
        height=2,
        command=ejecutar_procesamiento,
        bg="#4CAF50",
        fg="white",
        font=("Helvetica", 12, "bold"),
        state=tk.DISABLED,
    )
    boton_ejecutar.pack(side=tk.LEFT, padx=5)

    boton_cerrar = tk.Button(
        frame_botones,
        text="❌ Salir",
        width=20,
        height=2,
        command=root.quit,
        bg="#541b17",
        fg="white",
        font=("Helvetica", 14, "bold"),
    )
    boton_cerrar.pack(side=tk.RIGHT, padx=5)

    frame_log = tk.Frame(root, bg="#0D5B8C")
    frame_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    tk.Label(
        frame_log,
        text="ESTADO DE PROCESAMIENTO",
        bg="#0D5B8C",
        fg="white",
        font=("Helvetica", 10, "bold"),
    ).pack(anchor=tk.W, pady=5)

    ventana_log = scrolledtext.ScrolledText(
        frame_log,
        height=20,
        bg="#1a1a1a",
        fg="#00FF00",
        font=("Courier", 11),
        insertbackground="#00FF00",
        wrap=tk.WORD,
    )
    ventana_log.pack(fill=tk.BOTH, expand=True)
    ventana_log.config(state=tk.DISABLED)
    log_mensaje(
        "Sistema en espera...... \n\nSeleccione un archivo XLSX de PRESENTISMO para comenzar."
    )
    root.mainloop()


if __name__ == "__main__":
    iniciar_interfaz()
