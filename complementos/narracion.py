"""
MÓDULO: narracion.py
DESCRIPCION: Clase que encapsula el procesamiento de archivos de Control Diario (Narración).
             Extrae columnas, compara contra CMU y agrega los resultados al archivo unificado.
VERSION: 1.0.0
AUTOR: JANKOWICZ ALEXIS
FECHA: 2026-06-24

HISTORIAL DE CAMBIOS:
---------------------
[2026-06-24] v1.0.0: Extracción desde principal_eventos.pyw.
                     Conversión de funciones sueltas a Clase (Objeto) NarracionProcessor.
"""

# ─────────────────────────────────────────────
#  CONCEPTO CLAVE: ¿Qué es una Clase / Objeto?
# ─────────────────────────────────────────────
#
#  Una CLASE es como un molde o plantilla.
#  Un OBJETO es la "copia concreta" creada a partir de ese molde.
#
#  Ventajas de agrupar las funciones de narración en una clase:
#    • Todo el estado (datos_dir, función de log) vive junto → no hay variables globales.
#    • Se puede reutilizar: basta instanciar NarracionProcessor() desde cualquier script.
#    • Es fácil de testear de forma aislada.
#    • Sigue el patrón ya usado por ContraventoresProcessor, IntegradosProcessor, etc.
# ─────────────────────────────────────────────

import pandas as pd
from pathlib import Path
from typing import Optional, Callable, Tuple
from tkinter import simpledialog, messagebox


class NarracionProcessor:
    """
    Procesa archivos de Control Diario (Narración) y los integra
    al archivo unificado de eventos.

    Parámetros del constructor
    --------------------------
    log_fn : Callable[[str], None], opcional
        Función que recibe un string y lo muestra en el log visual.
        Por defecto se usa `print` (útil para pruebas o consola).
    datos_dir : Path | str, opcional
        Ruta a la carpeta DATOS donde se buscan los archivos CMU.
        Si no se pasa, se usa la carpeta 'DATOS' relativa a este módulo.

    Uso básico
    ----------
    >>> proc = NarracionProcessor(log_fn=log_mensaje, datos_dir=DATOS_DIR)
    >>> ok, msg = proc.procesar_narracion_file(ruta_narr, mes, ruta_unificado)
    """

    # ── Constructor ──────────────────────────────────────────────────────────
    def __init__(
        self,
        log_fn: Optional[Callable[[str], None]] = None,
        datos_dir: Optional[Path] = None,
    ):
        # Si no se pasa una función de log, usamos print como fallback.
        # Esto evita depender de variables globales del script principal.
        self.log = log_fn if log_fn is not None else print

        # Ruta a la carpeta donde están los CMU.
        # __file__ es la ruta de ESTE archivo (narracion.py).
        # .parent sube un nivel → carpeta "complementos".
        # .parent nuevamente → carpeta raíz del proyecto.
        if datos_dir is not None:
            self.datos_dir = Path(datos_dir)
        else:
            self.datos_dir = Path(__file__).parent.parent / "DATOS"

    # ── Métodos privados (auxiliares) ────────────────────────────────────────
    # La convención en Python es usar un guion bajo "_" al inicio para indicar
    # que el método es de uso interno de la clase.

    def _extraer_mes_de_nombre(self, nombre: str) -> Optional[str]:
        """
        Detecta el nombre del mes dentro de un nombre de archivo.

        Ejemplo: "CMU - MAYO.xlsx"  →  "mayo"
        """
        meses = [
            "enero",
            "febrero",
            "marzo",
            "abril",
            "mayo",
            "junio",
            "julio",
            "agosto",
            "septiembre",
            "octubre",
            "noviembre",
            "diciembre",
        ]
        ln = nombre.lower()
        for m in meses:
            if m in ln:
                return m
        return None

    def _agregar_narracion_a_unificado(
        self, ruta_unificado: str, df_narracion: pd.DataFrame
    ) -> Tuple[bool, str]:
        """
        Agrega el DataFrame de narración al archivo unificado existente.

        Retorna
        -------
        (True, mensaje_ok) o (False, mensaje_error)
        """
        try:
            df_existente = pd.read_excel(ruta_unificado, engine="openpyxl")
            df_combinado = pd.concat([df_existente, df_narracion], ignore_index=True)

            with pd.ExcelWriter(ruta_unificado, engine="openpyxl") as writer:
                df_combinado.to_excel(writer, sheet_name="eventos", index=False)

            return True, f"Narración agregada exitosamente a {ruta_unificado}"
        except Exception as e:
            return False, f"Error al agregar narración: {str(e)}"

    # ── Método público principal ──────────────────────────────────────────────
    # Los métodos SIN guion bajo son la "interfaz pública": los que llaman
    # desde afuera de la clase.

    def procesar_narracion_file(
        self,
        ruta_archivo_narracion: str,
        mes_cmu: Optional[str] = None,
        ruta_unificado: Optional[str] = None,
    ) -> Tuple[bool, str]:
        """
        Procesa un archivo de narración completo:
          1. Abre el Excel y pide al usuario que elija la hoja.
          2. Limpia y extrae columnas OPERADOR, GAP, SAE.
          3. Hace merge contra el archivo CMU correspondiente al mes.
          4. Agrega el resultado al archivo unificado (o lo guarda por separado).

        Parámetros
        ----------
        ruta_archivo_narracion : str
            Ruta al archivo Excel de Control Diario.
        mes_cmu : str, opcional
            Nombre del mes para buscar el CMU ("mayo", "junio", etc.).
        ruta_unificado : str, opcional
            Ruta al archivo eventos_unificados.xlsx para agregar los datos.

        Retorna
        -------
        (True, mensaje) si todo fue bien.
        (False, mensaje_error) si hubo algún problema.
        """
        # ── PASO 1: Abrir el archivo ──────────────────────────────────────────
        self.log(
            "\n═══════════════════════════════════════════════════════════════════════"
        )
        self.log("📋 PROCESANDO ARCHIVO DE NARRACIÓN")
        self.log(
            "═══════════════════════════════════════════════════════════════════════"
        )
        self.log(f"Archivo: {Path(ruta_archivo_narracion).name}")

        try:
            xls = pd.ExcelFile(ruta_archivo_narracion, engine="openpyxl")
            hojas = xls.sheet_names
            self.log("✓ Archivo abierto exitosamente")
            self.log(f"  Hojas disponibles: {len(hojas)}")
        except Exception as e:
            messagebox.showerror(
                "Error", f"No se pudo abrir el archivo de narración: {e}"
            )
            return False, str(e)

        # ── PASO 2: Pedir hoja al usuario ─────────────────────────────────────
        hoja_idx = simpledialog.askinteger(
            "Seleccionar hoja",
            "Seleccione una hoja:\n\n"
            + "\n".join(f"{i} - {h}" for i, h in enumerate(hojas)),
            minvalue=0,
            maxvalue=len(hojas) - 1,
        )

        if hoja_idx is None:
            return False, "Selección de hoja cancelada"

        self.log(f"✓ Hoja seleccionada: {hoja_idx} - {hojas[hoja_idx]}")

        try:
            df_n = pd.read_excel(
                ruta_archivo_narracion, sheet_name=hoja_idx, engine="openpyxl"
            )
            self.log("✓ Hoja leída exitosamente")
            self.log(f"  Dimensiones: {df_n.shape[0]} filas × {df_n.shape[1]} columnas")
        except Exception as e:
            return False, f"No se pudo leer la hoja seleccionada: {e}"

        # ── PASO 3: Limpiar GAP y SAE ─────────────────────────────────────────
        self.log("→ Limpiando columnas GAP y SAE...")
        for col in (3, 4):
            if col < len(df_n.columns):
                col_name = df_n.columns[col]
                df_n[col_name] = (
                    df_n.iloc[:, col]
                    .fillna("")
                    .astype(str)
                    .str.split("/", n=1)
                    .str[0]
                    .str.strip()
                )
        self.log("✓ Columnas GAP/SAE limpiadas")

        # ── PASO 4: Validar columnas mínimas ──────────────────────────────────
        columnas_a_extraer = [2, 3, 4]
        if df_n.shape[1] <= max(columnas_a_extraer):
            return (
                False,
                f"La hoja no tiene las columnas requeridas (necesita al menos {max(columnas_a_extraer) + 1}).",
            )

        # ── PASO 5: Filtrar filas válidas ─────────────────────────────────────
        # Condición 1: columna 5 debe decir "si"
        condicion_si = (
            df_n.iloc[:, 5].fillna("").astype(str).str.strip().str.lower() == "si"
        )
        # Condición 2: columna 2 (OPERADOR) no debe estar vacía
        condicion_no_vacia = df_n.iloc[:, 2].fillna("").astype(str).str.strip() != ""

        mask = condicion_si & condicion_no_vacia
        filas_cumplen = mask.sum()
        self.log("→ Aplicando filtros...")
        self.log(f"  Filas que cumplen condiciones: {filas_cumplen} de {len(df_n)}")

        df_eventos = (
            df_n.loc[mask].iloc[:, columnas_a_extraer].copy().reset_index(drop=True)
        )
        df_eventos.columns = ["OPERADOR", "GAP", "SAE"]
        self.log("✓ Columnas extraídas: OPERADOR, GAP, SAE")

        # ── PASO 6: Convertir SAE a numérico ──────────────────────────────────
        self.log("→ Convirtiendo SAE a numérico...")
        df_eventos["SAE"] = pd.to_numeric(df_eventos["SAE"], errors="coerce")

        filas_antes = len(df_eventos)
        df_eventos = df_eventos.dropna(subset=["SAE"])
        filas_eliminadas = filas_antes - len(df_eventos)
        if filas_eliminadas > 0:
            self.log(f"  Eliminadas {filas_eliminadas} filas con SAE inválido")

        df_eventos["SAE"] = df_eventos["SAE"].astype("Int64")
        self.log("✓ SAE convertido a formato numérico")
        self.log(f"  Filas restantes: {len(df_eventos)}")

        # ── PASO 7: Cargar archivo CMU ────────────────────────────────────────
        try:
            cmu_path = None

            if mes_cmu:
                candidate = self.datos_dir / f"CMU - {mes_cmu.upper()}.xlsx"
                if candidate.exists():
                    cmu_path = candidate
                    self.log(f"→ Buscando archivo CMU del mes: {mes_cmu.upper()}")
                    self.log(f"  Encontrado: {candidate.name}")

            if cmu_path is None:
                self.log("→ Buscando cualquier archivo CMU disponible...")
                matches = list(self.datos_dir.glob("CMU - *.xlsx"))
                if matches:
                    cmu_path = matches[0]
                    self.log(f"  Usando: {cmu_path.name}")
                else:
                    raise FileNotFoundError(
                        f"No se encontró archivo CMU en {self.datos_dir}"
                    )

            maestro = pd.read_excel(cmu_path, engine="openpyxl")
            self.log("✓ Archivo CMU cargado exitosamente")
            self.log(
                f"  Dimensiones: {maestro.shape[0]} filas * {maestro.shape[1]} columnas"
            )
        except Exception as e:
            return False, f"No se pudo abrir CMU: {e}"

        if maestro.shape[1] < 33:
            return (
                False,
                f"El archivo CMU no tiene suficientes columnas ({maestro.shape[1]}).",
            )

        # ── PASO 8: Preparar CMU y hacer merge ────────────────────────────────
        col_sae_cmu = maestro.columns[13]
        col_origen = maestro.columns[31]
        self.log("→ Identificando columnas en CMU...")
        self.log(f"  Columna SAE: {col_sae_cmu}")
        self.log(f"  Columna Origen: {col_origen}")

        self.log("→ Preparando SAE del CMU para merge...")
        maestro["SAE"] = pd.to_numeric(maestro[col_sae_cmu], errors="coerce")
        filas_antes_cmu = len(maestro)
        maestro = maestro.dropna(subset=["SAE"])
        filas_eliminadas_cmu = filas_antes_cmu - len(maestro)
        if filas_eliminadas_cmu > 0:
            self.log(
                f"  Eliminadas {filas_eliminadas_cmu} filas con SAE inválido en CMU"
            )

        maestro["SAE"] = maestro["SAE"].astype("Int64")

        if maestro["SAE"].duplicated().any():
            duplicados = maestro["SAE"].duplicated().sum()
            maestro = maestro.drop_duplicates(subset=["SAE"], keep="first")
            self.log(f"  Eliminados {duplicados} duplicados de SAE en CMU")

        self.log("→ Realizando merge entre narración y CMU...")
        try:
            resultado = df_eventos.merge(
                maestro[["SAE", col_origen]], on="SAE", how="inner", validate="m:1"
            )
        except Exception:
            resultado = df_eventos.merge(
                maestro[["SAE", col_origen]], on="SAE", how="inner"
            )

        self.log("✓ Merge completado")
        self.log(f"  Coincidencias encontradas: {len(resultado)} de {len(df_eventos)}")

        if resultado.empty:
            return (
                False,
                "No se encontraron coincidencias entre SAE(narracion) y SAE(CMU).",
            )

        # ── PASO 9: Estructurar resultado ─────────────────────────────────────
        resultado = resultado.rename(columns={col_origen: "ORIGEN"})
        resultado["FECHA"] = ""
        resultado["HORA"] = ""
        resultado["TIPO DE EVENTO"] = "NARRACION COMPLETA DE EVENTO"
        resultado["BARRIO"] = "ADICIONAL"
        resultado["GAP"] = resultado["GAP"].astype("Int64")

        column_order = [
            "FECHA",
            "HORA",
            "OPERADOR",
            "TIPO DE EVENTO",
            "BARRIO",
            "ORIGEN",
            "GAP",
            "SAE",
        ]
        resultado = resultado[column_order]

        self.log("✓ Estructura de datos preparada")
        self.log(f"  Columnas: {', '.join(resultado.columns)}")

        # ── PASO 10: Guardar resultado ────────────────────────────────────────
        if ruta_unificado and Path(ruta_unificado).exists():
            self.log("→ Agregando narración al archivo unificado existente...")
            ok, msg = self._agregar_narracion_a_unificado(ruta_unificado, resultado)
            if ok:
                self.log(
                    "═══════════════════════════════════════════════════════════════════════"
                )
                self.log("✅ NARRACIÓN PROCESADA Y AGREGADA AL ARCHIVO UNIFICADO")
                self.log(
                    "═══════════════════════════════════════════════════════════════════════"
                )
            return ok, msg
        else:
            # Guardar individualmente si no hay unificado disponible
            salida = Path(ruta_archivo_narracion).parent / "narracion_origen.xlsx"
            resultado.to_excel(salida, index=False)
            self.log(
                "═══════════════════════════════════════════════════════════════════════"
            )
            self.log("✅ ARCHIVO DE NARRACIÓN GENERADO")
            self.log(
                "═══════════════════════════════════════════════════════════════════════"
            )
            return True, f"Archivo generado: {salida}"

    # ── Método público auxiliar ───────────────────────────────────────────────
    # También exponemos extraer_mes_de_nombre como método público
    # para que el script principal pueda usarlo sin duplicar lógica.

    def extraer_mes_de_nombre(self, nombre: str) -> Optional[str]:
        """
        Detecta el nombre del mes dentro del nombre de un archivo.

        Ejemplo
        -------
        >>> proc.extraer_mes_de_nombre("ADIA - MAYO 2026.xlsx")
        'mayo'
        """
        return self._extraer_mes_de_nombre(nombre)
