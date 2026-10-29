"""
cruzar_colaboradores.py
=======================
GUI con PyQt6 para cruzar dos archivos Excel por columna SAE
e insertar COLABORADOR 1..5 entre columnas 7 y 8 del Archivo 1.

Dependencias:
    pip install pandas openpyxl PyQt6
"""

import sys
from pathlib import Path

import pandas as pd
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

# ══════════════════════════════════════════════════════════════════
#  COLORES Y ESTILOS
# ══════════════════════════════════════════════════════════════════
COLOR_FONDO = "#1E1E2E"
COLOR_PANEL = "#2A2A3E"
COLOR_BORDE = "#3D3D5C"
COLOR_ACENTO = "#7C6AF7"  # violeta
COLOR_ACENTO_HOVER = "#9B8FFB"
COLOR_TEXTO = "#E0E0F0"
COLOR_TEXTO_SUAVE = "#9090B0"
COLOR_EXITO = "#4CAF82"
COLOR_ERROR = "#F7706A"
COLOR_ADVERTENCIA = "#F7C26A"
COLOR_HEADER_TABLA = "#3A2A6E"
COLOR_FILA_ALT = "#252535"

ESTILO_GLOBAL = f"""
    QMainWindow, QWidget {{
        background-color: {COLOR_FONDO};
        color: {COLOR_TEXTO};
        font-family: 'Segoe UI', 'Inter', sans-serif;
        font-size: 13px;
    }}
    QGroupBox {{
        border: 1px solid {COLOR_BORDE};
        border-radius: 8px;
        margin-top: 12px;
        padding: 12px;
        font-weight: 600;
        color: {COLOR_TEXTO_SUAVE};
    }}
    QGroupBox::title {{
        subcontrol-origin: margin;
        left: 12px;
        padding: 0 6px;
        color: {COLOR_ACENTO};
    }}
    QLineEdit {{
        background-color: {COLOR_PANEL};
        border: 1px solid {COLOR_BORDE};
        border-radius: 6px;
        padding: 7px 10px;
        color: {COLOR_TEXTO};
        font-size: 12px;
    }}
    QLineEdit:focus {{
        border: 1px solid {COLOR_ACENTO};
    }}
    QPushButton {{
        background-color: {COLOR_PANEL};
        border: 1px solid {COLOR_BORDE};
        border-radius: 6px;
        padding: 8px 16px;
        color: {COLOR_TEXTO};
        font-size: 12px;
    }}
    QPushButton:hover {{
        background-color: {COLOR_BORDE};
        border: 1px solid {COLOR_ACENTO};
    }}
    QPushButton#btn_procesar {{
        background-color: {COLOR_ACENTO};
        border: none;
        color: white;
        font-weight: 700;
        font-size: 14px;
        padding: 12px 24px;
        border-radius: 8px;
        min-width: 200px;
    }}
    QPushButton#btn_procesar:hover {{
        background-color: {COLOR_ACENTO_HOVER};
    }}
    QPushButton#btn_procesar:disabled {{
        background-color: {COLOR_BORDE};
        color: {COLOR_TEXTO_SUAVE};
    }}
    QProgressBar {{
        border: 1px solid {COLOR_BORDE};
        border-radius: 6px;
        background-color: {COLOR_PANEL};
        height: 14px;
        text-align: center;
        color: white;
        font-size: 11px;
    }}
    QProgressBar::chunk {{
        background-color: {COLOR_ACENTO};
        border-radius: 5px;
    }}
    QLabel#lbl_titulo {{
        font-size: 20px;
        font-weight: 700;
        color: {COLOR_TEXTO};
    }}
    QLabel#lbl_subtitulo {{
        font-size: 12px;
        color: {COLOR_TEXTO_SUAVE};
    }}
    QTableWidget {{
        background-color: {COLOR_PANEL};
        border: 1px solid {COLOR_BORDE};
        border-radius: 6px;
        gridline-color: {COLOR_BORDE};
        color: {COLOR_TEXTO};
        font-size: 11px;
    }}
    QHeaderView::section {{
        background-color: {COLOR_HEADER_TABLA};
        color: {COLOR_TEXTO};
        padding: 6px 8px;
        border: none;
        border-right: 1px solid {COLOR_BORDE};
        font-weight: 600;
        font-size: 11px;
    }}
    QTableWidget::item:alternate {{
        background-color: {COLOR_FILA_ALT};
    }}
    QScrollBar:vertical {{
        background: {COLOR_PANEL};
        width: 8px;
        border-radius: 4px;
    }}
    QScrollBar::handle:vertical {{
        background: {COLOR_BORDE};
        border-radius: 4px;
    }}
"""


def seleccionar_hoja_excel(hojas_disponibles: list[str], indice: int) -> str:
    """Devuelve la hoja seleccionada por número (1-based) desde la lista disponible."""
    if not hojas_disponibles:
        raise ValueError("No hay hojas disponibles para seleccionar.")
    if indice < 1 or indice > len(hojas_disponibles):
        raise IndexError("El número de hoja está fuera del rango disponible.")
    return hojas_disponibles[indice - 1]


# ══════════════════════════════════════════════════════════════════
#  HILO DE PROCESAMIENTO (evita que la GUI se congele)
# ══════════════════════════════════════════════════════════════════
class HiloProcesamiento(QThread):
    """
    QThread: ejecuta el procesamiento Excel en segundo plano.
    Emite señales para actualizar la GUI sin bloquearla.
    """

    progreso = pyqtSignal(int, str)  # (valor 0-100, mensaje)
    terminado = pyqtSignal(dict)  # estadísticas finales
    error = pyqtSignal(str)  # mensaje de error

    # Nombre de columna en Archivo 1
    COL_SAE_A1 = "SAE"

    # Nombres de columnas en Archivo 2
    COL_SAE_A2 = "SAE"
    COL_COLABORADOR_1 = "COLABORADORES 1"
    COL_COLABORADOR_2 = "COLABORADORES 2"
    COL_COLABORADOR_3 = "COLABORADORES 3"
    COL_COLABORADOR_4 = "COLABORADORES 4"
    COL_COLABORADOR_5 = "COLABORADORES 5"

    def __init__(self, ruta1, ruta2, ruta_salida, hoja_archivo2):
        super().__init__()
        self.ruta1 = ruta1
        self.ruta2 = ruta2
        self.ruta_salida = ruta_salida
        self.hoja_archivo2 = hoja_archivo2

    def run(self):
        try:
            # ── Paso 1: Leer archivos ──────────────────────────
            self.progreso.emit(10, "Leyendo Archivo 1...")
            df1 = pd.read_excel(
                self.ruta1,
                header=0,
                converters={
                    self.COL_SAE_A1: lambda v: str(v).strip() if not pd.isna(v) else v,
                },
            )
            # Convertir nombres de columna a str (Excel puede generar float
            # en columnas sin encabezado: 0.0, 1.0, etc.)
            df1.columns = [str(c).strip() for c in df1.columns]

            self.progreso.emit(25, f"Leyendo Archivo 2 (hoja: {self.hoja_archivo2})...")
            df2 = pd.read_excel(
                self.ruta2,
                sheet_name=self.hoja_archivo2,
                header=0,
                converters={
                    self.COL_SAE_A2: lambda v: str(v).strip() if not pd.isna(v) else v,
                },
            )
            df2.columns = [str(c).strip() for c in df2.columns]

            # ── Paso 2: Validar columnas ───────────────────────
            self.progreso.emit(35, "Validando columnas...")
            requeridas_a2 = [
                self.COL_SAE_A2,
                self.COL_COLABORADOR_1,
                self.COL_COLABORADOR_2,
                self.COL_COLABORADOR_3,
                self.COL_COLABORADOR_4,
                self.COL_COLABORADOR_5,
            ]
            for col in requeridas_a2:
                if col not in df2.columns:
                    raise KeyError(
                        f"Columna '{col}' no encontrada en Archivo 2.\n\n"
                        f"Columnas disponibles:\n  {', '.join(str(c) for c in df2.columns)}"
                    )
            if self.COL_SAE_A1 not in df1.columns:
                raise KeyError(
                    f"Columna 'SAE' no encontrada en Archivo 1.\n\n"
                    f"Columnas disponibles:\n  {', '.join(str(c) for c in df1.columns)}"
                )

            # ── Paso 3: Construir lookup ───────────────────────
            self.progreso.emit(50, "Construyendo tabla de cruce...")

            # Normalización robusta del SAE:
            # Excel puede leer números como float → "48291057.0"
            # Convertimos a float primero para quitar decimales, luego a str entero.
            # Si el valor no es numérico (texto puro) se conserva tal cual.
            def normalizar_sae(serie: pd.Series) -> pd.Series:
                def _limpiar(v):
                    if pd.isna(v):
                        return v
                    v = str(v).strip()
                    try:
                        # "48291057.0" → 48291057 → "48291057"
                        return str(int(float(v)))
                    except ValueError:
                        return v  # texto no numérico: devolver limpio

                return serie.map(_limpiar)

            df1[self.COL_SAE_A1] = normalizar_sae(df1[self.COL_SAE_A1])
            df2[self.COL_SAE_A2] = normalizar_sae(df2[self.COL_SAE_A2])

            lookup = (
                df2[
                    [
                        self.COL_SAE_A2,
                        self.COL_COLABORADOR_1,
                        self.COL_COLABORADOR_2,
                        self.COL_COLABORADOR_3,
                        self.COL_COLABORADOR_4,
                        self.COL_COLABORADOR_5,
                    ]
                ]
                .drop_duplicates(subset=self.COL_SAE_A2)
                .set_index(self.COL_SAE_A2)
            )

            # ── Paso 4: Mapear e insertar columnas ─────────────
            self.progreso.emit(65, "Mapeando colaboradores...")
            nuevas_cols = {
                "COLABORADOR 1": df1[self.COL_SAE_A1].map(
                    lookup[self.COL_COLABORADOR_1]
                ),
                "COLABORADOR 2": df1[self.COL_SAE_A1].map(
                    lookup[self.COL_COLABORADOR_2]
                ),
                "COLABORADOR 3": df1[self.COL_SAE_A1].map(
                    lookup[self.COL_COLABORADOR_3]
                ),
                "COLABORADOR 4": df1[self.COL_SAE_A1].map(
                    lookup[self.COL_COLABORADOR_4]
                ),
                "COLABORADOR 5": df1[self.COL_SAE_A1].map(
                    lookup[self.COL_COLABORADOR_5]
                ),
            }

            self.progreso.emit(78, "Insertando columnas entre posición 7 y 8...")
            mitad_izq = df1.iloc[:, :7]
            mitad_der = df1.iloc[:, 7:]
            nuevas_df = pd.DataFrame(nuevas_cols, index=df1.index)
            df1 = pd.concat([mitad_izq, nuevas_df, mitad_der], axis=1)

            # ── Paso 5: Guardar ────────────────────────────────
            self.progreso.emit(90, "Guardando archivo...")
            df1.to_excel(self.ruta_salida, index=False)

            # ── Estadísticas ───────────────────────────────────
            # int() fuerza tipos nativos de Python (no numpy.int64)
            # para que pyqtSignal(dict) los serialice sin error.
            total = int(len(df1))
            encontrados = int(
                df1["COLABORADOR 1"].notna().sum()
                + df1["COLABORADOR 2"].notna().sum()
                + df1["COLABORADOR 3"].notna().sum()
                + df1["COLABORADOR 4"].notna().sum()
                + df1["COLABORADOR 5"].notna().sum()
            )
            no_encontrados = int(total - encontrados)

            self.progreso.emit(100, "¡Proceso completado!")
            self.terminado.emit(
                {
                    "total": total,
                    "encontrados": encontrados,
                    "no_encontrados": no_encontrados,
                    "salida": self.ruta_salida,
                }
            )

        except Exception as e:
            self.error.emit(str(e))


# ══════════════════════════════════════════════════════════════════
#  WIDGET: SELECTOR DE ARCHIVO (reutilizable)
# ══════════════════════════════════════════════════════════════════
class SelectorArchivo(QWidget):
    """Fila con etiqueta + campo de texto + botón Examinar."""

    def __init__(self, etiqueta: str, placeholder: str = "", parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        lbl = QLabel(etiqueta)
        lbl.setFixedWidth(90)
        lbl.setStyleSheet(f"color: {COLOR_TEXTO_SUAVE}; font-size: 12px;")

        self.campo = QLineEdit()
        self.campo.setPlaceholderText(placeholder)
        self.campo.setReadOnly(True)

        self.btn = QPushButton("Examinar…")
        self.btn.setFixedWidth(100)

        layout.addWidget(lbl)
        layout.addWidget(self.campo, 1)
        layout.addWidget(self.btn)

    def ruta(self) -> str:
        return self.campo.text()

    def set_ruta(self, ruta: str):
        self.campo.setText(ruta)


# ══════════════════════════════════════════════════════════════════
#  VENTANA PRINCIPAL
# ══════════════════════════════════════════════════════════════════
class VentanaPrincipal(QMainWindow):
    def __init__(self):
        super().__init__()
        self.hilo = None
        self._construir_ui()

    # ── Construcción de la UI ────────────────────────────────
    def _construir_ui(self):
        self.setWindowTitle("Cruzar Colaboradores — PRODUCTIVIDAD GENERAL")
        self.setMinimumSize(800, 600)
        self.resize(600, 480)
        self.setStyleSheet(ESTILO_GLOBAL)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(24, 20, 24, 20)
        root.setSpacing(16)

        # ── Encabezado ────────────────────────────────────────
        titulo = QLabel("Cruzar Colaboradores")
        titulo.setObjectName("lbl_titulo")
        subtitulo = QLabel(
            "Inserta COLABORADOR 1 al 5 entre columnas 7 y 8 del Archivo 1, "
            "cruzando por SAE con el Archivo 2"
        )
        subtitulo.setObjectName("lbl_subtitulo")
        subtitulo.setWordWrap(True)
        root.addWidget(titulo)
        root.addWidget(subtitulo)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet(f"color: {COLOR_BORDE};")
        root.addWidget(sep)

        # ── Selección de archivos ─────────────────────────────
        grp_archivos = QGroupBox("Archivos de entrada")
        grp_layout = QVBoxLayout(grp_archivos)
        grp_layout.setSpacing(10)

        # Archivo 1
        self.sel_arch1 = SelectorArchivo(
            "Archivo 1",
            "Eventos diarios (encabezados en fila 1) — contiene columna SAE",
        )
        self.sel_arch1.btn.clicked.connect(
            lambda: self._abrir_archivo(self.sel_arch1, "Archivo 1")
        )
        grp_layout.addWidget(self.sel_arch1)

        # Archivo 2
        self.sel_arch2 = SelectorArchivo(
            "Archivo 2",
            "Colaboradores (encabezados en fila 2) — contiene SAE, COLABORADO, SISEP…",
        )
        self.sel_arch2.btn.clicked.connect(
            lambda: self._abrir_archivo(self.sel_arch2, "Archivo 2")
        )
        grp_layout.addWidget(self.sel_arch2)

        root.addWidget(grp_archivos)

        # ── Archivo de salida ─────────────────────────────────
        grp_salida = QGroupBox("Archivo de salida")
        grp_s_layout = QVBoxLayout(grp_salida)

        self.sel_salida = SelectorArchivo(
            "Guardar en", "Seleccioná dónde guardar el resultado (.xlsx)"
        )
        self.sel_salida.campo.setReadOnly(False)
        self.sel_salida.btn.clicked.connect(self._guardar_como)
        grp_s_layout.addWidget(self.sel_salida)
        root.addWidget(grp_salida)

        # ── Barra de progreso ─────────────────────────────────
        self.barra = QProgressBar()
        self.barra.setValue(0)
        self.barra.setTextVisible(True)
        self.barra.setFormat("En espera")
        self.barra.setFixedHeight(22)
        root.addWidget(self.barra)

        self.lbl_estado = QLabel("")
        self.lbl_estado.setStyleSheet(f"color: {COLOR_TEXTO_SUAVE}; font-size: 12px;")
        self.lbl_estado.setAlignment(Qt.AlignmentFlag.AlignCenter)
        root.addWidget(self.lbl_estado)

        # ── Botón procesar ────────────────────────────────────
        fila_btn = QHBoxLayout()
        fila_btn.addStretch()
        self.btn_procesar = QPushButton("▶  Procesar archivos")
        self.btn_procesar.setObjectName("btn_procesar")
        self.btn_procesar.clicked.connect(self._iniciar_procesamiento)
        fila_btn.addWidget(self.btn_procesar)
        fila_btn.addStretch()
        root.addLayout(fila_btn)
        root.addStretch()

    # ── Diálogos de archivo ───────────────────────────────────
    def _abrir_archivo(self, selector: SelectorArchivo, titulo: str):
        ruta, _ = QFileDialog.getOpenFileName(
            self, f"Seleccionar {titulo}", "", "Archivos Excel (*.xlsx *.xls *.xlsm)"
        )
        if ruta:
            selector.set_ruta(ruta)
            # Proponer salida automática al elegir Archivo 1
            if selector is self.sel_arch1 and not self.sel_salida.ruta():
                salida = str(Path(ruta).parent / "resultado_colaboradores.xlsx")
                self.sel_salida.set_ruta(salida)

    def _guardar_como(self):
        ruta, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar resultado como",
            self.sel_salida.ruta() or "",
            "Archivos Excel (*.xlsx)",
        )
        if ruta:
            if not ruta.endswith(".xlsx"):
                ruta += ".xlsx"
            self.sel_salida.set_ruta(ruta)

    def _seleccionar_hoja_archivo2(self, ruta2: str):
        try:
            excel_file = pd.ExcelFile(ruta2, engine="openpyxl")
            hojas = excel_file.sheet_names
        except Exception as exc:
            self._alerta(
                "No se pudo leer Archivo 2",
                f"No se pudo abrir el archivo Excel:\n{exc}",
            )
            return None

        if not hojas:
            self._alerta(
                "Archivo sin hojas", "El Archivo 2 no contiene hojas para leer."
            )
            return None

        if len(hojas) == 1:
            return hojas[0]

        lista = "\n".join(f"{i}. {nombre}" for i, nombre in enumerate(hojas, start=1))
        texto = (
            "Hojas disponibles en el Archivo 2:\n"
            f"{lista}\n\n"
            "Ingresá el número de la hoja a usar:"
        )

        numero, ok = QInputDialog.getInt(
            self,
            "Seleccionar hoja",
            texto,
            1,
            1,
            len(hojas),
            1,
        )
        if not ok:
            return None

        try:
            return seleccionar_hoja_excel(hojas, numero)
        except (IndexError, ValueError) as exc:
            self._alerta("Selección inválida", str(exc))
            return None

    # ── Inicio del procesamiento ──────────────────────────────
    def _iniciar_procesamiento(self):
        r1 = self.sel_arch1.ruta()
        r2 = self.sel_arch2.ruta()
        rs = self.sel_salida.ruta()

        if not r1:
            return self._alerta(
                "Falta Archivo 1", "Seleccioná el Archivo 1 (eventos diarios)."
            )
        if not r2:
            return self._alerta(
                "Falta Archivo 2", "Seleccioná el Archivo 2 (colaboradores)."
            )
        if not rs:
            return self._alerta(
                "Falta destino", "Indicá dónde guardar el archivo de salida."
            )

        hoja_archivo2 = self._seleccionar_hoja_archivo2(r2)
        if hoja_archivo2 is None:
            return

        self.btn_procesar.setEnabled(False)
        self.barra.setValue(0)
        self.barra.setFormat("Iniciando…")
        self.lbl_estado.setText(f"Usando hoja: {hoja_archivo2}")
        self.lbl_estado.setStyleSheet(f"color: {COLOR_TEXTO_SUAVE}; font-size: 12px;")

        self.hilo = HiloProcesamiento(r1, r2, rs, hoja_archivo2)
        self.hilo.progreso.connect(self._en_progreso)
        self.hilo.terminado.connect(self._en_terminado)
        self.hilo.error.connect(self._en_error)
        self.hilo.start()

    # ── Slots de señales del hilo ─────────────────────────────
    def _en_progreso(self, valor: int, mensaje: str):
        self.barra.setValue(valor)
        self.barra.setFormat(f"{valor}%")
        self.lbl_estado.setText(mensaje)

    def _en_terminado(self, stats: dict):
        self.btn_procesar.setEnabled(True)
        self.barra.setStyleSheet(
            f"QProgressBar::chunk {{ background-color: {COLOR_EXITO}; border-radius: 5px; }}"
        )
        self.lbl_estado.setText(f"✅  Guardado en: {stats['salida']}")
        self.lbl_estado.setStyleSheet(f"color: {COLOR_EXITO}; font-size: 12px;")

        msg = QMessageBox(self)
        msg.setWindowTitle("Proceso completado")
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setText(
            f"<b>¡Archivo generado correctamente!</b><br><br>"
            f"<table cellspacing='6'>"
            f"<tr><td>Total de filas en procesadas:</td><td><b>{stats['total']}</b></td></tr>"
            f"<tr><td>SAEs con coincidencia:</td>"
            f"<td><b style='color:{COLOR_EXITO}'>{stats['encontrados']}</b></td></tr>"
            f"<tr><td>SAEs sin coincidencia:</td>"
            f"<td><b style='color:{COLOR_ADVERTENCIA}'>{stats['no_encontrados']}</b></td></tr>"
            f"</table><br>"
            f"<small>Guardado en:<br>{stats['salida']}</small>"
        )
        msg.setStyleSheet(
            ESTILO_GLOBAL + f"QMessageBox {{ background: {COLOR_PANEL}; }}"
        )
        msg.exec()

    def _en_error(self, mensaje: str):
        self.btn_procesar.setEnabled(True)
        self.barra.setValue(0)
        self.barra.setFormat("Error")
        self.barra.setStyleSheet(
            f"QProgressBar::chunk {{ background-color: {COLOR_ERROR}; border-radius: 5px; }}"
        )
        self.lbl_estado.setText("❌  Ocurrió un error")
        self.lbl_estado.setStyleSheet(f"color: {COLOR_ERROR}; font-size: 12px;")

        msg = QMessageBox(self)
        msg.setWindowTitle("Error en el procesamiento")
        msg.setIcon(QMessageBox.Icon.Critical)
        msg.setText(f"<b>Error:</b><br><pre style='font-size:12px'>{mensaje}</pre>")
        msg.setStyleSheet(
            ESTILO_GLOBAL + f"QMessageBox {{ background: {COLOR_PANEL}; }}"
        )
        msg.exec()

    # ── Utilidad ──────────────────────────────────────────────
    def _alerta(self, titulo: str, texto: str):
        msg = QMessageBox(self)
        msg.setWindowTitle(titulo)
        msg.setIcon(QMessageBox.Icon.Warning)
        msg.setText(texto)
        msg.setStyleSheet(
            ESTILO_GLOBAL + f"QMessageBox {{ background: {COLOR_PANEL}; }}"
        )
        msg.exec()


# ══════════════════════════════════════════════════════════════════
#  PUNTO DE ENTRADA
# ══════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")  # base limpia para que los estilos Qt6 apliquen bien
    ventana = VentanaPrincipal()
    ventana.show()
    sys.exit(app.exec())
