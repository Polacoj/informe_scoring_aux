"""
Analizador de Personal (civiles) — interfaz Flet 1.0.

Carga el Diccionario Maestro y, opcionalmente, Fiscalización y/o Control Diario.
Valida columnas, reconcilia nombres de OPERADOR/colaboradores contra el diccionario
y puede revertir IDs a nombres usando la hoja VALIDACION.

Flujo:
  1. Seleccionar Diccionario Maestro y al menos un archivo de personal.
  2. Elegir hoja si el Excel tiene más de una.
  3. Validar columnas y reconciliar nombres → ID.
  4. Guardar *_PROCESADO en la carpeta MODIFICADOS (los originales no se tocan).
  5. Si faltan nombres, guardar una copia actualizada del diccionario en MODIFICADOS.
  6. Revertir IDs → nombres a *_REVERTIDO, también en MODIFICADOS.
"""

from __future__ import annotations

import asyncio
import shutil
from collections.abc import Awaitable, Callable
from dataclasses import dataclass, field
from pathlib import Path

import flet as ft
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

COL_ID_DICCIONARIO = "LP/DNI"
COL_NOMBRE_DICCIONARIO = "APELLIDO Y NOMBRE"
COL_OPERADOR = "OPERADOR"
COLUMNAS_FISCALIZACION_DIARIA = [COL_OPERADOR]
COLUMNAS_CONTROL_DIARIO = [
    COL_OPERADOR,
    "COLABORADORES 1",
    "COLABORADORES 2",
    "COLABORADORES 3",
    "COLABORADORES 4",
    "COLABORADORES 5",
]
CARPETA_MODIFICADOS = "MODIFICADOS"
HOJA_SHEET1 = "SHEET1"
HOJA_VALIDACION = "VALIDACION"

LogFn = Callable[[str], None]
ElegirHojaFn = Callable[[list[str], str], Awaitable[str | None]]


class ProcesoIncompleto(RuntimeError):
    """Coincidencia < 100%: el diccionario se actualizó o faltan IDs."""


@dataclass
class ResultadoColumna:
    evaluadas: int = 0
    coincidencias: int = 0
    no_encontrados: int = 0
    vacias: int = 0
    sin_id: int = 0
    nombres_nuevos: dict[str, list[str]] = field(default_factory=dict)
    nombres_sin_id: dict[str, list[str]] = field(default_factory=dict)

    @property
    def porcentaje(self) -> float:
        if self.evaluadas == 0:
            return 100.0
        return (self.coincidencias / self.evaluadas) * 100.0


def normalizar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(col).strip().upper() for col in df.columns]
    return df


def validar_columnas(
    df: pd.DataFrame, columnas_requeridas: list[str]
) -> tuple[bool, str]:
    present = set(df.columns)
    faltantes = [col for col in columnas_requeridas if col not in present]
    if faltantes:
        return (
            False,
            f"Columnas faltantes: {faltantes}. Columnas encontradas: {sorted(present)}",
        )
    return True, ""


def normalizar_nombre(valor) -> str:
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto == "" or texto.lower() == "nan":
        return ""
    return " ".join(texto.upper().split())


def _serie_id(serie: pd.Series) -> pd.Series:
    return serie.fillna("").astype(str).str.strip().replace({"nan": "", "None": ""})


def construir_mapa_nombres(df_diccionario: pd.DataFrame) -> dict[str, str]:
    nombres = df_diccionario[COL_NOMBRE_DICCIONARIO].map(normalizar_nombre)
    ids = _serie_id(df_diccionario[COL_ID_DICCIONARIO])
    mask = nombres != ""
    return dict(zip(nombres[mask], ids[mask], strict=False))


def construir_mapa_inverso(df_validacion: pd.DataFrame) -> dict[str, str]:
    ids = _serie_id(df_validacion[COL_ID_DICCIONARIO])
    nombres = df_validacion[COL_NOMBRE_DICCIONARIO].fillna("").astype(str).str.strip()
    mask = ids != ""
    return dict(zip(ids[mask], nombres[mask], strict=False))


def listar_hojas(ruta: Path) -> list[str]:
    with pd.ExcelFile(ruta) as libro:
        return list(libro.sheet_names)


def leer_hoja(ruta: Path, hoja: str) -> pd.DataFrame:
    return pd.read_excel(ruta, sheet_name=hoja, dtype=str)


def ruta_procesado(ruta: Path) -> Path:
    return ruta.parent / f"{ruta.stem}_PROCESADO{ruta.suffix}"


def ruta_revertido(ruta_procesado_path: Path) -> Path:
    stem = ruta_procesado_path.stem.replace("_PROCESADO", "")
    return ruta_procesado_path.parent / f"{stem}_REVERTIDO{ruta_procesado_path.suffix}"


def _buscar_hoja(nombres: list[str], objetivo: str) -> str | None:
    objetivo = objetivo.strip().upper()
    return next((n for n in nombres if n.strip().upper() == objetivo), None)


def guardar_diccionario_con_bak(df_actualizado: pd.DataFrame, ruta_original: Path) -> Path:
    """
    Copia el original a ANTIGUO_DICCIONARIO.xlsx (se pisa) y reescribe Sheet1
    conservando el resto de las hojas (p. ej. Validacion).
    """
    ruta_bak = ruta_original.parent / BACKUP_DICCIONARIO
    shutil.copy2(ruta_original, ruta_bak)

    wb = load_workbook(ruta_original)
    nombre_sheet1 = _buscar_hoja(wb.sheetnames, HOJA_SHEET1)
    ws = wb[nombre_sheet1] if nombre_sheet1 else wb.create_sheet("Sheet1")
    if ws.max_row:
        ws.delete_rows(1, ws.max_row)

    df_out = df_actualizado.where(df_actualizado.notna(), "")
    for fila_idx, fila in enumerate(dataframe_to_rows(df_out, index=False, header=True), start=1):
        for col_idx, valor in enumerate(fila, start=1):
            ws.cell(row=fila_idx, column=col_idx, value=valor)

    wb.save(ruta_original)
    return ruta_bak


def _ubicaciones(col: str, mask: pd.Series) -> list[str]:
    return [f"{col}:fila {int(i) + 2}" for i in mask.index[mask]]


def aplicar_mapa_a_columnas(
    df: pd.DataFrame,
    columnas: list[str],
    mapa_nombres: dict[str, str],
) -> tuple[pd.DataFrame, ResultadoColumna]:
    df_mod = df.copy().reset_index(drop=True)
    stats = ResultadoColumna()

    for col in columnas:
        if col not in df_mod.columns:
            raise RuntimeError(f"La hoja no contiene la columna '{col}'.")

        normalizados = df_mod[col].map(normalizar_nombre)
        vacias = normalizados.eq("")
        ids = normalizados.map(mapa_nombres)
        en_mapa = ~vacias & ids.notna()
        sin_id = en_mapa & ids.eq("")
        coincidencias = en_mapa & ids.ne("")
        no_encontrados = ~vacias & ids.isna()

        stats.vacias += int(vacias.sum())
        stats.coincidencias += int(coincidencias.sum())
        stats.no_encontrados += int(no_encontrados.sum())
        stats.sin_id += int(sin_id.sum())
        stats.evaluadas += int((~vacias).sum())

        for nombre in normalizados[no_encontrados].unique():
            stats.nombres_nuevos.setdefault(str(nombre), []).extend(
                _ubicaciones(col, normalizados.eq(nombre) & no_encontrados)
            )
        for nombre in normalizados[sin_id].unique():
            stats.nombres_sin_id.setdefault(str(nombre), []).extend(
                _ubicaciones(col, normalizados.eq(nombre) & sin_id)
            )

        df_mod.loc[coincidencias, col] = ids[coincidencias].astype(str)

    return df_mod, stats


def _agregar_nombres_al_diccionario(
    df_diccionario: pd.DataFrame, nombres: list[str]
) -> pd.DataFrame:
    if not nombres:
        return df_diccionario
    nuevas = [{col: "" for col in df_diccionario.columns} for _ in nombres]
    for fila, nombre in zip(nuevas, nombres, strict=True):
        fila[COL_NOMBRE_DICCIONARIO] = nombre
    return pd.concat([df_diccionario, pd.DataFrame(nuevas)], ignore_index=True)


def _manejar_coincidencia_incompleta(
    stats: ResultadoColumna,
    df_diccionario: pd.DataFrame,
    ruta_diccionario: Path,
    log: LogFn,
) -> None:
    if stats.nombres_sin_id:
        log("Hay nombres en el diccionario sin LP/DNI. Completá la columna ID y volvé a correr.")
        for nombre, ubicaciones in stats.nombres_sin_id.items():
            log(f"  Sin ID: '{nombre}' en {ubicaciones}")

    if stats.nombres_nuevos:
        log("Coincidencia incompleta. Actualizando diccionario original con los faltantes...")
        df_act = _agregar_nombres_al_diccionario(
            df_diccionario, list(stats.nombres_nuevos)
        )
        ruta_bak = guardar_diccionario_con_bak(df_act, ruta_diccionario)
        for nombre, ubicaciones in stats.nombres_nuevos.items():
            log(f"  Nombre nuevo: '{nombre}' en {ubicaciones}")
        log(f"  ✔ Backup guardado: {ruta_bak.name}")
        log(f"  ✔ Diccionario original actualizado: {ruta_diccionario.name}")

    partes = [f"Coincidencia incompleta ({stats.porcentaje:.1f}%)."]
    if stats.nombres_nuevos:
        partes.append(
            f"Se actualizó '{ruta_diccionario.name}' con los nombres faltantes."
        )
    if stats.nombres_sin_id:
        partes.append("Hay entradas del diccionario sin LP/DNI.")
    partes.append("Completá la columna ID y volvé a correr.")
    raise ProcesoIncompleto(" ".join(partes))


def reconciliar_archivo(
    df: pd.DataFrame,
    df_diccionario: pd.DataFrame,
    columnas: list[str],
    ruta_origen: Path,
    ruta_diccionario: Path,
    etiqueta: str,
    log: LogFn,
) -> tuple[pd.DataFrame, pd.DataFrame, Path]:
    if (
        COL_ID_DICCIONARIO not in df_diccionario.columns
        or COL_NOMBRE_DICCIONARIO not in df_diccionario.columns
    ):
        raise RuntimeError(
            f"El Diccionario Maestro debe contener las columnas "
            f"'{COL_ID_DICCIONARIO}' y '{COL_NOMBRE_DICCIONARIO}'."
        )

    log(f"Iniciando reconciliación de {etiqueta}...")
    mapa = construir_mapa_nombres(df_diccionario)
    df_mod, stats = aplicar_mapa_a_columnas(df, columnas, mapa)
    log(
        f"{etiqueta}: {stats.evaluadas} celdas evaluadas, "
        f"{stats.coincidencias} coincidencias, {stats.no_encontrados} no encontradas, "
        f"{stats.sin_id} sin ID, {stats.porcentaje:.1f}%."
    )

    if stats.porcentaje < 100.0:
        _manejar_coincidencia_incompleta(stats, df_diccionario, ruta_diccionario, log)

    salida = ruta_procesado(ruta_origen)
    df_mod.to_excel(salida, index=False)
    log(f"{etiqueta} procesada guardada en: {salida.name}")
    return df_mod, df_diccionario, salida


def revertir_archivo(
    ruta_procesado_path: Path,
    columnas: list[str],
    mapa_inverso: dict[str, str],
    etiqueta: str,
    log: LogFn,
) -> Path:
    log(f"\nProcesando {etiqueta}: {ruta_procesado_path.name}...")
    hojas = listar_hojas(ruta_procesado_path)
    if not hojas:
        raise RuntimeError(f"{etiqueta} no tiene hojas válidas.")
    df = normalizar_columnas(leer_hoja(ruta_procesado_path, hojas[0])).reset_index(drop=True)

    reemplazados = 0
    vacias = 0
    ids_sin_match: list[str] = []

    for col in columnas:
        if col not in df.columns:
            log(f"  ⚠ Columna '{col}' no encontrada en {ruta_procesado_path.name}, se omite.")
            continue
        valores = df[col].fillna("").astype(str).str.strip()
        vacio = valores.eq("") | valores.str.lower().eq("nan")
        vacias += int(vacio.sum())
        match = ~vacio & valores.isin(mapa_inverso)
        sin = ~vacio & ~valores.isin(mapa_inverso)
        df.loc[match, col] = valores[match].map(mapa_inverso)
        reemplazados += int(match.sum())
        for idx, valor in valores[sin].items():
            ids_sin_match.append(f"{col}:fila {int(idx) + 2} → '{valor}'")

    log(f"  ✔ Reemplazados:  {reemplazados}")
    log(f"  ○ Vacíos:        {vacias}")
    if ids_sin_match:
        log(f"  ⚠ Sin coincidencia en VALIDACION: {len(ids_sin_match)}")
        for detalle in ids_sin_match:
            log(f"     {detalle}")

    salida = ruta_revertido(ruta_procesado_path)
    df.to_excel(salida, index=False)
    log(f"  ✔ Guardado en: {salida.name}")
    return salida


def validar_ruta(ruta_text: str, etiqueta: str) -> Path:
    if not ruta_text.strip():
        raise RuntimeError(f"Debés seleccionar el archivo: {etiqueta}")
    ruta = Path(ruta_text.strip())
    if not ruta.exists():
        raise RuntimeError(f"No se encontró el archivo {etiqueta}: {ruta}")
    return ruta


async def cargar_excel(
    ruta: Path,
    etiqueta: str,
    log: LogFn,
    elegir_hoja: ElegirHojaFn,
) -> pd.DataFrame:
    log(f"Cargando '{ruta.name}'...")
    hojas = listar_hojas(ruta)
    if not hojas:
        raise RuntimeError(f"El archivo {etiqueta} no tiene hojas válidas.")
    if len(hojas) == 1:
        nombre = hojas[0]
        log(f"Hoja seleccionada automáticamente: {nombre}")
    else:
        nombre = await elegir_hoja(hojas, etiqueta)
        if not nombre:
            raise RuntimeError(f"Se canceló la selección de hoja para {etiqueta}.")
        log(f"Hoja seleccionada: {nombre}")
    return normalizar_columnas(leer_hoja(ruta, nombre))


class AnalizadorApp:
    def __init__(self, page: ft.Page) -> None:
        self.page = page
        self.picker = ft.FilePicker()
        self.ruta_trabajo_procesado: Path | None = None
        self.ruta_control_procesado: Path | None = None
        self.ruta_diccionario: Path | None = None

        self.dict_path = ft.TextField(expand=True, read_only=True, hint_text="Sin seleccionar")
        self.trabajo_path = ft.TextField(expand=True, read_only=True, hint_text="Sin seleccionar")
        self.control_path = ft.TextField(expand=True, read_only=True, hint_text="Sin seleccionar")

        self.trabajo_check = ft.Checkbox(
            label="Incluir Fiscalización Diaria/Mensual",
            value=True,
            on_change=self._on_toggle_trabajo,
        )
        self.control_check = ft.Checkbox(
            label="Incluir Control Diario Mes",
            value=True,
            on_change=self._on_toggle_control,
        )

        self.btn_dict = ft.Button(
            "Seleccionar",
            icon=ft.Icons.FOLDER_OPEN,
            on_click=lambda e: self._pick(self.dict_path),
        )
        self.btn_trabajo = ft.Button(
            "Seleccionar",
            icon=ft.Icons.FOLDER_OPEN,
            on_click=lambda e: self._pick(self.trabajo_path),
        )
        self.btn_control = ft.Button(
            "Seleccionar",
            icon=ft.Icons.FOLDER_OPEN,
            on_click=lambda e: self._pick(self.control_path),
        )

        self.start_button = ft.Button(
            "Validar y reconciliar",
            icon=ft.Icons.PLAY_ARROW,
            bgcolor="#2196F3",
            color=ft.Colors.WHITE,
            on_click=self.on_start,
        )
        self.reverse_button = ft.Button(
            "Revertir IDs → Nombres (hoja Validacion)",
            icon=ft.Icons.UNDO,
            bgcolor=ft.Colors.GREY_500,
            color=ft.Colors.WHITE,
            disabled=True,
            on_click=self.on_revertir,
        )
        self.exit_button = ft.Button(
            "Salir",
            icon=ft.Icons.EXIT_TO_APP,
            bgcolor="#eb7f82",
            color=ft.Colors.WHITE,
            on_click=lambda e: self.page.window.close(),
        )
        self.progress = ft.ProgressRing(visible=False, width=22, height=22)
        self.log_field = ft.TextField(
            multiline=True,
            read_only=True,
            expand=True,
            min_lines=14,
            text_size=14,
            label="Registro de ejecución",
        )

    def build(self) -> None:
        page = self.page
        page.title = "Analizador de Personal"
        page.theme_mode = ft.ThemeMode.LIGHT
        page.padding = 20
        page.window.width = 920
        page.window.height = 760
        page.services.append(self.picker)

        page.add(
            ft.Column(
                [
                    ft.Text("Analizador de Personal", size=24, weight=ft.FontWeight.BOLD),
                    self._file_row("Diccionario Maestro:", self.dict_path, self.btn_dict),
                    self.trabajo_check,
                    self._file_row(
                        "Archivo Fiscalizacion diaria/mensual:",
                        self.trabajo_path,
                        self.btn_trabajo,
                    ),
                    self.control_check,
                    self._file_row(
                        "Control Diario Mes:", self.control_path, self.btn_control
                    ),
                    ft.Row([self.start_button, self.progress], spacing=12),
                    self.reverse_button,
                    self.log_field,
                    self.exit_button,
                ],
                expand=True,
                spacing=10,
            )
        )

    def _file_row(self, label: str, field: ft.TextField, button: ft.Button) -> ft.Row:
        return ft.Row(
            [ft.Text(label, width=260), field, button],
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )

    def append_log(self, mensaje: str) -> None:
        actual = self.log_field.value or ""
        self.log_field.value = f"{actual}{mensaje}\n" if actual else f"{mensaje}\n"
        self.page.update()

    async def _dialogo(self, titulo: str, mensaje: str) -> None:
        loop = asyncio.get_running_loop()
        done: asyncio.Future[None] = loop.create_future()

        def cerrar(_e=None) -> None:
            if done.done():
                return
            done.set_result(None)
            self.page.pop_dialog()

        self.page.show_dialog(
            ft.AlertDialog(
                modal=True,
                title=ft.Text(titulo),
                content=ft.Text(mensaje),
                actions=[ft.Button("Aceptar", on_click=cerrar)],
                on_dismiss=lambda e: cerrar(),
            )
        )
        await done

    async def _elegir_hoja(self, titulos: list[str], etiqueta: str) -> str | None:
        loop = asyncio.get_running_loop()
        done: asyncio.Future[str | None] = loop.create_future()
        dropdown = ft.Dropdown(
            label="Hoja",
            options=[ft.DropdownOption(key=t, text=t) for t in titulos],
            value=titulos[0],
            width=420,
        )

        def resolver(valor: str | None) -> None:
            if not done.done():
                done.set_result(valor)
            self.page.pop_dialog()

        self.page.show_dialog(
            ft.AlertDialog(
                modal=True,
                title=ft.Text("Seleccionar hoja"),
                content=ft.Column(
                    [
                        ft.Text(f"{etiqueta} tiene varias hojas. Seleccioná una:"),
                        dropdown,
                    ],
                    tight=True,
                    width=440,
                ),
                actions=[
                    ft.TextButton("Cancelar", on_click=lambda e: resolver(None)),
                    ft.Button("Aceptar", on_click=lambda e: resolver(dropdown.value)),
                ],
                on_dismiss=lambda e: resolver(None) if not done.done() else None,
            )
        )
        return await done

    async def _pick(self, field: ft.TextField) -> None:
        archivos = await self.picker.pick_files(
            dialog_title="Seleccionar archivo Excel",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx", "xlsm"],
            allow_multiple=False,
        )
        if archivos and archivos[0].path:
            field.value = archivos[0].path
            self.page.update()

    async def _on_toggle_trabajo(self, e: ft.Event[ft.Checkbox]) -> None:
        if not self.trabajo_check.value and not self.control_check.value:
            self.trabajo_check.value = True
            self.page.update()
            await self._dialogo(
                "Selección requerida",
                "Debés incluir al menos un archivo: Fiscalización o Control Diario.",
            )
            return
        activo = bool(self.trabajo_check.value)
        self.trabajo_path.disabled = not activo
        self.btn_trabajo.disabled = not activo
        if not activo:
            self.trabajo_path.value = ""
        self.page.update()

    async def _on_toggle_control(self, e: ft.Event[ft.Checkbox]) -> None:
        if not self.control_check.value and not self.trabajo_check.value:
            self.control_check.value = True
            self.page.update()
            await self._dialogo(
                "Selección requerida",
                "Debés incluir al menos un archivo: Fiscalización o Control Diario.",
            )
            return
        activo = bool(self.control_check.value)
        self.control_path.disabled = not activo
        self.btn_control.disabled = not activo
        if not activo:
            self.control_path.value = ""
        self.page.update()

    def _habilitar_reverso(self) -> None:
        self.reverse_button.disabled = False
        self.reverse_button.bgcolor = "#4CAF50"
        self.page.update()

    async def on_start(self, e: ft.Event[ft.Button]) -> None:
        self.start_button.disabled = True
        self.progress.visible = True
        self.log_field.value = ""
        self.page.update()
        try:
            await self.run_process()
        except Exception as exc:
            self.append_log(f"ERROR: {exc}")
            await self._dialogo("Error", str(exc))
        finally:
            self.progress.visible = False
            self.start_button.disabled = False
            self.page.update()

    async def run_process(self) -> None:
        incluir_trabajo = bool(self.trabajo_check.value)
        incluir_control = bool(self.control_check.value)
        if not incluir_trabajo and not incluir_control:
            raise RuntimeError(
                "Debés incluir al menos un archivo: Fiscalización o Control Diario."
            )

        ruta_dict = validar_ruta(self.dict_path.value or "", "Diccionario Maestro")
        ruta_trabajo = (
            validar_ruta(self.trabajo_path.value or "", "Fiscalización diaria/mensual")
            if incluir_trabajo
            else None
        )
        ruta_control = (
            validar_ruta(self.control_path.value or "", "Control Diario Mes")
            if incluir_control
            else None
        )

        df_dict = await cargar_excel(
            ruta_dict, "Diccionario Maestro", self.append_log, self._elegir_hoja
        )
        df_trabajo = None
        if ruta_trabajo is not None:
            df_trabajo = await cargar_excel(
                ruta_trabajo,
                "Fiscalización diaria/mensual",
                self.append_log,
                self._elegir_hoja,
            )
        df_control = None
        if ruta_control is not None:
            df_control = await cargar_excel(
                ruta_control, "Control Diario Mes", self.append_log, self._elegir_hoja
            )

        errores: list[str] = []
        if df_trabajo is not None:
            ok, msg = validar_columnas(df_trabajo, COLUMNAS_FISCALIZACION_DIARIA)
            if not ok:
                errores.append(f"Fiscalización Diaria/Mensual: {msg}")
        if df_control is not None:
            ok, msg = validar_columnas(df_control, COLUMNAS_CONTROL_DIARIO)
            if not ok:
                errores.append(f"Control Diario Mes: {msg}")
        if errores:
            raise RuntimeError("\n".join(errores))

        self.append_log("Archivos cargados y validados correctamente.")
        self.append_log(f"Diccionario: {ruta_dict.name}\n")
        if ruta_trabajo:
            self.append_log(f"Fiscalización Diaria/Mensual: {ruta_trabajo.name}\n")
        else:
            self.append_log(
                "Fiscalización Diaria/Mensual: NO incluida (omitida por el usuario)\n"
            )
        if ruta_control:
            self.append_log(f"Control Diario: {ruta_control.name}\n")
        else:
            self.append_log("Control Diario: NO incluido (omitido por el usuario)\n")

        df_dict_vigente = df_dict
        self.ruta_trabajo_procesado = None
        self.ruta_control_procesado = None

        if df_trabajo is not None and ruta_trabajo is not None:
            _, df_dict_vigente, salida = reconciliar_archivo(
                df_trabajo,
                df_dict_vigente,
                COLUMNAS_FISCALIZACION_DIARIA,
                ruta_trabajo,
                ruta_dict,
                "Fiscalización Diaria/Mensual",
                self.append_log,
            )
            self.append_log("Módulo 2 completado.\n")
            self.ruta_trabajo_procesado = salida
        else:
            self.append_log("Módulo 2 omitido (Fiscalización no incluida).\n")

        if df_control is not None and ruta_control is not None:
            reconciliar_archivo(
                df_control,
                df_dict_vigente,
                COLUMNAS_CONTROL_DIARIO,
                ruta_control,
                ruta_dict,
                "Control Diario Mes",
                self.append_log,
            )
            self.append_log("Módulo 3 completado.\n")
            self.ruta_control_procesado = ruta_procesado(ruta_control)
        else:
            self.append_log("Módulo 3 omitido (Control Diario no incluido).\n")

        self.ruta_diccionario = ruta_dict
        self._habilitar_reverso()
        self.append_log(
            "\n✔ Proceso inverso disponible. Presioná 'Revertir IDs → Nombres'."
        )
        await self._dialogo("Finalizado", "Proceso completado correctamente.")

    async def on_revertir(self, e: ft.Event[ft.Button]) -> None:
        self.reverse_button.disabled = True
        self.page.update()
        try:
            await asyncio.to_thread(self.revertir_ids_a_nombres)
            await self._dialogo("Proceso inverso", "Reversión completada correctamente.")
        except Exception as exc:
            self.append_log(f"ERROR en proceso inverso: {exc}")
            await self._dialogo("Error", str(exc))
        finally:
            self._habilitar_reverso()

    def revertir_ids_a_nombres(self) -> None:
        self.append_log("\n" + "═" * 50)
        self.append_log("  PROCESO INVERSO: IDs → NOMBRES (hoja VALIDACION)")
        self.append_log("═" * 50)

        if not self.ruta_trabajo_procesado and not self.ruta_control_procesado:
            raise RuntimeError(
                "No se encontraron archivos _PROCESADO. "
                "Ejecutá primero el proceso de validación y reconciliación."
            )
        if self.ruta_trabajo_procesado and not self.ruta_trabajo_procesado.exists():
            raise RuntimeError(
                f"No se encontró el archivo: {self.ruta_trabajo_procesado.name}"
            )
        if self.ruta_control_procesado and not self.ruta_control_procesado.exists():
            raise RuntimeError(
                f"No se encontró el archivo: {self.ruta_control_procesado.name}"
            )
        if self.ruta_diccionario is None:
            raise RuntimeError("No hay un Diccionario Maestro asociado a este ciclo.")

        self.append_log(f"\nLeyendo hoja 'VALIDACION' de: {self.ruta_diccionario.name}...")
        nombres_hojas = listar_hojas(self.ruta_diccionario)
        hoja_val = _buscar_hoja(nombres_hojas, HOJA_VALIDACION)
        if hoja_val is None:
            raise RuntimeError(
                "No se encontró la hoja 'VALIDACION' en el Diccionario Maestro.\n"
                f"Hojas disponibles: {nombres_hojas}"
            )

        df_val = normalizar_columnas(leer_hoja(self.ruta_diccionario, hoja_val))
        self.append_log(f"  Hoja encontrada: '{hoja_val}' — {len(df_val)} filas.")
        for col_req in (COL_ID_DICCIONARIO, COL_NOMBRE_DICCIONARIO):
            if col_req not in df_val.columns:
                raise RuntimeError(
                    f"La hoja 'VALIDACION' no contiene la columna '{col_req}'.\n"
                    f"Columnas encontradas: {list(df_val.columns)}"
                )

        mapa_inverso = construir_mapa_inverso(df_val)
        self.append_log(f"  Mapa inverso construido: {len(mapa_inverso)} entradas.")

        if self.ruta_trabajo_procesado:
            revertir_archivo(
                self.ruta_trabajo_procesado,
                [COL_OPERADOR],
                mapa_inverso,
                "Fiscalización Diaria/Mensual",
                self.append_log,
            )
        else:
            self.append_log(
                "\nFiscalización Diaria/Mensual: omitida (no se procesó en este ciclo)."
            )

        if self.ruta_control_procesado:
            revertir_archivo(
                self.ruta_control_procesado,
                COLUMNAS_CONTROL_DIARIO,
                mapa_inverso,
                "Control Diario Mes",
                self.append_log,
            )
        else:
            self.append_log("\nControl Diario Mes: omitido (no se procesó en este ciclo).")

        self.append_log("\n" + "═" * 50)
        self.append_log("  ✔ PROCESO INVERSO COMPLETADO")
        self.append_log("═" * 50)


async def main(page: ft.Page) -> None:
    app = AnalizadorApp(page)
    app.build()


if __name__ == "__main__":
    ft.run(main)
