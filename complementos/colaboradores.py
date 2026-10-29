import pandas as pd
import unicodedata
from pathlib import Path


class ColaboradoresProcessor:
    """Procesa un DataFrame para extraer filas con eventos de colaboración.

    Detecta colaboradores listados en columnas tipo `COLABORADOR 1..N` y genera
    dos tipos de filas: apoyo óptico (cuando `TIPIFICACIÓN` == 'COLABORACION')
    y colaboración en evento (en caso contrario).
    """

    def __init__(
        self,
        colaboradores_cols=None,
        tip_col="TIPIFICACION",
        org_col="ORIGEN",
        gap_col="GAP",
        sae_col="SAE",
        op_col="OPERADOR",
    ):
        if colaboradores_cols is None:
            colaboradores_cols = [
                "COLABORADOR 1",
                "COLABORADOR 2",
                "COLABORADOR 3",
                "COLABORADOR 4",
                "COLABORADOR 5",
            ]
        self.colaboradores_cols = colaboradores_cols
        self.tip_col = tip_col
        self.org_col = org_col
        self.gap_col = gap_col
        self.sae_col = sae_col
        self.op_col = op_col

    def remove_accents(self, s):
        s = str(s)
        return "".join(
            ch
            for ch in unicodedata.normalize("NFKD", s)
            if not unicodedata.combining(ch)
        )

    def process_df(self, df: pd.DataFrame) -> pd.DataFrame:
        colaborador_optico = []
        colaborador_evento = []

        for idx, fila in df.iterrows():
            tip_val = self.remove_accents(fila.get(self.tip_col, "")).strip().upper()
            es_colab = tip_val == "COLABORACION"

            for col in self.colaboradores_cols:
                valor = fila.get(col)
                if pd.notna(valor) and str(valor).strip() != "":
                    if es_colab:
                        colaborador_optico.append(
                            {
                                "FECHA": "",
                                "HORA": "",
                                "OPERADOR": valor,
                                "TIPO DE EVENTO": "COLABORACION EN APOYO OPTICO",
                                "BARRIO": "ADICIONAL",
                                "ORIGEN": fila.get(self.org_col),
                                "GAP": fila.get(self.gap_col),
                                "SAE": fila.get(self.sae_col),
                            }
                        )
                    else:
                        colaborador_evento.append(
                            {
                                "FECHA": "",
                                "HORA": "",
                                "OPERADOR": valor,
                                "TIPO DE EVENTO": "COLABORACION EN EVENTO",
                                "BARRIO": "ADICIONAL",
                                "ORIGEN": fila.get(self.org_col),
                                "GAP": fila.get(self.gap_col),
                                "SAE": fila.get(self.sae_col),
                            }
                        )

        resultado = pd.concat(
            [pd.DataFrame(colaborador_optico), pd.DataFrame(colaborador_evento)],
            ignore_index=True,
        )

        return resultado

    def process_file(self, ruta_archivo: str) -> pd.DataFrame:
        df = pd.read_excel(ruta_archivo, engine="openpyxl")
        return self.process_df(df)


def process_file(ruta_archivo, **kwargs):
    proc = ColaboradoresProcessor(**kwargs)
    return proc.process_file(ruta_archivo)


if __name__ == "__main__":
    default_path = Path.cwd() / "CMU - SEPTIEMBRE.xlsx"
    if default_path.exists():
        resultado = process_file(default_path)
        salida = Path.cwd() / "colaboradores.xlsx"
        resultado.to_excel(salida, index=False)
        print(f"Archivo '{salida}' generado correctamente.")
    else:
        print(
            "Ejemplo: importar ColaboradoresProcessor desde otro script y usar process_file()."
        )
