import csv
from io import BytesIO

import pandas as pd


def ler_csv_flexivel(conteudo: bytes) -> pd.DataFrame:
    """Lê CSV com detecção de separador, codificação e aspas malformadas."""
    ultimo_erro = None
    for codificacao in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return pd.read_csv(
                BytesIO(conteudo), dtype=str, sep=None, engine="python", encoding=codificacao
            )
        except UnicodeDecodeError as erro:
            ultimo_erro = erro
            continue
        except (pd.errors.ParserError, csv.Error) as erro:
            ultimo_erro = erro

        try:
            texto = conteudo.decode(codificacao)
            linhas = [linha for linha in texto.splitlines() if linha.strip()][:20]
            separadores = (";", ",", "\t", "|")
            separador = max(
                separadores,
                key=lambda candidato: sum(linha.count(candidato) for linha in linhas),
            )
            if not linhas or not any(separador in linha for linha in linhas):
                continue
            return pd.read_csv(
                BytesIO(conteudo),
                dtype=str,
                sep=separador,
                engine="python",
                encoding=codificacao,
                quoting=csv.QUOTE_NONE,
            )
        except (UnicodeDecodeError, pd.errors.ParserError, csv.Error) as erro:
            ultimo_erro = erro

    raise ultimo_erro or ValueError("Não foi possível identificar o formato do CSV.")
