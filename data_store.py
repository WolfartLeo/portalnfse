# data_store.py
import os
import pandas as pd

COLUNAS_OBRIGATORIAS = [
    "EMPRESA",
    "CNPJ",
    "TIPO_ACESSO",
    "LOGIN",
    "SENHA",
    "ATIVO",
    "PREFEITURA",
    "IDENT_CERT",
    "IMG_CERT",
]

def garantir_planilha_modelo(caminho: str) -> None:
    pasta = os.path.dirname(caminho) or "."
    os.makedirs(pasta, exist_ok=True)
    if os.path.exists(caminho):
        return
    df = pd.DataFrame(columns=COLUNAS_OBRIGATORIAS)
    df.to_excel(caminho, index=False)

def ler_clientes(caminho: str) -> pd.DataFrame:
    garantir_planilha_modelo(caminho)
    df = pd.read_excel(caminho, dtype=str).fillna("")
    for c in COLUNAS_OBRIGATORIAS:
        if c not in df.columns:
            df[c] = ""
    return df[COLUNAS_OBRIGATORIAS].copy()

def salvar_clientes(caminho: str, df: pd.DataFrame) -> None:
    pasta = os.path.dirname(caminho) or "."
    os.makedirs(pasta, exist_ok=True)
    df.to_excel(caminho, index=False)
