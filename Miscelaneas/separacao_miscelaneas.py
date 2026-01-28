import pandas as pd

def ler_btp(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo)

    # Padronizar nomes das colunas
    df.columns = df.columns.str.strip().str.upper()

    # Selecionar colunas relevantes
    df = df[[
        "TÉCNICO",
        "MATERIAL",
        "DESCR. MATERIAL",
        "QTD BAIXADA"
    ]]

    # Renomear para padrão do sistema
    df = df.rename(columns={
        "DESCR. MATERIAL": "DESCR_MATERIAL",
        "QTD BAIXADA": "QTD_BAIXADA"
    })

    # Garantir quantidade numérica
    df["QTD_BAIXADA"] = pd.to_numeric(df["QTD_BAIXADA"], errors="coerce")

    # Remover linhas sem quantidade válida
    df = df[df["QTD_BAIXADA"].notna()]

    return df

df_btp = ler_btp("consolidado.xlsx")
print(df_btp.head())
print(df_btp.info())