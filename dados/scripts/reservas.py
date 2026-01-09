import pandas as pd
from pathlib import Path

BASE = Path(__file__).parent.parent

ENTRADA = BASE / "entrada"
REGRAS = BASE / "regras"
SAIDA = BASE / "saida"
RELATORIO = BASE / "relatorios"

pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)

# Ler arquivo
df = pd.read_excel(ENTRADA / "RS 09_01.xlsx")

# Filtrar linhas
df_reservas_porto_alegre = df[
    (df["NM_MUNICIPIO_SAP"] == "PORTO ALEGRE") &
    (df["STATUS_NF"] != "ENTREGA CONCLUIDA") &
    (df["NM_FORN_SAP"].astype(str).str.contains("PROCISA", case=False, na=False))
]

# ✅ MANTER SOMENTE AS COLUNAS NECESSÁRIAS
colunas_desejadas = [
    'NM_MUNICIPIO_SAP',
    'NM_FORN_SAP',
    'TIPO_MATERIAL',
    'MATERIAL',
    'FAMILIA',
    'TEXTO_BREVE_MATERIAL',
    'QUANTIDADE',
    'DT_CRIACAO_RESERVA',
    'RESERVA',
    'NF_CORRIGIDA',
    'STATUS_NF',
    'TRANSPORTADORA',
    'DT_EXPEDICAO',
    'DT_LIMITE_ORIGINAL',

]

df_reservas_porto_alegre = df_reservas_porto_alegre.loc[:, colunas_desejadas]

# Conferência final
print(df_reservas_porto_alegre.columns.tolist())

# Exportar
df_reservas_porto_alegre.to_excel(RELATORIO/
    "reservas_porto_alegre_procisa.xlsx",
    index=False
)
