import pandas as pd
from paths import ENTRADA_RESERVA,RELATORIO_DE_RESERVAS

pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)

def carregar_reservas(nome_arquivo):
    caminho = ENTRADA_RESERVA / "reserva.xlsx"
    return pd.read_excel(caminho)

def gerar_reservas_porto_alegre(nome_arquivo):
    df = carregar_reservas(nome_arquivo)

    df.columns = df.columns.str.strip().str.upper()

    df_filtrado = df[
        (df["NM_MUNICIPIO_SAP"] == "PORTO ALEGRE") &
        (df["STATUS_NF"] != "ENTREGA CONCLUIDA") &
        (df["NM_FORN_SAP"].astype(str).str.contains("PROCISA", case=False, na=False))
    ]

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

    df_final = df_filtrado.loc[:, colunas_desejadas]

    # 🔹 Debug
    print(df_final.columns.tolist())

    # 🔹 Exportar
    RELATORIO_DE_RESERVAS.mkdir(exist_ok=True)
    df_final.to_excel(
        RELATORIO_DE_RESERVAS / "reservas_porto_alegre_procisa.xlsx",
        index=False
    )

    print("✔ Relatório de reservas Porto Alegre gerado com sucesso!")

if __name__ == "__main__":
    gerar_reservas_porto_alegre("")
