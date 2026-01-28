import pandas as pd
import matplotlib.pyplot as plt
import unicodedata
from paths import (
    BAIXAS_DE_MISCELANEAS,
    ENTRADA_EQUIPAMENTOS,
    ENTRADA_ESCALA_COMPLETA,
    RELATORIO_DE_MISCELANEAS
)

# ==============================
# PATHS
# ==============================
ARQ_FONTES = BAIXAS_DE_MISCELANEAS / "baixa_de_miscelaneas.xlsx"
ARQ_INSTALACOES = ENTRADA_EQUIPAMENTOS / "relatorio_equipamento.xlsx"
ARQ_VINCULO = ENTRADA_ESCALA_COMPLETA / "lista_atualizada.xlsx"

# ==============================
# FUNÇÕES AUXILIARES
# ==============================
def remover_acentos(texto):
    return "".join(
        c for c in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(c)
    )


def normalizar_colunas(df):
    df.columns = [
        remover_acentos(col)
        .strip()
        .upper()
        .replace(" ", "_")
        for col in df.columns
    ]
    return df


# ==============================
# EXECUÇÃO
# ==============================
def executar():

    # ==============================
    # LEITURA DOS ARQUIVOS
    # ==============================
    df_ft = pd.read_excel(ARQ_FONTES)
    df_eq = pd.read_excel(ARQ_INSTALACOES)
    df_vinculo = pd.read_excel(ARQ_VINCULO)

    df_ft = normalizar_colunas(df_ft)
    df_eq = normalizar_colunas(df_eq)
    df_vinculo = normalizar_colunas(df_vinculo)

    # ==============================
    # PADRONIZAÇÃO
    # ==============================
    df_ft["TECNICO"] = df_ft["TECNICO"].str.upper().str.strip()
    df_eq["TECNICO"] = df_eq["TECNICO"].str.upper().str.strip()

    # Técnico e Supervisor vêm do vínculo
    df_vinculo["TECNICO"] = df_vinculo["NOME"].str.upper().str.strip()
    df_vinculo["SUPERVISOR"] = df_vinculo["SUPORTE_DE_PRODUCAO"].str.upper().str.strip()

    df_vinculo = df_vinculo[["TECNICO", "SUPERVISOR"]]

    # ==============================
    # CRITÉRIO CORRETO
    # STATUS = INSTALADO → EQUIPAMENTO INSTALADO
    # ==============================
    instalacoes = (
        df_eq[df_eq["STATUS"] == "INSTALADO"]
        .groupby("TECNICO")
        .size()
        .rename("INSTALACOES")
    )

    baixas_fontes = (
        df_ft.groupby("TECNICO")["QUANTIDADE"]
        .sum()
        .rename("BAIXAS_FONTES")
    )

    # ==============================
    # RELATÓRIO BASE
    # ==============================
    relatorio = pd.concat(
        [instalacoes, baixas_fontes],
        axis=1
    ).fillna(0)

    relatorio["DIFERENCA"] = (
        relatorio["INSTALACOES"] - relatorio["BAIXAS_FONTES"]
    )

    relatorio["STATUS_AUDITORIA"] = relatorio["DIFERENCA"].apply(
        lambda x: "OK" if x <= 0 else "DIVERGENTE"
    )

    # ==============================
    # VINCULAR SUPERVISOR
    # ==============================
    relatorio = relatorio.join(
        df_vinculo.set_index("TECNICO"),
        how="left"
    )

    # ==============================
    # FILTRAR DIVERGENTES
    # ==============================
    divergentes = relatorio[
        relatorio["STATUS_AUDITORIA"] == "DIVERGENTE"
    ]

    # ==============================
    # AGRUPAR POR SUPERVISOR
    # ==============================
    divergencias_por_supervisor = (
        divergentes
        .groupby("SUPERVISOR")["DIFERENCA"]
        .sum()
        .sort_values(ascending=False)
    )

    # ==============================
    # GRÁFICO FINAL
    # ==============================
    # ==============================
    # GRÁFICO FINAL
    # ==============================
    if divergencias_por_supervisor.empty:
        print("ℹ️ Nenhuma divergência encontrada. Gráfico não gerado.")
    else:
        plt.figure(figsize=(10, 6))
        divergencias_por_supervisor.plot(kind="bar")
        plt.title("Total de Divergências por Supervisor")
        plt.xlabel("Supervisor")
        plt.ylabel("Fontes não baixadas")
        plt.xticks(rotation=45, ha="right")
        plt.tight_layout()
        plt.show()

    # ==============================
    # EXPORTAR RELATÓRIO
    # ==============================
    relatorio.to_excel(
        RELATORIO_DE_MISCELANEAS / "relatorio_baixa_fontes.xlsx"
    )

    print("✔ Relatório e gráfico gerados com sucesso")


# ==============================
# MAIN
# ==============================
if __name__ == "__main__":
    executar()
