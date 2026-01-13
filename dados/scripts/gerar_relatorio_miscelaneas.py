import pandas as pd
from datetime import datetime
from paths import ENTRADA_MISCELANEA,SAIDA,RELATORIO_DE_RESERVAS,RELATORIO_DE_MISCELANEAS
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)

# =============================
# DATA DE REFERÊNCIA
# =============================
HOJE = pd.to_datetime(datetime.today().date())

# =============================
# EXECUÇÃO PRINCIPAL
# =============================
def executar():

    # -------------------------
    # 1. LEITURA DOS ARQUIVOS
    # -------------------------
    estoque = pd.read_excel(ENTRADA_MISCELANEA / "estoque_miscelanea.xlsx")
    separacao = pd.read_excel(SAIDA / "separacao_miscelanea.xlsx")
    reservas = pd.read_excel(RELATORIO_DE_RESERVAS / "reservas_porto_alegre_procisa.xlsx")

    # -------------------------
    # 2. PADRONIZAÇÃO
    # -------------------------
    for df in (estoque, separacao):
        df["descricao"] = df["descricao"].str.upper().str.strip()

    reservas["descricao"] = (
        reservas["TEXTO_BREVE_MATERIAL"]
        .str.upper()
        .str.strip()
    )

    # -------------------------
    # 3. CONSOLIDAR ESTOQUE ATUAL
    # -------------------------
    estoque_base = (
        estoque
        .groupby("descricao", as_index=False)
        .agg(estoque_atual=("quantidade", "sum"))
    )

    # -------------------------
    # 4. CONSOLIDAR DEMANDA (SEPARAÇÃO)
    # -------------------------
    demanda = (
        separacao
        .groupby("descricao", as_index=False)
        .agg(qtd_repor=("qtd_entregar", "sum"))
    )

    # -------------------------
    # 5. TRATAR DATAS DAS RESERVAS
    # -------------------------
    for col in ["DT_EXPEDICAO", "DT_LIMITE_ORIGINAL"]:
        reservas[col] = pd.to_datetime(reservas[col], errors="coerce")

    def classificar_reserva(row):
        if pd.notna(row["DT_EXPEDICAO"]):
            return "CHEGANDO_NO_PRAZO"
        if pd.isna(row["DT_LIMITE_ORIGINAL"]):
            return "SEM_DATA"
        if HOJE <= row["DT_LIMITE_ORIGINAL"]:
            return "A_CAMINHO"
        return "ATRASADA"

    reservas["STATUS_ENTREGA"] = reservas.apply(classificar_reserva, axis=1)

    # -------------------------
    # 6. VALIDAR QUANTIDADE DE RESERVA
    # -------------------------
    def qtd_validada(row):
        if row["STATUS_ENTREGA"] == "CHEGANDO_NO_PRAZO":
            return row["QUANTIDADE"]
        if row["STATUS_ENTREGA"] == "A_CAMINHO":
            return row["QUANTIDADE"] * 0.5
        return 0

    reservas["QTD_VALIDADA"] = reservas.apply(qtd_validada, axis=1)

    reservas_consolidadas = (
        reservas
        .groupby("descricao", as_index=False)
        .agg(qtd_reserva_validada=("QTD_VALIDADA", "sum"))
    )

    # -------------------------
    # 7. CONSOLIDAR BASE FINAL
    # -------------------------
    base = (
        estoque_base
        .merge(demanda, on="descricao", how="left")
        .merge(reservas_consolidadas, on="descricao", how="left")
    )

    base["qtd_repor"] = base["qtd_repor"].fillna(0)
    base["qtd_reserva_validada"] = base["qtd_reserva_validada"].fillna(0)

    # -------------------------
    # 8. MÉTRICAS DE RISCO
    # -------------------------
    base["saldo_projetado"] = (
        base["estoque_atual"]
        + base["qtd_reserva_validada"]
        - base["qtd_repor"]
    )

    base["indice_risco"] = base.apply(
        lambda row: row["qtd_repor"] / max(
            row["estoque_atual"] + row["qtd_reserva_validada"], 1
        ),
        axis=1
    )

    # -------------------------
    # 9. CLASSIFICAÇÃO
    # -------------------------
    def classificar_status(indice):
        if indice >= 1:
            return "🔴 CRÍTICO"
        if indice >= 0.7:
            return "🟠 ALERTA"
        if indice >= 0.4:
            return "🟡 ATENÇÃO"
        return "🟢 OK"

    base["STATUS"] = base["indice_risco"].apply(classificar_status)

    # -------------------------
    # 10. EXPORTAR RELATÓRIO
    # -------------------------
    RELATORIO_DE_MISCELANEAS.mkdir(exist_ok=True)

    base.sort_values("indice_risco", ascending=False).to_excel(
        RELATORIO_DE_MISCELANEAS / "relatorio_risco_estoque.xlsx",
        index=False
    )

    print("✔ Relatório preditivo de risco de estoque gerado com sucesso")


# =============================
# ENTRY POINT
# =============================
if __name__ == "__main__":
    executar()
