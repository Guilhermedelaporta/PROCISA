import pandas as pd
from paths import ENTRADA_MISCELANEA,REGRAS_MISCELANEAS,RELATORIO_DE_MISCELANEAS,SAIDA,MISCELANEA_EM_CAMPO

# =============================
# CONFIGURAÇÕES
# =============================
LIMITE_MINIMO_ESTOQUE = 100

# =============================
# REGRAS DE CABOS (BOBINAS)
# =============================
REGRAS_CABOS = {
    "CABO COAXIAL RG6 TRISH SEM MENSAG BRANCO": {"metade": 150},
    "CABO COAXIAL RG6 TRISH COM MENSAG PRETO": {"metade": 152.5},
    "CABO DROP FO LOW F FIG0 BRANCO": {"metade": 150},
    "CABO DROP 1FO LOW F FIG8 LOW CINZA": {"metade": 250},
}

# =============================
# EXECUÇÃO PRINCIPAL
# =============================
def executar():

    # -------------------------
    # 1. LEITURA DOS ARQUIVOS
    # -------------------------
    df_mov = pd.read_excel(MISCELANEA_EM_CAMPO / "movimentacao_tecnico.xlsx")
    df_kit = pd.read_excel(REGRAS_MISCELANEAS / "kit_minimo_de_miscelanea.xlsx")
    df_estoque = pd.read_excel(ENTRADA_MISCELANEA / "estoque_de_miscelanea.xlsx")

    # -------------------------
    # 2. PADRONIZAÇÃO
    # -------------------------
    for df in (df_mov, df_kit, df_estoque):
        df["descricao"] = df["descricao"].astype(str).str.upper().str.strip()

    # -------------------------
    # 3. AJUSTAR KIT MÍNIMO
    # -------------------------
    # 🔴 SUA PLANILHA TEM "quantidade", NÃO "QTD_KIT"
    if "QTD_KIT" not in df_kit.columns:
        df_kit = df_kit.rename(columns={"quantidade": "QTD_KIT"})

    df_kit["QTD_KIT"] = df_kit["QTD_KIT"].fillna(0)

    # -------------------------
    # 4. CONSOLIDAR ESTOQUE DO TÉCNICO
    # -------------------------
    estoque_tecnico = (
        df_mov
        .groupby(["tecnico", "descricao"], as_index=False)
        .agg({"quantidade": "sum"})
        .rename(columns={"quantidade": "QTD_TECNICO"})
    )

    # -------------------------
    # 5. CRUZAR COM KIT MÍNIMO
    # -------------------------
    base = estoque_tecnico.merge(
        df_kit[["descricao", "QTD_KIT"]],
        on="descricao",
        how="left"
    )

    base["QTD_KIT"] = base["QTD_KIT"].fillna(0)

    # -------------------------
    # 6. CALCULAR NECESSIDADE
    # -------------------------
    def calcular_qtd_repor(row):
        descricao = row["descricao"]
        qtd_tecnico = row["QTD_TECNICO"]

        # 👉 REGRA ESPECIAL PARA CABOS
        if descricao in REGRAS_CABOS:
            metade_bobina = REGRAS_CABOS[descricao]["metade"]

            # Entrega somente se tiver MENOS que metade
            return 1 if qtd_tecnico < metade_bobina else 0

        # 👉 REGRA PADRÃO DE MISCELÂNEA
        return max(row["QTD_KIT"] - qtd_tecnico, 0)

    base["qtd_repor"] = base.apply(calcular_qtd_repor, axis=1)

    # -------------------------
    # 7. NECESSIDADE TOTAL POR ITEM
    # -------------------------
    आवश्यकता_total = (
        base
        .groupby("descricao", as_index=False)
        .agg({"qtd_repor": "sum"})
    )

    # -------------------------
    # 8. CRUZAR COM ESTOQUE DO ALMOX
    # -------------------------
    necessidade_estoque = आवश्यकता_total.merge(
        df_estoque[["descricao", "quantidade"]],
        on="descricao",
        how="left"
    ).rename(columns={"quantidade": "qtd_estoque"})

    necessidade_estoque["qtd_estoque"] = necessidade_estoque["qtd_estoque"].fillna(0)

    necessidade_estoque["qtd_separavel"] = necessidade_estoque[
        ["qtd_repor", "qtd_estoque"]
    ].min(axis=1)

    # -------------------------
    # 9. SEPARAÇÃO POR TÉCNICO
    # -------------------------
    separacao = base.merge(
        necessidade_estoque[["descricao", "qtd_separavel"]],
        on="descricao",
        how="left"
    )

    separacao["qtd_entregar"] = separacao[
        ["qtd_repor", "qtd_separavel"]
    ].min(axis=1)

    separacao_final = separacao[
        separacao["qtd_entregar"] > 0
    ][["tecnico", "descricao", "qtd_entregar"]]

    separacao_final.to_excel(
        SAIDA / "separacao_miscelanea.xlsx",
        index=False
    )

    # -------------------------
    # 10. ALERTA DE ESTOQUE BAIXO
    # -------------------------
    estoque_pos = necessidade_estoque.copy()

    estoque_pos["estoque_restante"] = (
        estoque_pos["qtd_estoque"] - estoque_pos["qtd_separavel"]
    )

    alerta = estoque_pos[
        estoque_pos["estoque_restante"] <= LIMITE_MINIMO_ESTOQUE
    ][["descricao", "estoque_restante"]]

    alerta.to_excel(
        RELATORIO_DE_MISCELANEAS / "alerta_estoque_baixo.xlsx",
        index=False
    )

    # -------------------------
    # 11. LOG FINAL
    # -------------------------
    print("✔ Separação de miscelânea e cabos gerada com sucesso")
    print("⚠️ Relatório de alerta de estoque baixo gerado")


# =============================
# ENTRY POINT
# =============================
if __name__ == "__main__":
    executar()
