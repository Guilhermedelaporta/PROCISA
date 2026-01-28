import pandas as pd

LOCAL_CORRETO = "PROCISA DO BRASIL PROJETOS, CONSTRU"
ESTADO_CORRETO = "INICIALIZADO"


def executar():
    # =============================
    # Leitura da NF
    # =============================
    nf_raw = pd.read_excel("relatorio_nf.xlsx", header=None)

    # localizar linha do cabeçalho
    linha_header = None
    for i, row in nf_raw.iterrows():
        if row.astype(str).str.contains("Endereçável Principal", case=False).any():
            linha_header = i
            break

    if linha_header is None:
        raise Exception("❌ Cabeçalho da NF não encontrado")

    # promover cabeçalho
    nf = nf_raw.iloc[linha_header + 1:].copy()
    nf.columns = nf_raw.iloc[linha_header]

    # normalização
    nf.columns = (
        nf.columns
        .astype(str)
        .str.strip()
        .str.upper()
    )

    nf = nf.dropna(axis=1, how="all")

    print("Colunas encontradas:")
    print(nf.columns.tolist())
    print(nf.head())

    # =============================
    # Leitura dos seriais físicos
    # =============================
    fisico = pd.read_excel("seriais_fisicos.xlsx")

    # =============================
    # Padronização
    # =============================
    for df in [nf, fisico]:
        df.columns = df.columns.str.strip().str.upper()

    nf["ENDEREÇÁVEL PRINCIPAL"] = nf["ENDEREÇÁVEL PRINCIPAL"].astype(str).str.strip()
    nf["LOCAL"] = nf["LOCAL"].astype(str).str.strip().str.upper()
    nf["ESTADO EQUIPAMENTO"] = nf["ESTADO EQUIPAMENTO"].astype(str).str.strip().str.upper()
    fisico["SERIAL"] = fisico["SERIAL"].astype(str).str.strip()

    # =============================
    # Comparações
    # =============================
    seriais_nf = set(nf["ENDEREÇÁVEL PRINCIPAL"])
    seriais_fisico = set(fisico["SERIAL"])

    fora_da_nf = fisico[~fisico["SERIAL"].isin(seriais_nf)]
    nao_recebidos = nf[~nf["ENDEREÇÁVEL PRINCIPAL"].isin(seriais_fisico)]

    local_invalido = nf[nf["LOCAL"] != LOCAL_CORRETO]
    estado_invalido = nf[nf["ESTADO EQUIPAMENTO"] != ESTADO_CORRETO]

    # =============================
    # Relatório de inconsistências
    # =============================
    if (
        not fora_da_nf.empty
        or not nao_recebidos.empty
        or not local_invalido.empty
        or not estado_invalido.empty
    ):
        with pd.ExcelWriter("relatorio_inconsistencias.xlsx") as writer:

            if not fora_da_nf.empty:
                fora_da_nf.assign(MOTIVO="Serial não consta na NF") \
                    .to_excel(writer, sheet_name="Serial fora da NF", index=False)

            if not nao_recebidos.empty:
                nao_recebidos[["ENDEREÇÁVEL PRINCIPAL"]] \
                    .rename(columns={"ENDEREÇÁVEL PRINCIPAL": "SERIAL"}) \
                    .assign(MOTIVO="Não recebido fisicamente") \
                    .to_excel(writer, sheet_name="Não recebido", index=False)

            if not local_invalido.empty:
                local_invalido.assign(MOTIVO="Local incorreto") \
                    .to_excel(writer, sheet_name="Local incorreto", index=False)

            if not estado_invalido.empty:
                estado_invalido.assign(MOTIVO="Estado inválido") \
                    .to_excel(writer, sheet_name="Estado inválido", index=False)

        print("❌ Inconsistências encontradas.")
        return

    # =============================
    # Entrada Massiva
    # =============================
    entrada = nf[[
        "ENDEREÇÁVEL PRINCIPAL",
        "MATERIAL SAP",
        "DMT"
    ]].copy()

    entrada.columns = ["serial", "codigo", "obs"]
    entrada.insert(2, "mac", "")
    entrada.insert(4, "novo", "")

    entrada.to_excel("entrada_massiva_connect.xlsx", index=False)

    print("✅ Entrada massiva gerada com sucesso!")


if __name__ == "__main__":
    executar()
