import pandas as pd
from paths import REGRAS_EQUIPAMENTOS,REGRAS_ESCALA,RELATORIO_DE_EQUIPAMENTOS,SAIDA

def executar():

    kit_minimo = pd.read_excel(REGRAS_EQUIPAMENTOS / "kit_minimo.xlsx")
    tecnicos = pd.read_excel(REGRAS_ESCALA / "escala_almoxarifado.xlsx")
    campo = pd.read_excel(RELATORIO_DE_EQUIPAMENTOS / "equipamentos_em_campo.xlsx", sheet_name="Resumo")

    for df in [kit_minimo, tecnicos, campo]:
        df.columns = df.columns.str.strip().str.upper()

    campo = campo.rename(columns={"DESCRICAO": "EQUIPAMENTO"})

    kit_minimo["TECNOLOGIA"] = kit_minimo["TECNOLOGIA"].str.strip().str.upper()
    kit_minimo["DESCRICAO"] = kit_minimo["DESCRICAO"].str.strip().str.upper()

    tecnicos["TECNICO"] = tecnicos["TECNICO"].str.strip().str.upper()
    tecnicos["SKILL"] = tecnicos["SKILL"].str.strip().str.upper()
    tecnicos["ALMOXARIFADO"] = tecnicos["ALMOXARIFADO"].str.strip().str.upper()

    campo["TECNICO"] = campo["TECNICO"].str.strip().str.upper()
    campo["EQUIPAMENTO"] = campo["EQUIPAMENTO"].str.strip().str.upper()

    # 🔹 Normalização de acentos
    tecnicos["ALMOXARIFADO"] = (
        tecnicos["ALMOXARIFADO"]
        .str.replace("Ç", "C", regex=False)
        .str.replace("Á", "A", regex=False)
        .str.replace("Ã", "A", regex=False)
        .str.replace("É", "E", regex=False)
    )

    tecnicos_hoje = tecnicos.copy()

    # 🔹 Montagem da separação
    linhas = []

    # 🔹 Remove linhas inválidas (NaN)
    tecnicos_hoje = tecnicos_hoje[
        tecnicos_hoje["SKILL"].notna() &
        tecnicos_hoje["TECNICO"].notna()
        ]

    for _, tecnico in tecnicos_hoje.iterrows():
        skills = str(tecnico["SKILL"]).split("+")

        for skill in skills:
            skill = skill.strip()

            kit = kit_minimo[kit_minimo["TECNOLOGIA"] == skill.strip()]

            for _, item in kit.iterrows():
                filtro = (
                    (campo["TECNICO"] == tecnico["TECNICO"]) &
                    (campo["EQUIPAMENTO"] == item["DESCRICAO"])
                )

                qtd_em_campo = campo.loc[filtro, "QUANTIDADE"].sum()
                qtd_separar = max(item["QTD_KIT"] - qtd_em_campo, 0)

                if qtd_separar > 0:
                    linhas.append({
                        "TECNICO": tecnico["TECNICO"],
                        "TECNOLOGIA": skill.strip(),
                        "EQUIPAMENTO": item["DESCRICAO"],
                        "QUANTIDADE": qtd_separar
                    })

    df_final = pd.DataFrame(linhas)

    # 🔹 Contar técnicos únicos
    tecnicos_unicos = df_final["TECNICO"].unique()

    separadores = ["GUILHERME", "FABRICIO", "ERICK"]

    # 🔹 Mapear técnico → separador
    mapa_separadores = {
        tecnico: separadores[i % len(separadores)]
        for i, tecnico in enumerate(tecnicos_unicos)
    }

    df_final["SEPARADOR"] = df_final["TECNICO"].map(mapa_separadores)

    # 🔹 Gerar Excel com 5 abas
    arquivo_saida = SAIDA / "separacao_de_equipamentos.xlsx"

    with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
        for separador in separadores:
            aba = df_final[df_final["SEPARADOR"] == separador]

            # Se não tiver nada pra pessoa, cria aba vazia mesmo
            aba.to_excel(
                writer,
                sheet_name=separador.title(),
                index=False
            )

    print("✔ Separação gerada com sucesso!")
    print(f"✔ Total de técnicos: {len(tecnicos_unicos)}")
    print(f"✔ Arquivo criado com {len(separadores)} abas (uma por separador)")


if __name__ == "__main__":
    executar()
