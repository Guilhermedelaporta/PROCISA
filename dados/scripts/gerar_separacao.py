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

    pd.DataFrame(linhas).to_excel(
        SAIDA / f"separacao_de_equipamentos.xlsx",
        index=False
    )

    print(f"✔ Separação gerada com sucesso para!")

if __name__ == "__main__":
    executar()
