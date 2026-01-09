from pathlib import Path
import pandas as pd
from dados.scripts.reservas import RELATORIO


def executar(dia_alvo=None):
    from datetime import datetime

    DEBUG = True

    BASE = Path(__file__).resolve().parent.parent
    REGRAS = BASE / "regras"
    SAIDA = BASE / "saida"
    SAIDA.mkdir(exist_ok=True)

    # 🔹 Ler arquivos
    kit_minimo = pd.read_excel(REGRAS / "kit_minimo.xlsx")
    tecnicos = pd.read_excel(REGRAS / "escala_almoxarifado.xlsx")
    campo = pd.read_excel(RELATORIO / "equipamentos_em_campo.xlsx", sheet_name="Resumo")

    # 🔹 Padronização
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

    # 🔹 Dia alvo
    dias_semana = {
        0: "SEGUNDA",
        1: "TERÇA",
        2: "QUARTA",
        3: "QUINTA",
        4: "SEXTA",
        5: "SÁBADO",
        6: "DOMINGO"
    }

    DIA_ALVO = dias_semana[datetime.today().weekday()] if dia_alvo is None else dia_alvo.upper()
    DIA_ALVO = (
        DIA_ALVO
        .replace("Ç", "C")
        .replace("Á", "A")
        .replace("Ã", "A")
        .replace("É", "E")
    )

    if DEBUG:
        print(f"Gerando separação para o dia: {DIA_ALVO}")

    # 🔹 Filtro correto
    tecnicos_hoje = tecnicos[
        tecnicos["ALMOXARIFADO"].str.contains(
            rf"(?:^|[/, ]){DIA_ALVO}(?:$|[/, ])",
            regex=True,
            na=False
        )
    ]

    if DEBUG:
        print("=== TÉCNICOS SELECIONADOS ===")
        print(tecnicos_hoje[["TECNICO", "ALMOXARIFADO"]])

    # 🔹 Montagem da separação
    linhas = []

    for _, tecnico in tecnicos_hoje.iterrows():
        for skill in tecnico["SKILL"].split("+"):
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
        SAIDA / f"separacao_equipamentos_{DIA_ALVO}.xlsx",
        index=False
    )

    print(f"✔ Separação gerada com sucesso para {DIA_ALVO}!")


if __name__ == "__main__":
    executar("QUARTA")
