import pandas as pd
from paths import ENTRADA_EQUIPAMENTOS, RELATORIO_DE_EQUIPAMENTOS

def executar():

    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)

    df = pd.read_excel(ENTRADA_EQUIPAMENTOS / "relatorio_equipamento.xlsx")

    df_campo = df[
        df["STATUS"].isin(["Com Técnico", "confirmação do técnico"])
    ]

    colunas_desejadas = [
        'TECNICO',
        'DESCRICAO',
    ]

    df_equipamentos_em_campo = df_campo.loc[:, colunas_desejadas]

    resumo_por_tecnico = (
        df_equipamentos_em_campo
        .groupby(['TECNICO', 'DESCRICAO'], as_index=False)
        .size()
        .rename(columns={'size': 'QUANTIDADE'})
    )

    tabela_pivot = resumo_por_tecnico.pivot_table(
        index='TECNICO',
        columns='DESCRICAO',
        values='QUANTIDADE',
        fill_value=0
    )

    with pd.ExcelWriter(RELATORIO_DE_EQUIPAMENTOS / "equipamentos_em_campo.xlsx", engine="openpyxl") as writer:
        df_equipamentos_em_campo.to_excel(
            writer,
            sheet_name="Detalhado",
            index=False
        )
        resumo_por_tecnico.to_excel(
            writer,
            sheet_name="Resumo",
            index=False
        )
        tabela_pivot.to_excel(
            writer,
            sheet_name="Por Técnico"
        )


if __name__ == "__main__":
    executar()