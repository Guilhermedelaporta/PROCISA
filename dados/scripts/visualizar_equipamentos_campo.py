def executar():

    import pandas as pd
    from pathlib import Path

    BASE = Path(__file__).parent.parent

    ENTRADA = BASE / "entrada"
    RELATORIO = BASE / "relatorios"
    SAIDA = BASE / "saida"


    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)

    # Ler arquivo
    df = pd.read_excel(ENTRADA / "relatorio_equipamento.xlsx")

    # Filtrar linhas corretamente
    df_campo = df[
        df["STATUS"].isin(["Com Técnico", "confirmação do técnico"])
    ]

    # Manter somente as colunas necessárias
    colunas_desejadas = [
        'TECNICO',
        'DESCRICAO',
    ]

    df_equipamentos_em_campo = df_campo.loc[:, colunas_desejadas]

    # Conferência final
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

    with pd.ExcelWriter(RELATORIO / "equipamentos_em_campo.xlsx", engine="openpyxl") as writer:
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