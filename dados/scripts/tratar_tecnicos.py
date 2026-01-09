def executar():

    import pandas as pd
    from pathlib import Path

    BASE = Path(__file__).parent.parent

    ENTRADA = BASE / "entrada"
    REGRAS = BASE / "regras"
    SAIDA = BASE / "saida"


    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)

    # Ler arquivo de escala
    escala = pd.read_excel(ENTRADA / "escala.xlsx")
    print("COLUNAS DISPONÍVEIS NO EXCEL:")
    print(escala.columns.tolist())

    # Renomear colunas para padrão do sistema
    escala = escala.rename(columns={
        'NOME': 'TECNICO',
        'LOGIN': 'LOGIN_CLARO',
        'RE': 'RE_PROCISA',
        'SUPORTE DE PRODUÇÃO': 'SUPORTE_PRODUCAO'
    })
    print("COLUNAS DISPONÍVEIS NO EXCEL:")
    print(escala.columns.tolist())

    # Manter somente colunas necessárias
    base_tecnicos = escala[
        ['TECNICO', 'LOGIN_CLARO', 'RE_PROCISA', 'SUPORTE_PRODUCAO', 'SKILL', 'ALMOXARIFADO']
    ]

    # Padronização (PASSO CRÍTICO)
    base_tecnicos['TECNICO'] = base_tecnicos['TECNICO'].str.strip().str.upper()
    base_tecnicos['LOGIN'] = base_tecnicos['LOGIN_CLARO'].str.strip().str.lower()
    base_tecnicos['SUPORTE DE PRODUÇAO'] = base_tecnicos['SUPORTE_PRODUCAO'].str.strip().str.upper()
    base_tecnicos['SKILL'] = base_tecnicos['SKILL'].str.strip().str.upper()
    base_tecnicos['ALMOXARIFADO'] = base_tecnicos['ALMOXARIFADO'].str.strip().str.upper()

    # Remover duplicados
    base_tecnicos = base_tecnicos.drop_duplicates()

    # Exportar base final
    base_tecnicos.to_excel(REGRAS /
        "escala_almoxarifado.xlsx",
        index=False
    )

    print("base_tecnicos.xlsx gerado com sucesso!")
    print(base_tecnicos.head())


if __name__ == "__main__":
    executar()