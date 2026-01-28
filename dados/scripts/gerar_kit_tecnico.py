import pandas as pd
from paths import REGRAS_EQUIPAMENTOS

def executar():

    dados_kit = [
        {'TECNOLOGIA': 'HFC', 'DESCRICAO': '3.1 DUAL BAND (200mb em diante)', 'QTD_KIT': 2},
        {'TECNOLOGIA': 'HFC', 'DESCRICAO': 'DECODER HD CABO IP  VOZ (TV BOX HFC IP)', 'QTD_KIT': 1},
        {'TECNOLOGIA': 'HFC', 'DESCRICAO': '3.1 WiFI6', 'QTD_KIT': 2},
        {'TECNOLOGIA': 'HFC', 'DESCRICAO': '3.0 DUAL BAND (01mb a 199mb)', 'QTD_KIT': 0},
        {'TECNOLOGIA': 'HFC', 'DESCRICAO': 'SMART', 'QTD_KIT': 0},
        {'TECNOLOGIA': 'HFC', 'DESCRICAO': 'DECODER 3P NG (LED AZUL)', 'QTD_KIT': 2},
        {'TECNOLOGIA': 'HFC', 'DESCRICAO':'CHIP CLARO', 'QTD_KIT': 3},
        {'TECNOLOGIA':'HFC', 'DESCRICAO': 'DECODER 4K STREAMING', 'QTD_KIT': 1},
        {'TECNOLOGIA':'HFC', 'DESCRICAO': 'DECODER HD 4K', 'QTD_KIT': 0},

        {'TECNOLOGIA': 'GPON', 'DESCRICAO': 'ONT - GPOM - WIFI6', 'QTD_KIT': 5},
        {'TECNOLOGIA': 'GPON', 'DESCRICAO': 'ONT - GPON', 'QTD_KIT': 8},
        {'TECNOLOGIA': 'GPON', 'DESCRICAO': 'ROTEADOR', 'QTD_KIT': 0},

        {'TECNOLOGIA': 'VT', 'DESCRICAO': '3.1 DUAL BAND (200mb em diante)', 'QTD_KIT': 3},
        {'TECNOLOGIA': 'VT', 'DESCRICAO': 'DECODER HD CABO IP  VOZ (TV BOX HFC IP)', 'QTD_KIT': 1},
        {'TECNOLOGIA': 'VT', 'DESCRICAO': '3.1 WiFI6', 'QTD_KIT': 3},
        {'TECNOLOGIA': 'VT', 'DESCRICAO': '3.0 DUAL BAND (01mb a 199mb)', 'QTD_KIT': 0},
        {'TECNOLOGIA': 'VT', 'DESCRICAO': 'SMART', 'QTD_KIT': 0},
        {'TECNOLOGIA': 'VT', 'DESCRICAO': 'DECODER 3P NG (LED AZUL)', 'QTD_KIT': 3},
        {'TECNOLOGIA': 'VT', 'DESCRICAO':'CHIP CLARO', 'QTD_KIT': 5},
        {'TECNOLOGIA':'VT', 'DESCRICAO': 'DECODER 4K STREAMING', 'QTD_KIT': 1},
        {'TECNOLOGIA':'VT', 'DESCRICAO': 'DECODER HD 4K', 'QTD_KIT': 0},
        {'TECNOLOGIA': 'VT', 'DESCRICAO': 'ONT - GPOM - WIFI6', 'QTD_KIT': 3},
        {'TECNOLOGIA': 'VT', 'DESCRICAO': 'ONT - GPON', 'QTD_KIT': 8},
        {'TECNOLOGIA': 'VT', 'DESCRICAO': 'ROTEADOR', 'QTD_KIT': 0},
        ]

    kit_minimo = pd.DataFrame(dados_kit)

    kit_minimo.to_excel(REGRAS_EQUIPAMENTOS / "kit_equipamento.xlsx", index=False)

    print("✔ Kit Mínimo gerado com sucesso")

if __name__ == "__main__":
    executar()