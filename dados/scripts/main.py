from visualizar_equipamentos_campo import executar as visualizar_equipamentos_campo
from tratar_tecnicos import executar as tratar_tecnicos
from gerar_kit_tecnico import executar as gerar_kit_tecnico
from gerar_separacao import executar as gerar_separacao
from recebimento_de_nf_serializada import executar as recebimento_de_nf_serializada

def main():
    visualizar_equipamentos_campo()
    tratar_tecnicos()
    gerar_kit_tecnico()
    gerar_separacao()
    # recebimento_de_nf_serializada()

if __name__ == "__main__":
    main()