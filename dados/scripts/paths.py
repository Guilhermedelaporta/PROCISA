from pathlib import Path

BASE = Path(__file__).resolve().parent.parent

REGRAS_MISCELANEAS = BASE / "regras" / "kit_miscelanea"
REGRAS_EQUIPAMENTOS = BASE / "regras" / "kit_equipamento"
REGRAS_ESCALA = BASE / "regras" / "escala_de_atendimento_do_dia"
ENTRADA_ESCALA = BASE / "entrada" / "escala_do_dia"
ENTRADA_RESERVA = BASE / "entrada" / "reserva"
ENTRADA_EQUIPAMENTOS = BASE / "entrada" / "equipamentos_em_campo"
ENTRADA_MISCELANEA = BASE / "entrada" / "estoque_miscelanea"
RELATORIO_DE_EQUIPAMENTOS = BASE / "relatorios" / "equipamentos"
RELATORIO_DE_MISCELANEAS = BASE / "relatorios" / "miscelaneas"
RELATORIO_DE_RESERVAS = BASE / "relatorios" / "reservas"
MISCELANEA_EM_CAMPO = BASE / "entrada" / "miscelaneas_em_campo"
SAIDA = BASE / "saida"
BAIXAS_DE_MISCELANEAS = BASE / "entrada" / "baixa_de_miscelaneas"
ENTRADA_ESCALA_COMPLETA = BASE / "entrada" / "escala_do_dia"

