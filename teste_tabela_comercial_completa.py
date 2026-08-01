from extrator_tabela_comercial import (
    calcular_metricas_tabela_comercial,
    extrair_tabela_comercial_completa,
)


TEXTO_PAGINA_25 = r"""
--- PÁGINA 25 OCR ---
ANEXO II - REMUNERAÇÃO
1. Pela prestação dos serviços, o EMISSOR pagará à ITAUCOR as seguintes quantias:
Valores em Reais (R$)
Custo de Implantação R$ 3.000,00
Custo fixos
Mensal R$ 4.000,00
Voto a distância isento
Taxa Mensal por acionista - Faixa
Até 5.000 acionistas R$ 0,75
De 5.001 a 15.000 acionistas R$ 0,70
De 15.001 a 35.000 acionistas R$ 0,65
De 35.001 a 70.000 acionistas R$ 0,60
De 70.001 a 120.000 acionistas R$ 0,55
De 120.001 a 220.000 acionistas R$ 0,50
De 200.001 a 300.000 acionistas R$ 0,45
Acima de 300.001 acionistas R$ 0,40
Eventos e Movimentações
Pagamento de Dividendos Clientes Itaú (acionistas em livro) R$ 2,55
Pagamento de Dividendos outros Bancos (acionistas em livro) R$ 2,55
Bonificação / desdobramento (acionistas em livro) R$ 1,40
Movimentação bolsa isento
Transferência/alteração cadastral/movimentação (acionistas livro) R$ 1,40
Subscrição:
Boletim emitido (acionistas em livro) R$ 1,40
Boletim efetivado (acionistas em livro) R$ 1,40
Emissão de avisos/extratos (acionistas em livro) R$ 1,40
Emissão de informes de rendimentos digital (acionistas bolsa e livro) R$ 1,40
Envio de correspondência (sob demanda do acionista)* Taxa Correio
Para informes de rendimentos digitais, será isenta a taxa de correio.
"""


itens = extrair_tabela_comercial_completa(TEXTO_PAGINA_25)

assert len(itens) == 21, f"Esperados 21 itens comerciais, encontrados {len(itens)}"
assert any(i["descricao"] == "Custo de Implantação" and i["valor_unitario"] == "R$ 3.000,00" for i in itens)
assert any(i["descricao"] == "Custo Fixo Mensal" and i["valor_unitario"] == "R$ 4.000,00" for i in itens)
assert any(i["descricao"] == "Voto a distância" and i["natureza_valor"] == "ISENTO" for i in itens)
assert any(i["faixa_condicao"] == "De 200.001 a 300.000 acionistas" and i["valor_unitario"] == "R$ 0,45" for i in itens)
assert any(i["natureza_valor"] == "REEMBOLSO" and "Correios" in i["valor_unitario"] for i in itens)

metricas = calcular_metricas_tabela_comercial(itens, itens)
assert metricas["itens_encontrados_documento"] == 21
assert metricas["itens_exibidos_auditor"] == 21
assert metricas["cobertura_tabela_percentual"] == 100

print("OK - Tabela comercial completa: 21/21 itens, cobertura 100%")
