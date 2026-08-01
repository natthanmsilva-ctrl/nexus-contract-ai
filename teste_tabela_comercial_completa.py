from extrator_tabela_comercial import (
    calcular_metricas_tabela_comercial,
    extrair_tabela_comercial_completa,
    mesclar_itens_comerciais,
)


# Texto reproduzido com os ruídos reais do OCR da página 25 do contrato SBF.
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
Até 5.000 acionistas R$075
De 5.001 a 15.000 acionistas R$070
De 15.001 a 35.000 acionistas R$065 |
De 35.001 a 70.000 acionistas R$0,60 5
De 70.001 a 120.000 acionistas R$055
De 120.001 a 200.000 acionistas R$050
De 200.001 a 300.000 acionistas R$045
Acima de 300.001 acionistas R$ 0,40
Eventos e Movimentações
Pagamento de Dividendos Clientes Itaú (acionistas em livro) R$2,55
Pagamento de Dividendos outros Bancos (acionistas em livro) R$2,55
Bonificação / desdobramento (acionistas em livro) | R$1,40
Movimentação bolsa isento
Transferêncialalteração cadastral/movimentação (acionistas livro) R$ 1,40
Subscrição:
Boletim emitido (acionistas em livro) R$1,40
Boletim efetivado (acionistas em livro) R$1,40
Emissão de avisos/extratos (acionistas em livro) R$ 1,40
Emissão de informes de rendimentos digital (acionistas bolsa e livro) R$1,40
Envio de correspondência (sob demanda do acionista) * Taxa Correio
*Para informes de rendimentos digitais, será isenta a taxa de correio.
"""


# Resumo equivalente que a IA pode devolver. As descrições são diferentes das
# células do OCR e, mesmo assim, devem ser reconciliadas em 21 linhas únicas.
DADOS_IA = [
    ("Taxa única de implantação dos serviços de escrituração", "IMPLANTACAO_UNICA", 3000.0),
    ("Mensalidade fixa para prestação dos serviços", "MENSAL_FIXO", 4000.0),
    ("Serviço de recepção e tratamento de instruções de voto a distância", "MENSAL_FIXO", 0.0),
    ("Tarifa mensal por acionista ativo na faixa até 5.000", "UNITARIO_VARIAVEL", 0.75),
    ("Tarifa mensal por acionista ativo na faixa de 5.001 a 15.000", "UNITARIO_VARIAVEL", 0.70),
    ("Tarifa mensal por acionista ativo na faixa de 15.001 a 35.000", "UNITARIO_VARIAVEL", 0.65),
    ("Tarifa mensal por acionista ativo na faixa de 35.001 a 70.000", "UNITARIO_VARIAVEL", 0.60),
    ("Tarifa mensal por acionista ativo na faixa de 70.001 a 120.000", "UNITARIO_VARIAVEL", 0.55),
    ("Tarifa mensal por acionista ativo na faixa de 120.001 a 200.000", "UNITARIO_VARIAVEL", 0.50),
    ("Tarifa mensal por acionista ativo na faixa de 200.001 a 300.000", "UNITARIO_VARIAVEL", 0.45),
    ("Tarifa mensal por acionista ativo na faixa acima de 300.001", "UNITARIO_VARIAVEL", 0.40),
    ("Tarifa por pagamento de dividendos para acionistas em livro clientes Itaú", "UNITARIO_VARIAVEL", 2.55),
    ("Tarifa por pagamento de dividendos para acionistas em livro clientes de outros bancos", "UNITARIO_VARIAVEL", 2.55),
    ("Tarifa por evento de bonificação ou desdobramento para acionistas em livro", "UNITARIO_VARIAVEL", 1.40),
    ("Tarifa para movimentação em bolsa", "UNITARIO_VARIAVEL", 0.0),
    ("Tarifa por transferência, alteração cadastral ou movimentação de acionistas em livro", "UNITARIO_VARIAVEL", 1.40),
    ("Tarifa por emissão de boletim de subscrição", "UNITARIO_VARIAVEL", 1.40),
    ("Tarifa por efetivação de boletim de subscrição", "UNITARIO_VARIAVEL", 1.40),
    ("Tarifa por emissão de avisos ou extratos para acionistas em livro", "UNITARIO_VARIAVEL", 1.40),
    ("Tarifa por emissão de informes de rendimentos digitais para acionistas bolsa e livro", "UNITARIO_VARIAVEL", 1.40),
    ("Tarifa de envio de correspondência sob demanda do acionista", "REEMBOLSO", "Não localizado"),
]

itens_ia = [
    {
        "Item": str(idx),
        "Descrição": descricao,
        "Natureza do valor": natureza,
        "Valor unitário": valor,
        "Grupo/Tabela": "Não localizado",
    }
    for idx, (descricao, natureza, valor) in enumerate(DADOS_IA, 1)
]

itens_documento = extrair_tabela_comercial_completa(TEXTO_PAGINA_25)
assert len(itens_documento) == 21, f"Esperados 21 itens comerciais, encontrados {len(itens_documento)}"
assert any(i["descricao"] == "Custo de Implantação" and i["valor_unitario"] == "R$ 3.000,00" for i in itens_documento)
assert any(i["descricao"] == "Custo Fixo Mensal" and i["valor_unitario"] == "R$ 4.000,00" for i in itens_documento)
assert any(i["descricao"] == "Voto a distância" and i["natureza_valor"] == "ISENTO" for i in itens_documento)
assert any(i["faixa_condicao"] == "De 35.001 a 70.000 acionistas" and i["valor_unitario"] == "R$ 0,60" for i in itens_documento)
assert any(i["faixa_condicao"] == "De 200.001 a 300.000 acionistas" and i["valor_unitario"] == "R$ 0,45" for i in itens_documento)
assert any(i["natureza_valor"] == "REEMBOLSO" and "Correios" in i["valor_unitario"] for i in itens_documento)

# Parser documental primeiro: a IA completa lacunas, mas não duplica nem altera
# a natureza confirmada das linhas de voto a distância e movimentação em bolsa.
itens_conciliados = mesclar_itens_comerciais(itens_documento, itens_ia)
assert len(itens_conciliados) == 21, f"A conciliação gerou {len(itens_conciliados)} itens; esperado: 21"
assert any(i.get("descricao") == "Voto a distância" and i.get("natureza_valor") == "ISENTO" for i in itens_conciliados)
assert any(i.get("descricao") == "Movimentação bolsa" and i.get("natureza_valor") == "ISENTO" for i in itens_conciliados)
variaveis = [
    i for i in itens_conciliados
    if str(i.get("natureza_valor") or i.get("Natureza do valor") or "").upper()
    in {"UNITARIO_VARIAVEL", "PERCENTUAL_VARIAVEL", "REEMBOLSO"}
]
assert len(variaveis) == 17, f"Esperadas 17 tarifas/condições variáveis, encontradas {len(variaveis)}"

metricas = calcular_metricas_tabela_comercial(itens_documento, itens_conciliados)
assert metricas["itens_encontrados_documento"] == 21
assert metricas["itens_exibidos_auditor"] == 21
assert metricas["cobertura_tabela_percentual"] == 100
assert metricas["divergencia_quantidade"] == 0

print("OK - OCR real conciliado: 21/21 itens, sem duplicidades, cobertura 100%")
