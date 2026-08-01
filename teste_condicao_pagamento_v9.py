from auditor_evidencias import _extrair_condicao_pagamento_dd_documental

casos = [
    ({1: "Pagamento da fatura em até 60 dias corridos após o recebimento."}, "60DD"),
    ({1: "A nota fiscal terá vencimento no prazo de 30 (trinta) dias úteis após aprovação."}, "30DD"),
    ({1: "Mensalmente será remetida fatura, com vencimento até o dia 15 (quinze) do mês subsequente."}, "15DD"),
]

for paginas, esperado in casos:
    valor, pagina, evidencia = _extrair_condicao_pagamento_dd_documental(paginas)
    assert valor == esperado, (valor, esperado, evidencia)
    assert pagina == "1"
    assert evidencia

# Não pode confundir aviso prévio com condição de pagamento.
valor, _, _ = _extrair_condicao_pagamento_dd_documental({1: "O contrato poderá ser denunciado com aviso prévio de 30 dias."})
assert valor == ""

print("OK - Condição de pagamento padronizada em 15DD/30DD/60DD sem capturar outros prazos")
