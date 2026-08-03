from auditor_evidencias import _extrair_condicao_pagamento_dd_documental

# A página operacional vem antes e contém a palavra pagamento em outro contexto.
paginas = {
    18: (
        "Valores relativos ao pagamento dos créditos. A ITAUCOR disponibilizará ao EMISSOR, "
        "em até 5 (cinco) dias úteis, mediante solicitação, a relação dos acionistas."
    ),
    26: (
        "Mensalmente, a ITAUCOR remeterá fatura para o EMISSOR, com vencimento até o dia "
        "15 (quinze) do mês subsequente. O EMISSOR pagará a remuneração."
    ),
}
valor, pagina, evidencia = _extrair_condicao_pagamento_dd_documental(paginas)
assert valor == "15DD", (valor, pagina, evidencia)
assert pagina == "26", (valor, pagina, evidencia)

# Prazo financeiro tradicional contado da nota fiscal.
valor, pagina, _ = _extrair_condicao_pagamento_dd_documental({
    4: "O pagamento será realizado no prazo de 30 dias corridos contados da emissão da nota fiscal."
})
assert valor == "30DD" and pagina == "4"

# Prazo puramente operacional não pode virar condição de pagamento.
valor, pagina, _ = _extrair_condicao_pagamento_dd_documental({
    2: "A contratada disponibilizará o relatório em até 5 dias úteis após a solicitação."
})
assert valor == "" and pagina == ""

print("OK - V10: cláusula financeira prevalece e prazos operacionais são rejeitados")
