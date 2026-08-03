from auditor_evidencias import aplicar_motor_evidencias_v4, CAMPOS_OFICIAIS_V4


def linha(campo, valor, status="CONFIRMADO", pagina="11", evidencia="Evidência documental", tipo="DADO_DOCUMENTAL"):
    return {
        "campo": campo,
        "rotulo": campo,
        "valor": valor,
        "status": status,
        "tipo_dado": tipo,
        "arquivo_fonte": "SBF.pdf" if status != "NÃO_LOCALIZADO" else "",
        "pagina": pagina if status != "NÃO_LOCALIZADO" else "",
        "clausula_secao": "Bloco de Assinaturas" if pagina == "11" else "",
        "trecho_evidencia": evidencia if status != "NÃO_LOCALIZADO" else "",
        "confianca": 100 if status != "NÃO_LOCALIZADO" else 0,
    }


texto = """
--- PÁGINA 11 OCR ---
E, por estarem assim justas e contratadas, firmam as Partes este Contrato em 2 vias de igual teor e forma,
na presença de 2 testemunhas abaixo identificadas.
São Paulo, 26 de outubro de 2023.
Pereira Henrique Muronaga Coordenador Elen Aparecida Pirolo Coordenadora 7338171
ITAÚ CORRETORA DE VALORES S.A.
GRUPO SBF. S.A.
Testemunhas: Nome: Gabriel Botelho Nascimento RG: 09627519-79
Reconheço por semelhança as firmas de (1) PEDRO DE SOUZA ZEMEL e (1) JOSE LUIS MAGALHAES SALAZAR.
São Paulo, 31 de outubro de 2023.

--- PÁGINA 18 OCR ---
Os valores bruto, líquido e do imposto de renda retido na fonte pelo EMISSOR, relativos ao pagamento dos créditos,
de acordo com a periodicidade exigida pela legislação tributária. A ITAUCOR disponibilizará ao EMISSOR,
em até 5 (cinco) dias úteis, mediante solicitação, a relação dos acionistas.

--- PÁGINA 26 OCR ---
Mensalmente, a ITAUCOR fará levantamento dos serviços efetivamente prestados e remeterá fatura para o EMISSOR,
com vencimento até o dia 15 (quinze) do mês subsequente.
O EMISSOR pagará a remuneração mediante disponibilização do valor na conta corrente indicada pelo EMISSOR.
"""

nomes = (
    "Elen Aparecida Pirolo; Henrique Muronaga Pereira; Pedro de Souza Zemel; "
    "Jose Luis Magalhaes Salazar; Gabriel Botelho Nascimento"
)

auditoria = [linha(campo, f"Valor confirmado de {campo}", pagina="1") for _, campo in CAMPOS_OFICIAIS_V4]
por_campo = {item["campo"]: item for item in auditoria}
por_campo.update({
    "local_prestacao": linha("local_prestacao", "Não identificado com segurança", "NÃO_LOCALIZADO", "", ""),
    "forma_pagamento": linha("forma_pagamento", "Faturamento mensal, com vencimento até o dia 15 do mês subsequente.", pagina="26", evidencia="vencimento até o dia 15 do mês subsequente"),
    "condicao_pagamento_dias": linha("condicao_pagamento_dias", "5DD", pagina="18", evidencia="disponibilizará em até 5 dias úteis a relação dos acionistas"),
    "data_contrato": linha("data_contrato", "26/10/2023", pagina="11", evidencia="São Paulo, 26 de outubro de 2023."),
    "data_assinatura": linha("data_assinatura", "26/10/2023", pagina="11", evidencia="São Paulo, 26 de outubro de 2023."),
    "pessoas_que_assinaram": linha("pessoas_que_assinaram", nomes, pagina="11", evidencia="Elen Aparecida Pirolo; Henrique Muronaga Pereira; firmas de Pedro de Souza Zemel e Jose Luis Magalhaes Salazar; Gabriel Botelho Nascimento"),
    "data_reconhecimento_firma": linha("data_reconhecimento_firma", "31/10/2023", pagina="11", evidencia="São Paulo, 31 de outubro de 2023."),
    "contrato_assinado": linha("contrato_assinado", "Não identificado com segurança", "NÃO_LOCALIZADO", "", ""),
    "alerta_assinatura": linha("alerta_assinatura", "Não identificado com segurança", "NÃO_LOCALIZADO", "", ""),
})
auditoria = list(por_campo.values())

base = {
    "auditoria_campos": auditoria,
    "texto_extraido": texto,
    "data_contrato": "26/10/2023",
    "data_assinatura": "26/10/2023",
    "assinaturas_contrato": [],
}
bruto = {
    "auditoria_campos": auditoria,
    "assinaturas_contrato": [],
    "paginas_processadas": 26,
    "total_paginas": 26,
}

resultado = aplicar_motor_evidencias_v4(base, bruto, texto)

assert resultado["condicao_pagamento_dias"] == "15DD", resultado["condicao_pagamento_dias"]
linha_dd = next(x for x in resultado["auditoria_campos"] if x["campo"] == "condicao_pagamento_dias")
assert linha_dd["pagina"] == "26", linha_dd
assert resultado["contrato_assinado"] == "Sim", resultado["contrato_assinado"]
assert len(resultado["assinaturas_contrato"]) == 5, resultado["assinaturas_contrato"]
assert "Elen Aparecida Pirolo" in resultado["pessoas_que_assinaram"]
assert "Gabriel Botelho Nascimento" in resultado["pessoas_que_assinaram"]
testemunhas = [a for a in resultado["assinaturas_contrato"] if a.get("categoria") == "TESTEMUNHA"]
representantes = [a for a in resultado["assinaturas_contrato"] if a.get("categoria") != "TESTEMUNHA"]
assert len(testemunhas) == 1 and len(representantes) == 4, resultado["assinaturas_contrato"]
reconhecidos = [a for a in resultado["assinaturas_contrato"] if a.get("data_reconhecimento_firma") == "31/10/2023"]
assert {a["nome"] for a in reconhecidos} == {"Pedro de Souza Zemel", "Jose Luis Magalhaes Salazar"}, reconhecidos
assert resultado["campos_nao_localizados"] == ["local_prestacao"], resultado["campos_nao_localizados"]
assert any("Segunda testemunha" in p.get("Pendência", "") for p in resultado.get("pendencias", []))
gabriel = next(a for a in resultado["assinaturas_contrato"] if a["nome"] == "Gabriel Botelho Nascimento")
assert gabriel["data_reconhecimento_firma"] == "Não aplicável", gabriel

print("OK - V10: 15DD na página 26, 5 assinaturas restauradas e apenas Local de Prestação ausente")
