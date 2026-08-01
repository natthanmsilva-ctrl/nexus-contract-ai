from auditor_evidencias import CAMPOS_OFICIAIS_V4, aplicar_motor_evidencias_v4


def ev(campo, valor, pagina="1", trecho="Trecho documental específico com conteúdo suficiente.", status="CONFIRMADO", tipo="DADO_DOCUMENTAL"):
    return {
        "campo": campo,
        "valor": valor,
        "status": status,
        "tipo_dado": tipo,
        "arquivo_fonte": "SBF.pdf" if status != "NÃO_LOCALIZADO" else "",
        "pagina": pagina if status != "NÃO_LOCALIZADO" else "",
        "clausula_secao": "Cláusula de teste" if status != "NÃO_LOCALIZADO" else "",
        "trecho_evidencia": trecho if status != "NÃO_LOCALIZADO" else "",
        "confianca": 95 if status != "NÃO_LOCALIZADO" else 0,
    }


faltantes_iniciais = {
    "local_prestacao",
    "descricao_servico_material",
    "descricao_breve_cadastro",
    "resumo_aditivos",
    "data_contrato",
    "data_conclusao_docusign",
}

auditoria = []
for _, campo in CAMPOS_OFICIAIS_V4:
    if campo in faltantes_iniciais:
        auditoria.append(ev(campo, "Não identificado com segurança", status="NÃO_LOCALIZADO"))
    else:
        auditoria.append(ev(campo, f"Valor confirmado de {campo}"))

# Valores que alimentam regras determinísticas.
substituicoes = {
    "tipo_contrato": "Contrato de Prestação de Serviços de Escrituração de Ações",
    "empresa_grupo_sbf": "GRUPO SBF S.A.",
    "cnpj_empresa_grupo": "13.217.285/0001-11",
    "contraparte": "ITAÚ CORRETORA DE VALORES S.A.",
    "cnpj_contraparte": "61.194.353/0001-64",
    "objetivo": "Contratação da ITAUCOR para prestação de serviços de escrituração de ações.",
    "vigencia_apos_assinatura": "Prazo indeterminado a partir da data de assinatura.",
    "data_assinatura": "26/10/2023",
    "rescisao_indenizacao": "Denúncia imotivada mediante aviso prévio de 30 dias. Limite de indenização por danos diretos de 12 vezes a remuneração do mês anterior.",
}
for item in auditoria:
    if item["campo"] in substituicoes:
        item["valor"] = substituicoes[item["campo"]]

texto = r"""
============================== ARQUIVO: SBF.pdf ==============================
--- PÁGINA 2 OCR ---
1.1 O objeto deste Contrato é a prestação de serviços de escrituração de Ações pela ITAUCOR ao EMISSOR.
2. DESCRIÇÃO DOS SERVIÇOS
2.1. Os serviços contemplam as seguintes atividades: (i) a abertura e manutenção, em sistemas informatizados, de livros de registro, conforme previsto na regulamentação em vigor; (ii) o registro das informações relativas à titularidade das Ações, assim como de direitos reais de fruição ou de garantia e de outros gravames incidentes sobre elas; (iii) o tratamento das instruções de movimentação recebidas do EMISSOR ou de pessoas legitimadas por contrato ou mandato; e (iv) o tratamento de eventos incidentes sobre as Ações.
2.1.1 Os serviços objeto deste Contrato compreenderão as Ações que não sejam objeto de depósito centralizado.
--- PÁGINA 4 OCR ---
8.2. Este Contrato vigorará por prazo indeterminado, podendo ser denunciado, sem ônus, por qualquer Parte, mediante aviso escrito com 30 (trinta) dias de antecedência.
--- PÁGINA 11 OCR ---
E, por estarem assim justas e contratadas, firmam as Partes este Contrato em 2 vias de igual teor e forma na presença de 2 testemunhas abaixo identificadas.
São Paulo, 26 de outubro de 2023.
ITAÚ CORRETORA DE VALORES S.A. GRUPO SBF S.A.
Testemunhas: Nome: Gabriel Botelho Nascimento RG: 09627519-79 Nome: RG:
--- PÁGINA 26 OCR ---
Mensalmente, a ITAUCOR fará levantamento dos serviços efetivamente prestados e remeterá fatura para o EMISSOR, com vencimento até o dia 15 (quinze) do mês subsequente.
O EMISSOR pagará a remuneração da ITAUCOR mediante disponibilização do valor na conta corrente indicada pelo EMISSOR na cláusula 6.2.
O EMISSOR reconhece e concorda que o valor da remuneração é oferecido considerando expectativa de que o EMISSOR não rescinda o Contrato de forma unilateral pelo prazo de 24 (vinte e quatro) meses (Prazo Mínimo). Caso rescinda antes do término do Prazo Mínimo, será devido valor equivalente à soma das remunerações dos meses remanescentes.
"""

raw = {
    "auditoria_campos": auditoria,
    "aditivos_contrato": [],
    "itens_contrato": [
        {
            "descricao": "Custo de Implantação",
            "tipo": "Serviço",
            "natureza_valor": "IMPLANTACAO_UNICA",
            "quantidade": 1,
            "unidade": "Implantação",
            "valor_unitario": "R$ 3.000,00",
            "valor_total": "R$ 3.000,00",
            "arquivo_fonte": "SBF.pdf",
            "pagina": "25",
            "trecho_evidencia": "Custo de Implantação - R$ 3.000,00",
            "status": "CONFIRMADO",
        },
        {
            "descricao": "Custo Fixo Mensal",
            "tipo": "Serviço",
            "natureza_valor": "MENSAL_FIXO",
            "quantidade": 1,
            "unidade": "Mês",
            "valor_unitario": "R$ 4.000,00",
            "arquivo_fonte": "SBF.pdf",
            "pagina": "25",
            "trecho_evidencia": "Custo Fixo Mensal - R$ 4.000,00",
            "status": "CONFIRMADO",
        },
    ],
    "assinaturas_contrato": [
        {"nome": "Henrique Muronaga Pereira", "papel_cargo": "Coordenador", "categoria": "REPRESENTANTE_CONTRAPARTE", "data_assinatura": "26/10/2023", "fonte": "SBF.pdf", "pagina": "11", "evidencia": "Henrique Muronaga Pereira Coordenador 7338171"},
        {"nome": "Elen Aparecida Pirolo", "papel_cargo": "Coordenadora", "categoria": "REPRESENTANTE_CONTRAPARTE", "data_assinatura": "26/10/2023", "fonte": "SBF.pdf", "pagina": "11", "evidencia": "Elen Aparecida Pirolo Coordenadora"},
        {"nome": "Pedro de Souza Zemel", "papel_cargo": "Representante Legal", "categoria": "REPRESENTANTE_CONTRATANTE", "data_assinatura": "26/10/2023", "data_reconhecimento_firma": "31/10/2023", "fonte": "SBF.pdf", "pagina": "11", "evidencia": "Reconheço por semelhança a firma de Pedro de Souza Zemel"},
        {"nome": "Jose Luis Magalhães Salazar", "papel_cargo": "Representante Legal", "categoria": "REPRESENTANTE_CONTRATANTE", "data_assinatura": "26/10/2023", "data_reconhecimento_firma": "31/10/2023", "fonte": "SBF.pdf", "pagina": "11", "evidencia": "Reconheço por semelhança a firma de Jose Luis Magalhães Salazar"},
        {"nome": "Gabriel Botelho Nascimento", "papel_cargo": "Testemunha", "categoria": "TESTEMUNHA", "data_assinatura": "26/10/2023", "fonte": "SBF.pdf", "pagina": "11", "evidencia": "Nome: Gabriel Botelho Nascimento RG: 09627519-79"},
    ],
    "paginas_processadas": 26,
    "total_paginas": 26,
}

resultado = aplicar_motor_evidencias_v4(raw, raw, texto)

assert resultado["descricao_servico_material"].startswith("Serviços de escrituração de ações")
assert resultado["descricao_breve_cadastro"] == "Serviços de escrituração de ações."
assert resultado["data_contrato"] == "26/10/2023"
assert "disponibilização do valor" in resultado["forma_pagamento"]
assert "débito em conta" not in resultado["forma_pagamento"].lower()
assert resultado["data_conclusao_docusign"] == "Não aplicável — assinatura física"
assert resultado["resumo_aditivos"].startswith("Nenhum aditivo identificado")
assert "24 meses" in resultado["rescisao_indenizacao"]
assert "12 vezes" in resultado["rescisao_indenizacao"]
assert resultado["campos_nao_localizados"] == ["local_prestacao"], resultado["campos_nao_localizados"]
assert resultado["indicadores_pendencias"]["pontos_atencao"] >= 1
assert any("Segunda testemunha" in p.get("Pendência", "") for p in resultado["pendencias"])
assert all(a.get("data_assinatura") == "Não localizada individualmente" for a in resultado["assinaturas_contrato"])
assert all(a.get("data_instrumento") == "26/10/2023" for a in resultado["assinaturas_contrato"])
assert resultado["score"] != resultado["confianca_extracao"]
assert resultado["score"] >= 95
assert resultado["confianca_extracao"] >= 95

print("OK - V8 recupera campos documentais, separa datas e mantém apenas 1 campo ausente")
