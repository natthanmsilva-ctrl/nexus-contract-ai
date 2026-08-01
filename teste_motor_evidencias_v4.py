from auditor_evidencias import aplicar_motor_evidencias_v4


def ev(campo, valor, pagina="1", trecho="Trecho documental específico com conteúdo suficiente.", status="CONFIRMADO"):
    return {
        "campo": campo,
        "valor": valor,
        "status": status,
        "tipo_dado": "DADO_DOCUMENTAL",
        "arquivo_fonte": "SBF.pdf",
        "pagina": pagina,
        "clausula_secao": "Cláusula de teste",
        "trecho_evidencia": trecho,
        "confianca": 95,
    }


raw = {
    "auditoria_campos": [
        ev("tipo_contrato", "Contrato de Prestação de Serviços"),
        ev("empresa_grupo_sbf", "GRUPO SBF S.A."),
        ev("cnpj_empresa_grupo", "13.217.285/0001-11"),
        ev("contraparte", "ITAU CORRETORA DE VALORES S.A."),
        ev("cnpj_contraparte", "61.194.353/0001-64"),
        ev("objetivo", "Prestação de serviços de escrituração de ações."),
        ev("descricao_servico_material", "Serviços de escrituração de ações."),
        ev("vigencia_apos_assinatura", "Prazo indeterminado a partir da assinatura.", pagina="4"),
        ev("data_assinatura", "26/10/2023", pagina="11"),
        ev("data_contrato", "26/10/2023", pagina="1"),
        ev("contrato_assinado", "Sim", pagina="11", trecho="Assinaturas físicas das partes localizadas na página 11."),
        ev("forma_pagamento", "Faturamento mensal.", pagina="26"),
        ev("condicao_pagamento_dias", "Até o dia 15 do mês subsequente.", pagina="26"),
    ],
    "itens_contrato": [
        {
            "descricao": "Custo de implantação",
            "tipo": "Serviço",
            "natureza_valor": "IMPLANTACAO_UNICA",
            "quantidade": 1,
            "unidade": "Taxa única",
            "valor_unitario": "R$ 3.000,00",
            "valor_total": "R$ 3.000,00",
            "arquivo_fonte": "SBF.pdf",
            "pagina": "25",
            "trecho_evidencia": "Custo de implantação no valor de R$ 3.000,00.",
            "status": "CONFIRMADO",
        },
        {
            "descricao": "Custo fixo mensal",
            "tipo": "Serviço",
            "natureza_valor": "MENSAL_FIXO",
            "quantidade": 1,
            "unidade": "Mês",
            "valor_unitario": "R$ 4.000,00",
            "valor_total": "R$ 4.000,00",
            "arquivo_fonte": "SBF.pdf",
            "pagina": "25",
            "trecho_evidencia": "Custo fixo mensal no valor de R$ 4.000,00.",
            "status": "CONFIRMADO",
        },
        {
            "descricao": "Tarifa por acionista",
            "tipo": "Serviço",
            "natureza_valor": "UNITARIO_VARIAVEL",
            "quantidade": "Não aplicável",
            "unidade": "Acionista/mês",
            "valor_unitario": "R$ 0,75",
            "valor_total": "Não localizado",
            "arquivo_fonte": "SBF.pdf",
            "pagina": "25",
            "trecho_evidencia": "Tarifa mensal por acionista no valor de R$ 0,75.",
            "status": "CONFIRMADO",
        },
    ],
    "assinaturas_contrato": [
        {
            "nome": "Representante Um",
            "papel_cargo": "Representante Contratante",
            "categoria": "REPRESENTANTE_CONTRATANTE",
            "data_assinatura": "26/10/2023",
            "fonte": "SBF.pdf",
            "pagina": "11",
            "evidencia": "Assinatura física do Representante Um na página 11.",
        },
        {
            "nome": "Representante Dois",
            "papel_cargo": "Representante Contraparte",
            "categoria": "REPRESENTANTE_CONTRAPARTE",
            "data_assinatura": "26/10/2023",
            "fonte": "SBF.pdf",
            "pagina": "11",
            "evidencia": "Assinatura física do Representante Dois na página 11.",
        },
    ],
    "paginas_processadas": 26,
    "total_paginas": 26,
}

resultado = aplicar_motor_evidencias_v4(raw, raw)
assert resultado["periodo_vigencia_formatado"] == "Início 26/10/2023 até 31/12/9999"
assert "situação operacional atual não confirmada" in resultado["status"].lower()
assert "R$ 3.000,00" in resultado["valor_total_materiais_servicos"]
assert "R$ 4.000,00/mês" in resultado["valor_total_materiais_servicos"]
assert "não calculável" in resultado["valor_total_estimado_vigencia"].lower()
assert resultado["contrato_assinado"] == "Sim"
assert len(resultado["assinaturas_contrato"]) == 2
assert resultado["score"] >= 70
print("OK - Motor de Evidências V4 validado")
