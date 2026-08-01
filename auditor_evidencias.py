"""Motor de Auditoria por Evidências V4.

Camada determinística aplicada depois da extração por IA. O objetivo é impedir
que cards, parecer, valores e assinaturas sejam confirmados sem prova documental.
O módulo não depende de Streamlit e pode ser testado isoladamente.
"""

from __future__ import annotations

import re
import unicodedata
from copy import deepcopy
from datetime import datetime
from typing import Any, Dict, Iterable, List, Mapping, MutableMapping, Optional, Tuple


CAMPOS_OFICIAIS_V4: List[Tuple[str, str]] = [
    ("Tipo de Contrato", "tipo_contrato"),
    ("Empresa do Grupo SBF", "empresa_grupo_sbf"),
    ("CNPJ Empresa do Grupo", "cnpj_empresa_grupo"),
    ("Local de Prestação Contraparte", "local_prestacao"),
    ("Contraparte", "contraparte"),
    ("CNPJ Contraparte", "cnpj_contraparte"),
    ("Objetivo", "objetivo"),
    ("Descrição do Serviço/ Material", "descricao_servico_material"),
    ("Descrição breve do cadastro", "descricao_breve_cadastro"),
    ("Forma de pagamento", "forma_pagamento"),
    ("Condição de Pagamento em Dias", "condicao_pagamento_dias"),
    ("Multa", "multa"),
    ("Vigência após a data de assinatura", "vigencia_apos_assinatura"),
    ("Tipo de Vigência", "tipo_vigencia"),
    ("Período de Vigência", "periodo_vigencia_formatado"),
    ("Status Contratual", "status_contratual"),
    ("Situação Operacional", "situacao_operacional"),
    ("Resumo de Aditivos", "resumo_aditivos"),
    ("Rescisão e Indenização", "rescisao_indenizacao"),
    ("Anticorrupção", "anticorrupcao"),
    ("Proteção de Dados LGPD", "protecao_dados_lgpd"),
    ("Data da Assinatura", "data_assinatura"),
    ("Data do Contrato", "data_contrato"),
    ("Data Conclusão DocuSign", "data_conclusao_docusign"),
    ("Valor do Contrato Original", "valor_contrato_original"),
    ("Valor Mensal Estimado", "valor_mensal_estimado"),
    ("Valor Total Estimado da Vigência", "valor_total_estimado_vigencia"),
    ("Valor Total dos Materiais e Serviços", "valor_total_materiais_servicos"),
    ("Pessoas que assinaram", "pessoas_que_assinaram"),
]

ROTULO_POR_CAMPO = {campo: rotulo for rotulo, campo in CAMPOS_OFICIAIS_V4}
CAMPO_POR_ROTULO = {rotulo: campo for rotulo, campo in CAMPOS_OFICIAIS_V4}

STATUS_CONFIRMADOS = {"CONFIRMADO", "VALIDADO", "LOCALIZADO", "CONFIRMADO_NO_DOCUMENTO"}
STATUS_CALCULADOS = {"CALCULADO", "CALCULADO_PELO_SISTEMA"}
STATUS_NAO_APLICAVEL = {"NAO_APLICAVEL", "NÃO_APLICÁVEL"}
STATUS_CONFLITANTE = {"CONFLITANTE", "CONFLITO"}
STATUS_INFERIDO = {"INFERIDO", "PROVAVEL", "PROVÁVEL"}

CAMPOS_CRITICOS = {
    "empresa_grupo_sbf",
    "cnpj_empresa_grupo",
    "contraparte",
    "cnpj_contraparte",
    "objetivo",
    "descricao_servico_material",
    "vigencia_apos_assinatura",
    "periodo_vigencia_formatado",
    "status_contratual",
    "data_assinatura",
    "contrato_assinado",
}

ALIASES_CAMPOS = {
    "tipo_de_contrato": "tipo_contrato",
    "empresa_do_grupo_sbf": "empresa_grupo_sbf",
    "cnpj_empresa_do_grupo": "cnpj_empresa_grupo",
    "cnpj_empresa_grupo_sbf": "cnpj_empresa_grupo",
    "local_de_prestacao_contraparte": "local_prestacao",
    "local_prestacao_contraparte": "local_prestacao",
    "fornecedor": "contraparte",
    "cnpj_fornecedor": "cnpj_contraparte",
    "descricao_do_servico_material": "descricao_servico_material",
    "descricao_servico_material": "descricao_servico_material",
    "vigencia_apos_a_data_de_assinatura": "vigencia_apos_assinatura",
    "periodo_de_vigencia": "periodo_vigencia_formatado",
    "periodo_vigencia": "periodo_vigencia_formatado",
    "vigencia_formatada": "periodo_vigencia_formatado",
    "tipo_de_vigencia": "tipo_vigencia",
    "status_contratual": "status_contratual",
    "situacao_operacional": "situacao_operacional",
    "situação_operacional": "situacao_operacional",
    "data_do_contrato": "data_contrato",
    "data_de_assinatura": "data_assinatura",
    "data_de_conclusao_docusign": "data_conclusao_docusign",
    "valor_do_contrato_original": "valor_contrato_original",
    "valor_mensal": "valor_mensal_estimado",
    "valor_total_da_vigencia": "valor_total_estimado_vigencia",
    "valor_total_dos_materiais_e_servicos": "valor_total_materiais_servicos",
    "valor_total_dos_materiais_servicos": "valor_total_materiais_servicos",
    "assinantes": "pessoas_que_assinaram",
    "pessoas_que_assinaram_o_contrato": "pessoas_que_assinaram",
}


PROMPT_EVIDENCIAS_V4 = r"""

MOTOR DE AUDITORIA POR EVIDÊNCIAS V4 — REGRA SUPERIOR

OBJETIVO
Produzir uma análise completa, fiel e rastreável. É preferível responder “NÃO LOCALIZADO” a preencher um campo por plausibilidade. Examine todas as páginas, anexos, tabelas, assinaturas, certificados e condições comerciais. Não encerre a análise antes de conferir as páginas finais e os anexos financeiros.

ORDEM DE AUTORIDADE DOCUMENTAL
1. Contrato final assinado e certificado de assinatura concluído.
2. Aditivo final assinado.
3. Contrato final não assinado.
4. Proposta comercial/técnica incorporada ou referenciada pelo contrato.
5. E-mails e documentos operacionais.
6. Minutas e versões antigas.
Quando houver conflito, registre em conflitos_documentais e aplique a fonte de maior autoridade. Nunca misture partes, CNPJs, valores ou datas de versões diferentes.

MATRIZ OBRIGATÓRIA DE EVIDÊNCIAS
Para CADA campo principal, retorne uma linha em auditoria_campos com:
- campo: chave técnica exata;
- valor;
- status: CONFIRMADO, CALCULADO, INFERIDO, NÃO_LOCALIZADO, NÃO_APLICÁVEL ou CONFLITANTE;
- tipo_dado: DADO_DOCUMENTAL, CALCULO_SISTEMA ou INTERPRETACAO;
- arquivo_fonte;
- pagina;
- clausula_secao;
- trecho_evidencia: trecho curto e fiel;
- confianca: 0 a 100.
Um dado CONFIRMADO exige arquivo, página/seção e trecho específico. “Conforme documento” não é evidência.

REGRAS DE EXATIDÃO
- Endereço da parte NÃO é local de prestação, salvo cláusula expressa de execução naquele local.
- “Prazo indeterminado” deve continuar indeterminado. O sistema exibirá tecnicamente 31/12/9999; não invente data final.
- Permanência mínima, fidelização ou multa de saída NÃO é prazo total do contrato.
- Status atual operacional não pode ser presumido. Informe apenas a situação documental da vigência.
- Data do instrumento, data da assinatura, conclusão DocuSign e reconhecimento de firma são eventos distintos.
- Nome no corpo do contrato não prova assinatura. Cada signatário exige evidência no bloco de assinatura/certificado.
- Testemunha em campo vazio não pode ser listada.
- Ausência de e-mail de signatário não é pendência contratual, salvo obrigação expressa.

FINANCEIRO — NATUREZAS INCOMPATÍVEIS
Classifique cada item em exatamente uma natureza:
VALOR_GLOBAL, IMPLANTACAO_UNICA, MENSAL_FIXO, UNITARIO_VARIAVEL, PERCENTUAL_VARIAVEL, REEMBOLSO ou OUTRO.
- Valor global: somente total fechado de todo o contrato.
- Implantação/setup/taxa única: valor pontual.
- Mensalidade/custo fixo mensal: recorrente por mês.
- Tarifa por acionista, operação, refeição, vaga, usuário, trabalhador, item ou unidade: variável.
- Percentual sobre salário/faturamento/base: variável.
Nunca some implantação com mensalidade como se fosse valor global. Nunca calcule total variável sem quantidade e prazo confirmados.
Cada item_contrato deve conter: descricao, tipo, grupo_tabela, natureza_valor, quantidade, unidade, valor_unitario, valor_total, periodicidade, faixa_condicao, condicao_comercial, arquivo_fonte, pagina, clausula_secao e trecho_evidencia.

EXTRAÇÃO COMPLETA DA TABELA COMERCIAL — OBRIGATÓRIA
- Localize todos os anexos de remuneração, preços, tarifas, percentuais, faixas, eventos e reembolsos.
- Retorne UMA LINHA para CADA linha comercial da tabela. Não resuma, não selecione exemplos e não limite a quantidade.
- Continue a leitura quando a tabela prosseguir em outra página.
- Preserve itens com valor textual, como ISENTO, TAXA CORREIO, conforme tabela, percentual ou faixa.
- Não converta ausência de valor em zero. R$ 0,00 somente quando o documento declarar isenção/gratuidade/zero.
- Diferencie implantação única, mensalidade fixa, tarifa por faixa, tarifa por evento, percentual e reembolso.
- Retorne metricas_tabela_comercial com itens_encontrados_documento, itens_exibidos_auditor, cobertura_tabela_percentual, paginas_tabela_comercial e grupos_tabela_comercial.
- Antes de concluir, conte as linhas da tabela e confira se a quantidade em itens_contrato é compatível.

INDICADORES SEPARADOS
- status_contratual descreve apenas a vigência documental.
- situacao_operacional não pode ser presumida; use “Não confirmada nos documentos analisados” quando faltar prova corporativa atual.
- confianca_extracao mede qualidade/cobertura da extração.
- risco descreve risco contratual; não use campos ausentes para inflar ou reduzir artificialmente o risco jurídico.

ASSINATURAS
Cada assinatura deve conter: nome, papel_cargo, categoria (REPRESENTANTE_CONTRATANTE, REPRESENTANTE_CONTRAPARTE, TESTEMUNHA, APROVADOR, OUTRO), email, data_assinatura, data_reconhecimento_firma, fonte, pagina e evidencia.
contrato_assinado = Sim somente quando houver assinatura eletrônica concluída/certificado válido ou assinaturas das partes no instrumento. Testemunha isolada não valida o contrato.

CHECKLIST E PENDÊNCIAS
Checklist só pode usar status CONCLUÍDO quando houver evidência específica. Pendência deve conter Arquivo, Página, Evidência, risco, criticidade e recomendação. Não crie pendência genérica sem prova documental.

COBERTURA
Retorne também:
- paginas_processadas;
- total_paginas;
- secoes_verificadas;
- tabelas_verificadas;
- assinaturas_verificadas;
- campos_nao_localizados;
- conflitos_documentais.

VALIDAÇÃO FINAL ANTES DO JSON
1. Verifique se todos os campos existentes no documento foram transportados.
2. Remova qualquer informação sem prova.
3. Separe fatos, cálculos e interpretações.
4. Confira valores e periodicidades.
5. Confira assinaturas e datas.
6. Retorne APENAS JSON válido e completo.
"""


PROMPT_VERIFICADOR_V4 = r"""

AUDITORIA INDEPENDENTE V4 — TESTE DE REPROVAÇÃO
Reprove ou corrija qualquer item que viole uma das regras abaixo:
- campo confirmado sem arquivo, página/seção e trecho literal;
- endereço usado como local de prestação sem cláusula expressa;
- valor mensal convertido em valor por unidade;
- implantação somada à mensalidade como valor global;
- tabela comercial resumida, truncada ou limitada a poucos exemplos;
- linha com “isento”, “taxa Correio”, faixa ou valor textual descartada;
- ausência de valor convertida artificialmente em R$ 0,00;
- tarifa variável totalizada sem quantidade/prazo;
- prazo mínimo confundido com vigência total;
- prazo indeterminado convertido em data real de encerramento;
- status “Ativo” presumido apenas porque não foi localizada rescisão;
- reconhecimento de firma usado como data da assinatura;
- signatário listado sem evidência no bloco de assinaturas/certificado;
- testemunha vazia inventada;
- pendência ou checklist sem evidência objetiva;
- fato de minuta antiga prevalecendo sobre contrato/aditivo final assinado.
Garanta que auditoria_campos possua uma linha para cada campo principal e que conflitos_documentais registre divergências reais. Retorne somente o JSON integral corrigido.
"""


def _texto(valor: Any) -> str:
    if valor is None:
        return ""
    if isinstance(valor, float) and valor != valor:
        return ""
    texto = str(valor).replace("\x00", " ")
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def _sem_acento(valor: Any) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", _texto(valor)) if not unicodedata.combining(c)
    )


def _token(valor: Any) -> str:
    return re.sub(r"[^A-Z0-9]+", "_", _sem_acento(valor).upper()).strip("_")


def _campo_canonico(valor: Any) -> str:
    raw = _texto(valor)
    if not raw:
        return ""
    if raw in ROTULO_POR_CAMPO:
        return raw
    if raw in CAMPO_POR_ROTULO:
        return CAMPO_POR_ROTULO[raw]
    snake = re.sub(r"[^a-z0-9]+", "_", _sem_acento(raw).lower()).strip("_")
    return ALIASES_CAMPOS.get(snake, snake)


def _valor_util(valor: Any) -> bool:
    texto = _texto(valor)
    if not texto:
        return False
    low = _sem_acento(texto).lower()
    bloqueios = (
        "nao localizado",
        "nao identificad",
        "nao informado",
        "sem informacao",
        "none",
        "n/a",
        "erro de leitura",
    )
    return not any(low == b or low.startswith(b) for b in bloqueios)


def _evidencia_util(valor: Any) -> bool:
    texto = _texto(valor)
    low = _sem_acento(texto).lower()
    if len(texto) < 12:
        return False
    genericos = (
        "conforme documento",
        "documento analisado",
        "informacao localizada",
        "termo localizado",
        "cadastro identificado",
        "data identificada conforme",
        "nao localizado",
    )
    return not any(low == x or low.startswith(x) for x in genericos)


def _status(valor: Any) -> str:
    t = _token(valor)
    if t in {"NAO_LOCALIZADO", "NAO_IDENTIFICADO", "NAO_CONFIRMADO", "NAO_VALIDADO", ""}:
        return "NÃO_LOCALIZADO"
    if t in {"NAO_APLICAVEL"}:
        return "NÃO_APLICÁVEL"
    if t in {"CONFIRMADO", "VALIDADO", "LOCALIZADO", "CONFIRMADO_NO_DOCUMENTO"}:
        return "CONFIRMADO"
    if t in {"CALCULADO", "CALCULADO_PELO_SISTEMA"}:
        return "CALCULADO"
    if t in {"INFERIDO", "PROVAVEL"}:
        return "INFERIDO"
    if t in {"CONFLITANTE", "CONFLITO"}:
        return "CONFLITANTE"
    return _texto(valor).upper() or "NÃO_LOCALIZADO"


def _confianca(valor: Any) -> int:
    try:
        return max(0, min(100, int(float(str(valor).replace(",", ".")))))
    except Exception:
        return 0


def _normalizar_auditoria(raw: Any) -> List[Dict[str, Any]]:
    if isinstance(raw, Mapping):
        raw = [dict(v, campo=k) if isinstance(v, Mapping) else {"campo": k, "valor": v} for k, v in raw.items()]
    if not isinstance(raw, list):
        return []

    saida: List[Dict[str, Any]] = []
    por_campo: Dict[str, Dict[str, Any]] = {}
    for item in raw:
        if not isinstance(item, Mapping):
            continue
        campo = _campo_canonico(item.get("campo") or item.get("Campo") or item.get("chave") or item.get("field"))
        if not campo:
            continue
        normal = {
            "campo": campo,
            "rotulo": ROTULO_POR_CAMPO.get(campo, campo.replace("_", " ").title()),
            "valor": _texto(item.get("valor") or item.get("Valor") or item.get("resultado")),
            "status": _status(item.get("status") or item.get("Status")),
            "tipo_dado": _texto(item.get("tipo_dado") or item.get("Tipo de dado") or item.get("classificacao") or "DADO_DOCUMENTAL").upper(),
            "arquivo_fonte": _texto(item.get("arquivo_fonte") or item.get("Arquivo fonte") or item.get("arquivo") or item.get("fonte")),
            "pagina": _texto(item.get("pagina") or item.get("Página") or item.get("página")),
            "clausula_secao": _texto(item.get("clausula_secao") or item.get("clausula") or item.get("seção") or item.get("secao")),
            "trecho_evidencia": _texto(item.get("trecho_evidencia") or item.get("Trecho de evidência") or item.get("trecho") or item.get("evidencia") or item.get("Evidência")),
            "confianca": _confianca(item.get("confianca") or item.get("Confiança")),
        }
        anterior = por_campo.get(campo)
        if anterior is None or normal["confianca"] > anterior["confianca"]:
            por_campo[campo] = normal

    for _, campo in CAMPOS_OFICIAIS_V4:
        if campo in por_campo:
            saida.append(por_campo.pop(campo))
    saida.extend(por_campo.values())
    return saida


def _evidencia_confirma(item: Optional[Mapping[str, Any]]) -> bool:
    if not item:
        return False
    status = _status(item.get("status"))
    if status == "NÃO_APLICÁVEL":
        return _valor_util(item.get("valor"))
    if status == "CALCULADO":
        return _valor_util(item.get("valor")) and _evidencia_util(item.get("trecho_evidencia"))
    if status != "CONFIRMADO":
        return False
    return (
        _valor_util(item.get("valor"))
        and _valor_util(item.get("arquivo_fonte"))
        and (_valor_util(item.get("pagina")) or _valor_util(item.get("clausula_secao")))
        and _evidencia_util(item.get("trecho_evidencia"))
        and _confianca(item.get("confianca")) >= 70
    )


def _mapa_auditoria(auditoria: Iterable[Mapping[str, Any]]) -> Dict[str, Dict[str, Any]]:
    return {_campo_canonico(item.get("campo")): dict(item) for item in auditoria if _campo_canonico(item.get("campo"))}


def _linha_nao_localizada(campo: str) -> Dict[str, Any]:
    return {
        "campo": campo,
        "rotulo": ROTULO_POR_CAMPO.get(campo, campo.replace("_", " ").title()),
        "valor": "Não identificado com segurança",
        "status": "NÃO_LOCALIZADO",
        "tipo_dado": "DADO_DOCUMENTAL",
        "arquivo_fonte": "",
        "pagina": "",
        "clausula_secao": "",
        "trecho_evidencia": "",
        "confianca": 0,
    }


def _parse_moeda(valor: Any) -> Optional[float]:
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return float(valor)
    texto = _texto(valor)
    if not texto:
        return None
    match = re.search(r"(?:R\$\s*)?(-?\d{1,3}(?:\.\d{3})*(?:,\d{1,2})|-?\d+(?:[.,]\d{1,2})?)", texto)
    if not match:
        return None
    numero = match.group(1)
    if "," in numero:
        numero = numero.replace(".", "").replace(",", ".")
    elif numero.count(".") > 1:
        numero = numero.replace(".", "")
    try:
        return float(numero)
    except Exception:
        return None


def _moeda(valor: float) -> str:
    s = f"{valor:,.2f}"
    return "R$ " + s.replace(",", "X").replace(".", ",").replace("X", ".")


def _data_br(valor: Any) -> str:
    texto = _texto(valor)
    m = re.search(r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{4})\b", texto)
    if m:
        return f"{int(m.group(1)):02d}/{int(m.group(2)):02d}/{m.group(3)}"
    meses = {
        "janeiro": 1, "fevereiro": 2, "marco": 3, "março": 3, "abril": 4,
        "maio": 5, "junho": 6, "julho": 7, "agosto": 8, "setembro": 9,
        "outubro": 10, "novembro": 11, "dezembro": 12,
    }
    m = re.search(r"\b(\d{1,2})\s+de\s+([A-Za-zçÇãÃéÉ]+)\s+de\s+(\d{4})\b", texto, flags=re.I)
    if m:
        mes = meses.get(_sem_acento(m.group(2)).lower()) or meses.get(m.group(2).lower())
        if mes:
            return f"{int(m.group(1)):02d}/{mes:02d}/{m.group(3)}"
    return ""


def _somar_meses(data: datetime, meses: int) -> datetime:
    mes_zero = data.month - 1 + meses
    ano = data.year + mes_zero // 12
    mes = mes_zero % 12 + 1
    dias_mes = [31, 29 if ano % 4 == 0 and (ano % 100 != 0 or ano % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    dia = min(data.day, dias_mes[mes - 1])
    return datetime(ano, mes, dia)


def _meses_vigencia(base: Mapping[str, Any]) -> Optional[int]:
    candidatos = [base.get("vigencia_apos_assinatura"), base.get("periodo_vigencia_formatado")]
    texto = " ".join(_texto(x) for x in candidatos)
    if "indeterminado" in _sem_acento(texto).lower():
        return None
    m = re.search(r"\b(\d{1,3})\s*mes", _sem_acento(texto), flags=re.I)
    if m:
        return int(m.group(1))
    m = re.search(r"\b(\d{1,2})\s*ano", _sem_acento(texto), flags=re.I)
    if m:
        return int(m.group(1)) * 12
    return None


def _natureza_item(item: Mapping[str, Any]) -> str:
    explicita = _token(item.get("natureza_valor") or item.get("Natureza do valor") or item.get("tipo_valor"))
    mapa_exp = {
        "VALOR_GLOBAL": "VALOR_GLOBAL",
        "GLOBAL": "VALOR_GLOBAL",
        "IMPLANTACAO_UNICA": "IMPLANTACAO_UNICA",
        "PONTUAL": "IMPLANTACAO_UNICA",
        "MENSAL_FIXO": "MENSAL_FIXO",
        "FIXO_MENSAL": "MENSAL_FIXO",
        "UNITARIO_VARIAVEL": "UNITARIO_VARIAVEL",
        "PERCENTUAL_VARIAVEL": "PERCENTUAL_VARIAVEL",
        "REEMBOLSO": "REEMBOLSO",
        "ISENTO": "ISENTO",
    }
    if explicita in mapa_exp:
        return mapa_exp[explicita]

    texto = _sem_acento(" ".join(_texto(item.get(k)) for k in (
        "Descrição", "descricao", "Item", "item", "Unidade", "unidade", "Periodicidade", "periodicidade"
    ))).lower()
    if any(x in texto for x in ("implantacao", "setup", "taxa unica", "adesao", "ativacao")):
        return "IMPLANTACAO_UNICA"
    if any(x in texto for x in ("mensalidade", "custo fixo mensal", "remuneracao mensal", "fixo mensal")):
        return "MENSAL_FIXO"
    if any(x in texto for x in ("por acionista", "acionista/mes", "por operacao", "por refeicao", "por usuario", "por vaga", "por trabalhador", "por unidade", "por item")):
        return "UNITARIO_VARIAVEL"
    percentual = _texto(item.get("Taxa / Percentual") or item.get("taxa_percentual"))
    if "%" in percentual or "percentual" in texto:
        return "PERCENTUAL_VARIAVEL"
    if "reembolso" in texto:
        return "REEMBOLSO"
    periodicidade = _sem_acento(_texto(item.get("Periodicidade") or item.get("periodicidade") or item.get("Unidade") or item.get("unidade"))).lower()
    if periodicidade in {"mes", "mensal"}:
        return "MENSAL_FIXO"
    return "OUTRO"


def _normalizar_itens(raw: Any) -> List[Dict[str, Any]]:
    if not isinstance(raw, list):
        return []
    saida: List[Dict[str, Any]] = []
    vistos = set()
    for idx, item in enumerate(raw, 1):
        if not isinstance(item, Mapping):
            continue
        desc = _texto(item.get("Descrição") or item.get("descricao") or item.get("Item") or item.get("item"))
        fonte = _texto(item.get("Fonte") or item.get("fonte") or item.get("arquivo_fonte"))
        pagina = _texto(item.get("Página") or item.get("pagina"))
        evidencia = _texto(item.get("Evidência") or item.get("evidencia") or item.get("trecho_evidencia"))
        status = _status(item.get("status") or ("CONFIRMADO" if fonte and evidencia else "NÃO_LOCALIZADO"))
        if not _valor_util(desc):
            continue
        # Em análises V4, item financeiro confirmado precisa de fonte e evidência.
        if status == "CONFIRMADO" and not (fonte and (pagina or item.get("clausula_secao")) and _evidencia_util(evidencia)):
            status = "NÃO_LOCALIZADO"
        chave = (_token(desc), _texto(item.get("Valor unitário") or item.get("valor_unitario")), _texto(item.get("Taxa / Percentual") or item.get("taxa_percentual")))
        if chave in vistos:
            continue
        vistos.add(chave)
        novo = dict(item)
        novo.update({
            "Item": _texto(item.get("Item") or item.get("item") or idx),
            "Descrição": desc,
            "Tipo": _texto(item.get("Tipo") or item.get("tipo") or "Serviço"),
            "Quantidade": _texto(item.get("Quantidade") or item.get("quantidade") or "Não aplicável"),
            "Unidade": _texto(item.get("Unidade") or item.get("unidade") or "Não aplicável"),
            "Valor unitário": _texto(item.get("Valor unitário") or item.get("valor_unitario") or "Não localizado"),
            "Valor total": _texto(item.get("Valor total") or item.get("valor_total") or "Não localizado"),
            "Taxa / Percentual": _texto(item.get("Taxa / Percentual") or item.get("taxa_percentual") or "Não aplicável"),
            "Periodicidade": _texto(item.get("Periodicidade") or item.get("periodicidade") or "Não localizado"),
            "Natureza do valor": _natureza_item(item),
            "Fonte": fonte or "Não localizado",
            "Página": pagina or "Não localizado",
            "Evidência": evidencia or "Não localizado",
            "Status de evidência": status,
        })
        if status in {"CONFIRMADO", "CALCULADO"}:
            saida.append(novo)
    return saida


def _quantidade_num(valor: Any) -> Optional[float]:
    texto = _texto(valor).replace(".", "").replace(",", ".")
    m = re.search(r"-?\d+(?:\.\d+)?", texto)
    if not m:
        return None
    try:
        return float(m.group(0))
    except Exception:
        return None


def _financeiro(base: MutableMapping[str, Any], raw: Mapping[str, Any], auditoria: List[Dict[str, Any]]) -> None:
    # Une itens da IA e do extrator determinístico. A IA não pode reduzir uma
    # tabela comercial completa a poucos exemplos.
    itens_raw = raw.get("itens_contrato") if isinstance(raw.get("itens_contrato"), list) else []
    itens_base = base.get("itens_contrato") if isinstance(base.get("itens_contrato"), list) else []
    itens = _normalizar_itens(list(itens_base) + list(itens_raw))
    base["itens_contrato"] = itens
    if isinstance(raw.get("metricas_tabela_comercial"), Mapping):
        base["metricas_tabela_comercial"] = dict(raw.get("metricas_tabela_comercial"))

    estruturados = raw.get("valores_estruturados") if isinstance(raw.get("valores_estruturados"), Mapping) else {}
    mapa = _mapa_auditoria(auditoria)

    def bloco(*nomes: str) -> Mapping[str, Any]:
        for nome in nomes:
            val = estruturados.get(nome)
            if isinstance(val, Mapping):
                return val
            if _valor_util(val):
                return {"valor": val, "status": "CONFIRMADO"}
        return {}

    def bloco_confirmado(b: Mapping[str, Any]) -> bool:
        if not b:
            return False
        st = _status(b.get("status") or "CONFIRMADO")
        fonte = b.get("arquivo_fonte") or b.get("fonte")
        pagina = b.get("pagina") or b.get("clausula_secao")
        ev = b.get("trecho_evidencia") or b.get("evidencia")
        # Compatibilidade com IA antiga: quando existe auditoria confirmada do campo, o bloco é aceito.
        return st in {"CONFIRMADO", "CALCULADO"} and _valor_util(b.get("valor")) and (
            (_valor_util(fonte) and _valor_util(pagina) and _evidencia_util(ev))
            or not auditoria
        )

    global_b = bloco("valor_global", "valor_contrato_original")
    mensal_b = bloco("valor_fixo_mensal", "valor_mensal_fixo", "valor_mensal")
    global_num = _parse_moeda(global_b.get("valor")) if bloco_confirmado(global_b) else None
    mensal_num = _parse_moeda(mensal_b.get("valor")) if bloco_confirmado(mensal_b) else None

    pontuais: List[Tuple[float, Dict[str, Any]]] = []
    mensais: List[Tuple[float, Dict[str, Any]]] = []
    variaveis: List[Dict[str, Any]] = []
    totais_mesma_natureza: List[float] = []

    for item in itens:
        natureza = item.get("Natureza do valor")
        unit = _parse_moeda(item.get("Valor unitário"))
        total = _parse_moeda(item.get("Valor total"))
        qtd = _quantidade_num(item.get("Quantidade"))
        valor_calc = total
        if valor_calc is None and unit is not None and qtd is not None and qtd >= 0:
            valor_calc = unit * qtd
        if natureza == "IMPLANTACAO_UNICA" and (total is not None or unit is not None):
            pontuais.append((total if total is not None else unit, item))
        elif natureza == "MENSAL_FIXO" and (total is not None or unit is not None):
            mensais.append((total if total is not None else unit, item))
        elif natureza in {"UNITARIO_VARIAVEL", "PERCENTUAL_VARIAVEL", "REEMBOLSO"}:
            variaveis.append(item)
        elif natureza == "ISENTO":
            # Zero só é mantido quando o documento declara isenção; não entra em soma.
            continue
        elif valor_calc is not None:
            totais_mesma_natureza.append(valor_calc)

    if mensal_num is None and mensais:
        mensal_num = sum(v for v, _ in mensais)

    if global_num is not None:
        base["valor_contrato_original"] = f"{_moeda(global_num)}. Valor global expressamente definido no instrumento."
    else:
        base["valor_contrato_original"] = (
            "Sem valor global fixo. O instrumento apresenta valores pontuais, mensais, unitários ou variáveis, que não devem ser confundidos com um total fechado do contrato."
        )

    if mensal_num is not None:
        texto_mensal = f"{_moeda(mensal_num)}/mês. Valor fixo mensal confirmado."
        if variaveis:
            texto_mensal += f" Existem também {len(variaveis)} tarifa(s) variável(is), cobradas separadamente conforme uso ou quantidade."
        base["valor_mensal_estimado"] = texto_mensal
    elif variaveis:
        exemplos = []
        for item in variaveis[:4]:
            valor = _texto(item.get("Valor unitário"))
            taxa = _texto(item.get("Taxa / Percentual"))
            ref = valor if _valor_util(valor) else taxa
            if _valor_util(ref):
                exemplos.append(f"{item.get('Descrição')}: {ref}")
        detalhe = (" Exemplos: " + "; ".join(exemplos) + ".") if exemplos else ""
        base["valor_mensal_estimado"] = "Não há mensalidade fixa confirmada; o valor depende de utilização, quantidade ou base percentual." + detalhe
    else:
        base["valor_mensal_estimado"] = "Não identificado com segurança"

    soma_pontual = sum(v for v, _ in pontuais) if pontuais else None
    soma_mensal = sum(v for v, _ in mensais) if mensais else mensal_num
    partes = []
    if soma_pontual is not None:
        partes.append(f"Valores pontuais/implantação: {_moeda(soma_pontual)}")
    if soma_mensal is not None:
        partes.append(f"Mensalidades fixas: {_moeda(soma_mensal)}/mês")
    if totais_mesma_natureza and not pontuais and not mensais and not variaveis:
        partes.append(f"Total dos itens calculáveis da mesma natureza: {_moeda(sum(totais_mesma_natureza))}")
    if variaveis:
        partes.append(f"{len(variaveis)} tarifa(s) variável(is) sem total calculável antes da quantidade/uso")
    if partes:
        base["valor_total_materiais_servicos"] = ". ".join(partes) + "."
    else:
        base["valor_total_materiais_servicos"] = "Não identificado com segurança"

    meses = _meses_vigencia(base)
    if global_num is not None:
        base["valor_total_estimado_vigencia"] = f"{_moeda(global_num)}. O próprio contrato define valor global; não foi necessário projetar."
    elif meses and mensal_num is not None and not variaveis:
        total_vig = mensal_num * meses + (soma_pontual or 0)
        formula = f"{_moeda(mensal_num)} x {meses} meses"
        if soma_pontual:
            formula += f" + {_moeda(soma_pontual)} de valor pontual"
        base["valor_total_estimado_vigencia"] = f"{_moeda(total_vig)}. Cálculo do sistema: {formula}."
    elif meses and mensal_num is not None and variaveis:
        minimo = mensal_num * meses + (soma_pontual or 0)
        base["valor_total_estimado_vigencia"] = (
            f"Estimativa mínima da parcela fixa: {_moeda(minimo)} para {meses} meses. O total definitivo não é calculável porque existem tarifas variáveis sem quantidade confirmada."
        )
    else:
        base["valor_total_estimado_vigencia"] = (
            "Não calculável com precisão. A vigência é indeterminada ou faltam quantidade, período ou base de consumo confirmados."
        )

    # Atualiza auditoria dos quatro campos financeiros como cálculo/consolidação do sistema.
    mapa = _mapa_auditoria(auditoria)
    for campo in (
        "valor_contrato_original", "valor_mensal_estimado",
        "valor_total_estimado_vigencia", "valor_total_materiais_servicos",
    ):
        linha = mapa.get(campo) or _linha_nao_localizada(campo)
        linha.update({
            "valor": _texto(base.get(campo)),
            "status": "CALCULADO" if campo != "valor_contrato_original" or global_num is not None else "CONFIRMADO",
            "tipo_dado": "CALCULO_SISTEMA" if campo != "valor_contrato_original" else "CONSOLIDACAO_DOCUMENTAL",
            "arquivo_fonte": linha.get("arquivo_fonte") or "Múltiplas evidências financeiras",
            "pagina": linha.get("pagina") or "Conforme itens/tabela comercial",
            "clausula_secao": linha.get("clausula_secao") or "Consolidação financeira por natureza",
            "trecho_evidencia": linha.get("trecho_evidencia") or "Valores mantidos separados por periodicidade e natureza; cálculos executados somente com bases confirmadas.",
            "confianca": max(_confianca(linha.get("confianca")), 90 if partes else 60),
        })
        mapa[campo] = linha
    auditoria[:] = _ordenar_auditoria(mapa.values())


def _ordenar_auditoria(itens: Iterable[Mapping[str, Any]]) -> List[Dict[str, Any]]:
    mapa = {_campo_canonico(i.get("campo")): dict(i) for i in itens if _campo_canonico(i.get("campo"))}
    saida = []
    for _, campo in CAMPOS_OFICIAIS_V4:
        saida.append(mapa.pop(campo, _linha_nao_localizada(campo)))
    saida.extend(mapa.values())
    return saida


def _vigencia_status(base: MutableMapping[str, Any], auditoria: List[Dict[str, Any]]) -> None:
    vigencia = " ".join(_texto(base.get(c)) for c in ("vigencia_apos_assinatura", "periodo_vigencia_formatado"))
    low = _sem_acento(vigencia).lower()
    inicio = _data_br(base.get("data_assinatura")) or _data_br(base.get("data_contrato")) or _data_br(vigencia)
    mapa = _mapa_auditoria(auditoria)
    fonte_vig = mapa.get("vigencia_apos_assinatura") or mapa.get("periodo_vigencia_formatado") or _linha_nao_localizada("periodo_vigencia_formatado")

    if "indeterminado" in low:
        base["tipo_vigencia"] = "Prazo indeterminado"
        base["periodo_vigencia_formatado"] = f"Início {inicio} até 31/12/9999" if inicio else "Prazo indeterminado; data de início não identificada com segurança"
        base["status_contratual"] = "Vigente por prazo indeterminado"
    else:
        base["tipo_vigencia"] = "Prazo determinado" if re.search(r"\b\d{2}/\d{2}/\d{4}\b", vigencia) else "Não identificado com segurança"
        datas = re.findall(r"\b\d{2}/\d{2}/\d{4}\b", _texto(base.get("periodo_vigencia_formatado")))
        if len(datas) >= 2:
            try:
                fim = datetime.strptime(datas[-1], "%d/%m/%Y")
                base["status_contratual"] = f"Vigência documental encerrada em {datas[-1]}" if fim.date() < datetime.now().date() else f"Vigência documental prevista até {datas[-1]}"
            except Exception:
                base["status_contratual"] = "Situação contratual não identificada com segurança"
        else:
            base["status_contratual"] = "Situação contratual não identificada com segurança"

    base["situacao_operacional"] = "Não confirmada nos documentos analisados"
    base["status"] = base["status_contratual"]  # compatibilidade com histórico/dashboard

    derivados = {
        "tipo_vigencia": (base.get("tipo_vigencia"), "INTERPRETACAO"),
        "periodo_vigencia_formatado": (base.get("periodo_vigencia_formatado"), "CALCULO_SISTEMA" if "31/12/9999" in _texto(base.get("periodo_vigencia_formatado")) else "DADO_DOCUMENTAL"),
        "status_contratual": (base.get("status_contratual"), "INTERPRETACAO"),
        "situacao_operacional": (base.get("situacao_operacional"), "INTERPRETACAO"),
    }
    for campo, (valor, tipo_dado) in derivados.items():
        linha = mapa.get(campo) or _linha_nao_localizada(campo)
        linha.update({
            "valor": _texto(valor),
            "status": "CALCULADO" if tipo_dado != "DADO_DOCUMENTAL" else linha.get("status", "CONFIRMADO"),
            "tipo_dado": tipo_dado,
            "arquivo_fonte": linha.get("arquivo_fonte") or fonte_vig.get("arquivo_fonte") or "Contrato principal",
            "pagina": linha.get("pagina") or fonte_vig.get("pagina") or "Cláusula de vigência",
            "clausula_secao": linha.get("clausula_secao") or fonte_vig.get("clausula_secao") or "Vigência e término",
            "trecho_evidencia": linha.get("trecho_evidencia") or fonte_vig.get("trecho_evidencia") or "Prazo documental consolidado sem confundir permanência mínima com término da vigência.",
            "confianca": max(_confianca(linha.get("confianca")), 90 if inicio else 75),
        })
        mapa[campo] = linha
    auditoria[:] = _ordenar_auditoria(mapa.values())


def _normalizar_assinaturas(raw: Any) -> List[Dict[str, Any]]:
    if not isinstance(raw, list):
        return []
    saida = []
    vistos = set()
    for item in raw:
        if not isinstance(item, Mapping):
            continue
        nome = _texto(item.get("nome") or item.get("Nome") or item.get("signatario") or item.get("assinante"))
        fonte = _texto(item.get("fonte") or item.get("Fonte") or item.get("arquivo_fonte"))
        pagina = _texto(item.get("pagina") or item.get("Página"))
        evidencia = _texto(item.get("evidencia") or item.get("Evidência") or item.get("trecho_evidencia"))
        if not (_valor_util(nome) and _valor_util(fonte) and (pagina or _valor_util(item.get("clausula_secao"))) and _evidencia_util(evidencia)):
            continue
        chave = _token(nome)
        if chave in vistos:
            continue
        vistos.add(chave)
        papel = _texto(item.get("papel_cargo") or item.get("Papel/Cargo") or item.get("cargo") or item.get("papel") or "Não localizado")
        cat = _token(item.get("categoria"))
        papel_token = _token(papel)
        if not cat:
            if "TESTEMUNHA" in papel_token:
                cat = "TESTEMUNHA"
            elif any(x in papel_token for x in ("CONTRATANTE", "GRUPO_SBF", "EMISSOR")):
                cat = "REPRESENTANTE_CONTRATANTE"
            elif any(x in papel_token for x in ("CONTRAPARTE", "FORNECEDOR", "CONTRATADA")):
                cat = "REPRESENTANTE_CONTRAPARTE"
            else:
                cat = "OUTRO"
        saida.append({
            "nome": nome,
            "papel_cargo": papel,
            "categoria": cat,
            "email": _texto(item.get("email") or item.get("e-mail") or "Não localizado"),
            "data_assinatura": _data_br(item.get("data_assinatura") or item.get("Data da assinatura")) or _texto(item.get("data_assinatura") or item.get("Data da assinatura")),
            "data_reconhecimento_firma": _data_br(item.get("data_reconhecimento_firma") or item.get("Data do reconhecimento de firma")) or _texto(item.get("data_reconhecimento_firma") or item.get("Data do reconhecimento de firma") or "Não aplicável"),
            "fonte": fonte,
            "pagina": pagina or _texto(item.get("clausula_secao")),
            "status": "Assinado",
            "evidencia": evidencia,
        })
    return saida


def _assinaturas(base: MutableMapping[str, Any], raw: Mapping[str, Any], auditoria: List[Dict[str, Any]]) -> None:
    assinaturas = _normalizar_assinaturas(raw.get("assinaturas_contrato") or base.get("assinaturas_contrato"))
    base["assinaturas_contrato"] = assinaturas

    docusign = _data_br(base.get("data_conclusao_docusign"))
    if not docusign:
        txt_docu = _sem_acento(base.get("data_conclusao_docusign")).lower()
        if "assinatura fisica" in txt_docu or "nao aplicavel" in txt_docu:
            base["data_conclusao_docusign"] = "Não aplicável — assinatura física"

    representantes = [a for a in assinaturas if a.get("categoria") != "TESTEMUNHA"]
    testemunhas = [a for a in assinaturas if a.get("categoria") == "TESTEMUNHA"]
    datas = [a.get("data_assinatura") for a in assinaturas if _data_br(a.get("data_assinatura"))]
    reconhecimentos = [a.get("data_reconhecimento_firma") for a in assinaturas if _data_br(a.get("data_reconhecimento_firma"))]

    mapa = _mapa_auditoria(auditoria)
    assinatura_aud = mapa.get("contrato_assinado")
    prova_auditoria = _evidencia_confirma(assinatura_aud) and _sem_acento(assinatura_aud.get("valor")).lower() in {"sim", "assinado", "contrato assinado"}
    assinado = bool(representantes) or bool(docusign) or prova_auditoria

    if assinado:
        base["contrato_assinado"] = "Sim"
        if not docusign and not _data_br(base.get("data_conclusao_docusign")):
            base["data_conclusao_docusign"] = "Não aplicável — assinatura física"
        nomes = [a["nome"] for a in assinaturas]
        base["pessoas_que_assinaram"] = "; ".join(nomes)
        if datas:
            base["data_assinatura"] = _data_br(datas[0])
        if reconhecimentos:
            base["data_reconhecimento_firma"] = "; ".join(dict.fromkeys(_data_br(x) for x in reconhecimentos if _data_br(x)))
        partes = ["Contrato assinado"]
        if representantes:
            partes.append(f"{len(representantes)} representante(s) das partes")
        if testemunhas:
            partes.append(f"{len(testemunhas)} testemunha(s)")
        if _data_br(base.get("data_assinatura")):
            partes.append(f"data principal: {_data_br(base.get('data_assinatura'))}")
        if _valor_util(base.get("data_reconhecimento_firma")):
            partes.append(f"reconhecimento de firma: {_texto(base.get('data_reconhecimento_firma'))}")
        base["alerta_assinatura"] = ". ".join(partes) + "."
    else:
        base["contrato_assinado"] = "Não validado"
        base["pessoas_que_assinaram"] = "Não identificado com segurança"
        base["alerta_assinatura"] = "Não foi localizada evidência documental suficiente para validar a assinatura das partes."

    for campo, valor in (
        ("pessoas_que_assinaram", base.get("pessoas_que_assinaram")),
        ("data_assinatura", base.get("data_assinatura")),
        ("contrato_assinado", base.get("contrato_assinado")),
        ("alerta_assinatura", base.get("alerta_assinatura")),
    ):
        linha = mapa.get(campo) or _linha_nao_localizada(campo)
        if assinado and _valor_util(valor):
            fonte = assinaturas[0]["fonte"] if assinaturas else linha.get("arquivo_fonte")
            pagina = assinaturas[0]["pagina"] if assinaturas else linha.get("pagina")
            evidencia = assinaturas[0]["evidencia"] if assinaturas else linha.get("trecho_evidencia")
            linha.update({
                "valor": _texto(valor),
                "status": "CONFIRMADO",
                "tipo_dado": "CONSOLIDACAO_DOCUMENTAL",
                "arquivo_fonte": fonte or "Documento de assinatura",
                "pagina": pagina or "Bloco de assinaturas/certificado",
                "trecho_evidencia": evidencia or "Assinaturas validadas individualmente.",
                "confianca": max(_confianca(linha.get("confianca")), 90),
            })
        mapa[campo] = linha
    auditoria[:] = _ordenar_auditoria(mapa.values())


def _filtrar_checklist(raw: Any) -> List[Dict[str, Any]]:
    if not isinstance(raw, list):
        return []
    saida = []
    for item in raw:
        if not isinstance(item, Mapping):
            continue
        evidencia = _texto(item.get("Evidência") or item.get("evidencia") or item.get("trecho_evidencia"))
        fonte = _texto(item.get("Arquivo") or item.get("arquivo") or item.get("Fonte") or item.get("fonte"))
        pagina = _texto(item.get("Página") or item.get("pagina"))
        status = _texto(item.get("Status") or item.get("status"))
        concluido = _token(status) in {"CONCLUIDO", "CONFIRMADO", "VALIDADO"}
        if concluido and not (fonte and (pagina or re.search(r"pagina|clausula|anexo", evidencia, flags=re.I)) and _evidencia_util(evidencia)):
            continue
        novo = dict(item)
        novo["Validação"] = _texto(item.get("Validação") or item.get("validacao") or item.get("item") or "Validação documental")
        novo["Status"] = status or "Atenção"
        novo["Peso de risco"] = _texto(item.get("Peso de risco") or item.get("peso_risco") or item.get("Risco") or "Baixo")
        novo["Crítico"] = _texto(item.get("Crítico") or item.get("critico") or "Não")
        novo["Arquivo"] = fonte or "Não localizado"
        novo["Página"] = pagina or "Não localizado"
        novo["Evidência"] = evidencia or "Não localizado"
        saida.append(novo)
    return saida


def _filtrar_pendencias(raw: Any) -> List[Dict[str, Any]]:
    if not isinstance(raw, list):
        return []
    saida = []
    for item in raw:
        if not isinstance(item, Mapping):
            continue
        texto_total = _sem_acento(" ".join(_texto(v) for v in item.values())).lower()
        if "email" in texto_total and any(x in texto_total for x in ("signatario", "assinante", "representante")) and not any(x in texto_total for x in ("contrato exige", "clausula exige", "obrigatorio")):
            continue
        evidencia = _texto(item.get("Evidência") or item.get("evidencia") or item.get("trecho_evidencia"))
        fonte = _texto(item.get("Arquivo") or item.get("arquivo") or item.get("Fonte") or item.get("fonte"))
        pagina = _texto(item.get("Página") or item.get("pagina"))
        recomendacao = _texto(item.get("Recomendação") or item.get("recomendacao"))
        ref_objetiva = bool(re.search(r"\b(cl[aá]usula|anexo|p[aá]gina|item\s+\d+)\b", f"{evidencia} {recomendacao}", flags=re.I))
        if not ((_evidencia_util(evidencia) and fonte and (pagina or ref_objetiva)) or (fonte and ref_objetiva)):
            continue
        novo = dict(item)
        novo["Pendência"] = _texto(item.get("Pendência") or item.get("pendencia") or "Ponto de atenção documental")
        novo["Crítico"] = _texto(item.get("Crítico") or item.get("critico") or "Não")
        novo["Risco"] = _texto(item.get("Risco") or item.get("risco") or "Baixo")
        novo["Recomendação"] = recomendacao or "Validar antes de seguir."
        novo["Arquivo"] = fonte
        novo["Página"] = pagina or "Conforme referência da evidência"
        novo["Evidência"] = evidencia or recomendacao
        saida.append(novo)
    return saida


def _conflitos(raw: Any) -> List[Dict[str, Any]]:
    if not isinstance(raw, list):
        return []
    saida = []
    for item in raw:
        if not isinstance(item, Mapping):
            continue
        campo = _campo_canonico(item.get("campo") or item.get("Campo"))
        valores = item.get("valores_conflitantes") or item.get("valores")
        if not campo or not valores:
            continue
        saida.append({
            "Campo": campo,
            "Valores conflitantes": valores,
            "Arquivos": item.get("arquivos") or item.get("fontes") or "Não localizado",
            "Regra aplicada": item.get("regra_aplicada") or "Fonte de maior autoridade documental",
            "Decisão": item.get("decisao") or item.get("decisão") or "Revisão necessária",
        })
    return saida



def _gerar_checklist_deterministico(base: Mapping[str, Any], auditoria: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Reconstrói o checklist a partir da matriz de evidências, sem depender do resumo da IA."""
    mapa = _mapa_auditoria(auditoria)

    grupos = [
        ("Identificação das partes e CNPJs", ["empresa_grupo_sbf", "cnpj_empresa_grupo", "contraparte", "cnpj_contraparte"], "Alto", "Sim"),
        ("Definição do objeto e escopo", ["objetivo", "descricao_servico_material"], "Alto", "Sim"),
        ("Vigência contratual", ["vigencia_apos_assinatura", "periodo_vigencia_formatado", "tipo_vigencia"], "Alto", "Sim"),
        ("Assinaturas das partes", ["data_assinatura", "pessoas_que_assinaram"], "Alto", "Sim"),
        ("Valores, preços e tarifas", ["valor_contrato_original", "valor_mensal_estimado", "valor_total_materiais_servicos"], "Médio", "Sim"),
        ("Forma e condição de pagamento", ["forma_pagamento", "condicao_pagamento_dias"], "Médio", "Não"),
        ("Multas e juros", ["multa"], "Médio", "Não"),
        ("Rescisão e indenização", ["rescisao_indenizacao"], "Médio", "Sim"),
        ("Proteção de dados (LGPD)", ["protecao_dados_lgpd"], "Médio", "Não"),
        ("Anticorrupção", ["anticorrupcao"], "Médio", "Não"),
    ]
    saida: List[Dict[str, Any]] = []
    for titulo, campos, peso, critico in grupos:
        evidencias = [mapa.get(c) for c in campos if mapa.get(c)]
        confirmadas = [e for e in evidencias if _evidencia_confirma(e) or (_status(e.get("status")) == "CALCULADO" and _valor_util(e.get("valor")))]
        faltantes = [c for c in campos if not (mapa.get(c) and (_evidencia_confirma(mapa.get(c)) or (_status(mapa.get(c).get("status")) == "CALCULADO" and _valor_util(mapa.get(c).get("valor")))))]
        ref = confirmadas[0] if confirmadas else (evidencias[0] if evidencias else {})
        if not faltantes:
            status = "Confirmado"
        elif confirmadas:
            status = "Atenção - confirmação parcial"
        else:
            status = "Não localizado"
        ev_textos = []
        for e in confirmadas[:3]:
            trecho = _texto(e.get("trecho_evidencia"))
            if _evidencia_util(trecho):
                ev_textos.append(trecho)
        evidencia = " | ".join(dict.fromkeys(ev_textos)) or ("Campos não localizados: " + ", ".join(faltantes))
        saida.append({
            "Validação": titulo,
            "Status": status,
            "Peso de risco": peso,
            "Crítico": critico,
            "Arquivo": _texto(ref.get("arquivo_fonte")) or "Não localizado",
            "Página": _texto(ref.get("pagina")) or _texto(ref.get("clausula_secao")) or "Não localizado",
            "Evidência": evidencia,
        })

    aditivos = base.get("aditivos_contrato") if isinstance(base.get("aditivos_contrato"), list) else []
    saida.append({
        "Validação": "Aditivos",
        "Status": "Confirmado" if aditivos else "Não aplicável - nenhum aditivo identificado",
        "Peso de risco": "Médio",
        "Crítico": "Não",
        "Arquivo": _texto(aditivos[0].get("Anexo do aditivo") if aditivos else "Não localizado"),
        "Página": _texto(aditivos[0].get("Página") if aditivos else "Não localizado"),
        "Evidência": f"{len(aditivos)} aditivo(s) com evidência documental." if aditivos else "Nenhum termo aditivo foi identificado no pacote analisado.",
    })

    faltantes = [i.get("campo") for i in auditoria if _status(i.get("status")) == "NÃO_LOCALIZADO"]
    if faltantes:
        saida.append({
            "Validação": "Campos não localizados",
            "Status": "Atenção",
            "Peso de risco": "Baixo",
            "Crítico": "Não",
            "Arquivo": "Matriz de evidências",
            "Página": "Múltiplas",
            "Evidência": ", ".join(ROTULO_POR_CAMPO.get(c, c) for c in faltantes),
        })
    conflitos = base.get("conflitos_documentais") if isinstance(base.get("conflitos_documentais"), list) else []
    saida.append({
        "Validação": "Conflitos documentais",
        "Status": "Atenção" if conflitos else "Confirmado - nenhum conflito",
        "Peso de risco": "Alto" if conflitos else "Baixo",
        "Crítico": "Sim" if conflitos else "Não",
        "Arquivo": "Múltiplos documentos" if conflitos else "Matriz de evidências",
        "Página": "Múltiplas" if conflitos else "Não aplicável",
        "Evidência": f"{len(conflitos)} conflito(s) documental(is) registrado(s)." if conflitos else "Nenhum conflito documental confirmado.",
    })
    return saida


def _classificar_indicadores_pendencias(base: MutableMapping[str, Any]) -> None:
    pendencias = base.get("pendencias") if isinstance(base.get("pendencias"), list) else []
    criticas = [p for p in pendencias if _token(p.get("Crítico")) in {"SIM", "TRUE", "1"}]
    pontos = [p for p in pendencias if p not in criticas]
    campos = base.get("campos_nao_localizados") if isinstance(base.get("campos_nao_localizados"), list) else []
    base["indicadores_pendencias"] = {
        "pendencias_criticas": len(criticas),
        "pontos_atencao": len(pontos),
        "campos_nao_localizados": len(campos),
    }

def _score_risco(base: MutableMapping[str, Any], auditoria: List[Dict[str, Any]]) -> None:
    """Calcula confiança de extração separadamente do risco contratual."""
    mapa = _mapa_auditoria(auditoria)
    pesos_total = 0
    pesos_confirmados = 0
    inferidos = 0
    nao_localizados = 0
    for _, campo in CAMPOS_OFICIAIS_V4:
        # Campos derivados não devem reduzir a cobertura quando a base documental existe.
        peso = 3 if campo in CAMPOS_CRITICOS else 1
        pesos_total += peso
        item = mapa.get(campo)
        if _evidencia_confirma(item) or (item and _status(item.get("status")) == "CALCULADO" and _valor_util(item.get("valor"))):
            pesos_confirmados += peso
        elif item and _status(item.get("status")) == "INFERIDO":
            inferidos += 1
        else:
            nao_localizados += 1

    cobertura = round((pesos_confirmados / pesos_total) * 100) if pesos_total else 0
    paginas_processadas = int(base.get("paginas_processadas") or 0) if str(base.get("paginas_processadas") or "").isdigit() else 0
    total_paginas = int(base.get("total_paginas") or 0) if str(base.get("total_paginas") or "").isdigit() else 0
    paginas_pct = round((paginas_processadas / total_paginas) * 100) if total_paginas else 100
    conflitos = len(base.get("conflitos_documentais") or [])
    confianca = max(0, min(100, round(cobertura * 0.85 + paginas_pct * 0.15 - inferidos - conflitos * 2)))

    pendencias = base.get("pendencias") if isinstance(base.get("pendencias"), list) else []
    tem_critico_alto = any(_token(p.get("Crítico")) in {"SIM", "TRUE", "1"} and _token(p.get("Risco")) == "ALTO" for p in pendencias)
    tem_critico = any(_token(p.get("Crítico")) in {"SIM", "TRUE", "1"} for p in pendencias)
    tem_alto = any(_token(p.get("Risco")) == "ALTO" for p in pendencias)
    if tem_critico_alto:
        risco_final = "ALTO"
    elif tem_critico or tem_alto or conflitos:
        risco_final = "MÉDIO"
    else:
        risco_final = "BAIXO"

    base["confianca_extracao"] = confianca
    base["score"] = confianca  # compatibilidade com histórico antigo
    base["risco"] = risco_final
    base["metricas_confianca"] = {
        "cobertura_campos_percentual": cobertura,
        "paginas_processadas_percentual": paginas_pct,
        "cobertura_paginas_percentual": paginas_pct,
        "campos_confirmados_ponderados": pesos_confirmados,
        "campos_totais_ponderados": pesos_total,
        "campos_nao_localizados": nao_localizados,
        "campos_inferidos": inferidos,
        "conflitos": conflitos,
        "pendencias_com_evidencia": len(pendencias),
        "confianca_extracao_percentual": confianca,
        "score_final": confianca,
    }


def _resumo_parecer(base: MutableMapping[str, Any]) -> None:
    tipo = _texto(base.get("tipo_contrato"))
    contratante = _texto(base.get("empresa_grupo_sbf"))
    contraparte = _texto(base.get("contraparte"))
    objeto = _texto(base.get("objetivo") or base.get("descricao_servico_material"))
    periodo = _texto(base.get("periodo_vigencia_formatado") or base.get("vigencia_apos_assinatura"))
    status = _texto(base.get("status_contratual") or base.get("status"))
    operacional = _texto(base.get("situacao_operacional"))

    partes = []
    if _valor_util(tipo):
        partes.append(tipo.rstrip("."))
    if _valor_util(contratante) and _valor_util(contraparte):
        partes.append(f"firmado entre {contratante} e {contraparte}")
    if _valor_util(objeto):
        partes.append(f"objeto: {objeto.rstrip('.')}" )
    if _valor_util(periodo):
        partes.append(f"período técnico: {periodo.rstrip('.')}" )
    if _valor_util(status):
        partes.append(f"status contratual: {status.rstrip('.')}" )
    if _valor_util(operacional):
        partes.append(f"situação operacional: {operacional.rstrip('.')}" )
    partes.append(f"assinatura validada: {_texto(base.get('contrato_assinado')) or 'Não validada'}")
    base["resumo_executivo"] = ". ".join(partes).strip() + "."

    financeiros = [
        _texto(base.get("valor_contrato_original")),
        _texto(base.get("valor_mensal_estimado")),
        _texto(base.get("valor_total_materiais_servicos")),
        _texto(base.get("valor_total_estimado_vigencia")),
    ]
    financeiros = [x for x in financeiros if _valor_util(x)]
    ind = base.get("indicadores_pendencias") if isinstance(base.get("indicadores_pendencias"), Mapping) else {}
    parecer = []
    if financeiros:
        parecer.append("Financeiro: " + " ".join(financeiros))
    parecer.append(
        f"Pendências críticas: {ind.get('pendencias_criticas', 0)}; "
        f"pontos de atenção: {ind.get('pontos_atencao', 0)}; "
        f"campos não localizados: {ind.get('campos_nao_localizados', 0)}."
    )
    parecer.append("Valores de implantação, mensalidade e tarifas variáveis permanecem separados por natureza e periodicidade.")
    parecer.append("A situação operacional atual deve ser confirmada nos sistemas corporativos quando o pacote documental não trouxer prova atual de encerramento ou continuidade.")
    base["parecer"] = " ".join(parecer)


def _mapa_cards(auditoria: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    saida = {}
    for item in auditoria:
        campo = _campo_canonico(item.get("campo"))
        if not campo:
            continue
        saida[campo] = {
            "status": _status(item.get("status")),
            "arquivo": _texto(item.get("arquivo_fonte")),
            "pagina": _texto(item.get("pagina")),
            "secao": _texto(item.get("clausula_secao")),
            "evidencia": _texto(item.get("trecho_evidencia")),
            "confianca": _confianca(item.get("confianca")),
        }
    return saida


def aplicar_motor_evidencias_v4(
    resultado: Mapping[str, Any] | None,
    resultado_bruto: Mapping[str, Any] | None = None,
    texto_extraido: str = "",
) -> Dict[str, Any]:
    """Aplica a consolidação determinística final.

    A trava rígida é ativada somente quando a resposta contém auditoria_campos.
    Registros históricos antigos permanecem em compatibilidade, sem apagamento.
    """
    base: Dict[str, Any] = deepcopy(dict(resultado or {}))
    bruto: Dict[str, Any] = dict(resultado_bruto or {})
    raw_audit = bruto.get("auditoria_campos") or base.get("auditoria_campos")
    auditoria = _normalizar_auditoria(raw_audit)
    if not auditoria:
        base["motor_evidencias_v4"] = "Modo compatibilidade — análise antiga sem matriz de evidências"
        return base

    mapa = _mapa_auditoria(auditoria)

    # Campos diretos: prevalece exclusivamente a evidência auditada.
    campos_derivados = {
        "periodo_vigencia_formatado",
        "valor_contrato_original",
        "valor_mensal_estimado",
        "valor_total_estimado_vigencia",
        "valor_total_materiais_servicos",
        "pessoas_que_assinaram",
    }
    for _, campo in CAMPOS_OFICIAIS_V4:
        if campo in campos_derivados:
            continue
        item = mapa.get(campo)
        if _evidencia_confirma(item):
            base[campo] = _texto(item.get("valor"))
        elif item and _status(item.get("status")) == "NÃO_APLICÁVEL" and _valor_util(item.get("valor")):
            base[campo] = _texto(item.get("valor"))
        else:
            base[campo] = "Não identificado com segurança"

    base["conflitos_documentais"] = _conflitos(bruto.get("conflitos_documentais") or base.get("conflitos_documentais"))
    base["paginas_processadas"] = bruto.get("paginas_processadas") or base.get("paginas_processadas") or 0
    base["total_paginas"] = bruto.get("total_paginas") or base.get("total_paginas") or 0
    base["secoes_verificadas"] = bruto.get("secoes_verificadas") or base.get("secoes_verificadas") or []
    base["tabelas_verificadas"] = bruto.get("tabelas_verificadas") or base.get("tabelas_verificadas") or []
    base["assinaturas_verificadas"] = bruto.get("assinaturas_verificadas") or base.get("assinaturas_verificadas") or []

    _vigencia_status(base, auditoria)
    _financeiro(base, bruto, auditoria)
    _assinaturas(base, bruto, auditoria)

    base["aditivos_contrato"] = [a for a in (bruto.get("aditivos_contrato") or base.get("aditivos_contrato") or []) if isinstance(a, Mapping) and _valor_util(a.get("Anexo do aditivo") or a.get("arquivo_fonte") or a.get("fonte")) and _evidencia_util(a.get("trecho_evidencia") or a.get("Evidência") or a.get("evidencia"))]
    if base["aditivos_contrato"]:
        base["resumo_aditivos"] = f"{len(base['aditivos_contrato'])} aditivo(s) identificado(s) com evidência documental."
    else:
        base["resumo_aditivos"] = "Nenhum aditivo identificado com evidência documental no pacote analisado."

    base["pendencias"] = _filtrar_pendencias(bruto.get("pendencias") or base.get("pendencias"))

    auditoria = _ordenar_auditoria(auditoria)
    base["auditoria_campos"] = auditoria
    base["mapa_evidencias_cards"] = _mapa_cards(auditoria)
    base["campos_nao_localizados"] = [i["campo"] for i in auditoria if _status(i.get("status")) == "NÃO_LOCALIZADO"]
    base["checklist"] = _gerar_checklist_deterministico(base, auditoria)
    _classificar_indicadores_pendencias(base)
    _score_risco(base, auditoria)
    _resumo_parecer(base)
    base["motor_evidencias_v4"] = "Ativo — evidência obrigatória, consolidação determinística e score calculado"
    base["texto_extraido"] = base.get("texto_extraido") or texto_extraido
    return base


def linhas_auditoria_para_tela(resultado: Mapping[str, Any]) -> List[Dict[str, Any]]:
    auditoria = _normalizar_auditoria(resultado.get("auditoria_campos"))
    linhas = []
    for item in auditoria:
        linhas.append({
            "Campo": item.get("rotulo") or ROTULO_POR_CAMPO.get(item.get("campo"), item.get("campo")),
            "Valor consolidado": item.get("valor") or "Não identificado com segurança",
            "Status": item.get("status"),
            "Tipo": item.get("tipo_dado"),
            "Arquivo": item.get("arquivo_fonte") or "Não localizado",
            "Página/Seção": " • ".join(x for x in (item.get("pagina"), item.get("clausula_secao")) if _valor_util(x)) or "Não localizado",
            "Evidência": item.get("trecho_evidencia") or "Não localizado",
            "Confiança": item.get("confianca", 0),
        })
    return linhas
