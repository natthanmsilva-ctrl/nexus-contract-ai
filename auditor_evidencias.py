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
- Foro eleito, comarca competente ou cidade escolhida para dirimir conflitos NÃO são local de prestação.
- condicao_pagamento_dias deve retornar somente no padrão executivo DD, por exemplo: 15DD, 30DD, 60DD ou 90DD. Use exclusivamente cláusula financeira de fatura, vencimento ou pagamento da remuneração. Não use prazos operacionais de entrega de informações, relatórios, atendimento ou execução. A frase completa permanece em forma_pagamento.
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



def _paginas_do_texto_extraido(texto_extraido: Any) -> Dict[int, str]:
    """Separa o texto OCR por página sem depender do formato do PDF original."""
    texto = str(texto_extraido or "").replace("\x00", " ")
    partes = re.split(r"---\s*P[ÁA]GINA\s+(\d+)\s+OCR\s*---", texto, flags=re.I)
    paginas: Dict[int, str] = {}
    for idx in range(1, len(partes), 2):
        try:
            numero = int(partes[idx])
        except Exception:
            continue
        paginas[numero] = _texto(partes[idx + 1])
    return paginas


def _trecho_limitado(texto: Any, limite: int = 520) -> str:
    valor = _texto(texto)
    if len(valor) <= limite:
        return valor
    cortado = valor[:limite].rsplit(" ", 1)[0].rstrip(" ,;:")
    return cortado + "..."


def _atualizar_auditoria(
    auditoria: List[Dict[str, Any]],
    campo: str,
    valor: Any,
    *,
    status: str,
    tipo_dado: str,
    arquivo: str,
    pagina: str,
    secao: str,
    evidencia: str,
    confianca: int,
) -> None:
    mapa = _mapa_auditoria(auditoria)
    linha = mapa.get(campo) or _linha_nao_localizada(campo)
    linha.update({
        "valor": _texto(valor),
        "status": _status(status),
        "tipo_dado": _texto(tipo_dado).upper(),
        "arquivo_fonte": _texto(arquivo),
        "pagina": _texto(pagina),
        "clausula_secao": _texto(secao),
        "trecho_evidencia": _trecho_limitado(evidencia),
        "confianca": _confianca(confianca),
    })
    mapa[campo] = linha
    auditoria[:] = _ordenar_auditoria(mapa.values())


def _extrair_descricao_servico_documental(paginas: Mapping[int, str]) -> Tuple[str, str, str, str]:
    """Recupera objeto detalhado e descrição breve a partir da cláusula de serviços."""
    candidatos = list(paginas.items())
    candidatos.sort(key=lambda x: (0 if x[0] <= 6 else 1, x[0]))
    for numero, pagina in candidatos:
        sem = _sem_acento(pagina).lower()
        if "servicos contemplam as seguintes atividades" not in sem:
            continue
        m = re.search(
            r"os\s+servi[cç]os\s+contemplam\s+as\s+seguintes\s+atividades\s*:\s*(.*?)"
            r"(?=(?:\b2[\.\sÃA]*1[\.\s]*1\b|\bos\s+servi[cç]os\s+objeto\s+deste\s+contrato\b|\b2[\.\s]*1[\.\s]*2\b|\|?3\.?\s+[-—]?\s*confidencialidade))",
            pagina,
            flags=re.I | re.S,
        )
        bloco = _texto(m.group(1) if m else pagina)
        atividades = []
        for item in re.findall(r"\((?:i|ii|iii|iv|v|vi)\)\s*(.*?)(?=\((?:i|ii|iii|iv|v|vi)\)|$)", bloco, flags=re.I | re.S):
            limpo = _texto(item).strip(" ;,.")
            limpo = re.sub(r"^e\s+", "", limpo, flags=re.I)
            limpo = re.sub(r"[;,]\s*e$", "", limpo, flags=re.I).strip(" ;,.")
            # Remove resíduos de numeração da cláusula seguinte, comuns no OCR
            # (ex.: 2ÃA antes de “Os serviços objeto deste Contrato”).
            limpo = re.sub(r"\s+\d+[A-Za-zÀ-ÿÃÂ.]*$", "", limpo).strip(" ;,.")
            if len(limpo) >= 12:
                sem_limpo = _sem_acento(limpo).lower()
                if "abertura e manutencao" in sem_limpo and "livros de registro" in sem_limpo:
                    limpo = "abertura e manutenção de livros de registro"
                elif "registro das informacoes relativas a titularidade" in sem_limpo or ("titularidade" in sem_limpo and "gravames" in sem_limpo):
                    limpo = "registro de titularidade e gravames"
                elif "tratamento das instrucoes de movimentacao" in sem_limpo:
                    limpo = "tratamento de instruções de movimentação"
                elif "tratamento de eventos incidentes" in sem_limpo:
                    limpo = "eventos incidentes sobre as ações"
                atividades.append(limpo)
        if atividades:
            if len(atividades) == 1:
                lista = atividades[0]
            else:
                lista = ", ".join(atividades[:-1]) + " e " + atividades[-1]
            descricao = "Serviços de escrituração de ações, incluindo " + lista.rstrip(".") + "."
        else:
            descricao = _trecho_limitado(bloco, 900)

        objetivo = ""
        mo = re.search(
            r"objeto\s+deste\s+contrato\s+[ée]\s+a\s+presta[cç][aã]o\s+de\s+servi[cç]os\s+de\s+(.*?)\s+pela\s+",
            pagina,
            flags=re.I | re.S,
        )
        if mo:
            nucleo = _texto(mo.group(1)).strip(" .")
            breve = "Serviços de " + nucleo.lower()
            breve = breve[:1].upper() + breve[1:]
        elif "escrituração de ações" in pagina.lower() or "escrituracao de acoes" in sem:
            breve = "Serviços de escrituração de ações"
        else:
            breve = "Serviços contratados conforme cláusula de objeto e escopo"
        evidencia = ("Os serviços contemplam as seguintes atividades: " + bloco).replace("Anexo |", "Anexo I")
        return descricao, breve.rstrip(".") + ".", str(numero), _trecho_limitado(evidencia)
    return "", "", "", ""


def _extrair_data_instrumento_documental(paginas: Mapping[int, str]) -> Tuple[str, str, str]:
    """Prioriza a data escrita no bloco de assinatura do instrumento."""
    ordem = sorted(
        paginas.items(),
        key=lambda x: (0 if any(t in _sem_acento(x[1]).lower() for t in ("firmam as partes", "testemunhas")) else 1, x[0]),
    )
    for numero, pagina in ordem:
        data = _data_br(pagina)
        if not data:
            continue
        sem = _sem_acento(pagina).lower()
        if "firmam as partes" in sem or "testemunhas" in sem or "sao paulo" in sem:
            m = re.search(r"(?:S[aã]o\s+Paulo\s*,?\s*)?\d{1,2}\s+de\s+[A-Za-zÀ-ÿ]+\s+de\s+\d{4}", pagina, flags=re.I)
            evidencia = m.group(0) if m else data
            return data, str(numero), evidencia
    return "", "", ""




def _extrair_condicao_pagamento_dd_documental(paginas: Mapping[int, str]) -> Tuple[str, str, str]:
    """Extrai a condição de pagamento em formato executivo ``15DD``/``30DD``.

    V10: procura primeiro cláusulas financeiras explícitas em *todas* as páginas
    e só depois considera construções genéricas. Isso evita que um prazo
    operacional anterior (por exemplo, entrega de informação em 5 dias úteis)
    vença uma cláusula posterior de faturamento/vencimento.
    """
    candidatos: List[Tuple[int, int, int, str, str]] = []

    # Maior prioridade: vencimento em dia definido do mês subsequente.
    # Ex.: "vencimento até o dia 15 (quinze) do mês subsequente".
    padrao_dia_mes = re.compile(
        r"vencimento\s+at[eé]\s+o\s+dia\s+(\d{1,3})(?:\s*\([^)]*\))?"
        r"\s+do\s+m[eê]s\s+subsequente",
        flags=re.I | re.S,
    )

    for numero, pagina in paginas.items():
        for m in padrao_dia_mes.finditer(pagina):
            try:
                dias = int(m.group(1))
            except Exception:
                continue
            if not 1 <= dias <= 365:
                continue
            inicio = max(0, m.start() - 180)
            fim = min(len(pagina), m.end() + 220)
            contexto = pagina[inicio:fim]
            contexto_sem = _sem_acento(contexto).lower()
            # A cláusula deve estar inserida em contexto financeiro real.
            if not any(x in contexto_sem for x in ("fatura", "pagamento", "remuneracao", "pagar", "vencimento")):
                continue
            candidatos.append((100, numero, m.start(), f"{dias}DD", _trecho_limitado(contexto, 420)))

    # Se existe uma cláusula explícita de vencimento, ela prevalece sobre todos
    # os demais prazos do documento, independentemente da ordem das páginas.
    if candidatos:
        candidatos.sort(key=lambda x: (-x[0], x[1], x[2]))
        _, numero, _, valor, evidencia = candidatos[0]
        return valor, str(numero), evidencia

    padroes_genericos = [
        re.compile(
            r"(?:pagamento|vencimento|fatura|nota\s+fiscal|remunera[cç][aã]o).{0,140}?"
            r"(?:em\s+at[eé]|no\s+prazo\s+de|prazo\s+de|at[eé])\s+"
            r"(\d{1,3})(?:\s*\([^)]*\))?\s+dias?",
            flags=re.I | re.S,
        ),
        re.compile(
            r"(\d{1,3})(?:\s*\([^)]*\))?\s+dias?\s*(?:corridos|[uú]teis)?"
            r".{0,140}?(?:emiss[aã]o\s+da\s+nota|nota\s+fiscal|recebimento\s+da\s+fatura|aprova[cç][aã]o\s+da\s+fatura)",
            flags=re.I | re.S,
        ),
    ]
    marcadores_operacionais = (
        "relacao dos acionistas",
        "disponibilizara ao emissor",
        "entrega de informacao",
        "fornecimento de informacao",
        "prazo de resposta",
        "relatorio",
        "documentos solicitados",
        "execucao dos servicos",
    )
    marcadores_financeiros_fortes = (
        "fatura",
        "nota fiscal",
        "vencimento",
        "pagamento da remuneracao",
        "pagara a remuneracao",
        "recebimento da fatura",
        "emissao da nota",
    )

    for numero, pagina in paginas.items():
        sem_pagina = _sem_acento(pagina).lower()
        if not any(t in sem_pagina for t in ("pagamento", "vencimento", "fatura", "nota fiscal", "remuneracao")):
            continue
        for indice, padrao in enumerate(padroes_genericos):
            for m in padrao.finditer(pagina):
                try:
                    dias = int(m.group(1))
                except Exception:
                    continue
                if not 1 <= dias <= 365:
                    continue
                inicio = max(0, m.start() - 90)
                fim = min(len(pagina), m.end() + 120)
                contexto = pagina[inicio:fim]
                sem_contexto = _sem_acento(contexto).lower()
                tem_financeiro_forte = any(x in sem_contexto for x in marcadores_financeiros_fortes)
                tem_operacional = any(x in sem_contexto for x in marcadores_operacionais)
                # Rejeita prazo operacional que apenas esteja próximo de uma
                # menção genérica a pagamento em outro assunto da cláusula.
                if tem_operacional and not tem_financeiro_forte:
                    continue
                score = 80 if indice == 0 else 75
                if tem_financeiro_forte:
                    score += 10
                candidatos.append((score, numero, m.start(), f"{dias}DD", _trecho_limitado(contexto, 420)))

    if not candidatos:
        return "", "", ""
    candidatos.sort(key=lambda x: (-x[0], x[1], x[2]))
    _, numero, _, valor, evidencia = candidatos[0]
    return valor, str(numero), evidencia

def _validar_local_prestacao_semantico(
    base: MutableMapping[str, Any],
    auditoria: List[Dict[str, Any]],
) -> None:
    """Impede que foro, sede ou endereço cadastral virem local de execução.

    O campo só permanece confirmado quando a evidência contém linguagem explícita
    de prestação/execução dos serviços naquele local.
    """
    mapa = _mapa_auditoria(auditoria)
    item = mapa.get("local_prestacao")
    if not item:
        return

    status = _status(item.get("status"))
    if status == "NÃO_LOCALIZADO":
        base["local_prestacao"] = "Não localizado com segurança"
        return

    valor = _texto(item.get("valor") or base.get("local_prestacao"))
    evidencia = _texto(item.get("trecho_evidencia"))
    secao = _texto(item.get("clausula_secao"))
    contexto = _sem_acento(" ".join((valor, evidencia, secao))).lower()

    marcadores_execucao = (
        "local de execucao",
        "execucao dos servicos",
        "prestacao dos servicos ocorrera",
        "servicos serao prestados",
        "servicos deverao ser prestados",
        "atividades serao executadas",
        "nas dependencias da",
        "unidade onde os servicos",
        "estabelecimento onde os servicos",
    )
    marcadores_forum = (
        "foro",
        "comarca",
        "dirimir",
        "competente para",
        "elegem o foro",
        "foro eleito",
        "capital do estado",
    )
    marcadores_endereco = (
        "endereco",
        "preambulo",
        "qualificacao das partes",
        "cadastro da parte",
        "sede",
        "domicilio",
        "cep",
        "logradouro",
        "avenida",
        " rua ",
    )

    execucao_explicita = any(x in contexto for x in marcadores_execucao)
    referencia_forum = any(x in contexto for x in marcadores_forum)
    referencia_endereco = any(x in contexto for x in marcadores_endereco)

    if not execucao_explicita and (referencia_forum or referencia_endereco):
        motivo = (
            "A referência localizada trata de foro eleito e não comprova o local de execução dos serviços."
            if referencia_forum
            else "O endereço cadastral/sede da parte não comprova o local de execução dos serviços."
        )
        base["local_prestacao"] = "Não localizado com segurança"
        _atualizar_auditoria(
            auditoria,
            "local_prestacao",
            "Não localizado com segurança",
            status="NÃO_LOCALIZADO",
            tipo_dado="DADO_DOCUMENTAL",
            arquivo=_texto(item.get("arquivo_fonte")),
            pagina=_texto(item.get("pagina")),
            secao="Validação semântica do local de prestação",
            evidencia=motivo,
            confianca=100,
        )

def _recuperar_campos_documentais(
    base: MutableMapping[str, Any],
    bruto: Mapping[str, Any],
    auditoria: List[Dict[str, Any]],
    texto_extraido: str,
) -> None:
    """Recupera fatos presentes no OCR quando a segunda passagem da IA os omite."""
    paginas = _paginas_do_texto_extraido(texto_extraido)
    if not paginas:
        return
    arquivo = _texto(bruto.get("arquivo_principal") or bruto.get("nome_arquivo") or "Contrato principal")
    # Quando há uma única fonte, o nome real costuma aparecer nas evidências/itens.
    for item in auditoria:
        fonte = _texto(item.get("arquivo_fonte"))
        if _valor_util(fonte) and "múltipl" not in _sem_acento(fonte).lower():
            arquivo = fonte
            break

    descricao, breve, pag_desc, ev_desc = _extrair_descricao_servico_documental(paginas)
    if descricao:
        base["descricao_servico_material"] = descricao
        _atualizar_auditoria(
            auditoria, "descricao_servico_material", descricao,
            status="CONFIRMADO", tipo_dado="DADO_DOCUMENTAL", arquivo=arquivo,
            pagina=pag_desc, secao="Cláusula 2.1 — Descrição dos serviços",
            evidencia=ev_desc, confianca=100,
        )
        base["descricao_breve_cadastro"] = breve
        _atualizar_auditoria(
            auditoria, "descricao_breve_cadastro", breve,
            status="CALCULADO", tipo_dado="INTERPRETACAO", arquivo=arquivo,
            pagina=pag_desc, secao="Síntese da cláusula 2.1",
            evidencia=ev_desc, confianca=95,
        )

    data_contrato, pag_data, ev_data = _extrair_data_instrumento_documental(paginas)
    if data_contrato:
        base["data_contrato"] = data_contrato
        _atualizar_auditoria(
            auditoria, "data_contrato", data_contrato,
            status="CONFIRMADO", tipo_dado="DADO_DOCUMENTAL", arquivo=arquivo,
            pagina=pag_data, secao="Bloco de assinaturas do instrumento",
            evidencia=ev_data, confianca=100,
        )

    # Forma de pagamento da remuneração deve vir do anexo financeiro. Não
    # confundir com o procedimento operacional de débito de créditos aos acionistas.
    for numero, pagina in paginas.items():
        sem_pag = _sem_acento(pagina).lower()
        if "remetera fatura" in sem_pag and "disponibilizacao do valor" in sem_pag and "conta corrente indicada" in sem_pag:
            forma = (
                "Faturamento mensal, com vencimento até o dia 15 do mês subsequente, "
                "mediante disponibilização do valor na conta corrente indicada pelo Emissor."
            )
            m_pag = re.search(
                r"Mensalmente.*?vencimento\s+at[eé]\s+o\s+dia\s+15.*?m[eê]s\s+subsequente.*?"
                r"O\s+EMISSOR\s+pagar[aá].*?conta\s+corrente\s+indicada\s+pelo\s+EMISSOR.*?(?=\b5\.|$)",
                pagina, flags=re.I | re.S,
            )
            evidencia_pag = m_pag.group(0) if m_pag else pagina
            base["forma_pagamento"] = forma
            _atualizar_auditoria(
                auditoria, "forma_pagamento", forma,
                status="CONFIRMADO", tipo_dado="DADO_DOCUMENTAL", arquivo=arquivo,
                pagina=str(numero), secao="Anexo II — itens 3 e 4",
                evidencia=evidencia_pag, confianca=100,
            )
            break

    # Card executivo separado: somente o prazo em formato DD (15DD, 30DD, 60DD...).
    condicao_dd, pag_dd, ev_dd = _extrair_condicao_pagamento_dd_documental(paginas)
    if condicao_dd:
        base["condicao_pagamento_dias"] = condicao_dd
        _atualizar_auditoria(
            auditoria, "condicao_pagamento_dias", condicao_dd,
            status="CONFIRMADO", tipo_dado="DADO_DOCUMENTAL", arquivo=arquivo,
            pagina=pag_dd, secao="Condição de pagamento / vencimento",
            evidencia=ev_dd, confianca=100,
        )

    # Validação semântica obrigatória: foro, sede e endereço cadastral não são
    # local de prestação sem cláusula expressa de execução dos serviços.
    _validar_local_prestacao_semantico(base, auditoria)

    texto_total_sem = _sem_acento(" ".join(paginas.values())).lower()
    paginas_assinatura = [
        (n, p) for n, p in paginas.items()
        if "firmam as partes" in _sem_acento(p).lower() or "testemunhas" in _sem_acento(p).lower()
    ]
    tem_assinatura_fisica = bool(paginas_assinatura)
    tem_certificado_docusign = any(x in texto_total_sem for x in ("certificate of completion", "envelope id", "docusign"))
    if tem_assinatura_fisica and not tem_certificado_docusign:
        pag_ass, txt_ass = paginas_assinatura[0]
        valor_docu = "Não aplicável — assinatura física"
        base["data_conclusao_docusign"] = valor_docu
        m_fisica = re.search(
            r"E,?\s+por\s+estarem.*?firmam\s+as\s+Partes.*?(?:S[aã]o\s+Paulo\s*,?\s*\d{1,2}\s+de\s+[A-Za-zÀ-ÿ]+\s+de\s+\d{4})",
            txt_ass, flags=re.I | re.S,
        )
        evidencia_fisica = m_fisica.group(0) if m_fisica else "Bloco de assinaturas físicas das partes localizado no instrumento."
        _atualizar_auditoria(
            auditoria, "data_conclusao_docusign", valor_docu,
            status="NÃO_APLICÁVEL", tipo_dado="INTERPRETACAO", arquivo=arquivo,
            pagina=str(pag_ass), secao="Bloco de assinaturas físicas",
            evidencia=evidencia_fisica, confianca=100,
        )

    # V10: se a matriz já confirmou os nomes no bloco físico, reconstrói os
    # registros individuais antes da consolidação final.
    _recuperar_assinaturas_documentais(base, bruto, auditoria, paginas, arquivo)

    # Ausência de aditivo é uma conclusão sobre o pacote, não um campo documental ausente.
    aditivos = bruto.get("aditivos_contrato") if isinstance(bruto.get("aditivos_contrato"), list) else []
    if not aditivos:
        resumo = "Nenhum aditivo identificado com evidência documental no pacote analisado."
        base["resumo_aditivos"] = resumo
        _atualizar_auditoria(
            auditoria, "resumo_aditivos", resumo,
            status="NÃO_APLICÁVEL", tipo_dado="INTERPRETACAO", arquivo="Pacote documental",
            pagina="Todas as páginas", secao="Triagem e consolidação dos anexos",
            evidencia="Nenhum arquivo classificado como termo aditivo foi identificado no pacote analisado.",
            confianca=100,
        )

    # Condição comercial mínima de 24 meses deve permanecer separada da vigência indeterminada.
    pagina_prazo = ""
    texto_prazo = ""
    for numero, pagina in paginas.items():
        sem = _sem_acento(pagina).lower()
        if "24 (vinte e quatro) meses" in sem or ("prazo minimo" in sem and "meses remanescentes" in sem):
            pagina_prazo, texto_prazo = str(numero), pagina
            break
    if texto_prazo:
        mapa_rescisao = _mapa_auditoria(auditoria)
        linha_rescisao = mapa_rescisao.get("rescisao_indenizacao") or {}
        atual = _texto(
            base.get("rescisao_indenizacao")
            or bruto.get("rescisao_indenizacao")
            or linha_rescisao.get("valor")
        )
        if not _valor_util(atual):
            atual = "Denúncia imotivada mediante aviso prévio de 30 dias."
        if "24 meses" not in _sem_acento(atual).lower() and "vinte e quatro" not in _sem_acento(atual).lower():
            atual = atual.rstrip(" .") + ". Compromisso comercial mínimo de 24 meses; em caso de rescisão unilateral antecipada pelo Emissor, será devido valor equivalente às mensalidades restantes até o fim do prazo mínimo."
        base["rescisao_indenizacao"] = atual
        mapa = _mapa_auditoria(auditoria)
        linha_atual = mapa.get("rescisao_indenizacao") or {}
        pag_atual = _texto(linha_atual.get("pagina"))
        secao_atual = _texto(linha_atual.get("clausula_secao"))
        evidencia_atual = _texto(linha_atual.get("trecho_evidencia"))
        m_prazo = re.search(
            r"O\s+EMISSOR\s+reconhece\s+e\s+concorda.*?24\s*\(vinte\s+e\s+quatro\)\s+meses.*?meses\s+remanescentes.*?(?:disposi[cç][oõ]es\s+espec[ií]ficas|$)",
            texto_prazo, flags=re.I | re.S,
        )
        evidencia_prazo = m_prazo.group(0) if m_prazo else _trecho_limitado(texto_prazo, 420)
        _atualizar_auditoria(
            auditoria, "rescisao_indenizacao", atual,
            status="CONFIRMADO", tipo_dado="DADO_DOCUMENTAL", arquivo=_texto(linha_atual.get("arquivo_fonte")) or arquivo,
            pagina=" e ".join(dict.fromkeys(x for x in (pag_atual, pagina_prazo) if x)),
            secao="; ".join(dict.fromkeys(x for x in (secao_atual, "Anexo II — condição comercial mínima") if x)),
            evidencia=" | ".join(dict.fromkeys(x for x in (evidencia_atual, evidencia_prazo) if x)),
            confianca=100,
        )


def _sincronizar_resumo_aditivos_auditoria(base: MutableMapping[str, Any], auditoria: List[Dict[str, Any]]) -> None:
    aditivos = base.get("aditivos_contrato") if isinstance(base.get("aditivos_contrato"), list) else []
    if aditivos:
        valor = f"{len(aditivos)} aditivo(s) identificado(s) com evidência documental."
        _atualizar_auditoria(
            auditoria, "resumo_aditivos", valor,
            status="CONFIRMADO", tipo_dado="CONSOLIDACAO_DOCUMENTAL", arquivo="Múltiplos documentos",
            pagina="Conforme aditivos", secao="Consolidação dos aditivos",
            evidencia=f"{len(aditivos)} termo(s) aditivo(s) validado(s) na matriz documental.", confianca=95,
        )
    else:
        valor = "Nenhum aditivo identificado com evidência documental no pacote analisado."
        _atualizar_auditoria(
            auditoria, "resumo_aditivos", valor,
            status="NÃO_APLICÁVEL", tipo_dado="INTERPRETACAO", arquivo="Pacote documental",
            pagina="Todas as páginas", secao="Triagem e consolidação dos anexos",
            evidencia="Nenhum arquivo classificado como termo aditivo foi identificado no pacote analisado.", confianca=100,
        )
    base["resumo_aditivos"] = valor


def _adicionar_pendencia_segunda_testemunha(base: MutableMapping[str, Any], texto_extraido: str) -> None:
    paginas = _paginas_do_texto_extraido(texto_extraido)
    pagina_ass = ""
    texto_ass = ""
    esperado = 0
    for numero, pagina in paginas.items():
        sem = _sem_acento(pagina).lower()
        m = re.search(r"presen[cç]a\s+de\s+(\d+)\s+testemunhas", pagina, flags=re.I)
        if m:
            esperado = int(m.group(1))
            pagina_ass, texto_ass = str(numero), pagina
            break
        if "testemunhas:" in sem:
            esperado = max(esperado, 2)
            pagina_ass, texto_ass = str(numero), pagina
    if esperado <= 0:
        return
    assinaturas = base.get("assinaturas_contrato") if isinstance(base.get("assinaturas_contrato"), list) else []
    testemunhas = [a for a in assinaturas if _token(a.get("categoria")) == "TESTEMUNHA"]
    if len(testemunhas) >= esperado:
        return
    pendencias = base.get("pendencias") if isinstance(base.get("pendencias"), list) else []
    if any("SEGUNDA_TESTEMUNHA" in _token(p.get("Pendência") or p.get("pendencia")) for p in pendencias if isinstance(p, Mapping)):
        return
    pendencias.append({
        "Pendência": "Segunda testemunha não identificada/assinada no instrumento",
        "Crítico": "Não",
        "Risco": "Baixo",
        "Recomendação": "Submeter ao Jurídico para confirmar se é necessário complementar a segunda testemunha para fins probatórios/executivos.",
        "Arquivo": next((_texto(a.get("fonte")) for a in assinaturas if _valor_util(a.get("fonte"))), "Contrato principal"),
        "Página": pagina_ass or "Bloco de assinaturas",
        "Evidência": "O instrumento declara assinatura na presença de 2 testemunhas, mas somente uma testemunha foi identificada no bloco de assinaturas.",
    })
    base["pendencias"] = pendencias


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



def _separar_nomes_assinatura(valor: Any) -> List[str]:
    """Separa nomes de signatários preservando a grafia retornada pela auditoria."""
    texto = _texto(valor)
    if not _valor_util(texto):
        return []
    texto = re.sub(r"\s*[|•·]+\s*", ";", texto)
    texto = re.sub(r"[\r\n]+", ";", texto)
    partes = [p.strip(" .,:-\t") for p in texto.split(";")]
    saida: List[str] = []
    vistos = set()
    invalidos = {
        "NAO IDENTIFICADO COM SEGURANCA", "NAO LOCALIZADO", "NAO VALIDADO",
        "CONTRATO ASSINADO", "SIM", "NAO",
    }
    for parte in partes:
        parte = re.sub(r"^(?:nome|assinante|signat[aá]rio)\s*:\s*", "", parte, flags=re.I)
        token = _token(parte)
        palavras = [x for x in re.findall(r"[A-Za-zÀ-ÿ]+", parte) if len(x) > 1]
        if token in invalidos or len(palavras) < 2 or token in vistos:
            continue
        vistos.add(token)
        saida.append(parte)
    return saida


def _span_aproximado_nome(texto: str, nome: str, alcance: int = 220) -> Tuple[int, int]:
    """Retorna o intervalo dos tokens do nome, mesmo com ordem invertida no OCR."""
    texto_norm = _sem_acento(texto).lower()
    nome_norm = _sem_acento(nome).lower()
    pos_exata = texto_norm.find(nome_norm)
    if pos_exata >= 0:
        return pos_exata, pos_exata + len(nome_norm)
    tokens = [
        x for x in re.findall(r"[a-z]+", nome_norm)
        if len(x) >= 3 and x not in {"dos", "das", "de", "da", "do"}
    ]
    if not tokens:
        return -1, -1
    posicoes: List[int] = []
    for token in tokens:
        pos = texto_norm.find(token)
        if pos >= 0:
            posicoes.append(pos)
    minimo = max(2, len(tokens) - 1)
    if len(posicoes) < minimo or max(posicoes) - min(posicoes) > alcance:
        return -1, -1
    return min(posicoes), max(posicoes) + max(len(x) for x in tokens)


def _contexto_aproximado_nome(texto: str, nome: str, raio: int = 100) -> Tuple[str, int]:
    """Localiza os tokens de um nome mesmo quando o OCR inverte sua ordem."""
    inicio_nome, fim_nome = _span_aproximado_nome(texto, nome)
    if inicio_nome < 0:
        return "", -1
    inicio = max(0, inicio_nome - raio)
    fim = min(len(texto), fim_nome + raio)
    return texto[inicio:fim], inicio_nome


def _recuperar_assinaturas_documentais(
    base: MutableMapping[str, Any],
    bruto: Mapping[str, Any],
    auditoria: List[Dict[str, Any]],
    paginas: Mapping[int, str],
    arquivo_padrao: str,
) -> None:
    """Recupera a tabela de assinaturas quando a IA trouxe os nomes na matriz,
    mas omitiu evidência/fonte nos registros individuais.

    A recuperação só é ativada quando existe bloco físico de assinaturas no OCR
    e uma lista auditada de pessoas ou registros individuais com nomes. Assim,
    uma testemunha isolada ou nomes citados apenas no corpo não validam o contrato.
    """
    pagina_ass = 0
    texto_ass = ""
    for numero, pagina in paginas.items():
        sem = _sem_acento(pagina).lower()
        if "firmam as partes" in sem and "testemunhas" in sem:
            pagina_ass, texto_ass = numero, pagina
            break
    if not texto_ass:
        return

    mapa = _mapa_auditoria(auditoria)
    linha_pessoas = mapa.get("pessoas_que_assinaram") or {}
    nomes: List[str] = []
    registros_origem: List[Mapping[str, Any]] = []
    for origem in (bruto.get("assinaturas_contrato"), base.get("assinaturas_contrato")):
        if isinstance(origem, list):
            for item in origem:
                if isinstance(item, Mapping):
                    registros_origem.append(item)
                    nome = _texto(item.get("nome") or item.get("Nome") or item.get("signatario") or item.get("assinante"))
                    if _valor_util(nome):
                        nomes.append(nome)

    if _evidencia_confirma(linha_pessoas):
        nomes.extend(_separar_nomes_assinatura(linha_pessoas.get("valor")))
    for origem in (base.get("pessoas_que_assinaram"), bruto.get("pessoas_que_assinaram")):
        nomes.extend(_separar_nomes_assinatura(origem))

    # Deduplicação preservando a ordem documental/auditada.
    nomes_unicos: List[str] = []
    vistos = set()
    for nome in nomes:
        chave = _token(nome)
        if not chave or chave in vistos:
            continue
        vistos.add(chave)
        nomes_unicos.append(nome)

    # Exige ao menos duas pessoas auditadas no bloco de assinatura das partes.
    if len(nomes_unicos) < 2:
        return

    origem_por_nome: Dict[str, Mapping[str, Any]] = {}
    for item in registros_origem:
        nome = _texto(item.get("nome") or item.get("Nome") or item.get("signatario") or item.get("assinante"))
        if _valor_util(nome):
            origem_por_nome[_token(nome)] = item

    data_instrumento = _data_br(base.get("data_contrato")) or _data_br(base.get("data_assinatura"))
    linha_rec = mapa.get("data_reconhecimento_firma") or {}
    data_rec_global = _data_br(linha_rec.get("valor") or base.get("data_reconhecimento_firma"))
    evidencia_pessoas = _texto(linha_pessoas.get("trecho_evidencia"))
    evidencia_rec = _texto(linha_rec.get("trecho_evidencia"))
    evidencia_nomes_norm = _sem_acento(evidencia_pessoas).lower()

    diretos_com_papel = 0
    contextos: Dict[str, Tuple[str, int, int]] = {}
    texto_ass_norm = _sem_acento(texto_ass).lower()
    for nome in nomes_unicos:
        inicio_nome, fim_nome = _span_aproximado_nome(texto_ass, nome)
        contexto, _ = _contexto_aproximado_nome(texto_ass, nome)
        contextos[_token(nome)] = (contexto, inicio_nome, fim_nome)
        if inicio_nome >= 0:
            proximo = texto_ass_norm[max(0, fim_nome - 5):min(len(texto_ass_norm), fim_nome + 45)]
            if any(x in proximo for x in ("coordenador", "coordenadora", "diretor", "diretora", "procurador", "representante")):
                diretos_com_papel += 1

    recuperadas: List[Dict[str, Any]] = []
    for indice, nome in enumerate(nomes_unicos):
        origem = dict(origem_por_nome.get(_token(nome)) or {})
        contexto, inicio_nome, fim_nome = contextos.get(_token(nome), ("", -1, -1))
        sem_contexto = _sem_acento(contexto).lower()
        antes = texto_ass_norm[max(0, inicio_nome - 120):inicio_nome] if inicio_nome >= 0 else ""
        depois = texto_ass_norm[fim_nome:min(len(texto_ass_norm), fim_nome + 70)] if fim_nome >= 0 else ""
        trecho_curto = texto_ass_norm[max(0, inicio_nome - 35):min(len(texto_ass_norm), fim_nome + 45)] if inicio_nome >= 0 else ""
        em_reconhecimento = bool(inicio_nome >= 0 and "firmas de" in antes and "reconhe" in texto_ass_norm[max(0, inicio_nome - 260):inicio_nome])
        em_campo_testemunha = bool(
            inicio_nome >= 0
            and "testemunhas" in antes
            and "nome" in antes[-70:]
            and not em_reconhecimento
        )

        papel = _texto(origem.get("papel_cargo") or origem.get("Papel/Cargo") or origem.get("cargo") or origem.get("papel"))
        categoria = _token(origem.get("categoria"))
        if not _valor_util(papel):
            if em_reconhecimento:
                papel = "Representante Legal"
            elif em_campo_testemunha:
                papel = "Testemunha"
            elif "coordenadora" in depois[:45]:
                papel = "Coordenadora"
            elif "coordenador" in depois[:45]:
                papel = "Coordenador"
            elif "diretora" in depois[:45]:
                papel = "Diretora"
            elif "diretor" in depois[:45]:
                papel = "Diretor"
            elif len(nomes_unicos) == 5 and diretos_com_papel >= 1 and indice == len(nomes_unicos) - 1 and "testemunhas" in texto_ass_norm:
                # Fallback restrito para OCR ilegível do único campo preenchido
                # de testemunha, mantendo os dois nomes reconhecidos em cartório
                # como representantes legais.
                papel = "Testemunha"
            else:
                papel = "Representante Legal"

        papel_tok = _token(papel)
        if not categoria:
            if "TESTEMUNHA" in papel_tok:
                categoria = "TESTEMUNHA"
            elif em_reconhecimento:
                categoria = "REPRESENTANTE_CONTRATANTE"
            elif inicio_nome >= 0 and "ITAU CORRETORA DE VALORES" in _sem_acento(texto_ass[fim_nome:]).upper():
                categoria = "REPRESENTANTE_CONTRAPARTE"
            else:
                categoria = "REPRESENTANTE_CONTRATANTE"

        fonte = _texto(origem.get("fonte") or origem.get("Fonte") or origem.get("arquivo_fonte")) or _texto(linha_pessoas.get("arquivo_fonte")) or arquivo_padrao
        pagina = _texto(origem.get("pagina") or origem.get("Página")) or _texto(linha_pessoas.get("pagina")) or str(pagina_ass)
        evidencia = _texto(origem.get("evidencia") or origem.get("Evidência") or origem.get("trecho_evidencia"))
        if not _evidencia_util(evidencia):
            evidencia = contexto.strip() if contexto else evidencia_pessoas
        if not _evidencia_util(evidencia):
            evidencia = f"Nome identificado no bloco de assinaturas físicas do instrumento: {nome}."

        data_individual = _texto(origem.get("data_assinatura") or origem.get("Data da assinatura"))
        if not _data_br(data_individual):
            data_individual = "Não localizada individualmente"

        data_rec = _data_br(origem.get("data_reconhecimento_firma") or origem.get("Data do reconhecimento de firma"))
        nome_norm = _sem_acento(nome).lower()
        m_firmas = re.search(r"firmas\s+de\s+(.*?)(?:\]|;|s[aã]o\s+paulo|$)", evidencia_nomes_norm, flags=re.I | re.S)
        trecho_firmas = m_firmas.group(1) if m_firmas else ""
        citado_em_firmas = bool(nome_norm and nome_norm in trecho_firmas)
        if not data_rec and data_rec_global and citado_em_firmas:
            data_rec = data_rec_global
        if not data_rec:
            data_rec = "Não aplicável"

        recuperadas.append({
            "nome": nome,
            "papel_cargo": papel,
            "categoria": categoria,
            "email": _texto(origem.get("email") or origem.get("e-mail")) or "Não localizado",
            "data_assinatura": data_individual,
            "data_instrumento": data_instrumento or "Não localizado",
            "data_reconhecimento_firma": data_rec,
            "fonte": fonte,
            "pagina": pagina,
            "status": "Assinado",
            "evidencia": evidencia,
        })

    base["assinaturas_contrato"] = recuperadas

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
    fontes_assinatura: List[Dict[str, Any]] = []
    for origem in (raw.get("assinaturas_contrato"), base.get("assinaturas_contrato")):
        if isinstance(origem, list):
            fontes_assinatura.extend(item for item in origem if isinstance(item, Mapping))
    assinaturas = _normalizar_assinaturas(fontes_assinatura)

    # Em assinatura física, a data geral do instrumento não deve ser apresentada
    # como se fosse um carimbo individual de cada signatário.
    data_instrumento = _data_br(base.get("data_contrato")) or _data_br(base.get("data_assinatura"))
    for assinatura in assinaturas:
        data_individual = _data_br(assinatura.get("data_assinatura"))
        evidencia_sem = _sem_acento(assinatura.get("evidencia")).lower()
        evidencia_tem_data = bool(data_individual and (data_individual in evidencia_sem or re.search(r"\d{1,2}\s+de\s+[a-z]+\s+de\s+\d{4}", evidencia_sem)))
        assinatura["data_instrumento"] = data_instrumento or "Não localizado"
        if data_individual and data_instrumento and data_individual == data_instrumento and not evidencia_tem_data:
            assinatura["data_assinatura"] = "Não localizada individualmente"
            assinatura["observacao_data"] = f"Data geral do instrumento: {data_instrumento}"

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
        elif data_instrumento:
            # Data principal do instrumento; não representa data individual de cada assinatura.
            base["data_assinatura"] = data_instrumento
        if reconhecimentos:
            base["data_reconhecimento_firma"] = "; ".join(dict.fromkeys(_data_br(x) for x in reconhecimentos if _data_br(x)))
        partes = ["Contrato assinado"]
        if representantes:
            partes.append(f"{len(representantes)} representante(s) das partes")
        if testemunhas:
            partes.append(f"{len(testemunhas)} testemunha(s)")
        if _data_br(base.get("data_assinatura")):
            partes.append(f"data do instrumento: {_data_br(base.get('data_assinatura'))}")
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
            nomes_evidencia = "; ".join(a.get("nome", "") for a in assinaturas[:6] if _valor_util(a.get("nome")))
            evidencia = f"Bloco de assinaturas físicas: {nomes_evidencia}." if nomes_evidencia else _texto(linha.get("trecho_evidencia"))
            if campo == "data_assinatura" and data_instrumento:
                evidencia = f"Data geral escrita no bloco de assinaturas do instrumento: {data_instrumento}."
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

    # Se o checklist possuir atenção parcial sem uma pendência equivalente, o
    # indicador não pode continuar zerado. O bloco específico de campos não
    # localizados permanece separado para não contar a mesma ausência duas vezes.
    checklist = base.get("checklist") if isinstance(base.get("checklist"), list) else []
    atencoes_checklist = [
        c for c in checklist
        if "ATEN" in _token(c.get("Status"))
        and _token(c.get("Validação")) not in {"CAMPOS_NAO_LOCALIZADOS", "CONFLITOS_DOCUMENTAIS"}
    ]
    pontos_count = max(len(pontos), len(atencoes_checklist))
    metricas_tabela = base.get("metricas_tabela_comercial") if isinstance(base.get("metricas_tabela_comercial"), Mapping) else {}
    try:
        divergencia_itens = int(metricas_tabela.get("divergencia_quantidade") or 0)
    except Exception:
        divergencia_itens = 0
    if divergencia_itens:
        pontos_count = max(pontos_count, 1)

    base["indicadores_pendencias"] = {
        "pendencias_criticas": len(criticas),
        "pontos_atencao": pontos_count,
        "campos_nao_localizados": len(campos),
        "divergencia_itens_comerciais": divergencia_itens,
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
    metricas_tabela = base.get("metricas_tabela_comercial") if isinstance(base.get("metricas_tabela_comercial"), Mapping) else {}
    try:
        divergencia_itens = abs(int(metricas_tabela.get("divergencia_quantidade") or 0))
    except Exception:
        divergencia_itens = 0
    try:
        cobertura_tabela = float(metricas_tabela.get("cobertura_tabela_percentual") or 100)
    except Exception:
        cobertura_tabela = 100.0
    penalidade_tabela = min(30.0, divergencia_itens * 5.0 + max(0.0, 100.0 - cobertura_tabela) * 0.30)
    confianca = max(0, min(100, round(cobertura * 0.85 + paginas_pct * 0.15 - inferidos - conflitos * 2 - penalidade_tabela)))

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

    # Score de completude e confiança técnica são indicadores diferentes.
    completos_simples = 0.0
    total_simples = len(CAMPOS_OFICIAIS_V4)
    for _, campo in CAMPOS_OFICIAIS_V4:
        item = mapa.get(campo)
        st = _status(item.get("status")) if item else "NÃO_LOCALIZADO"
        if _evidencia_confirma(item) or (item and st == "CALCULADO" and _valor_util(item.get("valor"))) or st == "NÃO_APLICÁVEL":
            completos_simples += 1
        elif st == "INFERIDO":
            completos_simples += 0.5
    score_completude = round((completos_simples / total_simples) * 100) if total_simples else 0
    if divergencia_itens:
        score_completude = max(0, score_completude - min(20, divergencia_itens * 5))

    base["confianca_extracao"] = confianca
    base["score"] = score_completude
    base["score_completude"] = score_completude
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
        "divergencia_itens_comerciais": divergencia_itens,
        "cobertura_tabela_comercial_percentual": cobertura_tabela,
        "penalidade_tabela_comercial": round(penalidade_tabela, 2),
        "pendencias_com_evidencia": len(pendencias),
        "confianca_extracao_percentual": confianca,
        "score_final": score_completude,
        "score_completude_percentual": score_completude,
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

    _recuperar_campos_documentais(base, bruto, auditoria, texto_extraido)
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
        elif item and _status(item.get("status")) == "NÃO_LOCALIZADO":
            base[campo] = _texto(item.get("valor")) or "Não localizado com segurança"
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
    _sincronizar_resumo_aditivos_auditoria(base, auditoria)

    base["pendencias"] = _filtrar_pendencias(bruto.get("pendencias") or base.get("pendencias"))
    _adicionar_pendencia_segunda_testemunha(base, texto_extraido)

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
