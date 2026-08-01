"""Extração determinística e completa de tabelas comerciais.

O módulo complementa a resposta da IA. Ele não substitui a auditoria jurídica:
apenas garante que linhas de preços, faixas, isenções e tarifas textuais não sejam
resumidas ou descartadas quando o OCR/texto do documento as contém.
"""
from __future__ import annotations

import re
import unicodedata
from typing import Any, Dict, Iterable, List, Mapping, Sequence, Tuple

NAO_LOCALIZADO = "Não localizado"
NAO_APLICAVEL = "Não aplicável"


def _texto(valor: Any) -> str:
    return re.sub(r"\s+", " ", str(valor or "")).strip()


def _sem_acento(valor: Any) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", _texto(valor))
        if unicodedata.category(c) != "Mn"
    )


def _token(valor: Any) -> str:
    return re.sub(r"[^A-Z0-9]+", " ", _sem_acento(valor).upper()).strip()


def _valor_util(valor: Any) -> bool:
    return _token(valor) not in {"", "NAO LOCALIZADO", "NAO APLICAVEL", "N A", "NONE", "NAN"}


def _pagina_blocos(texto: str) -> List[Tuple[str, str, str]]:
    """Retorna (arquivo, pagina, conteúdo) preservando blocos de OCR/texto."""
    arquivo_atual = "Contrato/anexo"
    pagina_atual = "Não localizado"
    blocos: List[Tuple[str, str, str]] = []
    acumulado: List[str] = []

    def fechar() -> None:
        nonlocal acumulado
        conteudo = "\n".join(acumulado).strip()
        if conteudo:
            blocos.append((arquivo_atual, pagina_atual, conteudo))
        acumulado = []

    for linha in str(texto or "").splitlines():
        m_arq = re.match(r"\s*ARQUIVO\s*:\s*(.+?)\s*$", linha, flags=re.I)
        if m_arq:
            fechar()
            arquivo_atual = _texto(m_arq.group(1)) or "Contrato/anexo"
            pagina_atual = "Não localizado"
            continue
        m_pag = re.match(r"\s*---\s*P[AÁ]GINA\s+(\d+)(?:\s+(?:OCR|TABELA\s+\d+))?\s*---\s*$", linha, flags=re.I)
        if m_pag:
            fechar()
            pagina_atual = m_pag.group(1)
            continue
        acumulado.append(linha)
    fechar()
    if not blocos and str(texto or "").strip():
        blocos.append((arquivo_atual, pagina_atual, str(texto)))
    return blocos


def _normalizar_valor_ocr(valor: str) -> str:
    """Corrige formatos frequentes do OCR: R$075 -> R$ 0,75; R$2,55 -> R$ 2,55."""
    txt = _texto(valor).replace("RS", "R$").replace("R S", "R$")
    txt = re.sub(r"R\$\s*", "R$ ", txt, flags=re.I)
    corpo = re.sub(r"(?i)R\$\s*", "", txt).strip(" |;.")
    corpo = corpo.replace(" ", "")
    # OCR de centavos sem vírgula, comum em tabelas escaneadas: 075, 070, 040.
    if re.fullmatch(r"0?\d{2}", corpo):
        corpo = "0," + corpo[-2:]
    elif re.fullmatch(r"\d{3}", corpo) and corpo.startswith("0"):
        corpo = "0," + corpo[-2:]
    elif re.fullmatch(r"\d+", corpo):
        # Inteiros monetários explícitos permanecem inteiros; completa centavos.
        corpo = corpo + ",00"
    elif re.fullmatch(r"\d+[.,]\d{1}", corpo):
        inteiro, dec = re.split(r"[.,]", corpo)
        corpo = f"{inteiro},{dec}0"
    else:
        corpo = corpo.replace(".", "#").replace(",", ".").replace("#", ",") if False else corpo
        # Mantém separador brasileiro; converte ponto decimal simples para vírgula.
        if re.fullmatch(r"\d+\.\d{2}", corpo):
            corpo = corpo.replace(".", ",")
    return f"R$ {corpo}"


def _extrair_valor_final(linha: str) -> Tuple[str, str] | None:
    """Separa descrição e condição/valor, tolerando pequenos ruídos de OCR.

    Tabelas escaneadas frequentemente acrescentam caracteres soltos depois do
    valor (por exemplo ``R$0,60 5`` ou ``R$ 0,65 |``). Esses caracteres não
    fazem parte do preço e não podem provocar a perda da linha comercial.
    """
    linha = _texto(linha).strip("| ")

    # Condições textuais também são linhas comerciais válidas.
    m_txt = re.match(
        r"^(.*?)(?:\s+)(isento|isenta|taxa\s+correio|tarifa\s+oficial\s+dos\s+correios)\s*$",
        linha,
        flags=re.I,
    )
    if m_txt and len(_texto(m_txt.group(1))) >= 3:
        return _texto(m_txt.group(1)).rstrip(" *"), _texto(m_txt.group(2))

    # Aceita até três resíduos curtos após o valor. Não aceita palavras, para
    # evitar interpretar como preço uma linha narrativa comum do contrato.
    m = re.match(
        r"^(.*?)(?:\s+)(R\s*\$|RS)\s*([0-9][0-9.,]*)(?:\s+(?:[|+\-=]+|[0-9]{1,2})){0,3}\s*$",
        linha,
        flags=re.I,
    )
    if m and len(_texto(m.group(1))) >= 3:
        return _texto(m.group(1)).rstrip(" *"), _normalizar_valor_ocr("R$ " + m.group(3))
    return None


def _natureza(descricao: str, valor: str, grupo: str) -> Tuple[str, str, str, str]:
    low = _sem_acento(f"{descricao} {grupo}").lower()
    vlow = _sem_acento(valor).lower()
    if "isento" in vlow or "isenta" in vlow:
        return "ISENTO", "R$ 0,00", "Conforme ocorrência", "Isento conforme tabela comercial"
    if "taxa correio" in vlow or "tarifa oficial" in vlow:
        return "REEMBOLSO", "Tarifa oficial dos Correios", "Por envio", "Conforme tarifa oficial vigente"
    if "implantacao" in low:
        return "IMPLANTACAO_UNICA", valor, "Única", "Cobrança única"
    if "mensal" in low and "acionista" not in low:
        return "MENSAL_FIXO", valor, "Mensal", "Mensalidade fixa"
    if "acionista" in low and ("ate " in low or "acima" in low or re.search(r"\bde\s+[\d.]", low)):
        return "UNITARIO_VARIAVEL", valor, "Mensal", "Tarifa mensal por acionista conforme faixa"
    if any(x in low for x in ("dividendo", "bonificacao", "desdobramento", "movimentacao", "transferencia", "subscricao", "boletim", "aviso", "extrato", "informe", "correspondencia", "voto a distancia")):
        return "UNITARIO_VARIAVEL", valor, "Por evento", "Cobrança conforme serviço efetivamente utilizado"
    return "UNITARIO_VARIAVEL", valor, "Conforme utilização", "Conforme tabela comercial"


def _unidade(descricao: str, natureza: str) -> str:
    low = _sem_acento(descricao).lower()
    if natureza == "IMPLANTACAO_UNICA":
        return "Implantação"
    if natureza == "MENSAL_FIXO":
        return "Mês"
    if "acionista" in low and ("ate " in low or "acima" in low or re.search(r"\bde\s+[\d.]", low)):
        return "Acionista/mês"
    if "correspondencia" in low:
        return "Correspondência"
    if "voto a distancia" in low:
        return "Serviço"
    return "Operação/evento"


def _grupo_por_linha(linha: str, grupo_atual: str) -> str:
    token = _token(linha)
    if "TAXA MENSAL POR ACIONISTA" in token:
        return "Taxa mensal por acionista - faixas"
    if token == "EVENTOS E MOVIMENTACOES":
        return "Eventos e movimentações"
    if token.startswith("SUBSCRICAO"):
        return "Subscrição"
    if token.startswith("CUSTO FIXO"):
        return "Custos fixos"
    return grupo_atual


def _eh_cabecalho_ignorado(linha: str) -> bool:
    token = _token(linha)
    ignorar = (
        "VALORES EM REAIS", "ANEXO II REMUNERACAO", "PELA PRESTACAO DOS SERVICOS",
        "SOLUCOES PARA O", "MERCADO DE CAPITAIS", "CONTRATO DE PRESTACAO DE SERVICOS",
        "FORMALIZACAO CONTRATOS", "EMISSOR SOMENTE INICIARA", "PARA INFORMES DE RENDIMENTOS",
    )
    return not token or any(x in token for x in ignorar)


def extrair_tabela_comercial_completa(texto: str) -> List[Dict[str, Any]]:
    """Extrai todas as linhas comerciais encontradas nas páginas de preços/remuneração.

    Não limita quantidade e preserva linhas com valor textual (isento/taxa Correio).
    """
    itens: List[Dict[str, Any]] = []
    vistos = set()
    tabela_ativa = False
    arquivo_anterior = ""

    for arquivo, pagina, conteudo in _pagina_blocos(texto):
        if arquivo != arquivo_anterior:
            tabela_ativa = False
            arquivo_anterior = arquivo
        low = _sem_acento(conteudo).lower()
        linhas_bloco = [_texto(l) for l in conteudo.splitlines() if _texto(l)]
        candidatos = sum(1 for l in linhas_bloco if _extrair_valor_final(l))
        gatilho_forte = (
            ("anexo" in low and "remuneracao" in low)
            or "valores em reais" in low
            or "taxa mensal por acionista" in low
            or "eventos e movimentacoes" in low
            or "tabela comercial" in low
            or "precos e tarifas" in low
        )
        gatilho = gatilho_forte or (tabela_ativa and candidatos >= 2)
        tabela_ativa = gatilho
        if not gatilho:
            continue

        grupo = "Tabela comercial"
        prefixo_pendente = ""
        linhas = linhas_bloco
        for linha in linhas:
            grupo_novo = _grupo_por_linha(linha, grupo)
            if grupo_novo != grupo:
                grupo = grupo_novo
                if _token(linha) in {"CUSTO FIXOS", "CUSTOS FIXOS"}:
                    prefixo_pendente = "Custo Fixo"
                continue
            if _eh_cabecalho_ignorado(linha):
                continue
            if _token(linha) in {"CUSTO FIXOS", "CUSTOS FIXOS"}:
                grupo = "Custos fixos"
                prefixo_pendente = "Custo Fixo"
                continue
            if _token(linha) in {"EVENTOS E MOVIMENTACOES", "SUBSCRICAO", "TAXA MENSAL POR ACIONISTA FAIXA"}:
                continue

            separado = _extrair_valor_final(linha)
            if not separado:
                continue
            descricao, valor_original = separado
            if prefixo_pendente and _token(descricao) == "MENSAL":
                descricao = f"{prefixo_pendente} Mensal"
                prefixo_pendente = ""
            descricao = re.sub(r"(?i)transfer[eê]ncialaltera[cç][aã]o", "Transferência/alteração", descricao)
            descricao = re.sub(r"\s+", " ", descricao).strip(" -:;|")
            if len(descricao) < 3:
                continue

            natureza, valor, periodicidade, condicao = _natureza(descricao, valor_original, grupo)
            faixa = descricao if grupo.startswith("Taxa mensal por acionista") else NAO_APLICAVEL
            unidade = _unidade(descricao, natureza)
            evidencia = f"{descricao} - {valor_original}"
            chave = (_token(descricao), _token(valor), _token(grupo))
            if chave in vistos:
                continue
            vistos.add(chave)

            itens.append({
                "item": str(len(itens) + 1),
                "descricao": descricao,
                "tipo": "Serviço",
                "grupo_tabela": grupo,
                "natureza_valor": natureza,
                "quantidade": "1" if natureza in {"IMPLANTACAO_UNICA", "MENSAL_FIXO"} else NAO_APLICAVEL,
                "unidade": unidade,
                "valor_unitario": valor,
                "valor_total": valor if natureza == "IMPLANTACAO_UNICA" else NAO_LOCALIZADO,
                "periodicidade": periodicidade,
                "faixa_condicao": faixa,
                "condicao_comercial": condicao,
                "taxa_percentual": NAO_APLICAVEL,
                "total_encargos": NAO_APLICAVEL,
                "fonte": arquivo,
                "pagina": pagina,
                "status_evidencia": "CONFIRMADO",
                "evidencia": evidencia,
            })
    return itens


def _campo_item(item: Mapping[str, Any], *nomes: str) -> Any:
    """Lê um campo independentemente de caixa, acento ou estilo de chave."""
    mapa = {
        re.sub(r"[^A-Z0-9]+", "_", _sem_acento(k).upper()).strip("_"): v
        for k, v in item.items()
    }
    for nome in nomes:
        chave = re.sub(r"[^A-Z0-9]+", "_", _sem_acento(nome).upper()).strip("_")
        if chave in mapa:
            return mapa[chave]
    return None


def _numero_identidade(valor: Any) -> str:
    """Normaliza 3000.0 e R$ 3.000,00 para a mesma identidade."""
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        numero = float(valor)
        return (f"{numero:.6f}").rstrip("0").rstrip(".")
    txt = _texto(valor)
    if not txt:
        return ""
    token = _token(txt)
    if "CORREIO" in token:
        return "TARIFA_CORREIOS"
    if token in {"ISENTO", "ISENTA"}:
        return "0"
    m = re.search(r"-?[0-9][0-9.]*([,][0-9]+|[.][0-9]+)?", txt)
    if not m:
        return token
    numero_txt = m.group(0)
    if "," in numero_txt:
        numero_txt = numero_txt.replace(".", "").replace(",", ".")
    elif numero_txt.count(".") > 1:
        numero_txt = numero_txt.replace(".", "")
    try:
        numero = float(numero_txt)
        return (f"{numero:.6f}").rstrip("0").rstrip(".")
    except ValueError:
        return token


def _numeros_descricao(descricao: str) -> List[int]:
    valores: List[int] = []
    for numero in re.findall(r"\d[\d.]*", descricao):
        try:
            valores.append(int(numero.replace(".", "")))
        except ValueError:
            continue
    return valores


def _assinatura_semantica(item: Mapping[str, Any]) -> Tuple[str, ...]:
    """Cria uma identidade comercial estável para reconciliar IA e OCR.

    A IA costuma ampliar a descrição (``Tarifa mensal por acionista...``),
    enquanto o OCR preserva apenas a célula da tabela (``Até 5.000 acionistas``).
    Esta assinatura reconhece que ambos representam a mesma linha.
    """
    descricao = _texto(_campo_item(item, "Descrição", "descricao", "Item", "item"))
    low = _sem_acento(descricao).lower()
    natureza = _token(_campo_item(item, "Natureza do valor", "natureza_valor"))
    valor = _numero_identidade(
        _campo_item(item, "Valor unitário", "valor_unitario", "Taxa / Percentual", "taxa_percentual")
    )

    if "implantacao" in low or "taxa unica" in low or "setup" in low:
        return ("IMPLANTACAO", valor)
    if "voto" in low and "distancia" in low:
        return ("VOTO_DISTANCIA",)
    if "mensalidade" in low or "custo fixo mensal" in low or (natureza == "MENSAL FIXO" and "acionista" not in low):
        return ("MENSAL_FIXO", valor)
    if "movimentacao" in low and "bolsa" in low:
        return ("MOVIMENTACAO_BOLSA",)

    if "acionista" in low and ("ate" in low or "acima" in low or "faixa" in low or re.search(r"\bde\s+\d", low)):
        numeros = _numeros_descricao(low)
        if "acima" in low and numeros:
            return ("FAIXA_ACIONISTA", "ACIMA", str(numeros[0]))
        if "ate" in low and numeros:
            return ("FAIXA_ACIONISTA", "ATE", str(numeros[-1]))
        if len(numeros) >= 2:
            return ("FAIXA_ACIONISTA", "DE", str(numeros[-2]), str(numeros[-1]))

    if "dividendo" in low:
        if "outros bancos" in low or "outro banco" in low:
            return ("DIVIDENDOS_OUTROS_BANCOS",)
        if "itau" in low:
            return ("DIVIDENDOS_ITAU",)
    if "bonificacao" in low or "desdobramento" in low:
        return ("BONIFICACAO_DESDOBRAMENTO",)
    if "transferencia" in low and ("cadastral" in low or "movimentacao" in low):
        return ("TRANSFERENCIA_CADASTRAL",)
    # A IA pode devolver a mesma linha em forma verbal ou nominal:
    # "boletim emitido" == "emissão de boletim" e
    # "boletim efetivado" == "efetivação de boletim".
    # A assinatura semântica precisa reconhecer ambos para não duplicar a tabela.
    if "boletim" in low and (re.search(r"\bemit", low) or re.search(r"\bemiss", low)):
        return ("BOLETIM_EMITIDO",)
    if "boletim" in low and re.search(r"\befetiv", low):
        return ("BOLETIM_EFETIVADO",)
    if ("aviso" in low or "avisos" in low) and ("extrato" in low or "extratos" in low):
        return ("AVISOS_EXTRATOS",)
    if "informe" in low and "rendimento" in low and "digita" in low:
        return ("INFORMES_RENDIMENTOS_DIGITAL",)
    if "correspondencia" in low:
        return ("CORRESPONDENCIA",)

    return ("LITERAL", _token(descricao), valor, natureza)


def _chave_campo(campo: Any) -> str:
    return re.sub(r"[^A-Z0-9]+", "_", _sem_acento(campo).upper()).strip("_")


def _mesclar_campos_sem_sobrescrever_alias(destino: Dict[str, Any], origem: Mapping[str, Any]) -> None:
    """Preenche lacunas sem criar aliases que substituam o dado documental."""
    existentes = {_chave_campo(k): k for k in destino}
    for campo, valor in origem.items():
        chave = _chave_campo(campo)
        if chave in existentes:
            # A primeira lista é prioritária. Mesmo "Não localizado" pode ser
            # uma decisão documental intencional (ex.: valor total de uma
            # mensalidade não deve receber o valor unitário devolvido pela IA).
            continue
        if _valor_util(valor):
            destino[campo] = valor
            existentes[chave] = campo


def _chave_item(item: Mapping[str, Any]) -> Tuple[str, ...]:
    return _assinatura_semantica(item)


def mesclar_itens_comerciais(*listas: Any) -> List[Dict[str, Any]]:
    """Reconcilia parser documental + IA sem duplicar a tabela comercial.

    A ordem importa: a primeira lista é a fonte prioritária. No fluxo principal,
    o parser determinístico deve vir primeiro e a IA serve para completar lacunas.
    """
    saida: List[Dict[str, Any]] = []
    pos: Dict[Tuple[str, ...], int] = {}
    for lista in listas:
        if not isinstance(lista, list):
            continue
        for item in lista:
            if not isinstance(item, Mapping):
                continue
            chave = _chave_item(item)
            if not chave or not chave[0]:
                continue
            if chave in pos:
                _mesclar_campos_sem_sobrescrever_alias(saida[pos[chave]], item)
                continue
            pos[chave] = len(saida)
            saida.append(dict(item))

    # Renumeração estável para tela/Excel.
    for idx, item in enumerate(saida, 1):
        campo_item = next((k for k in item if _chave_campo(k) in {"ITEM", "NUMERO", "N"}), None)
        if campo_item:
            item[campo_item] = str(idx)
        else:
            item["item"] = str(idx)
    return saida

def calcular_metricas_tabela_comercial(itens_documento: Any, itens_exibidos: Any) -> Dict[str, Any]:
    doc = [i for i in (itens_documento or []) if isinstance(i, Mapping)]
    exib = [i for i in (itens_exibidos or []) if isinstance(i, Mapping)]
    ch_doc = {_chave_item(i) for i in doc if _chave_item(i)[0]}
    ch_exib = {_chave_item(i) for i in exib if _chave_item(i)[0]}
    encontrados = len(ch_doc)
    cobertos = len(ch_doc & ch_exib)
    cobertura = round((cobertos / encontrados) * 100) if encontrados else (100 if exib else 0)
    paginas = sorted({_texto(i.get("Página") or i.get("pagina")) for i in doc if _valor_util(i.get("Página") or i.get("pagina"))})
    grupos = []
    for i in doc:
        g = _texto(i.get("Grupo/Tabela") or i.get("grupo_tabela"))
        if _valor_util(g) and g not in grupos:
            grupos.append(g)
    return {
        "itens_encontrados_documento": encontrados,
        "itens_exibidos_auditor": len(exib),
        "itens_exibidos_unicos": len(ch_exib),
        "divergencia_quantidade": len(exib) - encontrados,
        "itens_documentais_cobertos": cobertos,
        "cobertura_tabela_percentual": cobertura,
        "paginas_tabela_comercial": ", ".join(paginas) if paginas else "Não localizado",
        "grupos_tabela_comercial": grupos,
    }
