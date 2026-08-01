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
    """Separa descrição e condição/valor no final da linha."""
    linha = _texto(linha).strip("| ")
    # Condições textuais também são linhas comerciais válidas.
    m_txt = re.match(r"^(.*?)(?:\s+)(isento|isenta|taxa\s+correio|tarifa\s+oficial\s+dos\s+correios)\s*$", linha, flags=re.I)
    if m_txt and len(_texto(m_txt.group(1))) >= 3:
        return _texto(m_txt.group(1)), _texto(m_txt.group(2))

    m = re.match(r"^(.*?)(?:\s+)(R\s*\$|RS)\s*([0-9][0-9.,]*)(?:\s*[|+\-=]+)*\s*$", linha, flags=re.I)
    if m and len(_texto(m.group(1))) >= 3:
        return _texto(m.group(1)), _normalizar_valor_ocr("R$ " + m.group(3))
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


def _chave_item(item: Mapping[str, Any]) -> Tuple[str, str, str]:
    desc = item.get("Descrição") or item.get("descricao") or item.get("Item") or item.get("item")
    valor = item.get("Valor unitário") or item.get("valor_unitario") or item.get("Taxa / Percentual") or item.get("taxa_percentual")
    natureza = item.get("Natureza do valor") or item.get("natureza_valor") or ""
    # Grupo/Tabela é metadado de exibição, não parte da identidade. Assim a linha
    # resumida pela IA é substituída/complementada pela linha documentada local.
    return _token(desc), _token(valor), _token(natureza)


def mesclar_itens_comerciais(*listas: Any) -> List[Dict[str, Any]]:
    """Une IA + parser local, sem limite e sem perder linhas textuais."""
    saida: List[Dict[str, Any]] = []
    pos: Dict[Tuple[str, str, str], int] = {}
    for lista in listas:
        if not isinstance(lista, list):
            continue
        for item in lista:
            if not isinstance(item, Mapping):
                continue
            chave = _chave_item(item)
            if not chave[0]:
                continue
            if chave in pos:
                atual = saida[pos[chave]]
                # O item com evidência/página mais completa prevalece, sem apagar dados úteis.
                for k, v in item.items():
                    if _valor_util(v) and not _valor_util(atual.get(k)):
                        atual[k] = v
                continue
            pos[chave] = len(saida)
            saida.append(dict(item))
    # Renumeração estável para tela/Excel.
    for idx, item in enumerate(saida, 1):
        if "Item" in item:
            item["Item"] = str(idx)
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
        "itens_exibidos_auditor": len(ch_exib),
        "itens_documentais_cobertos": cobertos,
        "cobertura_tabela_percentual": cobertura,
        "paginas_tabela_comercial": ", ".join(paginas) if paginas else "Não localizado",
        "grupos_tabela_comercial": grupos,
    }
