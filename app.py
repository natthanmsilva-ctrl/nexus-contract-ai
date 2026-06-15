# =========================================================
# NEXUS CONTRACT AI - VERSÃO REVISADA
# Ajustes aplicados: visual profissional, carregamento seguro
# do histórico, proteção de HTML no Assistente IA e refinamentos
# de usabilidade sem alterar o fluxo principal.
# =========================================================

import io
import os
import re
import json
import html
import base64
from datetime import datetime
from typing import Any, Dict, List

import pandas as pd
import sqlite3
import streamlit as st
import pdfplumber
import pytesseract
from pdf2image import convert_from_bytes
from docx import Document
from dotenv import load_dotenv

from database import criar_banco, salvar_analise, listar_analises, limpar_historico

from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


# =========================================================
# CONFIGURAÇÕES INICIAIS
# =========================================================
load_dotenv()
criar_banco()

TESSERACT_CMD = os.getenv("TESSERACT_CMD", r"C:\Users\124034\AppData\Local\Programs\Tesseract-OCR\tesseract.exe")
if os.path.exists(TESSERACT_CMD):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

st.set_page_config(
    page_title="Auditor DE Contratos",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Campos oficiais exibidos nos cards e no Excel
CAMPOS_OFICIAIS = [
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
    ("Rescisão e Indenização", "rescisao_indenizacao"),
    ("Anticorrupção", "anticorrupcao"),
    ("Proteção de Dados LGPD", "protecao_dados_lgpd"),
    ("Data da Assinatura", "data_assinatura"),
    ("Valor do Contrato Original", "valor_contrato_original"),
]

CAMPOS_JSON_OBRIGATORIOS = ", ".join([campo for _, campo in CAMPOS_OFICIAIS] + [
    "contraparte", "fornecedor",
    "contrato_assinado", "alerta_assinatura", "status", "risco", "score",
    "resumo_executivo", "parecer", "checklist", "pendencias"
])


# =========================================================
# ESTILO VISUAL
# =========================================================
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap');

:root{
    --bg:#070b10;
    --panel:#101822;
    --panel-2:#0b1118;
    --green:#003c2f;
    --green-2:#065f46;
    --gold:#d7bf75;
    --gold-soft:#f3e6b3;
    --text:#f8fafc;
    --muted:#94a3b8;
    --line:rgba(215,191,117,.22);
    --danger:#dc2626;
    --warn:#f59e0b;
    --ok:#16a34a;
}

html, body, [class*="css"]{
    font-family:'Inter', sans-serif;
}

.stApp{
    background:
        radial-gradient(circle at top left, rgba(0,60,47,.32), transparent 34%),
        radial-gradient(circle at top right, rgba(215,191,117,.18), transparent 32%),
        var(--bg);
    color:var(--text);
}

.block-container{
    padding-top:1.8rem;
    padding-bottom:2rem;
}

section[data-testid="stSidebar"]{
    background:linear-gradient(180deg,#202631,#111821);
    border-right:1px solid var(--line);
}

section[data-testid="stSidebar"] *{
    color:#f8fafc;
}

.stRadio label, .stFileUploader label{
    font-weight:800 !important;
}

.hero{
    padding:34px 40px;
    border-radius:28px;
    background:
        linear-gradient(115deg, rgba(0,60,47,.98), rgba(10,88,65,.96) 52%, rgba(215,191,117,.95));
    margin:0 0 26px;
    box-shadow:0 22px 55px rgba(0,0,0,.38);
    border:1px solid rgba(255,255,255,.10);
}

.hero .eyebrow{
    display:inline-block;
    padding:8px 13px;
    border-radius:999px;
    background:rgba(255,255,255,.13);
    color:#fff;
    font-size:12px;
    font-weight:900;
    letter-spacing:.9px;
    text-transform:uppercase;
    margin-bottom:14px;
}

.hero h1{
    margin:0;
    color:#fff;
    font-size:44px;
    line-height:1.05;
    font-weight:900;
    letter-spacing:.4px;
}

.hero p{
    margin:14px 0 0;
    color:rgba(255,255,255,.92);
    font-size:17px;
    font-weight:600;
    max-width:980px;
}

.section-title{
    color:var(--gold);
    font-size:22px;
    font-weight:900;
    margin:28px 0 14px;
    letter-spacing:.2px;
}

.subtle{
    color:var(--muted);
    font-size:13px;
    line-height:1.6;
}

.metric-card{
    background:linear-gradient(145deg,var(--panel),var(--panel-2));
    border:1px solid var(--line);
    border-radius:22px;
    padding:22px;
    min-height:122px;
    box-shadow:0 16px 38px rgba(0,0,0,.35);
}

.metric-card small{
    color:var(--gold);
    font-weight:900;
    text-transform:uppercase;
    letter-spacing:.9px;
    font-size:11px;
}

.metric-card h2{
    color:#fff;
    margin:13px 0 0;
    font-size:34px;
    line-height:1.05;
    font-weight:900;
    word-break:break-word;
}

.metric-link{
    display:block;
    text-decoration:none !important;
    color:inherit !important;
}

.metric-card.clickable{
    cursor:pointer;
    transition:all .18s ease;
    position:relative;
    overflow:hidden;
}

.metric-card.clickable:hover{
    transform:translateY(-3px);
    border-color:rgba(215,191,117,.70);
    box-shadow:0 20px 48px rgba(0,0,0,.42);
}

.metric-card.clickable.active{
    border:2px solid var(--gold);
    background:linear-gradient(145deg,rgba(0,60,47,.98),rgba(11,17,24,.98));
}

.metric-card.clickable.active::after{
    content:"Filtro ativo";
    position:absolute;
    top:14px;
    right:14px;
    background:rgba(215,191,117,.18);
    color:var(--gold);
    border:1px solid rgba(215,191,117,.45);
    border-radius:999px;
    padding:5px 9px;
    font-size:10px;
    font-weight:900;
    text-transform:uppercase;
    letter-spacing:.5px;
}

.filter-note{
    margin:16px 0 4px;
    padding:12px 16px;
    border-radius:14px;
    background:rgba(0,60,47,.42);
    border:1px solid rgba(215,191,117,.25);
    color:#f8fafc;
    font-size:13px;
    font-weight:700;
}

.executive-box{
    background:linear-gradient(135deg, rgba(0,60,47,.96), rgba(17,24,33,.96));
    border:1px solid rgba(215,191,117,.35);
    border-radius:20px;
    padding:22px;
    color:#fff;
    margin:12px 0;
    box-shadow:0 14px 35px rgba(0,0,0,.25);
}

.executive-box h3{
    margin:0 0 8px;
    color:var(--gold);
    font-weight:900;
}

.info-grid{
    display:grid;
    grid-template-columns:repeat(4, minmax(0, 1fr));
    gap:16px;
}

.info-card{
    background:linear-gradient(145deg,var(--panel),var(--panel-2));
    border:1px solid rgba(215,191,117,.22);
    border-radius:18px;
    padding:18px;
    min-height:112px;
    box-shadow:0 10px 28px rgba(0,0,0,.24);
}

.info-card small{
    display:block;
    color:var(--gold);
    font-weight:900;
    text-transform:uppercase;
    letter-spacing:.7px;
    font-size:11px;
}

.info-card p{
    color:#fff;
    font-size:15px;
    font-weight:700;
    margin:10px 0 0;
    line-height:1.45;
    overflow-wrap:anywhere;
}

.flow-grid{
    display:grid;
    grid-template-columns:repeat(5, minmax(0, 1fr));
    gap:14px;
    margin-top:14px;
}

.flow-card{
    background:linear-gradient(145deg,var(--panel),var(--panel-2));
    border:1px solid rgba(215,191,117,.24);
    border-radius:20px;
    padding:18px;
    min-height:150px;
}

.flow-number{
    width:36px;
    height:36px;
    border-radius:999px;
    background:var(--gold);
    color:var(--green);
    display:flex;
    align-items:center;
    justify-content:center;
    font-weight:900;
    margin-bottom:12px;
}

.flow-card h4{
    color:#fff;
    margin:0;
    font-size:16px;
    font-weight:900;
}

.flow-card p{
    color:var(--muted);
    font-size:13px;
    margin:8px 0 0;
    line-height:1.45;
}

.contract-card{
    background:linear-gradient(145deg,var(--panel),var(--panel-2));
    border:1px solid rgba(215,191,117,.18);
    border-radius:22px;
    padding:22px;
    margin:16px 0;
    box-shadow:0 16px 38px rgba(0,0,0,.32);
}

.contract-top{
    display:flex;
    justify-content:space-between;
    align-items:flex-start;
    gap:16px;
}

.contract-title{
    color:var(--gold);
    font-size:21px;
    font-weight:900;
    margin:0;
}

.contract-file{
    color:#93c5fd;
    font-size:13px;
    margin-top:6px;
    overflow-wrap:anywhere;
}

.contract-grid{
    display:grid;
    grid-template-columns:2fr 1fr 1.5fr .7fr;
    gap:20px;
    margin-top:18px;
}

.contract-label{
    color:var(--muted);
    font-size:11px;
    font-weight:900;
    text-transform:uppercase;
    letter-spacing:.7px;
}

.contract-value{
    color:#fff;
    font-size:14px;
    font-weight:800;
    margin-top:7px;
    line-height:1.5;
    overflow-wrap:anywhere;
}

.badge-risk{
    color:white;
    padding:9px 16px;
    border-radius:999px;
    font-weight:900;
    font-size:13px;
    white-space:nowrap;
}

.card-footer{
    display:flex;
    justify-content:space-between;
    align-items:center;
    gap:14px;
    margin-top:20px;
    padding-top:16px;
    border-top:1px solid rgba(255,255,255,.08);
}

.excel-link{
    background:var(--green);
    color:white !important;
    padding:9px 14px;
    border-radius:10px;
    font-size:12px;
    font-weight:900;
    text-decoration:none !important;
    border:1px solid rgba(215,191,117,.60);
}

.risk-row, .ok-row{
    padding:15px 16px;
    border-radius:14px;
    margin:10px 0;
    color:#fff;
    font-weight:700;
}

.risk-row{
    background:rgba(245,158,11,.14);
    border-left:5px solid var(--warn);
}

.ok-row{
    background:rgba(22,163,74,.14);
    border-left:5px solid var(--ok);
}

.pill{
    display:block;
    padding:14px 16px;
    border-radius:16px;
    font-weight:800;
    line-height:1.5;
}

.pill-ok{background:rgba(22,163,74,.14);color:#bbf7d0;border:1px solid rgba(22,163,74,.30);}
.pill-warn{background:rgba(245,158,11,.14);color:#fde68a;border:1px solid rgba(245,158,11,.30);}
.pill-danger{background:rgba(220,38,38,.14);color:#fecaca;border:1px solid rgba(220,38,38,.30);}

.stButton>button,
.stDownloadButton>button{
    border-radius:13px !important;
    border:1px solid var(--gold) !important;
    background:var(--green) !important;
    color:#fff !important;
    font-weight:900 !important;
    min-height:43px;
}

.stButton>button:hover,
.stDownloadButton>button:hover{
    border-color:#fff !important;
    color:var(--gold) !important;
}

[data-testid="stDataFrame"]{
    border-radius:16px;
    overflow:hidden;
    border:1px solid rgba(215,191,117,.18);
}

.stTextInput input,
.stTextArea textarea,
.stSelectbox div[data-baseweb="select"] > div,
.stMultiSelect div[data-baseweb="select"] > div{
    border-radius:13px !important;
    border:1px solid rgba(215,191,117,.24) !important;
    background:#0b1118 !important;
    color:#f8fafc !important;
}

.stFileUploader{
    background:linear-gradient(145deg,rgba(16,24,34,.82),rgba(11,17,24,.82));
    border:1px dashed rgba(215,191,117,.45);
    border-radius:20px;
    padding:14px;
}

button[kind="secondary"]{
    transition:all .18s ease !important;
}

button[kind="secondary"]:hover{
    transform:translateY(-1px);
}

[data-testid="stTabs"] button{
    font-weight:900 !important;
}

.footer{
    color:var(--muted);
    text-align:center;
    margin-top:32px;
    padding:20px;
    border-top:1px solid rgba(255,255,255,.10);
}

@media(max-width:1100px){
    .info-grid{grid-template-columns:repeat(2, minmax(0, 1fr));}
    .flow-grid{grid-template-columns:repeat(2, minmax(0, 1fr));}
    .contract-grid{grid-template-columns:1fr 1fr;}
}

@media(max-width:700px){
    .hero{padding:26px 24px;}
    .hero h1{font-size:32px;}
    .info-grid,.flow-grid,.contract-grid{grid-template-columns:1fr;}
    .contract-top,.card-footer{flex-direction:column;align-items:flex-start;}
}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# =========================================================
# FUNÇÕES AUXILIARES
# =========================================================
def safe(value: Any) -> str:
    """Escapa textos para exibição segura em HTML."""
    if value is None:
        return "Não localizado"
    value = str(value).strip()
    return html.escape(value if value else "Não localizado")


def clean_text(value: Any) -> str:
    """Remove tags HTML e entidades comuns."""
    if value is None:
        return "Não localizado"
    if not isinstance(value, str):
        return value
    value = re.sub(r"<[^>]+>", "", value)
    value = html.unescape(value)
    value = re.sub(r"\s+", " ", value).strip()
    return value or "Não localizado"


def normalize_risco(risco: Any) -> str:
    risco = str(risco or "N/A").upper().strip()
    if risco == "MEDIO":
        return "MÉDIO"
    return risco


def risco_cor(risco: Any) -> str:
    risco = normalize_risco(risco)
    if risco == "ALTO":
        return "#dc2626"
    if risco in ["MÉDIO", "MEDIO"]:
        return "#f59e0b"
    if risco == "BAIXO":
        return "#16a34a"
    return "#64748b"


def as_float_score(value: Any) -> float:
    try:
        return float(str(value).replace(",", "."))
    except Exception:
        return 0.0

def carregar_contratos_chat():
    try:
        df = listar_analises()
        return df if isinstance(df, pd.DataFrame) else pd.DataFrame()
    except Exception:
        return pd.DataFrame()


def carregar_historico_seguro() -> pd.DataFrame:
    """Carrega o histórico sem derrubar a tela caso o banco esteja vazio ou inacessível."""
    try:
        df = listar_analises()
        return df if isinstance(df, pd.DataFrame) else pd.DataFrame()
    except Exception as erro:
        st.error(f"Não foi possível carregar o histórico. Detalhe: {erro}")
        return pd.DataFrame()


# =========================================================
# LEITURA DOS ARQUIVOS
# =========================================================
def ler_pdf(file) -> str:
    texto = ""

    try:
        file.seek(0)
        with pdfplumber.open(file) as pdf:
            for i, page in enumerate(pdf.pages, 1):
                pagina = page.extract_text() or ""
                if pagina.strip():
                    texto += f"\n\n--- PÁGINA {i} ---\n{pagina}"
    except Exception:
        texto = ""

    if len(texto.strip()) > 300:
        return texto

    # OCR automático para PDFs escaneados
    try:
        file.seek(0)
        imagens = convert_from_bytes(file.read(), dpi=300)
        texto_ocr = ""

        for i, imagem in enumerate(imagens, 1):
            pagina = pytesseract.image_to_string(imagem, lang="por")
            texto_ocr += f"\n\n--- PÁGINA {i} OCR ---\n{pagina}"

        return texto_ocr
    except Exception as e:
        return f"Não foi possível extrair texto do PDF. Detalhe: {e}"


def ler_docx(file) -> str:
    file.seek(0)
    doc = Document(file)
    partes = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    for table in doc.tables:
        for row in table.rows:
            linha = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
            if linha:
                partes.append(linha)

    return "\n".join(partes)


# =========================================================
# ANÁLISE LOCAL
# =========================================================
def local_extract(texto: str) -> Dict[str, Any]:
    texto = texto or ""
    low = texto.lower()

    cnpjs = re.findall(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto)
    valores = re.findall(r"R\$\s?\d{1,3}(?:\.\d{3})*,\d{2}", texto)
    datas = re.findall(r"\b\d{2}/\d{2}/\d{4}\b", texto)

    fornecedor = "Não localizado"
    fornecedor_patterns = [
        r"contratada[:\s]+([^\n]{5,140})",
        r"fornecedor[:\s]+([^\n]{5,140})",
        r"licenciante[:\s]+([^\n]{5,140})",
        r"licenciada[:\s]+([^\n]{5,140})",
        r"([A-ZÁÉÍÓÚÂÊÔÃÕÇ0-9\s\.\-&]+S/A)",
        r"([A-ZÁÉÍÓÚÂÊÔÃÕÇ0-9\s\.\-&]+LTDA\.?)",
    ]

    for pattern in fornecedor_patterns:
        match = re.search(pattern, texto, re.IGNORECASE)
        if match:
            fornecedor = re.sub(r"\s+", " ", match.group(1)).strip(" -:;,.|/")[:180]
            break

    objeto = "Não localizado"
    objeto_patterns = [
        r"CLÁUSULA PRIMEIRA[:\s\-]+DO OBJETO DO CONTRATO(.{50,1200})",
        r"CLAUSULA PRIMEIRA[:\s\-]+DO OBJETO DO CONTRATO(.{50,1200})",
        r"OBJETO DO CONTRATO(.{50,1200})",
        r"objeto(?:\s+do\s+contrato)?[:\s]+(.{20,900})",
        r"escopo[:\s]+(.{20,900})",
    ]

    for pattern in objeto_patterns:
        match = re.search(pattern, texto, re.IGNORECASE | re.DOTALL)
        if match:
            objeto = re.sub(r"\s+", " ", match.group(1)).strip(" -:;,.|/")[:700]
            break

    valor_total = max(valores, key=len) if valores else "Não localizado"

    vigencia = "Não localizada"
    vigencia_match = re.search(r"(vig[eê]ncia|prazo)[^.\n]{0,180}", texto, re.IGNORECASE)
    if vigencia_match:
        vigencia = clean_text(vigencia_match.group(0))[:220]
    if "prazo de duração indeterminado" in low or "prazo indeterminado" in low:
        vigencia = "Prazo indeterminado"

    condicao_pagamento = "Não localizada"
    pagamento_patterns = [
        r"prazo de\s*(\d{1,3})\s*\(?\w*\)?\s*dias",
        r"em até\s*(\d{1,3})\s*\(?\w*\)?\s*dias",
        r"(\d{1,3})\s*\(?\w*\)?\s*dias",
        r"(\d{1,3})\s*dd",
        r"(\d{1,3})\s*d\.?d\.?”?",
    ]

    for pattern in pagamento_patterns:
        match = re.search(pattern, low, re.IGNORECASE)
        if match:
            condicao_pagamento = f"{match.group(1)}DD"
            break

    if condicao_pagamento == "Não localizada" and any(t in low for t in ["pagamento", "nota fiscal", "fatura", "mensalidade", "vencimento"]):
        condicao_pagamento = "Cláusula de pagamento localizada"

    multa = "Não localizada"
    match = re.search(r"multa[^.\n]{0,140}?(\d{1,3})\s*%", low, re.IGNORECASE)
    if match:
        multa = f"{match.group(1)}%"
    elif any(t in low for t in ["multa", "penalidade", "penalidades"]):
        multa = "Cláusula localizada"

    reajuste = "Não localizado"
    if "igp-m" in low or "igpm" in low:
        reajuste = "IGP-M"
    elif "ipca" in low:
        reajuste = "IPCA"
    elif "fgv" in low:
        reajuste = "FGV"
    elif "reajuste" in low:
        reajuste = "Cláusula localizada"

    rescisao = "Localizada" if any(t in low for t in ["rescisão", "rescisao", "resilir", "rescindir"]) else "Não localizada"
    sla = "Localizado" if any(t in low for t in ["sla", "nível de serviço", "nivel de serviço", "tempo de resposta", "disponibilidade"]) else "Não localizado"
    confidencialidade = "Localizada" if any(t in low for t in ["confidencialidade", "sigilo"]) else "Não localizada"
    lgpd = "Localizada" if any(t in low for t in ["lgpd", "proteção de dados", "dados pessoais"]) else "Não localizada"
    anticorrupcao = "Localizada" if any(t in low for t in ["anticorrupção", "anticorrupcao", "lei anticorrupção", "lei 12.846"]) else "Não localizada"

    assinatura_termos = [
        "assinado", "assinatura", "testemunhas", "docusign", "certificado digital",
        "partes assinam", "por estarem justas e contratadas"
    ]
    contrato_assinado = "Sim" if any(t in low for t in assinatura_termos) else "Não"
    alerta_assinatura = "Evidência de assinatura localizada" if contrato_assinado == "Sim" else "Contrato sem evidência de assinatura localizada"

    validacoes = [
        ("Contrato assinado", contrato_assinado == "Sim", 30, "Sim"),
        ("CNPJ do fornecedor", bool(cnpjs), 20, "Sim"),
        ("Fornecedor / Contratada", fornecedor != "Não localizado", 20, "Sim"),
        ("Objeto do contrato", objeto != "Não localizado", 20, "Sim"),
        ("Valor contratual", valor_total != "Não localizado", 20, "Sim"),
        ("Vigência", vigencia != "Não localizada", 25, "Sim"),
        ("Condição de pagamento", condicao_pagamento != "Não localizada", 15, "Sim"),
        ("Multa / penalidade", multa != "Não localizada", 10, "Não"),
        ("Reajuste", reajuste != "Não localizado", 10, "Não"),
        ("Rescisão", rescisao == "Localizada", 15, "Sim"),
        ("SLA / nível de serviço", sla == "Localizado", 15, "Não"),
        ("Confidencialidade", confidencialidade == "Localizada", 10, "Não"),
        ("LGPD / proteção de dados", lgpd == "Localizada", 15, "Não"),
        ("Anticorrupção", anticorrupcao == "Localizada", 15, "Não"),
    ]

    checklist: List[Dict[str, Any]] = []
    pendencias: List[Dict[str, Any]] = []
    score = 0

    for item, ok, peso, critico in validacoes:
        if not ok:
            score += peso
            pendencias.append({
                "Pendência": item,
                "Crítico": critico,
                "Risco": peso,
                "Recomendação": "Validar cláusula/informação antes de seguir com RC/PO."
            })

        checklist.append({
            "Validação": item,
            "Status": "OK" if ok else "Pendente",
            "Peso de risco": 0 if ok else peso,
            "Crítico": critico,
            "Evidência": "Termo localizado" if ok else "Não localizado no texto extraído",
        })

    score = min(score, 100)
    risco = "BAIXO" if score <= 25 else "MÉDIO" if score <= 60 else "ALTO"
    status = "Aprovado" if score <= 25 else "Aprovado com ressalvas" if score <= 60 else "Contrato pendente de revisão"

    return {
        "tipo_contrato": "Não localizado",
        "empresa_grupo_sbf": "SBF COMÉRCIO DE PRODUTOS ESPORTIVOS LTDA." if "sbf comércio" in low or "sbf comercio" in low else "Não localizado",
        "cnpj_empresa_grupo": "06.347.409/0001-65" if "06.347.409/0001-65" in texto else "Não localizado",
        "local_prestacao": "Não localizado",
        "contraparte": fornecedor,
        "cnpj_contraparte": cnpjs[0] if cnpjs else "Não localizado",
        "objetivo": objeto,
        "descricao_servico_material": objeto,
        "descricao_breve_cadastro": objeto[:120] if objeto != "Não localizado" else "Não localizado",
        "forma_pagamento": condicao_pagamento,
        "condicao_pagamento_dias": condicao_pagamento,
        "multa": multa,
        "vigencia_apos_assinatura": vigencia,
        "rescisao_indenizacao": rescisao,
        "anticorrupcao": anticorrupcao,
        "indice_reajuste_anual": reajuste,
        "protecao_dados_lgpd": "Há cláusula de proteção de dados pessoais em conformidade com a LGPD." if lgpd == "Localizada" else "Não localizado",
        "data_assinatura": datas[-1] if datas else "Não localizada",
        "unidade_organizacional": "Não localizado",
        "area_solicitante": "Não localizado",
        "cadastrado_por": "Não localizado",
        "fim_contrato": "Não localizado",
        "prazo_contrato_dias": "Não localizado",
        "data_assinatura_aditivo": "Não localizado",
        "valor_contrato_original": valor_total,
        "aditivo_escopo": "Não localizado",
        "aditivo_quitacao": "Não localizado",
        "aditivo_valor": "Não localizado",
        "aditivo_prazo": "Não localizado",
        "quantidade_renovacoes": "Não localizado",
        "data_atualizacao": datetime.now().strftime("%d/%m/%Y"),
        "contrato_assinado": contrato_assinado,
        "alerta_assinatura": alerta_assinatura,
        "fornecedor": fornecedor,
        "cnpj": cnpjs[0] if cnpjs else "Não localizado",
        "cnpjs_encontrados": ", ".join(cnpjs[:10]) or "Não localizado",
        "valor_total": valor_total,
        "valores_encontrados": ", ".join(valores[:15]) or "Não localizado",
        "datas_encontradas": ", ".join(datas[:15]) or "Não localizado",
        "objeto": objeto,
        "vigencia": vigencia,
        "condicao_pagamento": condicao_pagamento,
        "multa": multa,
        "reajuste": reajuste,
        "rescisao": rescisao,
        "sla": sla,
        "lgpd": lgpd,
        "confidencialidade": confidencialidade,
        "assinatura": "Localizada" if contrato_assinado == "Sim" else "Não localizada",
        "score": score,
        "risco": risco,
        "status": status,
        "resumo_executivo": f"Análise concluída. Status: {status}. Risco: {risco}. Score: {score}. Pendências: {len(pendencias)}.",
        "parecer": "Recomenda-se revisar as pendências antes de seguir com RC/PO." if pendencias else "Contrato aparentemente possui os itens essenciais para continuidade.",
        "checklist": checklist,
        "pendencias": pendencias,
    }


# =========================================================
# IA
# =========================================================
def prompt_ia(texto: str) -> str:
    return f"""
Você é um analista sênior de contratos da área de Suprimentos/Jurídico.
Sua tarefa é consolidar CONTRATO PRINCIPAL + ANEXOS + ORÇAMENTOS enviados no mesmo texto.
Retorne APENAS JSON válido, sem markdown, sem comentários e sem texto fora do JSON.

OBJETIVO PRINCIPAL
Extrair os dados comerciais e jurídicos reais do contrato, priorizando o documento principal e usando anexos/orçamentos apenas quando o contrato fizer referência a eles.

CAMPOS OBRIGATÓRIOS
Retorne exatamente estas chaves principais e internas:
{CAMPOS_JSON_OBRIGATORIOS}

Os campos principais que serão exibidos ao usuário são exatamente:
- Tipo de Contrato = tipo_contrato
- Empresa do Grupo SBF = empresa_grupo_sbf
- CNPJ Empresa do Grupo = cnpj_empresa_grupo
- Local de Prestação Contraparte = local_prestacao
- Contraparte = contraparte
- CNPJ Contraparte = cnpj_contraparte
- Objetivo = objetivo
- Descrição do Serviço/ Material = descricao_servico_material
- Descrição breve do cadastro = descricao_breve_cadastro
- Forma de pagamento = forma_pagamento
- Condição de Pagamento em Dias = condicao_pagamento_dias
- Multa = multa
- Vigência após a data de assinatura = vigencia_apos_assinatura
- Rescisão e Indenização = rescisao_indenizacao
- Anticorrupção = anticorrupcao
- Proteção de Dados LGPD = protecao_dados_lgpd
- Data da Assinatura = data_assinatura
- Valor do Contrato Original = valor_contrato_original

REGRAS DE EXTRAÇÃO OBRIGATÓRIAS
1. Não invente dados. Se não encontrar, retorne "Não localizado".
2. Nunca confunda CNPJ com número sem máscara. Sempre formate CNPJ como 00.000.000/0000-00 quando houver 14 dígitos.
3. contraparte deve ser somente a empresa contratada/fornecedor. Não inclua CNPJ no mesmo campo.
4. local_prestacao deve ser cidade/UF/local de execução, se existir. Não misture com contraparte.
5. objetivo deve responder "para quê é o contrato" em uma frase curta.
6. descricao_servico_material deve descrever o serviço/material contratado, não histórico societário, alteração cadastral ou qualificação da empresa.
7. descricao_breve_cadastro deve ser curta e própria para cadastro de material/serviço. Exemplo: "Fornecimento e instalação de baterias para nobreaks".
8. condicao_pagamento_dias deve retornar em formato claro, exemplo: "90 dias". Não retorne apenas número solto.
9. forma_pagamento deve informar a forma descrita no contrato, exemplo: "Mediante emissão de nota fiscal".
10. valor_contrato_original deve ser o VALOR TOTAL DO CONTRATO/ORÇAMENTO REFERENCIADO, não valor unitário, não parcela e não subtotal parcial. Se houver quantitativo x valor unitário, calcule o total. Se existir orçamento anexo prevalecente citado pelo contrato, use o total do orçamento.
11. Se houver mais de um valor monetário, priorize nesta ordem:
    a) valor total do contrato;
    b) valor total do orçamento/proposta anexa citada no contrato;
    c) soma/calculo por quantidade x valor unitário;
    d) maior valor total claramente vinculado ao objeto.
12. vigencia_apos_assinatura deve manter a redação objetiva do contrato, exemplo: "Prazo indeterminado".
13. data_assinatura deve ser a data da assinatura do contrato, não data de orçamento, não data de proposta, não data de alteração cadastral.
14. contrato_assinado deve retornar "Sim" ou "Não".
15. Se não encontrar assinatura, alerta_assinatura deve retornar "Contrato sem evidência de assinatura localizada".
16. anticorrupcao e protecao_dados_lgpd devem resumir a cláusula específica quando houver, explicando a obrigação principal.
17. risco deve ser BAIXO, MÉDIO ou ALTO.
18. score deve ser número inteiro de 0 a 100, onde maior é melhor.
19. checklist deve ser lista de objetos com exatamente: Validação, Status, Peso de risco, Crítico, Evidência.
20. pendencias deve ser lista de objetos com exatamente: Pendência, Crítico, Risco, Recomendação.
21. Os campos exibidos ao usuário devem vir em RESUMO EXECUTIVO: nem curtos demais, nem cópia integral da cláusula.
22. Para cada campo principal, retorne de 1 a 3 linhas objetivas, com informação suficiente para entendimento corporativo.
23. Nunca retorne apenas: "Localizado", "Localizada", "Sim", "Possui cláusula" ou "Cláusula localizada" nos campos principais. Sempre explique resumidamente o conteúdo encontrado.
24. Só use "Não localizado" quando realmente não houver informação no contrato/anexo.
25. Anticorrupção e Proteção de Dados LGPD devem dizer que há cláusula e resumir a obrigação/finalidade.
26. Multa deve informar percentual, juros, base de cálculo e situação de aplicação quando constar.
27. Rescisão e Indenização deve resumir aviso prévio, multa/ônus, hipóteses de rescisão e indenização quando constar.
28. Vigência deve explicar prazo e início/encerramento de forma objetiva.
29. Forma de pagamento deve explicar o gatilho do pagamento, por exemplo: emissão/aprovação de nota fiscal, aceite, medição ou conclusão do serviço.
30. Descrição breve do cadastro deve ser uma frase própria para cadastro de serviço/material, sem copiar histórico societário ou qualificação jurídica.
31. valor_contrato_original deve estar formatado em reais, exemplo: "R$ 43.468,16".

PADRÃO ESPERADO PARA ESTE TIPO DE CONTRATO
Se o texto indicar contrato de prestação de serviços de substituição/fornecimento de baterias de nobreaks:
- objetivo deve ser algo como "Substituição das baterias dos nobreaks da unidade indicada".
- descricao_servico_material deve citar fornecimento, instalação, testes, recolhimento/descarte das baterias, se constar.
- descrição breve de cadastro não deve falar de transformação societária/EIRELI; deve falar do serviço/material.

EXEMPLOS DE NÍVEL DE DETALHE ESPERADO
- objetivo: "Contratação para substituição das baterias dos nobreaks instalados na unidade da Centauro."
- descricao_servico_material: "Fornecimento, instalação e testes de baterias para nobreaks, incluindo recolhimento das baterias antigas quando previsto."
- forma_pagamento: "Pagamento mediante emissão de nota fiscal após conclusão/aceite dos serviços pela contratante."
- condicao_pagamento_dias: "90 dias após aprovação da nota fiscal."
- multa: "Multa de 2% sobre valores em atraso, acrescida de juros de 1% ao mês, se previsto."
- rescisao_indenizacao: "Rescisão mediante aviso prévio de 30 dias, sem ônus ou multa, conforme previsto no contrato."
- anticorrupcao: "Há cláusula de anticorrupção exigindo conduta ética e conformidade com a Lei 12.846/13."
- protecao_dados_lgpd: "Há cláusula de proteção de dados pessoais, com obrigações de segurança, confidencialidade e conformidade com a LGPD."

CONTRATO E ANEXOS EXTRAÍDOS:
{texto[:120000]}
"""
MODELOS_GEMINI = {
    "Automático recomendado": [
        "gemini-3.5-thinking",
        "gemini-3.1-pro",
        "gemini-3.5-flash",
    ],
    "Gemini 3.5 Thinking": ["gemini-3.5-thinking"],
    "Gemini 3.1 Pro": ["gemini-3.1-pro"],
    "Gemini 3.5 Flash": ["gemini-3.5-flash"],
}


def analisar_gemini(texto: str, api_key: str, opcao_modelo: str) -> Dict[str, Any]:
    import google.generativeai as genai

    genai.configure(api_key=api_key)

    modelos = MODELOS_GEMINI.get(opcao_modelo, MODELOS_GEMINI["Automático recomendado"])

    ultimo_erro = None

    for nome in modelos:
        try:
            model = genai.GenerativeModel(nome)

            resp = model.generate_content(
                prompt_ia(texto),
                generation_config={
                    "temperature": 0.0,
                    "top_p": 0.2,
                    "response_mime_type": "application/json",
                },
            )

            content = (
                resp.text
                .strip()
                .replace("```json", "")
                .replace("```", "")
                .strip()
            )

            st.success(f"IA utilizada: {nome}")
            resultado_json = json.loads(content)
            if isinstance(resultado_json, dict):
                resultado_json["modelo_ia"] = nome
            return resultado_json

        except Exception as e:
            ultimo_erro = e
            if opcao_modelo != "Automático recomendado":
                raise Exception(f"Erro ao usar o modelo {nome}. Detalhe: {e}")
            continue

    raise Exception(f"Nenhum modelo Gemini disponível. Detalhe: {ultimo_erro}")


def formatar_cnpj(valor: Any) -> str:
    txt = clean_text(valor)
    if not txt or txt == "Não localizado":
        return "Não localizado"
    m = re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", txt)
    if m:
        return m.group(0)
    digits = re.sub(r"\D", "", txt)
    if len(digits) == 14:
        return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"
    return txt


def resumir_campo(valor: Any, limite: int = 220) -> str:
    txt = clean_text(valor)
    if len(txt) <= limite:
        return txt
    corte = txt[:limite].rsplit(" ", 1)[0]
    return corte + "..."


def padronizar_resultado_ia(base: Dict[str, Any]) -> Dict[str, Any]:
    """Ajustes finais para evitar cards quebrados e dados mal formatados."""
    base["cnpj_empresa_grupo"] = formatar_cnpj(base.get("cnpj_empresa_grupo"))
    base["cnpj_contraparte"] = formatar_cnpj(base.get("cnpj_contraparte"))
    base["cnpj"] = formatar_cnpj(base.get("cnpj") or base.get("cnpj_contraparte"))

    # Se a IA devolver somente número no pagamento, transforma em "X dias".
    pag = clean_text(base.get("condicao_pagamento_dias"))
    if re.fullmatch(r"\d{1,3}", pag):
        base["condicao_pagamento_dias"] = f"{pag} dias"

    # Mantém os cards com resumo executivo: detalhado o suficiente, sem virar cláusula inteira.
    limites = {
        "objetivo": 260,
        "descricao_servico_material": 360,
        "descricao_breve_cadastro": 180,
        "forma_pagamento": 240,
        "condicao_pagamento_dias": 160,
        "multa": 260,
        "vigencia_apos_assinatura": 280,
        "rescisao_indenizacao": 340,
        "anticorrupcao": 300,
        "protecao_dados_lgpd": 320,
        "alerta_assinatura": 220,
    }
    for campo, limite in limites.items():
        base[campo] = resumir_campo(base.get(campo), limite)

    # Corrige erro comum: descrição de cadastro com histórico societário em vez do serviço.
    desc_cad = clean_text(base.get("descricao_breve_cadastro"))
    desc_serv = clean_text(base.get("descricao_servico_material"))
    if any(t in desc_cad.lower() for t in ["transformação", "transformacao", "eireli", "societária", "societaria"]):
        base["descricao_breve_cadastro"] = resumir_campo(desc_serv, 120)

    return base

def normalizar(resultado: Dict[str, Any]) -> Dict[str, Any]:
    base = local_extract("")
    base.update(resultado or {})

    if clean_text(base.get("contraparte")) in ("", "Não localizado", "Não localizada", "N/A", "None"):
        base["contraparte"] = base.get("fornecedor", "Não localizado")
    if clean_text(base.get("fornecedor")) in ("", "Não localizado", "Não localizada", "N/A", "None"):
        base["fornecedor"] = base.get("contraparte", "Não localizado")

    # Compatibilidade com nomes que a IA pode devolver em português ou com variações.
    alias = {
        "local_de_prestacao_contraparte": "local_prestacao",
        "local_prestacao_contraparte": "local_prestacao",
        "cnpj_empresa_do_grupo": "cnpj_empresa_grupo",
        "empresa_do_grupo_sbf": "empresa_grupo_sbf",
        "descricao_do_servico_material": "descricao_servico_material",
        "descricao_servico_material": "descricao_servico_material",
        "vigencia_apos_a_data_de_assinatura": "vigencia_apos_assinatura",
        "valor_do_contrato_original": "valor_contrato_original",
    }
    def _vazio(valor: Any) -> bool:
        txt = clean_text(valor)
        return txt in ("", "Não localizado", "Não localizada", "N/A", "None")

    for origem, destino in alias.items():
        if origem in base and _vazio(base.get(destino)):
            base[destino] = base.get(origem)

    for chave, valor in list(base.items()):
        if isinstance(valor, str):
            base[chave] = clean_text(valor)

    if not isinstance(base.get("checklist"), list):
        base["checklist"] = []
    if not isinstance(base.get("pendencias"), list):
        base["pendencias"] = []

    for lista in ["checklist", "pendencias"]:
        for item in base.get(lista, []):
            if isinstance(item, dict):
                for k, v in list(item.items()):
                    if isinstance(v, str):
                        item[k] = clean_text(v)

    base["risco"] = normalize_risco(base.get("risco"))
    base["score"] = int(min(max(as_float_score(base.get("score")), 0), 100))
    base = padronizar_resultado_ia(base)
    return base




# =========================================================
# EXCEL - RELATÓRIO EXECUTIVO 100% PROFISSIONAL
# =========================================================
def excel_clean(value: Any, padrao: str = "Não localizado") -> str:
    """Texto seguro para Excel, sem HTML, sem caracteres inválidos e sem risco de fórmula."""
    value = clean_text(value)
    if value in (None, "", "Não localizado", "Não localizada", "None", "N/A"):
        value = padrao
    value = str(value).strip()
    value = re.sub(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    if value and value[0] in ("=", "+", "-", "@", "\t", "\r", "\n"):
        value = "'" + value
    return value[:32700]


def excel_risk_fill(risco: Any) -> str:
    risco_norm = normalize_risco(risco)
    if risco_norm == "ALTO":
        return "DC2626"
    if risco_norm in ["MÉDIO", "MEDIO"]:
        return "F59E0B"
    if risco_norm == "BAIXO":
        return "16A34A"
    return "64748B"


def _wb_styles():
    # Paleta visual do Excel
    # Fundo geral: verde claro | Cabeçalhos: verde escuro | Títulos: dourado
    return {
        "green": "004D40",
        "green2": "00695C",
        "gold": "D4AF37",
        "pale_gold": "F6E7A8",
        "white": "FFFFFF",
        "dark": "111827",
        "muted": "4B5563",
        "background": "EAF6F1",
        "light": "EEF5F2",
        "line": "8FA39A",
        "soft_green": "DCFCE7",
        "soft_warn": "FEF3C7",
        "soft_danger": "FEE2E2",
    }


def _thin(color: str = "8FA39A") -> Border:
    return Border(
        left=Side(style="thin", color=color),
        right=Side(style="thin", color=color),
        top=Side(style="thin", color=color),
        bottom=Side(style="thin", color=color),
    )


def _sheet_base(ws, title: str, subtitle: str, last_col: int = 8, zoom: int = 95) -> None:
    s = _wb_styles()
    end_col = get_column_letter(last_col)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = zoom
    ws.freeze_panes = None
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.35
    ws.page_margins.bottom = 0.35

    widths = {"A": 22, "B": 24, "C": 20, "D": 20, "E": 20, "F": 20, "G": 20, "H": 20}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    # Fundo geral verde claro para tirar o visual cinza/branco padrão do Excel.
    for row_cells in ws.iter_rows(min_row=1, max_row=120, min_col=1, max_col=26):
        for cell in row_cells:
            cell.fill = PatternFill("solid", fgColor=s["background"])

    ws.merge_cells(f"A1:{end_col}1")
    ws["A1"] = title
    ws["A1"].font = Font(name="Calibri", size=26, bold=True, color=s["white"])
    ws["A1"].fill = PatternFill("solid", fgColor=s["green"])
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 38

    ws.merge_cells(f"A2:{end_col}2")
    ws["A2"] = subtitle
    ws["A2"].font = Font(name="Calibri", size=11, bold=True, color=s["green"])
    ws["A2"].fill = PatternFill("solid", fgColor=s["gold"])
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22

    ws.merge_cells(f"A3:{end_col}3")
    ws["A3"] = ""
    ws["A3"].fill = PatternFill("solid", fgColor=s["background"])
    ws.row_dimensions[3].height = 14


def _merge(ws, cell_range: str, value: Any, *, fill: str = "FFFFFF", font_color: str = "111827", size: int = 10, bold: bool = False, align: str = "left", valign: str = "center") -> None:
    ws.merge_cells(cell_range)
    cell = ws[cell_range.split(":")[0]]
    cell.value = excel_clean(value, "")
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.font = Font(name="Calibri", size=size, bold=bold, color=font_color)
    cell.alignment = Alignment(horizontal=align, vertical=valign, wrap_text=True)
    cell.border = _thin()


def _section(ws, row: int, text: str, last_col: int = 8) -> int:
    s = _wb_styles()
    _merge(ws, f"A{row}:{get_column_letter(last_col)}{row}", text, fill=s["pale_gold"], font_color=s["green"], size=13, bold=True, align="left")
    ws.row_dimensions[row].height = 24
    return row + 2


def _metric_card(ws, cell_range: str, label: str, value: Any, fill: str) -> None:
    s = _wb_styles()
    _merge(ws, cell_range, f"{label}\n{excel_clean(value, '')}", fill=fill, font_color=s["white"], size=14, bold=True, align="center")


def _table_header(ws, row: int, headers: List[str], col_spans: List[int] | None = None) -> None:
    s = _wb_styles()
    if col_spans is None:
        col_spans = [1] * len(headers)
    col = 1
    for header, span in zip(headers, col_spans):
        start = get_column_letter(col)
        end = get_column_letter(col + span - 1)
        _merge(ws, f"{start}{row}:{end}{row}", header, fill=s["green"], font_color=s["white"], size=10, bold=True, align="center")
        col += span
    ws.row_dimensions[row].height = 28


def _write_kv_table(ws, start_row: int, rows: List[tuple[str, Any]], label_span: int = 2, value_span: int = 6, row_height: int = 34) -> int:
    s = _wb_styles()
    r = start_row
    for i, (campo, valor) in enumerate(rows):
        fill = "FFFFFF" if i % 2 == 0 else s["light"]
        _merge(ws, f"A{r}:{get_column_letter(label_span)}{r}", campo, fill=fill, font_color=s["green"], size=10, bold=True, align="left")
        _merge(ws, f"{get_column_letter(label_span+1)}{r}:{get_column_letter(label_span+value_span)}{r}", valor, fill=fill, font_color=s["dark"], size=10, align="left")
        ws.row_dimensions[r].height = row_height
        r += 1
    return r


def _write_dataframe_table(ws, start_row: int, df: pd.DataFrame, widths: Dict[str, int] | None = None, row_height: int = 34) -> int:
    s = _wb_styles()
    if df is None or df.empty:
        return start_row
    for c_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(start_row, c_idx)
        cell.value = str(col_name)
        cell.fill = PatternFill("solid", fgColor=s["green"])
        cell.font = Font(name="Calibri", size=10, bold=True, color=s["white"])
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _thin()
    ws.row_dimensions[start_row].height = 28

    for r_idx, (_, row) in enumerate(df.iterrows(), start_row + 1):
        fill = "FFFFFF" if (r_idx - start_row) % 2 else s["light"]
        for c_idx, col_name in enumerate(df.columns, 1):
            cell = ws.cell(r_idx, c_idx)
            cell.value = excel_clean(row.get(col_name, ""), "")
            cell.fill = PatternFill("solid", fgColor=fill)
            cell.font = Font(name="Calibri", size=10, color=s["dark"])
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            cell.border = _thin()
        ws.row_dimensions[r_idx].height = row_height

    if widths:
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
    return start_row + len(df) + 1


def gerar_excel(resultado: Dict[str, Any], texto: str) -> io.BytesIO:
    """Relatório Excel com visual executivo: capa, dashboard, dados, resumo, checklist, parecer, auditoria e texto."""
    from openpyxl import Workbook

    output = io.BytesIO()
    s = _wb_styles()

    def v(chave: str, padrao: str = "Não localizado") -> str:
        return excel_clean(resultado.get(chave, padrao), padrao)

    risco = normalize_risco(resultado.get("risco"))
    score = resultado.get("score", 0)
    risk_fill = excel_risk_fill(risco)
    pendencias_lista = resultado.get("pendencias", []) if isinstance(resultado.get("pendencias"), list) else []
    checklist_lista = resultado.get("checklist", []) if isinstance(resultado.get("checklist"), list) else []

    contraparte = v("contraparte", v("fornecedor"))
    cnpj_contraparte = v("cnpj_contraparte", v("cnpj"))
    valor_contrato = v("valor_contrato_original", v("valor_total"))
    vigencia = v("vigencia_apos_assinatura", v("vigencia"))
    pagamento = v("condicao_pagamento_dias", v("condicao_pagamento"))
    data_assinatura = v("data_assinatura")
    modelo_ia = v("modelo_ia", "Não informado")
    origem = v("tipo_origem", v("origem_contrato", "Não informado"))
    data_analise = v("data_analise", datetime.now().strftime("%d/%m/%Y %H:%M"))
    parecer = v("parecer")
    resumo = v("resumo_executivo")

    wb = Workbook()
    ws = wb.active
    ws.title = "Capa"

    # CAPA
    _sheet_base(ws, "NEXUS CONTRACT AI", "Relatório Executivo de Análise Contratual", 8, 100)
    row = _section(ws, 4, "Visão Geral da Análise")
    cards = [
        ("Status", v("status"), "A6:B8", s["green"]),
        ("Risco", risco, "C6:D8", risk_fill),
        ("Score", str(score), "E6:F8", s["green"]),
        ("Pendências", str(len(pendencias_lista)), "G6:H8", s["green"]),
    ]
    for label, value, cell_range, fill in cards:
        _metric_card(ws, cell_range, label, value, fill)
    ws.row_dimensions[6].height = 28
    ws.row_dimensions[7].height = 28
    ws.row_dimensions[8].height = 28

    row = _section(ws, 10, "Resumo do Contrato")
    row = _write_kv_table(ws, row, [
        ("Contraparte", contraparte),
        ("CNPJ Contraparte", cnpj_contraparte),
        ("Valor do Contrato", valor_contrato),
        ("Vigência", vigencia),
        ("Pagamento", pagamento),
        ("Data da Assinatura", data_assinatura),
        ("Contrato Assinado", v("contrato_assinado")),
        ("Modelo IA", modelo_ia),
    ], row_height=30)
    row = _section(ws, row + 1, "Parecer Automático")
    _merge(ws, f"A{row}:H{row+4}", parecer, fill=s["light"], font_color=s["dark"], size=11, align="left", valign="top")
    for rr in range(row, row+5):
        ws.row_dimensions[rr].height = 28

    # DASHBOARD EXECUTIVO
    ws = wb.create_sheet("Dashboard Executivo")
    _sheet_base(ws, "NEXUS CONTRACT AI", "Dashboard Executivo", 8, 100)
    _section(ws, 4, "Indicadores Principais")
    for label, value, cell_range, fill in cards:
        _metric_card(ws, cell_range, label, value, fill)
    _section(ws, 10, "Informações-Chave")
    _table_header(ws, 12, ["Indicador", "Valor"], [2, 6])
    _write_kv_table(ws, 13, [
        ("Contraparte", contraparte),
        ("CNPJ Contraparte", cnpj_contraparte),
        ("Valor do Contrato", valor_contrato),
        ("Condição de Pagamento", pagamento),
        ("Vigência", vigencia),
        ("Data da Assinatura", data_assinatura),
        ("Contrato Assinado", v("contrato_assinado")),
        ("Origem", origem),
        ("Modelo IA", modelo_ia),
        ("Data da Análise", data_analise),
    ], row_height=31)

    # DADOS EXTRAÍDOS
    ws = wb.create_sheet("Dados Extraídos")
    _sheet_base(ws, "NEXUS CONTRACT AI", "Relatório de Análise Contratual • Dados Extraídos", 8, 95)
    _table_header(ws, 5, ["Campo", "Informação"], [2, 6])
    dados_rows = [(label, v(chave)) for label, chave in CAMPOS_OFICIAIS]
    _write_kv_table(ws, 6, dados_rows, row_height=42)

    # RESUMO EXECUTIVO
    ws = wb.create_sheet("Resumo Executivo")
    _sheet_base(ws, "NEXUS CONTRACT AI", "Relatório de Análise Contratual • Resumo Executivo", 8, 95)
    _table_header(ws, 5, ["Seção", "Conteúdo"], [2, 6])
    _write_kv_table(ws, 6, [
        ("Resumo Executivo", resumo),
        ("Objeto / Escopo", v("descricao_servico_material", v("objetivo"))),
        ("Parecer Automático", parecer),
        ("Alerta de Assinatura", v("alerta_assinatura")),
    ], row_height=74)

    # CHECKLIST
    ws = wb.create_sheet("Checklist")
    _sheet_base(ws, "NEXUS CONTRACT AI", "Relatório de Análise Contratual • Checklist", 8, 95)
    checklist = pd.DataFrame(checklist_lista)
    if checklist.empty:
        checklist = pd.DataFrame([{"Validação": "Nenhum checklist retornado", "Status": "N/A", "Peso de risco": 0, "Crítico": "Não", "Evidência": "Não informado"}])
    _write_dataframe_table(ws, 5, checklist, {"A": 34, "B": 16, "C": 18, "D": 16, "E": 54}, 36)
    headers = {str(ws.cell(5, c).value).strip().lower(): c for c in range(1, ws.max_column + 1)}
    status_col = headers.get("status")
    critico_col = headers.get("crítico") or headers.get("critico")
    for rr in range(6, ws.max_row + 1):
        if status_col:
            cell = ws.cell(rr, status_col)
            txt = str(cell.value or "").upper()
            if "OK" in txt or "CONCLU" in txt:
                cell.fill = PatternFill("solid", fgColor=s["soft_green"])
                cell.font = Font(name="Calibri", size=10, bold=True, color="166534")
            elif "ATEN" in txt or "PEND" in txt:
                cell.fill = PatternFill("solid", fgColor=s["soft_warn"])
                cell.font = Font(name="Calibri", size=10, bold=True, color="92400E")
        if critico_col:
            cell = ws.cell(rr, critico_col)
            if str(cell.value or "").strip().upper() == "SIM":
                cell.fill = PatternFill("solid", fgColor=s["soft_danger"])
                cell.font = Font(name="Calibri", size=10, bold=True, color="991B1B")

    # PENDÊNCIAS
    ws = wb.create_sheet("Pendências")
    _sheet_base(ws, "NEXUS CONTRACT AI", "Relatório de Análise Contratual • Pendências", 8, 95)
    pendencias = pd.DataFrame(pendencias_lista)
    if pendencias.empty:
        pendencias = pd.DataFrame([{
            "Pendência": "Nenhuma pendência crítica localizada",
            "Crítico": "Não",
            "Risco": 0,
            "Recomendação": "Não foram identificadas pendências críticas na análise realizada.",
        }])
    _write_dataframe_table(ws, 5, pendencias, {"A": 42, "B": 16, "C": 14, "D": 70}, 46)
    for rr in range(6, ws.max_row + 1):
        if str(ws.cell(rr, 1).value or "").lower().startswith("nenhuma"):
            for cc in range(1, ws.max_column + 1):
                ws.cell(rr, cc).fill = PatternFill("solid", fgColor=s["soft_green"])
                ws.cell(rr, cc).font = Font(name="Calibri", size=10, bold=True, color="166534")

    # PARECER
    ws = wb.create_sheet("Parecer")
    _sheet_base(ws, "NEXUS CONTRACT AI", "Relatório de Análise Contratual • Parecer", 8, 95)
    _table_header(ws, 5, ["Item", "Descrição"], [2, 6])
    recomendacao = "Seguir com o processo caso as informações extraídas estejam de acordo com a documentação analisada." if risco == "BAIXO" else "Revisar as pendências e pontos de atenção antes de seguir com RC/PO."
    _write_kv_table(ws, 6, [
        ("Risco Geral", risco),
        ("Score", score),
        ("Parecer", parecer),
        ("Resumo Executivo", resumo),
        ("Recomendação", recomendacao),
    ], row_height=62)

    # AUDITORIA
    ws = wb.create_sheet("Auditoria")
    _sheet_base(ws, "NEXUS CONTRACT AI", "Relatório de Análise Contratual • Auditoria", 8, 95)
    _table_header(ws, 5, ["Campo", "Valor"], [2, 6])
    _write_kv_table(ws, 6, [
        ("Data da análise", data_analise),
        ("Modelo IA", modelo_ia),
        ("Origem", origem),
        ("Status", v("status")),
        ("Risco", risco),
        ("Score", score),
        ("Contrato Assinado", v("contrato_assinado")),
        ("Quantidade de Pendências", len(pendencias_lista)),
        ("Arquivos analisados", v("arquivos_analisados", v("arquivo", "Não informado"))),
    ], row_height=32)

    # TEXTO EXTRAÍDO
    ws = wb.create_sheet("Texto Extraído")
    _sheet_base(ws, "NEXUS CONTRACT AI", "Relatório de Análise Contratual • Texto Extraído", 8, 90)
    texto_limpo = clean_text(texto or resultado.get("texto_extraido", ""))
    if texto_limpo in ("", "Não localizado"):
        texto_limpo = "Texto extraído não disponível para este registro."
    blocos = [texto_limpo[i:i + 6000] for i in range(0, len(texto_limpo), 6000)] or [texto_limpo]
    _table_header(ws, 5, ["Bloco", "Texto Extraído"], [1, 7])
    r = 6
    for i, bloco in enumerate(blocos, 1):
        fill = "FFFFFF" if i % 2 else s["light"]
        _merge(ws, f"A{r}:A{r}", i, fill=fill, font_color=s["dark"], size=10, align="center")
        _merge(ws, f"B{r}:H{r}", excel_clean(bloco, ""), fill=fill, font_color=s["dark"], size=9, align="left", valign="top")
        ws.row_dimensions[r].height = 145
        r += 1

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
        ws.sheet_view.topLeftCell = "A1"
        ws.sheet_view.selection[0].sqref = "A1"
        ws.sheet_view.selection[0].activeCell = "A1"
        ws.sheet_properties.tabColor = s["green"]
        # Oculta colunas muito distantes para reduzir poluição visual.
        for col_idx in range(9, 27):
            ws.column_dimensions[get_column_letter(col_idx)].width = 3

    wb.active = 0
    wb.save(output)
    output.seek(0)
    return output


def gerar_excel_card_bytes(row: pd.Series) -> bytes:
    """Gera Excel completo do card do Dashboard."""
    resultado_completo: Dict[str, Any] = {}
    texto_extraido = ""

    raw_json = row.get("resultado_json") if "resultado_json" in row.index else None
    if raw_json not in (None, "", "Não informado"):
        try:
            resultado_completo = json.loads(raw_json)
            if not isinstance(resultado_completo, dict):
                resultado_completo = {}
        except Exception:
            resultado_completo = {}

    if resultado_completo:
        texto_extraido = resultado_completo.get("texto_extraido", "")
        return gerar_excel(resultado_completo, texto_extraido).getvalue()

    resultado_fallback = {
        "data_analise": row.get("data_analise"),
        "contraparte": row.get("fornecedor"),
        "fornecedor": row.get("fornecedor"),
        "cnpj_contraparte": row.get("cnpj"),
        "valor_contrato_original": row.get("valor_total"),
        "vigencia_apos_assinatura": row.get("vigencia"),
        "status": row.get("status"),
        "risco": row.get("risco"),
        "score": row.get("score"),
        "contrato_assinado": row.get("contrato_assinado"),
        "modelo_ia": row.get("modelo_ia"),
        "tipo_origem": row.get("tipo_origem"),
        "resumo_executivo": "Relatório gerado com os dados disponíveis no histórico.",
        "parecer": "Para gerar todas as informações detalhadas, refaça a análise do contrato nesta versão atualizada.",
        "checklist": [],
        "pendencias": [],
    }
    return gerar_excel(resultado_fallback, "Texto extraído não disponível para este registro antigo.").getvalue()



# Funções auxiliares para o Excel de Histórico.
# Mantidas para compatibilidade com a seção "Histórico".
def excel_apply_page(ws, zoom: int = 90) -> None:
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = zoom
    ws.freeze_panes = None
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.35
    ws.page_margins.bottom = 0.35


def excel_header(ws, titulo: str, subtitulo: str, last_col: int = 8) -> None:
    last_col = max(int(last_col or 1), 1)
    _sheet_base(ws, titulo, subtitulo, last_col=last_col, zoom=90)


def excel_style_table(ws, header_row: int = 6, data_start: int = 7) -> None:
    s = _wb_styles()
    # Cabeçalho
    for cell in ws[header_row]:
        if cell.value is not None:
            cell.fill = PatternFill("solid", fgColor=s["green"])
            cell.font = Font(name="Calibri", size=10, bold=True, color=s["white"])
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = _thin()
    ws.row_dimensions[header_row].height = 28

    # Linhas alternadas branco / verde claríssimo
    for r in range(data_start, ws.max_row + 1):
        fill = s["white"] if (r - data_start) % 2 == 0 else s["light"]
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            cell.fill = PatternFill("solid", fgColor=fill)
            cell.font = Font(name="Calibri", size=10, color=s["dark"])
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            cell.border = _thin()
        ws.row_dimensions[r].height = 34


def excel_auto_width(ws, max_width: int = 60) -> None:
    for idx, col_cells in enumerate(ws.columns, 1):
        max_len = 12
        for cell in col_cells:
            if cell.value is not None:
                max_len = max(max_len, min(len(str(cell.value)) + 2, max_width))
        ws.column_dimensions[get_column_letter(idx)].width = max_len



def _resumir_arquivos_excel(valor: Any) -> str:
    """Deixa a coluna de arquivos mais executiva no Excel do histórico."""
    txt = excel_clean(valor, "Não informado")
    if txt == "Não informado":
        return txt
    partes = [p.strip() for p in re.split(r"\s*\|\s*", txt) if p.strip()]
    if not partes:
        return txt
    if len(partes) == 1:
        return partes[0][:120]
    return f"{len(partes)} arquivos analisados"


def _preparar_historico_excel(df: pd.DataFrame) -> pd.DataFrame:
    """Padroniza colunas do histórico para exportação executiva."""
    colunas = [
        "ID", "Data da análise", "Contraparte", "CNPJ", "Valor total", "Vigência",
        "Status", "Risco", "Score", "Assinado", "Modelo IA", "Origem", "Arquivos analisados",
    ]
    base = df.copy() if df is not None else pd.DataFrame(columns=colunas)
    for col in colunas:
        if col not in base.columns:
            base[col] = "Não informado"

    base = base[colunas].copy()
    base["Score"] = pd.to_numeric(base["Score"], errors="coerce").fillna(0).astype(int)
    base["Risco"] = base["Risco"].astype(str).str.upper().replace({"MEDIO": "MÉDIO"})
    base["Arquivos analisados"] = base["Arquivos analisados"].apply(_resumir_arquivos_excel)

    for col in base.columns:
        if col != "Score":
            base[col] = base[col].apply(lambda x: excel_clean(x, "Não informado"))

    return base


def _aplicar_fundo_excel(ws, max_row: int = 80, max_col: int = 14) -> None:
    """Aplica fundo verde claro sem criar uma área visual gigante."""
    s = _wb_styles()
    for row_cells in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row_cells:
            if cell.value in (None, ""):
                cell.fill = PatternFill("solid", fgColor=s["background"])
            cell.border = _thin(s["line"])


def _formatar_risco_cell(cell, risco: Any) -> None:
    risco_norm = normalize_risco(risco)
    if risco_norm == "ALTO":
        cell.fill = PatternFill("solid", fgColor="FEE2E2")
        cell.font = Font(name="Calibri", size=10, bold=True, color="991B1B")
    elif risco_norm in ["MÉDIO", "MEDIO"]:
        cell.fill = PatternFill("solid", fgColor="FEF3C7")
        cell.font = Font(name="Calibri", size=10, bold=True, color="92400E")
    elif risco_norm == "BAIXO":
        cell.fill = PatternFill("solid", fgColor="DCFCE7")
        cell.font = Font(name="Calibri", size=10, bold=True, color="166534")
    else:
        cell.fill = PatternFill("solid", fgColor="E5E7EB")
        cell.font = Font(name="Calibri", size=10, bold=True, color="374151")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = _thin()


def gerar_excel_historico_profissional(export_df: pd.DataFrame, total_geral: int | None = None) -> io.BytesIO:
    """Gera um histórico executivo com dashboard, tabela filtrável e auditoria."""
    from openpyxl import Workbook
    from openpyxl.chart import PieChart, Reference
    from openpyxl.worksheet.table import Table, TableStyleInfo

    s = _wb_styles()
    output = io.BytesIO()
    df = _preparar_historico_excel(export_df)

    total_filtrado = int(len(df))
    total_geral = int(total_geral if total_geral is not None else total_filtrado)
    risco_series = df["Risco"].astype(str).str.upper().replace({"MEDIO": "MÉDIO"}) if not df.empty else pd.Series(dtype=str)
    qtd_alto = int((risco_series == "ALTO").sum())
    qtd_medio = int((risco_series == "MÉDIO").sum())
    qtd_baixo = int((risco_series == "BAIXO").sum())
    score_medio = round(float(pd.to_numeric(df["Score"], errors="coerce").fillna(0).mean()), 1) if total_filtrado else 0
    assinados = int(df["Assinado"].astype(str).str.upper().eq("SIM").sum()) if total_filtrado else 0
    data_geracao = datetime.now().strftime("%d/%m/%Y %H:%M")

    wb = Workbook()

    # =====================================================
    # ABA 1 - DASHBOARD
    # =====================================================
    ws = wb.active
    ws.title = "Dashboard"
    _sheet_base(ws, "NEXUS CONTRACT AI", "Dashboard Executivo do Histórico", 8, 95)
    _aplicar_fundo_excel(ws, 48, 10)
    ws.freeze_panes = None

    _section(ws, 4, "Indicadores Executivos", 8)
    cards = [
        ("Contratos filtrados", total_filtrado, "A6:B8", s["green"]),
        ("Base total", total_geral, "C6:D8", s["green2"]),
        ("Score médio", score_medio, "E6:F8", s["green"]),
        ("Assinados", assinados, "G6:H8", s["green2"]),
    ]
    for label, value, cell_range, fill in cards:
        _metric_card(ws, cell_range, label, value, fill)
        for rr in range(6, 9):
            ws.row_dimensions[rr].height = 27

    _section(ws, 10, "Distribuição por Risco", 8)
    risco_data = [("ALTO", qtd_alto), ("MÉDIO", qtd_medio), ("BAIXO", qtd_baixo)]
    _table_header(ws, 12, ["Risco", "Quantidade"], [2, 2])
    r = 13
    for risco_nome, qtd in risco_data:
        _merge(ws, f"A{r}:B{r}", risco_nome, fill=s["white"], font_color=s["dark"], size=10, bold=True, align="center")
        _merge(ws, f"C{r}:D{r}", qtd, fill=s["white"], font_color=s["dark"], size=10, bold=True, align="center")
        _formatar_risco_cell(ws[f"A{r}"], risco_nome)
        ws.row_dimensions[r].height = 25
        r += 1

    if total_filtrado:
        chart = PieChart()
        chart.title = "Risco dos contratos"
        labels = Reference(ws, min_col=1, min_row=13, max_row=15)
        data = Reference(ws, min_col=3, min_row=12, max_row=15)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(labels)
        chart.height = 7.5
        chart.width = 10
        chart.firstSliceAng = 270
        ws.add_chart(chart, "F12")

    _section(ws, 22, "Últimas análises", 8)
    ultimas = df.head(8).copy()
    resumo_cols = ["ID", "Data da análise", "Contraparte", "Valor total", "Risco", "Score"]
    _write_dataframe_table(ws, 24, ultimas[resumo_cols] if not ultimas.empty else pd.DataFrame(columns=resumo_cols),
                           {"A": 10, "B": 20, "C": 36, "D": 18, "E": 14, "F": 12}, 28)
    # Formata risco no resumo
    for rr in range(25, 25 + len(ultimas)):
        _formatar_risco_cell(ws.cell(rr, 5), ws.cell(rr, 5).value)
        ws.cell(rr, 6).alignment = Alignment(horizontal="center", vertical="center")

    _merge(ws, "A36:H39", "Leitura recomendada: use a aba Histórico Completo para filtros, auditoria e consulta detalhada dos contratos analisados.",
           fill=s["light"], font_color=s["dark"], size=11, align="left", valign="top")

    # =====================================================
    # ABA 2 - HISTÓRICO COMPLETO
    # =====================================================
    ws = wb.create_sheet("Histórico Completo")
    _sheet_base(ws, "NEXUS CONTRACT AI", "Histórico Completo de Análises", 13, 90)
    _aplicar_fundo_excel(ws, max(40, len(df) + 12), 14)
    ws.freeze_panes = "A7"

    start_row = 6
    # cabeçalho da tabela
    for c_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(start_row, c_idx)
        cell.value = col_name
        cell.fill = PatternFill("solid", fgColor=s["green"])
        cell.font = Font(name="Calibri", size=10, bold=True, color=s["white"])
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _thin()
    ws.row_dimensions[start_row].height = 30

    for r_idx, (_, row) in enumerate(df.iterrows(), start_row + 1):
        fill = s["white"] if (r_idx - start_row) % 2 else s["light"]
        for c_idx, col_name in enumerate(df.columns, 1):
            cell = ws.cell(r_idx, c_idx)
            cell.value = excel_clean(row[col_name], "") if col_name != "Score" else int(row[col_name])
            cell.fill = PatternFill("solid", fgColor=fill)
            cell.font = Font(name="Calibri", size=10, color=s["dark"])
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            cell.border = _thin()
        _formatar_risco_cell(ws.cell(r_idx, df.columns.get_loc("Risco") + 1), row["Risco"])
        ws.cell(r_idx, df.columns.get_loc("Score") + 1).alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[r_idx].height = 36

    if total_filtrado:
        last_row = start_row + total_filtrado
        table = Table(displayName="TabelaHistoricoNexus", ref=f"A{start_row}:M{last_row}")
        style = TableStyleInfo(name="TableStyleMedium4", showFirstColumn=False, showLastColumn=False, showRowStripes=False, showColumnStripes=False)
        table.tableStyleInfo = style
        ws.add_table(table)
        ws.auto_filter.ref = f"A{start_row}:M{last_row}"

    widths = {
        "A": 9, "B": 19, "C": 36, "D": 19, "E": 17, "F": 48,
        "G": 16, "H": 13, "I": 10, "J": 13, "K": 18, "L": 15, "M": 28,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    # =====================================================
    # ABA 3 - AUDITORIA
    # =====================================================
    ws = wb.create_sheet("Auditoria")
    _sheet_base(ws, "NEXUS CONTRACT AI", "Auditoria do Histórico", 8, 95)
    _aplicar_fundo_excel(ws, 45, 9)
    _section(ws, 4, "Informações da Exportação", 8)
    auditoria_rows = [
        ("Data de geração", data_geracao),
        ("Registros exportados", total_filtrado),
        ("Total geral no banco", total_geral),
        ("Score médio filtrado", score_medio),
        ("Risco alto", qtd_alto),
        ("Risco médio", qtd_medio),
        ("Risco baixo", qtd_baixo),
        ("Contratos assinados", assinados),
        ("Observação", "Relatório gerado com base nos filtros aplicados na aba Histórico do NEXUS Contract AI."),
    ]
    _table_header(ws, 6, ["Campo", "Valor"], [2, 6])
    _write_kv_table(ws, 7, auditoria_rows, row_height=32)

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = s["green"]
        ws.sheet_view.topLeftCell = "A1"
        ws.sheet_view.selection[0].sqref = "A1"
        ws.sheet_view.selection[0].activeCell = "A1"

    wb.active = 0
    wb.save(output)
    output.seek(0)
    return output

# =========================================================
# COMPONENTES DE TELA
# =========================================================
def render_hero(titulo: str, descricao: str, tag: str = "Suprimentos • Contratos") -> None:
    st.markdown(
        f"""
        <div class="hero">
            <span class="eyebrow">{safe(tag)}</span>
            <h1>{safe(titulo)}</h1>
            <p>{safe(descricao)}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_metric(label: str, value: Any) -> str:
    return f'<div class="metric-card"><small>{safe(label)}</small><h2>{safe(value)}</h2></div>'


def render_filter_metric(label: str, value: Any, filtro: str, ativo: bool = False) -> str:
    """Card clicável do Dashboard. Mantém o usuário na aba Dashboard ao filtrar."""
    classe_ativo = " active" if ativo else ""
    href = f"?pagina=dashboard&risco={filtro}"
    return f"""
    <a class="metric-link" href="{href}" target="_self">
        <div class="metric-card clickable{classe_ativo}">
            <small>{safe(label)}</small>
            <h2>{safe(value)}</h2>
        </div>
    </a>
    """

def render_info_card(label: str, value: Any) -> str:
    return f'<div class="info-card"><small>{safe(label)}</small><p>{safe(value)}</p></div>'


def render_contract_card(row: pd.Series) -> None:
    """Renderiza o card do Dashboard sem HTML bruto.

    A versão anterior montava um grande bloco HTML com link base64; em algumas
    versões/temas do Streamlit isso aparecia como texto na tela. Aqui usamos
    componentes nativos para garantir estabilidade visual.
    """
    risco = normalize_risco(row.get("risco"))
    cor = risco_cor(risco)

    contraparte = row.get("fornecedor") or row.get("contraparte") or "Não localizado"
    cnpj = row.get("cnpj") or row.get("cnpj_contraparte") or "Não localizado"
    valor = row.get("valor_total") or row.get("valor_contrato_original") or "Não localizado"
    vigencia = row.get("vigencia") or row.get("vigencia_apos_assinatura") or "Não localizada"
    status = row.get("status") or "Não localizado"
    score = row.get("score") or "0"
    assinado = row.get("contrato_assinado") or "Não informado"
    origem = row.get("tipo_origem") or row.get("origem") or "Não informado"
    modelo = row.get("modelo_ia") or "Não informado"
    data = row.get("data_analise") or "Não informado"
    arquivo = row.get("arquivo") or "Não informado"

    with st.container(border=True):
        topo_esq, topo_dir = st.columns([5, 1])

        with topo_esq:
            st.markdown("### 📄 Análise de Contrato")
            st.caption(str(arquivo))

        with topo_dir:
            st.markdown(
                f"""
                <div style="text-align:center;background:{cor};color:white;
                padding:10px 14px;border-radius:999px;font-weight:900;font-size:12px;">
                    {safe(risco)}
                </div>
                """,
                unsafe_allow_html=True,
            )

        c1, c2, c3, c4 = st.columns([1.6, 1.1, 1.2, .7])
        c1.markdown(f"**Contraparte**  \n{safe(contraparte)}", unsafe_allow_html=True)
        c2.markdown(f"**CNPJ**  \n{safe(cnpj)}", unsafe_allow_html=True)
        c3.markdown(f"**Valor do Contrato**  \n{safe(valor)}", unsafe_allow_html=True)
        c4.markdown(f"**Score**  \n{safe(score)}", unsafe_allow_html=True)

        c5, c6, c7, c8 = st.columns([1.6, 1.1, 1.2, .7])
        c5.markdown(f"**Vigência**  \n{safe(vigencia)}", unsafe_allow_html=True)
        c6.markdown(f"**Status**  \n{safe(status)}", unsafe_allow_html=True)
        c7.markdown(f"**Origem**  \n{safe(origem)}", unsafe_allow_html=True)
        c8.markdown(f"**Assinado**  \n{safe(assinado)}", unsafe_allow_html=True)

        rodape_esq, rodape_dir = st.columns([4, 1])
        rodape_esq.caption(f"Analisado em {data} • Modelo: {modelo}")

        excel_bytes = gerar_excel_card_bytes(row)
        rodape_dir.download_button(
            "📥 Relatório Excel",
            data=excel_bytes,
            file_name=f"analise_contrato_{row.get('id', 'historico')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"download_dashboard_{row.get('id', id(row))}",
        )



def obter_resultado_completo_historico(row: pd.Series) -> tuple[Dict[str, Any], str]:
    """Recupera do histórico o JSON completo salvo na análise.

    Quando o registro é antigo e não possui resultado_json, monta um fallback com
    os dados básicos do card para não quebrar a visualização.
    """
    resultado: Dict[str, Any] = {}
    texto_extraido = ""

    raw_json = row.get("resultado_json") if "resultado_json" in row.index else None
    if raw_json not in (None, "", "Não informado"):
        try:
            resultado = json.loads(raw_json)
            if not isinstance(resultado, dict):
                resultado = {}
        except Exception:
            resultado = {}

    if resultado:
        texto_extraido = str(resultado.get("texto_extraido") or row.get("texto_extraido") or "")
    else:
        texto_extraido = str(row.get("texto_extraido") or "")
        resultado = {
            "data_analise": row.get("data_analise"),
            "contraparte": row.get("fornecedor"),
            "fornecedor": row.get("fornecedor"),
            "cnpj_contraparte": row.get("cnpj"),
            "cnpj": row.get("cnpj"),
            "valor_contrato_original": row.get("valor_total"),
            "valor_total": row.get("valor_total"),
            "vigencia_apos_assinatura": row.get("vigencia"),
            "vigencia": row.get("vigencia"),
            "status": row.get("status"),
            "risco": row.get("risco"),
            "score": row.get("score"),
            "contrato_assinado": row.get("contrato_assinado"),
            "modelo_ia": row.get("modelo_ia"),
            "tipo_origem": row.get("tipo_origem"),
            "arquivos_analisados": row.get("arquivo"),
            "resumo_executivo": "Registro antigo: visualização montada com os dados disponíveis no histórico.",
            "parecer": "Para visualizar todos os campos com maior detalhe, refaça a análise do contrato nesta versão atualizada.",
            "checklist": [],
            "pendencias": [],
        }

    resultado.setdefault("texto_extraido", texto_extraido)
    resultado.setdefault("arquivos_analisados", row.get("arquivo", "Não informado"))
    resultado.setdefault("data_analise", row.get("data_analise", "Não informado"))
    resultado.setdefault("modelo_ia", row.get("modelo_ia", "Não informado"))
    resultado.setdefault("tipo_origem", row.get("tipo_origem", "Não informado"))

    try:
        resultado = normalizar(resultado)
    except Exception:
        pass

    return resultado, texto_extraido


def render_analise_completa_historico(row: pd.Series) -> None:
    """Mostra no Histórico a mesma visão completa exibida após uma nova análise."""
    resultado, texto_extraido = obter_resultado_completo_historico(row)
    risco = normalize_risco(resultado.get("risco"))
    pill = "pill-ok" if risco == "BAIXO" else "pill-warn" if risco == "MÉDIO" else "pill-danger"

    st.markdown('<div class="section-title">Resumo da análise</div>', unsafe_allow_html=True)
    m1, m2, m3, m4 = st.columns(4)
    m1.markdown(render_metric("Status", resultado.get("status")), unsafe_allow_html=True)
    m2.markdown(render_metric("Risco", risco), unsafe_allow_html=True)
    m3.markdown(render_metric("Score", resultado.get("score")), unsafe_allow_html=True)
    m4.markdown(render_metric("Pendências", len(resultado.get("pendencias", []))), unsafe_allow_html=True)

    if str(resultado.get("contrato_assinado", "")).upper() == "NÃO":
        st.error("⚠️ Contrato sem assinatura localizada. Revisar antes da criação da RC/PO.")

    st.markdown('<div class="section-title">Resumo executivo</div>', unsafe_allow_html=True)
    st.markdown(f'<span class="pill {pill}">{safe(resultado.get("resumo_executivo"))}</span>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Dados extraídos</div>', unsafe_allow_html=True)
    cards_html = "".join(
        render_info_card(label, resultado.get(chave))
        for label, chave in CAMPOS_OFICIAIS
    )
    st.markdown(f'<div class="info-grid">{cards_html}</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Objeto / Escopo</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="executive-box">{safe(resultado.get("descricao_servico_material") or resultado.get("objetivo"))}</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Checklist de validação</div>', unsafe_allow_html=True)
    df_checklist = pd.DataFrame(resultado.get("checklist", []))
    if df_checklist.empty:
        st.info("Checklist detalhado não disponível para este registro.")
    else:
        st.dataframe(df_checklist, use_container_width=True, hide_index=True)

    st.markdown('<div class="section-title">Pendências encontradas</div>', unsafe_allow_html=True)
    pendencias = resultado.get("pendencias", []) if isinstance(resultado.get("pendencias"), list) else []
    if pendencias:
        for i, pendencia in enumerate(pendencias, 1):
            st.markdown(
                f"""
                <div class="risk-row">
                    <b>{i}. {safe(pendencia.get('Pendência', 'Pendência'))}</b><br>
                    Crítico: {safe(pendencia.get('Crítico', 'N/A'))} • Risco: {safe(pendencia.get('Risco', 'N/A'))}<br>
                    <span class="subtle">{safe(pendencia.get('Recomendação', 'Validar antes de seguir.'))}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )
    else:
        st.markdown('<div class="ok-row">Nenhuma pendência crítica localizada.</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Parecer automático</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="executive-box"><b>Parecer:</b><br><br>{safe(resultado.get("parecer"))}</div>', unsafe_allow_html=True)

    with st.expander("📄 Ver texto extraído do contrato e anexos"):
        texto = texto_extraido or str(resultado.get("texto_extraido") or "Texto extraído não disponível para este registro.")
        st.text_area("Texto extraído", texto[:50000], height=320, key=f"texto_historico_{row.get('id', id(row))}")

    st.download_button(
        "📥 Baixar relatório Excel completo",
        data=gerar_excel(resultado, texto_extraido).getvalue(),
        file_name=f"analise_completa_contrato_{row.get('id', 'historico')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key=f"download_historico_completo_{row.get('id', id(row))}",
    )

# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    st.markdown("## ⚖️ NEXUS CONTRACT")
    st.caption("Análise inteligente de contratos")
    st.divider()

    query_pagina = str(st.query_params.get("pagina", "")).lower()
    abrir_dashboard = bool(st.query_params.get("risco")) or query_pagina == "dashboard"

    pagina = st.radio(
    "Menu",
    [
        "🏠 Dashboard",
        "📄 Nova Análise",
        "📚 Histórico",
        "🤖 Assistente IA"
    ],
    index=0 if abrir_dashboard else 1,
)

    st.divider()

    modo = st.radio(
        "Modo de análise",
        [
            "Análise Local",
            "Automático recomendado",
            "Gemini 3.5 Thinking",
            "Gemini 3.1 Pro",
            "Gemini 3.5 Flash",
        ],
        index=1,
    )

    gemini_key = os.getenv("GEMINI_API_KEY", "") if modo != "Análise Local" else ""

    if modo == "Automático recomendado":
        st.info("Modo automático: tenta Gemini 3.5 Thinking, depois 3.1 Pro e, por último, 3.5 Flash.")
    elif modo == "Análise Local":
        st.info("A análise local é objetiva e serve como fallback. Para melhor leitura jurídica, use Gemini.")
    else:
        st.info(f"Modelo selecionado: {modo}.")


# =========================================================
# DASHBOARD
# =========================================================
if pagina == "🏠 Dashboard":
    render_hero("Dashboard", "Visão executiva das análises de contratos, riscos e pendências.")

    historico = carregar_historico_seguro()

    col_a, col_b = st.columns([3, 1])
    with col_b:
        if st.button("🗑️ Limpar histórico", use_container_width=True):
            limpar_historico()
            st.success("Histórico limpo com sucesso.")
            st.rerun()

    total = len(historico)
    if total > 0:
        risco_series = historico["risco"].astype(str).str.upper().replace({"MEDIO": "MÉDIO"})
        alto = int((risco_series == "ALTO").sum())
        medio = int((risco_series == "MÉDIO").sum())
        baixo = int((risco_series == "BAIXO").sum())
        score_medio = round(historico["score"].apply(as_float_score).mean(), 1)
    else:
        alto = medio = baixo = score_medio = 0

    st.markdown(
        f"""
        <div class="executive-box">
            <h3>📊 Resumo executivo</h3>
            <p>O ambiente possui <b>{total}</b> contrato(s) analisado(s). Atualmente existem <b>{alto}</b> contrato(s) classificados como risco alto.</p>
            <p>A prioridade recomendada é revisar contratos de risco <b>ALTO</b> antes da criação ou continuidade de RC/PO.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    filtro_param = st.query_params.get("risco", "TODOS")
    if isinstance(filtro_param, list):
        filtro_param = filtro_param[0] if filtro_param else "TODOS"

    filtro_atual = normalize_risco(filtro_param)
    if filtro_atual in ["MEDIO", "MÉDIO"]:
        filtro_atual = "MÉDIO"
    if filtro_atual not in ["TODOS", "ALTO", "MÉDIO", "BAIXO"]:
        filtro_atual = "TODOS"

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.markdown(render_filter_metric("Contratos", total, "TODOS", filtro_atual == "TODOS"), unsafe_allow_html=True)
    c2.markdown(render_filter_metric("Risco alto", alto, "ALTO", filtro_atual == "ALTO"), unsafe_allow_html=True)
    c3.markdown(render_filter_metric("Risco médio", medio, "MEDIO", filtro_atual == "MÉDIO"), unsafe_allow_html=True)
    c4.markdown(render_filter_metric("Risco baixo", baixo, "BAIXO", filtro_atual == "BAIXO"), unsafe_allow_html=True)
    c5.markdown(render_filter_metric("Score médio", score_medio, "TODOS", False), unsafe_allow_html=True)

    historico_filtrado = historico.copy()
    if not historico_filtrado.empty and filtro_atual != "TODOS":
        risco_filtro = historico_filtrado["risco"].astype(str).str.upper().replace({"MEDIO": "MÉDIO"})
        historico_filtrado = historico_filtrado[risco_filtro == filtro_atual]

    texto_filtro = "todos os contratos" if filtro_atual == "TODOS" else f"contratos com risco {filtro_atual}"
    st.markdown(
        f'<div class="filter-note">🔎 Visualizando {len(historico_filtrado)} de {total} análise(s): <b>{safe(texto_filtro)}</b>.</div>',
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title">Últimas análises</div>', unsafe_allow_html=True)
    if historico.empty:
        st.info("Nenhuma análise registrada ainda.")
    elif historico_filtrado.empty:
        st.warning("Nenhuma análise encontrada para este filtro.")
    else:
        for _, row in historico_filtrado.head(10).iterrows():
            render_contract_card(row)

        st.markdown('<div class="footer">NEXUS CONTRACT AI • Suprimentos • Análise de Contratos</div>', unsafe_allow_html=True)
        st.stop()

# =========================================================
# ASSISTENTE IA LOCAL
# =========================================================
if pagina == "🤖 Assistente IA":

    render_hero(
        "Assistente Nexus",
        "Consulte informações dos contratos analisados no sistema de forma simples, visual e executiva."
    )

    contratos = carregar_contratos_chat()

    if contratos.empty:
        st.warning("Nenhum contrato encontrado no banco.")
        st.markdown('<div class="footer">NEXUS CONTRACT AI • Suprimentos • Análise de Contratos</div>', unsafe_allow_html=True)
        st.stop()

    st.markdown('<div class="section-title">Perguntas rápidas</div>', unsafe_allow_html=True)

    exemplos = [
        "Quantos contratos existem?",
        "Quais contratos são de risco alto?",
        "Quais contratos são de risco médio?",
        "Quais contratos são de risco baixo?",
        "Qual contrato tem maior valor?",
        "Qual contrato tem menor score?",
        "Qual é o score médio?",
        "Quais contratos estão assinados?",
        "Quais contratos não estão assinados?",
        "Quais contratos são do Projuris?",
        "Quais contratos são do Ariba?",
        "Quais contratos foram analisados pelo Gemini?",
        "Liste os últimos contratos analisados",
    ]

    pergunta_exemplo = st.selectbox(
        "Escolha uma pergunta pronta",
        [""] + exemplos,
    )

    pergunta_digitada = st.chat_input("Ou digite sua pergunta sobre os contratos...")
    pergunta = pergunta_digitada or pergunta_exemplo

    def texto_tem(texto: str, palavras: List[str]) -> bool:
        texto = str(texto).lower()
        return any(p in texto for p in palavras)

    def preparar_df_chat(df: pd.DataFrame) -> pd.DataFrame:
        base = df.copy()
        colunas_padrao = [
            "fornecedor", "cnpj", "valor_total", "vigencia", "status", "risco",
            "score", "contrato_assinado", "modelo_ia", "tipo_origem", "arquivo", "data_analise"
        ]
        for col in colunas_padrao:
            if col not in base.columns:
                base[col] = "Não informado"

        base["risco_norm"] = base["risco"].astype(str).str.upper().replace({"MEDIO": "MÉDIO"})
        base["score_num"] = pd.to_numeric(base["score"], errors="coerce").fillna(0)
        base["valor_num"] = (
            base["valor_total"]
            .astype(str)
            .str.replace("R$", "", regex=False)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        base["valor_num"] = pd.to_numeric(base["valor_num"], errors="coerce").fillna(0)
        return base

    def render_ai_resumo(df_resultado: pd.DataFrame) -> None:
        if df_resultado.empty:
            st.info("Nenhum contrato encontrado para esta busca.")
            return

        riscos = df_resultado["risco_norm"].astype(str)
        score_medio = round(float(df_resultado["score_num"].mean()), 1) if len(df_resultado) else 0
        alto = int((riscos == "ALTO").sum())
        medio = int((riscos == "MÉDIO").sum())
        baixo = int((riscos == "BAIXO").sum())

        st.markdown('<div class="section-title">Resumo executivo</div>', unsafe_allow_html=True)
        a1, a2, a3, a4, a5 = st.columns(5)
        a1.markdown(render_metric("Contratos", len(df_resultado)), unsafe_allow_html=True)
        a2.markdown(render_metric("Score médio", score_medio), unsafe_allow_html=True)
        a3.markdown(render_metric("Risco alto", alto), unsafe_allow_html=True)
        a4.markdown(render_metric("Risco médio", medio), unsafe_allow_html=True)
        a5.markdown(render_metric("Risco baixo", baixo), unsafe_allow_html=True)

    def render_ai_cards(df_resultado: pd.DataFrame, limite: int = 10) -> None:
        if df_resultado.empty:
            return

        st.markdown('<div class="section-title">Contratos encontrados</div>', unsafe_allow_html=True)
        for _, r in df_resultado.head(limite).iterrows():
            risco = normalize_risco(r.get("risco"))
            cor = risco_cor(risco)
            with st.container(border=True):
                topo1, topo2 = st.columns([5, 1])
                with topo1:
                    st.markdown(f"### 📄 {safe(r.get('fornecedor', 'Não informado'))}", unsafe_allow_html=True)
                    st.caption(str(r.get("arquivo") or "Não informado"))
                with topo2:
                    st.markdown(
                        f"""
                        <div style="text-align:center;background:{cor};color:white;
                        padding:10px 14px;border-radius:999px;font-weight:900;font-size:12px;">
                            {safe(risco)}
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                c1, c2, c3, c4, c5 = st.columns([1.2, 1.1, 1.2, .8, 1])
                c1.markdown(f"**CNPJ**  \n{safe(r.get('cnpj'))}", unsafe_allow_html=True)
                c2.markdown(f"**Score**  \n{safe(r.get('score'))}", unsafe_allow_html=True)
                c3.markdown(f"**Valor**  \n{safe(r.get('valor_total'))}", unsafe_allow_html=True)
                c4.markdown(f"**Assinado**  \n{safe(r.get('contrato_assinado'))}", unsafe_allow_html=True)
                c5.markdown(f"**Origem**  \n{safe(r.get('tipo_origem'))}", unsafe_allow_html=True)

                c6, c7 = st.columns([1.4, 2])
                c6.markdown(f"**Status**  \n{safe(r.get('status'))}", unsafe_allow_html=True)
                c7.markdown(f"**Vigência**  \n{safe(r.get('vigencia'))}", unsafe_allow_html=True)

        if len(df_resultado) > limite:
            st.info(f"Exibindo {limite} de {len(df_resultado)} contrato(s) encontrados.")

    def render_ai_destaque(titulo: str, row: pd.Series) -> None:
        st.markdown(f'<div class="section-title">{safe(titulo)}</div>', unsafe_allow_html=True)
        render_ai_cards(pd.DataFrame([row]), limite=1)

    def render_busca_vazia() -> None:
        st.warning("Não encontrei contratos relacionados a essa busca.")
        st.info("Você pode pesquisar por fornecedor, CNPJ, valor, risco, status, origem, modelo IA ou nome do arquivo.")

    if pergunta:
        pergunta_lower = pergunta.lower().strip()
        df = preparar_df_chat(contratos)

        st.chat_message("user").write(pergunta)

        with st.chat_message("assistant"):
            try:
                if texto_tem(pergunta_lower, ["oi", "olá", "ola", "bom dia", "boa tarde", "boa noite"]):
                    st.success(
                        "Olá! Eu sou o Assistente Nexus. Posso consultar quantidade de contratos, riscos, "
                        "scores, valores, origem, assinatura e histórico das análises."
                    )

                elif texto_tem(pergunta_lower, ["quantos contratos", "total de contratos", "quantidade de contratos", "qtd contratos"]):
                    st.success(f"Existem {len(df)} contrato(s) cadastrados no histórico.")
                    render_ai_resumo(df)

                elif texto_tem(pergunta_lower, ["risco alto", "alto risco", "contratos alto"]):
                    filtro = df[df["risco_norm"] == "ALTO"]
                    render_ai_resumo(filtro)
                    render_ai_cards(filtro, limite=20)

                elif texto_tem(pergunta_lower, ["risco médio", "risco medio", "médio risco", "medio risco"]):
                    filtro = df[df["risco_norm"] == "MÉDIO"]
                    render_ai_resumo(filtro)
                    render_ai_cards(filtro, limite=20)

                elif texto_tem(pergunta_lower, ["risco baixo", "baixo risco", "contratos baixo"]):
                    filtro = df[df["risco_norm"] == "BAIXO"]
                    render_ai_resumo(filtro)
                    render_ai_cards(filtro, limite=20)

                elif texto_tem(pergunta_lower, ["maior valor", "valor mais alto", "contrato mais caro", "maior contrato"]):
                    maior = df.sort_values("valor_num", ascending=False).iloc[0]
                    render_ai_destaque("Contrato com maior valor", maior)

                elif texto_tem(pergunta_lower, ["menor score", "pior score", "menor nota", "pior contrato"]):
                    menor = df.sort_values("score_num", ascending=True).iloc[0]
                    render_ai_destaque("Contrato com menor score", menor)

                elif texto_tem(pergunta_lower, ["maior score", "melhor score", "maior nota", "melhor contrato"]):
                    maior_score = df.sort_values("score_num", ascending=False).iloc[0]
                    render_ai_destaque("Contrato com maior score", maior_score)

                elif texto_tem(pergunta_lower, ["score médio", "score medio", "média de score", "media de score"]):
                    st.success(f"O score médio dos contratos é {round(float(df['score_num'].mean()), 1)}.")
                    render_ai_resumo(df)

                elif texto_tem(pergunta_lower, ["contratos assinados", "estão assinados", "assinados", "com assinatura"]):
                    filtro = df[df["contrato_assinado"].astype(str).str.upper() == "SIM"]
                    render_ai_resumo(filtro)
                    render_ai_cards(filtro, limite=20)

                elif texto_tem(pergunta_lower, ["não assinados", "nao assinados", "sem assinatura", "não estão assinados", "nao estao assinados"]):
                    filtro = df[df["contrato_assinado"].astype(str).str.upper() != "SIM"]
                    render_ai_resumo(filtro)
                    render_ai_cards(filtro, limite=20)

                elif texto_tem(pergunta_lower, ["projuris"]):
                    filtro = df[df["tipo_origem"].astype(str).str.lower().str.contains("projuris", na=False)]
                    render_ai_resumo(filtro)
                    render_ai_cards(filtro, limite=20)

                elif texto_tem(pergunta_lower, ["ariba"]):
                    filtro = df[df["tipo_origem"].astype(str).str.lower().str.contains("ariba", na=False)]
                    render_ai_resumo(filtro)
                    render_ai_cards(filtro, limite=20)

                elif texto_tem(pergunta_lower, ["gemini", "ia", "inteligência artificial", "inteligencia artificial"]):
                    filtro = df[df["modelo_ia"].astype(str).str.lower().str.contains("gemini", na=False)]
                    render_ai_resumo(filtro)
                    render_ai_cards(filtro, limite=20)

                elif texto_tem(pergunta_lower, ["últimos", "ultimos", "recentes", "últimas análises", "ultimas analises"]):
                    ultimos = df.head(10)
                    render_ai_resumo(ultimos)
                    render_ai_cards(ultimos, limite=10)

                else:
                    colunas_busca = [
                        "fornecedor", "cnpj", "valor_total", "vigencia", "status", "risco",
                        "contrato_assinado", "modelo_ia", "tipo_origem", "arquivo",
                    ]
                    filtro = pd.Series(False, index=df.index)
                    for coluna in colunas_busca:
                        if coluna in df.columns:
                            filtro = filtro | df[coluna].astype(str).str.lower().str.contains(pergunta_lower, na=False, regex=False)

                    resultado_busca = df[filtro]
                    if resultado_busca.empty:
                        render_busca_vazia()
                    else:
                        render_ai_resumo(resultado_busca)
                        render_ai_cards(resultado_busca, limite=50)

            except Exception as erro:
                st.error(f"Erro ao consultar o histórico: {erro}")

    st.markdown('<div class="footer">NEXUS CONTRACT AI • Suprimentos • Análise de Contratos</div>', unsafe_allow_html=True)
    st.stop()


# =========================================================
# NOVA ANÁLISE
# =========================================================
if pagina == "📄 Nova Análise":
    render_hero("NEXUS Contract AI", "Análise profissional e automatizada de contratos Projuris ou Ariba em PDF e Word.")

    st.markdown('<div class="section-title">Tipo de análise</div>', unsafe_allow_html=True)
    origem_contrato = st.radio("Origem do contrato", ["📘 Projuris", "🛒 Ariba"], horizontal=True, label_visibility="collapsed")

    st.markdown('<div class="section-title">Upload do contrato</div>', unsafe_allow_html=True)
    arquivos = st.file_uploader(
        "Envie o contrato principal e anexos",
        type=["pdf", "docx"],
        accept_multiple_files=True,
    )

    if not arquivos:
        cols = st.columns(4)
        cards = [("Formato", "PDF/Word"), ("Análise", "Local ou IA"), ("Saída", "Excel"), ("Risco", "Score")]
        for col, (label, value) in zip(cols, cards):
            col.markdown(render_metric(label, value), unsafe_allow_html=True)

        st.markdown('<div class="section-title">Como funciona</div>', unsafe_allow_html=True)
        st.markdown(
            """
            <div class="flow-grid">
                <div class="flow-card"><div class="flow-number">1</div><h4>Upload do contrato</h4><p>Envie o contrato principal e todos os anexos necessários.</p></div>
                <div class="flow-card"><div class="flow-number">2</div><h4>Extração dos dados</h4><p>O sistema consolida PDF, Word e anexos em uma análise única.</p></div>
                <div class="flow-card"><div class="flow-number">3</div><h4>Validação do risco</h4><p>As cláusulas são avaliadas por pendência, criticidade e score.</p></div>
                <div class="flow-card"><div class="flow-number">4</div><h4>Relatório executivo</h4><p>Gere o Excel com resumo, checklist, pendências e texto extraído.</p></div>
                <div class="flow-card"><div class="flow-number">5</div><h4>Histórico</h4><p>A análise fica salva para consulta posterior no dashboard.</p></div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.success(f"{len(arquivos)} arquivo(s) carregado(s). Clique abaixo para iniciar a análise.")

        if st.button("🚀 Analisar contrato e anexos", use_container_width=True):
            with st.spinner("Lendo contrato e anexos..."):
                texto_total = ""

                for arquivo in arquivos:
                    try:
                        if arquivo.name.lower().endswith(".pdf"):
                            texto_arquivo = ler_pdf(arquivo)
                        else:
                            texto_arquivo = ler_docx(arquivo)
                    except Exception as e:
                        texto_arquivo = f"Erro ao ler arquivo {arquivo.name}: {e}"

                    texto_total += "\n\n==============================\n"
                    texto_total += f"ARQUIVO: {arquivo.name}\n"
                    texto_total += "==============================\n"
                    texto_total += texto_arquivo

                texto = texto_total.strip()

                try:
                    if modo != "Análise Local":
                        if not gemini_key:
                            st.error("Configure a GEMINI_API_KEY no arquivo .env para usar a análise IA.")
                            st.stop()

                        resultado = analisar_gemini(
                            texto=texto,
                            api_key=gemini_key,
                            opcao_modelo=modo,
                        )
                    else:
                        resultado = local_extract(texto)
                except Exception as e:
                    st.warning(f"A IA falhou e o sistema usou análise local. Detalhe: {e}")
                    resultado = local_extract(texto)

                resultado = normalizar(resultado)
                resultado["data_analise"] = datetime.now().strftime("%d/%m/%Y %H:%M")
                resultado["origem_contrato"] = origem_contrato

                nomes_arquivos = " | ".join([arquivo.name for arquivo in arquivos])
                resultado["texto_extraido"] = texto
                resultado["arquivos_analisados"] = nomes_arquivos
                resultado["tipo_origem"] = origem_contrato.replace("📘", "").replace("🛒", "").strip()
                resultado["modelo_ia"] = resultado.get("modelo_ia", modo if modo != "Análise Local" else "Análise Local")
                salvar_analise(
                    resultado,
                    nomes_arquivos,
                    modelo_ia=resultado.get("modelo_ia"),
                    tipo_origem=resultado.get("tipo_origem"),
                    texto_extraido=texto,
                )

            risco = normalize_risco(resultado.get("risco"))
            pill = "pill-ok" if risco == "BAIXO" else "pill-warn" if risco == "MÉDIO" else "pill-danger"

            c1, c2, c3, c4 = st.columns(4)
            c1.markdown(render_metric("Status", resultado.get("status")), unsafe_allow_html=True)
            c2.markdown(render_metric("Risco", risco), unsafe_allow_html=True)
            c3.markdown(render_metric("Score", resultado.get("score")), unsafe_allow_html=True)
            c4.markdown(render_metric("Pendências", len(resultado.get("pendencias", []))), unsafe_allow_html=True)

            if str(resultado.get("contrato_assinado", "")).upper() == "NÃO":
                st.error("⚠️ Contrato sem assinatura localizada. Revisar antes da criação da RC/PO.")

            st.markdown('<div class="section-title">Resumo executivo</div>', unsafe_allow_html=True)
            st.markdown(f'<span class="pill {pill}">{safe(resultado.get("resumo_executivo"))}</span>', unsafe_allow_html=True)

            st.markdown('<div class="section-title">Dados extraídos</div>', unsafe_allow_html=True)
            cards_html = "".join(
                render_info_card(label, resultado.get(chave))
                for label, chave in CAMPOS_OFICIAIS
            )
            st.markdown(f'<div class="info-grid">{cards_html}</div>', unsafe_allow_html=True)

            st.markdown('<div class="section-title">Objeto / Escopo</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="executive-box">{safe(resultado.get("descricao_servico_material"))}</div>', unsafe_allow_html=True)

            st.markdown('<div class="section-title">Checklist de validação</div>', unsafe_allow_html=True)
            df_checklist = pd.DataFrame(resultado.get("checklist", []))
            st.dataframe(df_checklist, use_container_width=True, hide_index=True)

            st.markdown('<div class="section-title">Pendências encontradas</div>', unsafe_allow_html=True)
            if resultado.get("pendencias"):
                for pendencia in resultado.get("pendencias", []):
                    st.markdown(
                        f"""
                        <div class="risk-row">
                            <b>{safe(pendencia.get('Pendência', 'Pendência'))}</b><br>
                            Crítico: {safe(pendencia.get('Crítico', 'N/A'))} • Risco: {safe(pendencia.get('Risco', 'N/A'))}<br>
                            <span class="subtle">{safe(pendencia.get('Recomendação', 'Validar antes de seguir.'))}</span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
            else:
                st.markdown('<div class="ok-row">Nenhuma pendência crítica localizada.</div>', unsafe_allow_html=True)

            st.markdown('<div class="section-title">Parecer automático</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="executive-box"><b>Parecer:</b><br><br>{safe(resultado.get("parecer"))}</div>', unsafe_allow_html=True)

            with st.expander("Ver texto extraído do contrato e anexos"):
                st.text_area("Texto", texto[:50000], height=320)

            excel = gerar_excel(resultado, texto)
            st.download_button(
                "📥 Baixar relatório Excel",
                data=excel,
                file_name=f"relatorio_nexus_contract_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )


# =========================================================
# HISTÓRICO
# =========================================================
if pagina == "📚 Histórico":
    render_hero("Histórico", "Consulta executiva dos contratos analisados, com filtros, indicadores e relatórios.")

    historico = carregar_historico_seguro()

    if historico.empty:
        st.info("Nenhuma análise salva ainda.")
        st.markdown('<div class="footer">NEXUS CONTRACT AI • Suprimentos • Análise de Contratos</div>', unsafe_allow_html=True)
        st.stop()

    # -------------------------
    # Preparação dos dados
    # -------------------------
    hist = historico.copy()
    for col in ["fornecedor", "cnpj", "valor_total", "vigencia", "status", "risco", "score", "contrato_assinado", "modelo_ia", "tipo_origem", "arquivo", "data_analise"]:
        if col not in hist.columns:
            hist[col] = "Não informado"

    hist["risco_norm"] = hist["risco"].astype(str).str.upper().replace({"MEDIO": "MÉDIO"})
    hist["score_num"] = hist["score"].apply(as_float_score)
    hist["data_dt"] = pd.to_datetime(hist["data_analise"], format="%d/%m/%Y %H:%M", errors="coerce")

    total_hist = int(len(hist))
    qtd_alto = int((hist["risco_norm"] == "ALTO").sum())
    qtd_medio = int((hist["risco_norm"] == "MÉDIO").sum())
    qtd_baixo = int((hist["risco_norm"] == "BAIXO").sum())
    qtd_assinado = int(hist["contrato_assinado"].astype(str).str.upper().eq("SIM").sum())
    score_medio_hist = round(float(hist["score_num"].mean()), 1) if total_hist else 0

    # -------------------------
    # Indicadores executivos
    # -------------------------
    st.markdown('<div class="section-title">Painel do histórico</div>', unsafe_allow_html=True)
    h1, h2, h3, h4, h5 = st.columns(5)
    h1.markdown(render_metric("Contratos", total_hist), unsafe_allow_html=True)
    h2.markdown(render_metric("Risco alto", qtd_alto), unsafe_allow_html=True)
    h3.markdown(render_metric("Risco médio", qtd_medio), unsafe_allow_html=True)
    h4.markdown(render_metric("Risco baixo", qtd_baixo), unsafe_allow_html=True)
    h5.markdown(render_metric("Score médio", score_medio_hist), unsafe_allow_html=True)

    st.markdown(
        f"""
        <div class="executive-box">
            <h3>📚 Visão executiva do histórico</h3>
            <p>Existem <b>{total_hist}</b> análise(s) registrada(s), sendo <b>{qtd_baixo}</b> de risco baixo, <b>{qtd_medio}</b> de risco médio e <b>{qtd_alto}</b> de risco alto.</p>
            <p><b>{qtd_assinado}</b> contrato(s) possuem evidência de assinatura registrada no histórico.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # -------------------------
    # Filtros profissionais
    # -------------------------
    st.markdown('<div class="section-title">Filtros</div>', unsafe_allow_html=True)
    f1, f2, f3, f4 = st.columns([1.4, 1, 1, 1])

    with f1:
        busca = st.text_input(
            "Buscar",
            placeholder="Digite contraparte, CNPJ, arquivo, status ou modelo...",
        ).strip()

    riscos_disponiveis = [r for r in ["ALTO", "MÉDIO", "BAIXO"] if r in set(hist["risco_norm"].dropna().astype(str))]
    with f2:
        riscos_sel = st.multiselect(
            "Risco",
            options=riscos_disponiveis,
            default=riscos_disponiveis,
        )

    origens_disponiveis = sorted([x for x in hist["tipo_origem"].dropna().astype(str).unique() if x and x != "Não informado"])
    with f3:
        origens_sel = st.multiselect(
            "Origem",
            options=origens_disponiveis,
            default=origens_disponiveis,
        )

    assinaturas_disponiveis = sorted([x for x in hist["contrato_assinado"].dropna().astype(str).unique() if x])
    with f4:
        assinatura_sel = st.multiselect(
            "Assinatura",
            options=assinaturas_disponiveis,
            default=assinaturas_disponiveis,
        )

    f5, f6, f7, f8 = st.columns([1, 1, 1, 1])
    modelos_disponiveis = sorted([x for x in hist["modelo_ia"].dropna().astype(str).unique() if x])
    status_disponiveis = sorted([x for x in hist["status"].dropna().astype(str).unique() if x])

    with f5:
        modelos_sel = st.multiselect("Modelo IA", options=modelos_disponiveis, default=modelos_disponiveis)
    with f6:
        status_sel = st.multiselect("Status", options=status_disponiveis, default=status_disponiveis)
    with f7:
        score_min, score_max = st.slider("Score", 0, 100, (0, 100))
    with f8:
        ordenar_por = st.selectbox("Ordenar por", ["Mais recentes", "Maior score", "Menor score", "Risco", "Contraparte"])

    filtrado = hist.copy()

    if busca:
        busca_lower = busca.lower()
        cols_busca = ["fornecedor", "cnpj", "valor_total", "vigencia", "status", "risco", "modelo_ia", "tipo_origem", "arquivo", "data_analise"]
        mask = pd.Series(False, index=filtrado.index)
        for col in cols_busca:
            mask = mask | filtrado[col].astype(str).str.lower().str.contains(busca_lower, na=False)
        filtrado = filtrado[mask]

    if riscos_sel:
        filtrado = filtrado[filtrado["risco_norm"].isin(riscos_sel)]
    else:
        filtrado = filtrado.iloc[0:0]

    if origens_disponiveis and origens_sel:
        filtrado = filtrado[filtrado["tipo_origem"].astype(str).isin(origens_sel)]
    elif origens_disponiveis:
        filtrado = filtrado.iloc[0:0]

    if assinatura_sel:
        filtrado = filtrado[filtrado["contrato_assinado"].astype(str).isin(assinatura_sel)]
    else:
        filtrado = filtrado.iloc[0:0]

    if modelos_sel:
        filtrado = filtrado[filtrado["modelo_ia"].astype(str).isin(modelos_sel)]
    else:
        filtrado = filtrado.iloc[0:0]

    if status_sel:
        filtrado = filtrado[filtrado["status"].astype(str).isin(status_sel)]
    else:
        filtrado = filtrado.iloc[0:0]

    filtrado = filtrado[(filtrado["score_num"] >= score_min) & (filtrado["score_num"] <= score_max)]

    if ordenar_por == "Mais recentes":
        filtrado = filtrado.sort_values(["data_dt", "id"], ascending=[False, False], na_position="last")
    elif ordenar_por == "Maior score":
        filtrado = filtrado.sort_values("score_num", ascending=False)
    elif ordenar_por == "Menor score":
        filtrado = filtrado.sort_values("score_num", ascending=True)
    elif ordenar_por == "Risco":
        ordem_risco = {"ALTO": 1, "MÉDIO": 2, "BAIXO": 3}
        filtrado["ordem_risco"] = filtrado["risco_norm"].map(ordem_risco).fillna(9)
        filtrado = filtrado.sort_values(["ordem_risco", "score_num"], ascending=[True, True])
    elif ordenar_por == "Contraparte":
        filtrado = filtrado.sort_values("fornecedor", ascending=True)

    st.markdown(
        f'<div class="filter-note">🔎 Visualizando <b>{len(filtrado)}</b> de <b>{total_hist}</b> análise(s) do histórico.</div>',
        unsafe_allow_html=True,
    )

    # -------------------------
    # Exportação filtrada
    # -------------------------
    export_cols = [
        "id", "data_analise", "fornecedor", "cnpj", "valor_total", "vigencia",
        "status", "risco", "score", "contrato_assinado", "modelo_ia", "tipo_origem", "arquivo",
    ]
    export_df = filtrado[[c for c in export_cols if c in filtrado.columns]].copy()
    export_df = export_df.rename(columns={
        "id": "ID",
        "data_analise": "Data da análise",
        "fornecedor": "Contraparte",
        "cnpj": "CNPJ",
        "valor_total": "Valor total",
        "vigencia": "Vigência",
        "status": "Status",
        "risco": "Risco",
        "score": "Score",
        "contrato_assinado": "Assinado",
        "modelo_ia": "Modelo IA",
        "tipo_origem": "Origem",
        "arquivo": "Arquivos analisados",
    })

    excel_hist = gerar_excel_historico_profissional(export_df, total_geral=total_hist)

    d1, d2 = st.columns([3, 1])
    with d2:
        st.download_button(
            "📥 Baixar histórico filtrado",
            data=excel_hist,
            file_name=f"historico_nexus_contract_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # -------------------------
    # Visualização
    # -------------------------
    st.markdown('<div class="section-title">Resultado do histórico</div>', unsafe_allow_html=True)

    if filtrado.empty:
        st.warning("Nenhuma análise encontrada para os filtros selecionados.")
    else:
        tab_cards, tab_tabela, tab_auditoria = st.tabs(["📌 Cards executivos", "📋 Tabela executiva", "🧾 Auditoria técnica"])

        with tab_cards:
            for _, row in filtrado.head(25).iterrows():
                risco = normalize_risco(row.get("risco"))
                cor = risco_cor(risco)
                with st.container(border=True):
                    topo1, topo2 = st.columns([5, 1])
                    with topo1:
                        st.markdown(f"### 📄 {safe(row.get('fornecedor'))}", unsafe_allow_html=True)
                        st.caption(str(row.get("arquivo") or "Não informado"))
                    with topo2:
                        st.markdown(
                            f"""
                            <div style="text-align:center;background:{cor};color:white;
                            padding:10px 14px;border-radius:999px;font-weight:900;font-size:12px;">
                                {safe(risco)}
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )

                    cc1, cc2, cc3, cc4 = st.columns([1.2, 1.1, 1.2, .8])
                    cc1.markdown(f"**CNPJ**  \n{safe(row.get('cnpj'))}", unsafe_allow_html=True)
                    cc2.markdown(f"**Valor**  \n{safe(row.get('valor_total'))}", unsafe_allow_html=True)
                    cc3.markdown(f"**Status**  \n{safe(row.get('status'))}", unsafe_allow_html=True)
                    cc4.markdown(f"**Score**  \n{safe(row.get('score'))}", unsafe_allow_html=True)

                    cc5, cc6, cc7, cc8 = st.columns([1.4, 1, 1, 1])
                    cc5.markdown(f"**Vigência**  \n{safe(row.get('vigencia'))}", unsafe_allow_html=True)
                    cc6.markdown(f"**Assinado**  \n{safe(row.get('contrato_assinado'))}", unsafe_allow_html=True)
                    cc7.markdown(f"**Origem**  \n{safe(row.get('tipo_origem'))}", unsafe_allow_html=True)
                    cc8.markdown(f"**Modelo**  \n{safe(row.get('modelo_ia'))}", unsafe_allow_html=True)

                    b1, b2 = st.columns([4, 1])
                    b1.caption(f"Analisado em {row.get('data_analise') or 'Não informado'} • ID {row.get('id')}")
                    b2.download_button(
                        "📥 Excel",
                        data=gerar_excel_card_bytes(row),
                        file_name=f"analise_contrato_{row.get('id', 'historico')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key=f"download_historico_card_{row.get('id', id(row))}",
                    )

                    with st.expander("🔎 Abrir análise completa deste contrato", expanded=False):
                        render_analise_completa_historico(row)

            if len(filtrado) > 25:
                st.info("Exibindo os 25 primeiros registros filtrados. Use os filtros ou a tabela executiva para consultar os demais.")

        with tab_tabela:
            tabela = export_df.copy()
            st.dataframe(
                tabela,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Score": st.column_config.ProgressColumn("Score", min_value=0, max_value=100),
                    "Risco": st.column_config.TextColumn("Risco"),
                    "Valor total": st.column_config.TextColumn("Valor total"),
                    "Arquivos analisados": st.column_config.TextColumn("Arquivos analisados", width="large"),
                },
            )

        with tab_auditoria:
            auditoria_cols = [c for c in ["id", "data_criacao", "data_analise", "modelo_ia", "tipo_origem", "risco", "score", "status", "contrato_assinado", "arquivo"] if c in filtrado.columns]
            st.dataframe(
                filtrado[auditoria_cols].rename(columns={
                    "id": "ID",
                    "data_criacao": "Criado em",
                    "data_analise": "Data análise",
                    "modelo_ia": "Modelo IA",
                    "tipo_origem": "Origem",
                    "risco": "Risco",
                    "score": "Score",
                    "status": "Status",
                    "contrato_assinado": "Assinado",
                    "arquivo": "Arquivo",
                }),
                use_container_width=True,
                hide_index=True,
            )

    st.markdown('<div class="footer">NEXUS CONTRACT AI • Suprimentos • Análise de Contratos</div>', unsafe_allow_html=True)
    st.stop()

st.markdown('<div class="footer">NEXUS CONTRACT AI • Suprimentos • Análise de Contratos</div>', unsafe_allow_html=True)
