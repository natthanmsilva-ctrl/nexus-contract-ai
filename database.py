"""
Database - Auditor de Contratos - Grupo SBF
Histórico SQLite com compatibilidade para bancos antigos e armazenamento completo da análise.
"""

from __future__ import annotations

import json
import sqlite3
from pathlib import Path
from typing import Any, Dict, Optional

import pandas as pd

DB_NAME = "nexus_contract.db"
DB_PATH = Path(__file__).resolve().parent / DB_NAME

CAMPOS_TABELA = {
    "id": "INTEGER PRIMARY KEY AUTOINCREMENT",
    "data_analise": "TEXT",
    "fornecedor": "TEXT",
    "cnpj": "TEXT",
    "valor_total": "TEXT",
    "vigencia": "TEXT",
    "status": "TEXT",
    "risco": "TEXT",
    "score": "INTEGER",
    "arquivo": "TEXT",
    "modelo_ia": "TEXT",
    "tipo_origem": "TEXT",
    "contrato_assinado": "TEXT",
    "resultado_json": "TEXT",
    "texto_extraido": "TEXT",
    "data_criacao": "TEXT DEFAULT CURRENT_TIMESTAMP",
}


def conectar() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def criar_banco() -> None:
    colunas_sql = ",\n            ".join(f"{campo} {tipo}" for campo, tipo in CAMPOS_TABELA.items())
    with conectar() as conn:
        cursor = conn.cursor()
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS contratos (
                {colunas_sql}
            )
        """)
        _garantir_colunas(cursor)
        conn.commit()


def _garantir_colunas(cursor: sqlite3.Cursor) -> None:
    cursor.execute("PRAGMA table_info(contratos)")
    colunas_existentes = {linha[1] for linha in cursor.fetchall()}
    for campo, tipo in CAMPOS_TABELA.items():
        if campo != "id" and campo not in colunas_existentes:
            cursor.execute(f"ALTER TABLE contratos ADD COLUMN {campo} {tipo}")


def _valor(resultado: Dict[str, Any], *chaves: str, padrao: str = "Não localizado") -> Any:
    for chave in chaves:
        valor = resultado.get(chave)
        if valor not in (None, "", [], {}, "Não localizado", "Não localizada"):
            return valor
    return padrao


def _score_seguro(valor: Any) -> int:
    try:
        return int(float(str(valor).replace(",", ".")))
    except (TypeError, ValueError):
        return 0


def _json_seguro(resultado: Dict[str, Any], texto_extraido: Optional[str] = None) -> str:
    payload = dict(resultado or {})
    if texto_extraido and not payload.get("texto_extraido"):
        payload["texto_extraido"] = texto_extraido
    try:
        return json.dumps(payload, ensure_ascii=False, default=str)
    except Exception:
        return json.dumps({"erro": "Não foi possível serializar a análise completa."}, ensure_ascii=False)


def salvar_analise(
    resultado: Dict[str, Any],
    arquivo: str,
    modelo_ia: Optional[str] = None,
    tipo_origem: Optional[str] = None,
    texto_extraido: Optional[str] = None,
    **kwargs: Any,
) -> int:
    """Salva análise e mantém também o JSON completo para relatórios futuros."""
    criar_banco()

    modelo_final = modelo_ia or kwargs.get("modelo") or kwargs.get("modelo_usado")
    origem_final = tipo_origem or kwargs.get("origem") or kwargs.get("origem_contrato")
    texto_final = texto_extraido or kwargs.get("texto") or resultado.get("texto_extraido") or ""

    dados = {
        "data_analise": _valor(resultado, "data_analise"),
        "fornecedor": _valor(resultado, "contraparte", "fornecedor"),
        "cnpj": _valor(resultado, "cnpj_contraparte", "cnpj"),
        "valor_total": _valor(resultado, "valor_contrato_original", "valor_total"),
        "vigencia": _valor(resultado, "vigencia_apos_assinatura", "vigencia"),
        "status": _valor(resultado, "status"),
        "risco": str(_valor(resultado, "risco", padrao="N/A")).upper().replace("MEDIO", "MÉDIO"),
        "score": _score_seguro(_valor(resultado, "score", padrao=0)),
        "arquivo": arquivo or "Não informado",
        "modelo_ia": modelo_final or _valor(resultado, "modelo_ia", padrao="Não informado"),
        "tipo_origem": origem_final or _valor(resultado, "tipo_origem", "origem_contrato", "origem", padrao="Não informado"),
        "contrato_assinado": _valor(resultado, "contrato_assinado", padrao="Não informado"),
        "resultado_json": _json_seguro(resultado, texto_final),
        "texto_extraido": texto_final,
    }

    campos = ", ".join(dados.keys())
    placeholders = ", ".join(["?"] * len(dados))
    with conectar() as conn:
        cursor = conn.cursor()
        cursor.execute(f"INSERT INTO contratos ({campos}) VALUES ({placeholders})", tuple(dados.values()))
        conn.commit()
        return int(cursor.lastrowid)


def listar_analises(limite: Optional[int] = None) -> pd.DataFrame:
    criar_banco()
    query = """
        SELECT
            id,
            data_analise,
            fornecedor,
            cnpj,
            valor_total,
            vigencia,
            status,
            risco,
            score,
            contrato_assinado,
            modelo_ia,
            tipo_origem,
            arquivo,
            resultado_json,
            texto_extraido,
            data_criacao
        FROM contratos
        ORDER BY id DESC
    """
    params: tuple[Any, ...] = ()
    if limite is not None:
        query += " LIMIT ?"
        params = (int(limite),)
    with conectar() as conn:
        return pd.read_sql_query(query, conn, params=params)


def buscar_analise_por_id(id_analise: int) -> pd.DataFrame:
    criar_banco()
    with conectar() as conn:
        return pd.read_sql_query("SELECT * FROM contratos WHERE id = ?", conn, params=(id_analise,))


def excluir_analise(id_analise: int) -> None:
    criar_banco()
    with conectar() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM contratos WHERE id = ?", (id_analise,))
        conn.commit()


def limpar_historico() -> None:
    criar_banco()
    with conectar() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM contratos")
        cursor.execute("DELETE FROM sqlite_sequence WHERE name = 'contratos'")
        conn.commit()


def resumo_dashboard() -> Dict[str, Any]:
    df = listar_analises()
    if df.empty:
        return {"total": 0, "alto": 0, "medio": 0, "baixo": 0, "score_medio": 0}
    risco = df["risco"].astype(str).str.upper().replace({"MEDIO": "MÉDIO"})
    score = pd.to_numeric(df["score"], errors="coerce").fillna(0)
    return {
        "total": int(len(df)),
        "alto": int((risco == "ALTO").sum()),
        "medio": int((risco == "MÉDIO").sum()),
        "baixo": int((risco == "BAIXO").sum()),
        "score_medio": round(float(score.mean()), 1),
    }
