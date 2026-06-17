"""
Teste de _subir_arquivos_originais_gemini usando google-genai >= 2.x (novo SDK).

Uso:
    python teste_upload_gemini_v2.py <api_key> <arquivo1> [arquivo2 ...]

Exemplo:
    python teste_upload_gemini_v2.py AIza... contrato.pdf aditivo.docx
"""

import os
import sys
import tempfile
import time
from pathlib import Path


# ---------------------------------------------------------------------------
# Helpers (iguais ao app.py)
# ---------------------------------------------------------------------------

def _mime_type_arquivo(nome_arquivo: str) -> str:
    nome = str(nome_arquivo or "").lower()
    if nome.endswith(".pdf"):
        return "application/pdf"
    if nome.endswith(".docx"):
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    if nome.endswith(".doc"):
        return "application/msword"
    return "application/octet-stream"


class _ArquivoSimulado:
    """Simula o objeto de upload do Streamlit a partir de um caminho real."""

    def __init__(self, caminho: str):
        self.name = Path(caminho).name
        self._buf = open(caminho, "rb")

    def seek(self, pos):
        self._buf.seek(pos)

    def read(self):
        return self._buf.read()

    def close(self):
        self._buf.close()


# ---------------------------------------------------------------------------
# Código original do app.py — adaptado para google-genai >= 2.x
#
# Mudanças em relação ao app.py original (google-generativeai 0.8.x):
#   - Recebe `client` (google.genai.Client) em vez do módulo `genai`
#   - genai.upload_file(path=, ...)   →  client.files.upload(file=, config={...})
#   - genai.get_file(name)            →  client.files.get(name=name)
# ---------------------------------------------------------------------------

def _subir_arquivos_originais_gemini(client, arquivos_originais) -> tuple:
    """Salva temporariamente e envia os arquivos originais para o Gemini Files API."""
    uploaded_files = []
    temp_paths = []

    if not arquivos_originais:
        return uploaded_files, temp_paths

    for arquivo in arquivos_originais:
        try:
            nome = getattr(arquivo, "name", "documento")
            arquivo.seek(0)
            conteudo = arquivo.read()
            suffix = Path(nome).suffix or ".bin"

            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
            tmp.write(conteudo)
            tmp.flush()
            tmp.close()
            temp_paths.append(tmp.name)

            # ANTES (google-generativeai):
            #   uploaded = genai.upload_file(
            #       path=tmp.name,
            #       mime_type=_mime_type_arquivo(nome),
            #       display_name=nome,
            #   )
            uploaded = client.files.upload(
                file=tmp.name,
                config={
                    "mime_type": _mime_type_arquivo(nome),
                    "display_name": nome,
                },
            )

            # Alguns arquivos podem ficar em processamento por poucos segundos.
            for _ in range(45):
                state = getattr(getattr(uploaded, "state", None), "name", "")
                if state and state.upper() == "PROCESSING":
                    time.sleep(1)
                    # ANTES: uploaded = genai.get_file(uploaded.name)
                    uploaded = client.files.get(name=uploaded.name)
                else:
                    break

            uploaded_files.append(uploaded)
            print(f"  OK: '{nome}' → {uploaded.name}")

        except Exception as erro:
            print(
                f"  ERRO ao enviar '{getattr(arquivo, 'name', 'arquivo')}': {erro}"
            )

    return uploaded_files, temp_paths


def _limpar_uploads_gemini(client, uploaded_files: list, temp_paths: list) -> None:
    """Remove temporários locais e tenta remover arquivos da área temporária da API."""
    for uploaded in uploaded_files or []:
        try:
            if getattr(uploaded, "name", None):
                # ANTES: genai.delete_file(uploaded.name)
                client.files.delete(name=uploaded.name)
        except Exception:
            pass

    for caminho in temp_paths or []:
        try:
            os.remove(caminho)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    api_key = sys.argv[1]
    caminhos = sys.argv[2:]

    try:
        from google import genai
    except ImportError:
        print("ERRO: instale o pacote: pip install --upgrade google-genai")
        sys.exit(1)

    import importlib.metadata
    ver = importlib.metadata.version("google-genai")
    print(f"google-genai version: {ver}\n")

    client = genai.Client(api_key=api_key)

    arquivos = []
    for c in caminhos:
        if not Path(c).exists():
            print(f"Arquivo não encontrado: {c}")
            sys.exit(1)
        arquivos.append(_ArquivoSimulado(c))

    print(f"Enviando {len(arquivos)} arquivo(s) para o Gemini Files API...\n")
    uploaded_files, temp_paths = _subir_arquivos_originais_gemini(client, arquivos)

    for a in arquivos:
        a.close()

    print(f"\nResultado: {len(uploaded_files)} arquivo(s) enviados com sucesso.")

    if uploaded_files:
        resp = input("\nDeseja deletar os arquivos da API agora? [s/N] ").strip().lower()
        if resp == "s":
            _limpar_uploads_gemini(client, uploaded_files, temp_paths)
            print("Limpeza concluída.")
        else:
            print("Arquivos mantidos na API.")
            for u in uploaded_files:
                print(f"  {u.name}")


if __name__ == "__main__":
    main()
