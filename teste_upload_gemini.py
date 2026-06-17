"""
Teste isolado de _subir_arquivos_originais_gemini.

Uso:
    python teste_upload_gemini.py <api_key> <arquivo1> [arquivo2 ...]

Exemplo:
    python teste_upload_gemini.py AIza... contrato.pdf aditivo.docx
"""

import os
import sys
import tempfile
import time
from pathlib import Path


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
        self._caminho = caminho
        self._buf = open(caminho, "rb")

    def seek(self, pos):
        self._buf.seek(pos)

    def read(self):
        return self._buf.read()

    def close(self):
        self._buf.close()


def _subir_arquivos_originais_gemini(client, arquivos_originais) -> tuple:
    """Usa google-genai (novo SDK) — não depende do Discovery Service."""
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

            mime = _mime_type_arquivo(nome)
            print(f"  Enviando '{nome}' ({len(conteudo)} bytes, mime={mime})...")

            uploaded = client.files.upload(
                file=tmp.name,
                config={"mime_type": mime, "display_name": nome},
            )

            for i in range(45):
                state = getattr(getattr(uploaded, "state", None), "name", "")
                if state and state.upper() == "PROCESSING":
                    print(f"    [{i+1}s] ainda processando...")
                    time.sleep(1)
                    uploaded = client.files.get(name=uploaded.name)
                else:
                    break

            print(f"  OK: name={uploaded.name}, state={getattr(getattr(uploaded, 'state', None), 'name', '?')}")
            uploaded_files.append(uploaded)

        except Exception as erro:
            print(f"  ERRO ao enviar '{getattr(arquivo, 'name', 'arquivo')}': {erro}")

    return uploaded_files, temp_paths


def _limpar(client, uploaded_files, temp_paths):
    for uploaded in uploaded_files or []:
        try:
            if getattr(uploaded, "name", None):
                client.files.delete(name=uploaded.name)
                print(f"  Deletado da API: {uploaded.name}")
        except Exception as e:
            print(f"  Falha ao deletar da API: {e}")

    for caminho in temp_paths or []:
        try:
            os.remove(caminho)
        except Exception:
            pass


def _diagnostico(genai, api_key: str):
    print("=== DIAGNÓSTICO ===")

    import importlib.metadata
    try:
        ver = importlib.metadata.version("google-generativeai")
        print(f"  google-generativeai version : {ver}")
    except Exception:
        print("  google-generativeai          : não encontrado")
    try:
        ver2 = importlib.metadata.version("google-genai")
        print(f"  google-genai version        : {ver2}")
    except Exception:
        pass

    fmt = "AIzaSy"
    print(f"  Formato da chave            : {'OK (AIzaSy...)' if api_key.startswith(fmt) else f'INCOMUM (esperado AIzaSy..., recebido {api_key[:8]}...'}")

    print("  Testando list_models()...   ", end="", flush=True)
    try:
        modelos = list(genai.list_models())
        print(f"OK ({len(modelos)} modelos)")
    except Exception as e:
        print(f"ERRO: {e}")

    print("===================\n")


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    api_key = sys.argv[1]
    caminhos = sys.argv[2:]

    try:
        import google.generativeai as genai_legacy
        from google import genai
    except ImportError:
        print("ERRO: instale os pacotes: pip install google-generativeai google-genai")
        sys.exit(1)

    genai_legacy.configure(api_key=api_key)
    _diagnostico(genai_legacy, api_key)

    client = genai.Client(api_key=api_key)

    arquivos = []
    for c in caminhos:
        if not Path(c).exists():
            print(f"Arquivo não encontrado: {c}")
            sys.exit(1)
        arquivos.append(_ArquivoSimulado(c))

    print(f"\nEnviando {len(arquivos)} arquivo(s) para o Gemini Files API...\n")
    uploaded_files, temp_paths = _subir_arquivos_originais_gemini(client, arquivos)

    for a in arquivos:
        a.close()

    print(f"\nResultado: {len(uploaded_files)} arquivo(s) enviados com sucesso.\n")

    if uploaded_files:
        resp = input("Deseja deletar os arquivos da API agora? [s/N] ").strip().lower()
        if resp == "s":
            _limpar(client, uploaded_files, temp_paths)
            print("Limpeza concluída.")
        else:
            print("Arquivos mantidos na API.")
            for u in uploaded_files:
                print(f"  {u.name}")


if __name__ == "__main__":
    main()
