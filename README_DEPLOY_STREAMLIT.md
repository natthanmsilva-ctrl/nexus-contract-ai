# Deploy no Streamlit Cloud

1. Suba estes arquivos para o GitHub sem `.env`.
2. No Streamlit Cloud, configure o Secret:

```toml
GEMINI_API_KEY = "SUA_CHAVE_AQUI"
```

3. Main file path: `app.py`.

Esta versão usa `google-genai` com `client.files.upload` para a Files API e lê a chave tanto do `.env` local quanto do `st.secrets` no Streamlit Cloud.

O arquivo `packages.txt` instala `poppler-utils` e Tesseract em português no ambiente publicado, necessários para OCR de PDFs digitalizados.
