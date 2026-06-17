# Migração: upload de arquivos para o Gemini Files API

## Problema

A função `_subir_arquivos_originais_gemini` em `app.py` falhava ao tentar enviar arquivos para o Gemini com o erro:

```
HttpError 400 when requesting
https://generativelanguage.googleapis.com/$discovery/rest?version=v1beta&key=AQ.Ab8RN6...
returned "API key not valid. Please pass a valid API key."
```

O erro ocorria **mesmo com a chave API funcionando** para geração de texto.

---

## Causa raiz

Duas libs do Google coexistiam no ambiente:

| Lib | Versão instalada | Versão mais recente |
|---|---|---|
| `google-generativeai` (legada) | 0.8.5 | 0.8.6 |
| `google-genai` (nova) | 1.50.1 | **2.8.0** |

O `app.py` usava `google.generativeai` (legada) e chamava `genai.upload_file()`. Essa função depende do **Google Discovery Service** (`$discovery/rest`) para localizar o endpoint de upload. O Discovery Service rejeita chaves no formato `AQ.Ab8RN6...` (geradas via Vertex AI / OAuth), aceitando apenas chaves no formato `AIzaSy...` do Google AI Studio.

Já `genai.list_models()` funcionava porque usava internamente um caminho que não passa pelo Discovery Service.

**Diagnóstico chave:** o erro apontava para `$discovery/rest`, não para o endpoint de upload em si.

---

## Solução

Migrar para o cliente da nova SDK (`google-genai >= 2.x`), que não usa o Discovery Service.

### Atualizar as libs

```bash
python3 -m pip install --upgrade google-generativeai google-genai
```

### Três mudanças no código

```python
# ANTES — google-generativeai (legada)
genai.upload_file(
    path=tmp.name,
    mime_type=_mime_type_arquivo(nome),
    display_name=nome,
)
genai.get_file(uploaded.name)
genai.delete_file(uploaded.name)

# DEPOIS — google-genai >= 2.x
client.files.upload(
    file=tmp.name,
    config={
        "mime_type": _mime_type_arquivo(nome),
        "display_name": nome,
    },
)
client.files.get(name=uploaded.name)
client.files.delete(name=uploaded.name)
```

### Como criar o client

```python
from google import genai

client = genai.Client(api_key=api_key)
```

---

## Arquivos de teste

### `teste_upload_gemini.py` — diagnóstico + novo SDK

Inclui diagnóstico de versões e formato de chave. Usa `google-genai` via `genai.Client`.

```
python teste_upload_gemini.py <api_key> arquivo.pdf
```

Saída de diagnóstico esperada:

```
=== DIAGNÓSTICO ===
  google-generativeai version : 0.8.6
  google-genai version        : 2.8.0
  Formato da chave            : OK (AIzaSy...)   ← ou INCOMUM se for AQ.Ab8RN6...
  Testando list_models()...   OK (55 modelos)
===================
```

### `teste_upload_gemini_v2.py` — código do app.py adaptado para google-genai 2.x

Contém a função `_subir_arquivos_originais_gemini` idêntica ao `app.py`, com apenas as 3 linhas de API trocadas (comentadas lado a lado). Serve de referência direta para aplicar o fix no `app.py`.

```
python teste_upload_gemini_v2.py <api_key> arquivo.pdf
```

---

## Aplicar o fix no app.py

Em `app.py`, a função `analisar_gemini` importa e configura a lib legada:

```python
def analisar_gemini(...):
    import google.generativeai as genai
    genai.configure(api_key=api_key)
    ...
    uploaded_files, temp_paths = _subir_arquivos_originais_gemini(genai, arquivos_originais)
```

Substituir por:

```python
def analisar_gemini(...):
    from google import genai as genai_new
    import google.generativeai as genai
    genai.configure(api_key=api_key)

    client = genai_new.Client(api_key=api_key)
    ...
    uploaded_files, temp_paths = _subir_arquivos_originais_gemini(client, arquivos_originais)
```

E aplicar as 3 trocas de método em `_subir_arquivos_originais_gemini` e `_limpar_uploads_gemini` conforme a tabela acima.
