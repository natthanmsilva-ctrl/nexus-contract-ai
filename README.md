# Auditor de Contratos - Grupo SBF

Versão final corrigida:
- Assistente IA sem HTML/código aparecendo nas respostas.
- Perguntas rápidas renderizadas em cards visuais.
- Histórico com filtros limpos e alinhados.
- Files API via google-genai >= 2.8.0.
- DOC/DOCX convertido para TXT estruturado antes do upload.
- Excel com novos campos de valor, datas DocuSign e assinantes.
- Sem .env, sem banco local e sem arquivos sensíveis no pacote.

Como rodar:

```bash
pip install -r requirements.txt
streamlit run app.py
```

Crie seu .env localmente:

```env
GEMINI_API_KEY=SUA_CHAVE_AQUI
```
