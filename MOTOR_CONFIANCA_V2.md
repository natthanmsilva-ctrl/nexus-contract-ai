# Motor de Confiança V2

Esta versão mantém o visual e o fluxo atual do Auditor de Contratos, mas reforça a análise documental.

## Alterações principais

- Análise Gemini em duas passagens: extração e auditoria independente.
- Evidência obrigatória por campo: arquivo, página, trecho e confiança.
- Separação rígida entre valor global, mensalidade, implantação, tarifa unitária e percentual.
- Datas separadas: contrato, assinatura, conclusão DocuSign e reconhecimento de firma.
- Signatários, pendências, checklist, itens e aditivos sem evidência são descartados da consolidação.
- Score e risco são recalculados somente a partir de pendências com evidência.
- Resumo e parecer são reconstruídos usando apenas fatos validados.
- Nova aba `Auditoria de Campos` no relatório Excel.

## Teste local

```powershell
cd C:\IA_CONTRATOS\nexus_contract_ai
.\.venv\Scripts\python.exe -m py_compile app.py database.py
.\.venv\Scripts\python.exe -m streamlit run app.py
```

Faça uma análise nova. Registros antigos do histórico não possuem a nova matriz de evidências e continuam sendo exibidos em modo de compatibilidade.
