# Auditor de Contratos — Ajustes V5

## Regras implementadas

- Mantém o período técnico com término em `31/12/9999` para contratos por prazo indeterminado.
- Separa tipo de vigência, status contratual e situação operacional.
- Mantém implantação, mensalidade fixa, tarifas variáveis e valores isentos em naturezas distintas.
- Nunca converte valor ausente em `R$ 0,00`.
- Reconstrói checklist e indicadores de pendências a partir das evidências consolidadas.
- Separa confiança da extração de risco contratual.
- Exibe `Páginas processadas` em vez de uma cobertura enganosa.
- Extrai integralmente tabelas comerciais, sem `head`, limite ou resumo de linhas.
- Mostra quantidade encontrada, exibida e cobertura da tabela comercial.

## Validações locais

```powershell
python -m py_compile app.py auditor_evidencias.py extrator_tabela_comercial.py database.py
python teste_motor_evidencias_v4.py
python teste_tabela_comercial_completa.py
python -m streamlit run app.py
```
