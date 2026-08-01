# Motor de Auditoria por Evidências V4

Esta versão mantém o layout e o fluxo do Auditor de Contratos, mas altera a regra de consolidação final.

## O que mudou

- Todo campo principal precisa de arquivo, página ou seção, trecho de evidência e confiança.
- A análise continua em duas passagens no Gemini e recebe uma terceira validação determinística em Python.
- Os cards mostram fonte, status da evidência e confiança.
- Nova aba **Evidências** com cobertura dos campos, cobertura das páginas, conflitos e matriz completa.
- Local da prestação não é preenchido apenas pelo endereço de uma das partes.
- Prazo indeterminado aparece tecnicamente como `Início DD/MM/AAAA até 31/12/9999`, sem presumir que o contrato está operacionalmente ativo.
- Permanência mínima não é confundida com prazo final.
- Valor global, implantação, mensalidade, tarifa unitária e percentual ficam separados.
- Implantação não é somada à mensalidade como valor global.
- Total da vigência só é calculado quando prazo e bases financeiras estão confirmados.
- Signatários só entram na tabela quando existe evidência específica no bloco de assinaturas ou certificado.
- Data da assinatura, conclusão DocuSign e reconhecimento de firma permanecem separadas.
- Checklist e pendências sem evidência objetiva são descartados.
- Score é calculado pelo sistema com base em cobertura, páginas processadas, conflitos e pendências documentadas.
- Resumo e parecer são reconstruídos apenas com informações validadas.

## Arquivos novos

- `auditor_evidencias.py`: motor determinístico V4.
- `teste_motor_evidencias_v4.py`: teste automatizado das regras financeiras, vigência e assinaturas.

## Teste local

```powershell
cd C:\IA_CONTRATOS\nexus_contract_ai
.\.venv\Scripts\python.exe -m py_compile app.py database.py auditor_evidencias.py
.\.venv\Scripts\python.exe teste_motor_evidencias_v4.py
.\.venv\Scripts\python.exe -m streamlit run app.py
```

## Atualização no GitHub

```powershell
git add app.py auditor_evidencias.py teste_motor_evidencias_v4.py MOTOR_EVIDENCIAS_V4.md README.md requirements.txt
git commit -m "Adiciona Motor de Auditoria por Evidencias V4"
git pull --rebase origin main
git push origin main
```

O arquivo `.env` não deve ser enviado ao GitHub. No Streamlit Cloud, mantenha a chave em **Settings > Secrets**.
