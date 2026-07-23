# Auditor de Contratos - Motor de Confiança V3

Ajustes desta versão:

1. Assinatura consolidada sem contradição entre card, tabela, alerta e triagem.
2. Data da assinatura separada da data de reconhecimento de firma.
3. Triagem atualizada após a análise profunda quando a assinatura é confirmada.
4. Pendência por ausência de e-mail de signatário removida, salvo obrigação contratual expressa.
5. Implantação/taxa única não é mais somada com mensalidade recorrente.
6. Valor global, valor mensal, valores pontuais e tarifas variáveis permanecem separados.
7. Excel de assinaturas ganhou a coluna Data do reconhecimento de firma.
8. Reaplicação final da matriz de evidências depois das regras legadas, evitando que pós-processamentos sobrescrevam dados validados.

## Teste local

```powershell
cd C:\IA_CONTRATOS\nexus_contract_ai
.\.venv\Scripts\python.exe -m py_compile app.py database.py
.\.venv\Scripts\python.exe -m streamlit run app.py
```

## Resultado esperado no contrato de escrituração

- Contrato assinado: Sim
- Data da assinatura: 26/10/2023
- Reconhecimento de firma: 31/10/2023
- Implantação: R$ 3.000,00
- Mensalidade: R$ 4.000,00/mês
- Os dois valores não são somados como valor total do contrato.
- Ausência de e-mail não aparece como pendência contratual.
- A triagem muda para Assinado: Sim após a validação profunda.
