# Ajustes V9 — Pagamento, Reconciliação Comercial e Local de Prestação

## 1. Condição de pagamento em dias

- O campo `condicao_pagamento_dias` passa a ser exibido em formato executivo:
  - `15DD`
  - `30DD`
  - `60DD`
  - `90DD`
- A frase detalhada do contrato permanece no campo `forma_pagamento`.
- A extração é restrita ao contexto de pagamento, fatura, nota fiscal e vencimento para evitar captura de aviso prévio, vigência ou prazo de cura.
- No contrato SBF usado no teste, o resultado esperado é `15DD`.

## 2. Reconciliação da tabela comercial

- As expressões abaixo passam a ser tratadas como equivalentes:
  - `Boletim emitido` = `Emissão de boletim`
  - `Boletim efetivado` = `Efetivação de boletim`
- Resultado esperado para o Anexo II do contrato SBF:
  - 21 itens no documento
  - 21 itens exibidos
  - 17 tarifas/condições variáveis
  - 0 divergência de quantidade
- Divergências remanescentes reduzem Score e Confiança e geram ponto de atenção.

## 3. Local de prestação

- Foro eleito, comarca competente, sede, endereço cadastral e preâmbulo não confirmam local de execução.
- O campo somente é confirmado quando a evidência contém linguagem explícita de prestação ou execução dos serviços no local.
- Para o contrato SBF, o local de prestação deve permanecer `Não localizado com segurança`.

## Testes adicionados/atualizados

- `teste_tabela_comercial_completa.py`
- `teste_regressao_v9_sbf.py`
- `teste_condicao_pagamento_v9.py`
