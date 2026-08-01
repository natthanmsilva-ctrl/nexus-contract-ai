# Ajustes V7 — Correção da tabela comercial SBF

## Correções aplicadas

1. **OCR da página comercial**
   - O extrator agora tolera resíduos curtos após o valor, como `R$0,60 5`.
   - A faixa `De 35.001 a 70.000 acionistas — R$ 0,60` deixa de ser descartada.
   - O teste usa o texto real gerado pelo OCR da página 25.

2. **Reconciliação IA + parser documental**
   - O parser documental passou a ser a fonte prioritária.
   - Descrições equivalentes da IA e do OCR são reconciliadas semanticamente.
   - A tabela do contrato de teste passa de 41 linhas duplicadas para 21 linhas únicas.
   - A IA pode completar campos ausentes, mas não substitui a classificação e os valores documentais confirmados.

3. **Cálculos financeiros**
   - Implantação: R$ 3.000,00.
   - Mensalidade fixa: R$ 4.000,00/mês.
   - Tarifas/condições variáveis: 17.
   - Serviços isentos: 2.
   - A vigência indeterminada não é projetada como 24 meses; o total da vigência permanece não calculável com precisão.

4. **Indicadores da tabela**
   - Itens encontrados: 21.
   - Itens exibidos: 21.
   - Cobertura: 100%.
   - Foi adicionado indicador interno de divergência de quantidade.

5. **Aditivos**
   - Quando nenhum aditivo é identificado, o valor dos aditivos aparece como `Não aplicável`.

6. **Painel de processamento**
   - Ao concluir, o painel deixa de bloquear a tela, mas permanece visível com tempos e status finais.

7. **Segurança do pacote entregue**
   - O ZIP limpo não contém `.env`, chave Gemini, banco SQLite, `.git`, `.venv`, cache Python nem arquivo de backup legado.
   - A chave que estava no ZIP anterior deve ser revogada e substituída.

## Testes executados

- `python teste_tabela_comercial_completa.py`
- `python teste_motor_evidencias_v4.py`
- Compilação de `app.py` e `extrator_tabela_comercial.py` com `py_compile`.
