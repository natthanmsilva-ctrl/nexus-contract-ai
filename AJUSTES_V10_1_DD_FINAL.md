# Ajustes V10.1 - Condição de Pagamento em DD

## Regra oficial

- O card **Condição de Pagamento em Dias** sempre retorna no padrão `NDD`.
- Exemplos aceitos na origem e normalizados para a saída oficial:
  - `30 dias` -> `30DD`;
  - `30DD` -> `30DD`;
  - `30 DDL` -> `30DD`;
  - `30DDF` -> `30DD`;
  - `vencimento no/até o dia 15 do mês subsequente` -> `15DD`.

## Proteção contra falso positivo

- A abreviação ou o prazo só é aceito quando está em contexto financeiro de pagamento, fatura, nota fiscal, remuneração ou vencimento.
- Prazos operacionais de relatório, atendimento, execução e disponibilização de informações continuam rejeitados.

## Contrato SBF de validação

- A página 18 contém prazo operacional de 5 dias úteis e não pode alimentar o card.
- A página 26 contém o vencimento até o dia 15 do mês subsequente e deve gerar `15DD`.
