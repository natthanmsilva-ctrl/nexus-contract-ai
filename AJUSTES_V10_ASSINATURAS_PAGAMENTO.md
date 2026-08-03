# Ajustes V10 — Condição de Pagamento e Assinaturas

## Escopo cirúrgico

A V10 preserva integralmente as correções já aprovadas de:

- 21 itens comerciais exibidos;
- 17 tarifas/condições variáveis;
- ausência de duplicidades na tabela;
- Local de Prestação não confundido com foro, sede ou endereço cadastral.

## Condição de Pagamento em Dias

- A busca passa a priorizar, em todas as páginas, cláusulas financeiras explícitas de fatura e vencimento.
- O padrão “vencimento até o dia 15 do mês subsequente” retorna `15DD`.
- Prazos operacionais de entrega de informações, relatórios, atendimento ou execução são rejeitados.
- No contrato SBF, a fonte correta é a página 26, Anexo II, itens 3 e 4.

## Assinaturas físicas

- Quando a matriz de evidências confirma os nomes no bloco de assinaturas, a V10 recupera os registros individuais mesmo que a resposta da IA omita fonte/evidência em cada linha.
- A consolidação volta a usar uma única fonte final para cards, aba Assinaturas, triagem, checklist, pendências e Excel.
- Data geral do instrumento e data individual da assinatura permanecem separadas.
- Reconhecimento de firma permanece separado e é atribuído somente aos nomes indicados no selo/cartório.
- Segunda testemunha em branco permanece como ponto de atenção, sem apagar a validade das assinaturas das partes.

## Resultado esperado no SBF

- Condição de pagamento: `15DD`;
- Contrato assinado: `Sim`;
- 5 signatários identificados;
- 4 representantes das partes e 1 testemunha;
- reconhecimento de firma em 31/10/2023 para os dois representantes indicados no selo;
- campos não localizados: somente `Local de Prestação`.
