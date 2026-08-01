# Auditor de Contratos - Grupo SBF

Aplicação Streamlit para análise profissional de contratos Projuris ou Ariba, com leitura de PDF/DOCX, Gemini Files API, histórico SQLite e exportação Excel.

## Versão atual — Motor de Auditoria por Evidências V5

O fluxo atual possui três camadas:

1. Extração completa dos documentos originais pelo Gemini.
2. Segunda revisão independente do JSON contra os arquivos.
3. Consolidação determinística em Python, impedindo que cards, valores, assinaturas e parecer sejam confirmados sem evidência.

Principais recursos:

- matriz de evidências por campo;
- fonte, página/seção, trecho e confiança em cada card;
- nova aba de auditoria de evidências;
- separação entre valor global, implantação, mensalidade e tarifas variáveis;
- cálculo de vigência somente com bases confirmadas;
- assinatura física, DocuSign e reconhecimento de firma separados;
- checklist, pendências, confiança da extração, risco, resumo e parecer baseados em evidências;
- extração integral de tabelas comerciais, sem limitar ou resumir linhas;
- indicadores de itens encontrados, itens exibidos e cobertura da tabela comercial;
- histórico e Excel completos.

Consulte `MOTOR_EVIDENCIAS_V4.md` e `AJUSTES_V5_TABELA_COMPLETA.md` para as regras e o passo a passo de teste.

## Configuração local

Crie um arquivo `.env` apenas na sua máquina:

```env
GEMINI_API_KEY=SUA_CHAVE_AQUI
```

No Streamlit Community Cloud, configure a chave em **App settings > Secrets**:

```toml
GEMINI_API_KEY = "SUA_CHAVE_AQUI"
```

Nunca envie `.env`, banco local, `.venv` ou arquivos com segredos ao GitHub.
