# Acompanhamento de Carreira Técnica — OLFATI

Site de acompanhamento do plano de carreira dos Técnicos em Manutenção da OLFATI.

## Como continuar este projeto no Claude Code

1. Baixe e descompacte esta pasta no seu computador.
2. Abra um terminal dentro da pasta.
3. Rode `claude` (Claude Code precisa estar instalado: https://docs.claude.com/claude-code).
4. Peça para o Claude Code continuar o desenvolvimento — por exemplo: integrar com o Google Sheets (planilha de atendimentos/infrações), adicionar um backend com banco de dados, ajustar o layout, etc.

## Estrutura atual

- `index.html` — aplicação completa (frontend), single-file: HTML + CSS + JS.
  - Login (admin / 8 técnicos)
  - Importação de planilha do Auvo (.xlsx/.csv) com mapeamento de colunas
  - Cálculo automático de produtividade, bônus, penalidades, remuneração estimada
  - Progresso de promoção (Júnior → Pleno → Sênior)
  - Painel do técnico com acompanhamento diário
  - Painel do admin com visão consolidada, exportação CSV e configurações

## Limitação atual

O app roda inteiramente no navegador, sem backend: os dados ficam em memória e se perdem ao recarregar a página. Os próximos passos combinados são:
1. Conectar a planilha do Google Sheets (publicada como CSV) para leitura automática dos atendimentos e infrações já calculados.
2. Avaliar se vale migrar para um backend real (Node/Express + banco de dados) para persistência e login seguro por técnico, como descrito no prompt original (ver histórico da conversa).

## Publicação

Hoje o `index.html` pode ser publicado como site estático (ex: arrastando em app.netlify.com/drop). Não requer build.
