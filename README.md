# Script-Gmail-Auto

Automação em Google Apps Script para gerenciar e-mails de Não Conformidade (NC) no Gmail. O script monitora mensagens com assunto padrão, verifica se há resposta, reenvia alertas após 48 horas sem retorno, e marca os e-mails com labels para organizar o fluxo.

## Funcionalidades

- Verifica se a NC foi enviada por remetentes autorizados.
- Avalia o tempo desde a última mensagem para decidir reenviar o alerta (48h).
- Identifica respostas por terceiros e marca como respondida.
- Permite encerramento manual via comando especial no corpo do e-mail.
- Aplica labels para controlar o status: reenviada, respondida, cancelada.
- Log detalhado para acompanhar as ações realizadas.

## Como usar

1. Configure o script no Google Apps Script vinculado à sua conta Gmail.
2. Atualize os remetentes autorizados e labels conforme sua necessidade.
3. Agende a execução periódica (ex.: a cada 48h no horário desejado).
4. Acompanhe o log para verificar as ações do script.

## Estrutura

- `verificarNCsSemResposta.gs`: Script principal para gerenciamento das NCs.
- Labels usados: `NC_Reenviada`, `NC_Respondida`, `NC_Cancelada`.
