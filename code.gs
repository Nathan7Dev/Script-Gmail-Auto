
function verificarNCsSemResposta() {
  const remetentesNC = ["email.autorizado@dominio.com"];
  const assuntoInicio = "NÃO CONFORMIDADE: NC - ";
  const tempoMinimoHoras = 1;
 
  const labelReenviado = GmailApp.getUserLabelByName("NC_Reenviada") || GmailApp.createLabel("NC_Reenviada");
  const labelRespondido = GmailApp.getUserLabelByName("NC_Respondida") || GmailApp.createLabel("NC_Respondida");
  const labelCancelada = GmailApp.getUserLabelByName("NC_Cancelada") || GmailApp.createLabel("NC_Cancelada");
 
  const dataInicio = new Date("2025-08-14T00:00:00");
  Logger.log("Iniciando verificação das NCs a partir de " + dataInicio.toLocaleString());
 
  const query = 'subject:"${assuntoInicio}"';
  const threads = GmailApp.search(query);
  Logger.log("Total de threads encontradas com assunto contendo '" + assuntoInicio + "': " + threads.length);
 
  let totalVerificados = 0;
  let ignoradasCount = 0;
  let ignoradasMotivo = {};
  let reenviadas = [];
  let ignoradasPorData = 0;
 
  const agora = new Date();
 
  threads.forEach(thread => {
    totalVerificados++;
    const assunto = thread.getFirstMessageSubject();
    const labels = thread.getLabels().map(label => label.getName());
    const mensagens = thread.getMessages();
    const primeira = mensagens[0];
    const ultima = mensagens[mensagens.length - 1];
    const remetenteOriginal = primeira.getFrom().toLowerCase();
 
    if (ultima.getTime() < dataInicio.getTime()) {
      ignoradasCount++;
      ignoradasPorData++;
      ignoradasMotivo["E-mail anterior à data de início"] = (ignoradasMotivo["E-mail anterior à data de início"] || 0) + 1;
      return;
    }
 
    if (!assunto.startsWith(assuntoInicio)) {
      ignoradasCount++;
      ignoradasMotivo["Assunto não inicia com padrão esperado"] = (ignoradasMotivo["Assunto não inicia com padrão esperado"] || 0) + 1;
      return;
    }
 
    if (!remetentesNC.some(r => remetenteOriginal.includes(r))) {
      ignoradasCount++;
      ignoradasMotivo["Remetente não autorizado"] = (ignoradasMotivo["Remetente não autorizado"] || 0) + 1;
      return;
    }
 
    const diffHoras = (agora.getTime() - ultima.getTime()) / (1000 * 60 * 60);
    if (diffHoras < tempoMinimoHoras) {
      ignoradasCount++;
      ignoradasMotivo["Tempo mínimo não atingido"] = (ignoradasMotivo["Tempo mínimo não atingido"] || 0) + 1;
      return;
    }
 
    const comandoFechamento = "#fecharnc";
    const houveFechamentoManual = mensagens.slice(1).some(msg => {
      const from = msg.getFrom().toLowerCase();
      const corpo = msg.getPlainBody().toLowerCase();
      return remetentesNC.some(r => from.includes(r)) && corpo.includes(comandoFechamento);
    });
 
    if (houveFechamentoManual) {
      thread.addLabel(labelRespondido);
      thread.addLabel(labelCancelada);
      if (labels.includes("NC_Reenviada")) thread.removeLabel(labelReenviado);
      ignoradasCount++;
      ignoradasMotivo["NC encerrada manualmente"] = (ignoradasMotivo["NC encerrada manualmente"] || 0) + 1;
      return;
    }
 
    const houveResposta = mensagens.slice(1).some(msg => {
      const from = msg.getFrom().toLowerCase();
      const toAndCc = (msg.getTo() + "," + msg.getCc()).toLowerCase();
      return (!remetentesNC.some(r => from.includes(r))) && (toAndCc.includes(from));
    });
 
    if (houveResposta) {
      thread.addLabel(labelRespondido);
      if (labels.includes("NC_Reenviada")) thread.removeLabel(labelReenviado);
      ignoradasCount++;
      ignoradasMotivo["Respondida por destinatário ou cópia"] = (ignoradasMotivo["Respondida por destinatário ou cópia"] || 0) + 1;
      return;
    }
 
    const mensagem = "⚠️ Reforço automático: Ainda não identificamos resposta a esta Não Conformidade.\n\nPor favor, verificar o histórico deste e-mail e responder diretamente por aqui conforme orientações.";
    thread.replyAll(mensagem);
    if (!labels.includes("NC_Reenviada")) thread.addLabel(labelReenviado);
    reenviadas.push(assunto);
  });
 
  Logger.log("===== Resumo da execução =====");
  Logger.log("Total de e-mails verificados: " + totalVerificados);
  Logger.log("Total de e-mails ignorados: " + ignoradasCount);
  for (const motivo in ignoradasMotivo) {
    Logger.log("- " + motivo + ": " + ignoradasMotivo[motivo] + " e-mail(s)");
  }
  Logger.log("E-mails ignorados apenas por serem anteriores a " + dataInicio.toLocaleString() + ": " + ignoradasPorData + " e-mail(s)");
 
  if (reenviadas.length > 0) {
    Logger.log("E-mails reenviados (" + reenviadas.length + "):");
    reenviadas.forEach(a => Logger.log("- " + a));
  } else {
    Logger.log("Nenhum reenvio foi necessário nesta execução.");
  }
 
  Logger.log("Verificação concluída.");
}