function verificarNCsSemResposta() {
  const remetentesNC = ["supervisora.qualidade@telelaudo.com.br"];
  const assuntoInicio = "NÃO CONFORMIDADE: NC - "; // texto que deve ser o início do assunto
  const tempoMinimoHoras = 48;

  const labelReenviado = GmailApp.getUserLabelByName("NC_Reenviada") || GmailApp.createLabel("NC_Reenviada");
  const labelRespondido = GmailApp.getUserLabelByName("NC_Respondida") || GmailApp.createLabel("NC_Respondida");
  const labelCancelada = GmailApp.getUserLabelByName("NC_Cancelada") || GmailApp.createLabel("NC_Cancelada");

  Logger.log("Iniciando verificação das NCs");

  const query = `subject:"NÃO CONFORMIDADE: NC -"`; // busca geral, depois filtro exato abaixo
  const threads = GmailApp.search(query);

  Logger.log("Encontradas " + threads.length + " threads com o assunto contendo 'NÃO CONFORMIDADE: NC -'.");

  let ignoradasCount = 0;
  let ignoradasMotivo = {};
  let reenviadas = [];

  const agora = new Date();

  threads.forEach(thread => {
    const assunto = thread.getFirstMessageSubject();

    // Só processar se o assunto COMEÇA com o texto esperado
    if (!assunto.startsWith(assuntoInicio)) {
      // Ignorar se não começar com o padrão
      ignoradasCount++;
      ignoradasMotivo["Assunto não inicia com padrão esperado"] = (ignoradasMotivo["Assunto não inicia com padrão esperado"] || 0) + 1;
      return;
    }

    const labels = thread.getLabels().map(label => label.getName());

    if (labels.includes("NC_Respondida")) {
      ignoradasCount++;
      ignoradasMotivo["Já respondida"] = (ignoradasMotivo["Já respondida"] || 0) + 1;
      return;
    }

    const mensagens = thread.getMessages();
    const primeira = mensagens[0];
    const ultima = mensagens[mensagens.length - 1];

    const remetenteOriginal = primeira.getFrom().toLowerCase();

    if (!remetentesNC.some(r => remetenteOriginal.includes(r))) {
      ignoradasCount++;
      ignoradasMotivo["Remetente não autorizado"] = (ignoradasMotivo["Remetente não autorizado"] || 0) + 1;
      return;
    }

    const diffHoras = (agora.getTime() - ultima.getDate().getTime()) / (1000 * 60 * 60);

    if (diffHoras < tempoMinimoHoras) {
      ignoradasCount++;
      ignoradasMotivo["Tempo mínimo não atingido"] = (ignoradasMotivo["Tempo mínimo não atingido"] || 0) + 1;
      return;
    }

    const houveResposta = mensagens.slice(1).some(msg => {
      const from = msg.getFrom().toLowerCase();
      return !remetentesNC.some(r => from.includes(r));
    });

    if (houveResposta) {
      thread.addLabel(labelRespondido);
      if (labels.includes("NC_Reenviada")) {
        thread.removeLabel(labelReenviado);
      }
      ignoradasCount++;
      ignoradasMotivo["Respondida por terceiros"] = (ignoradasMotivo["Respondida por terceiros"] || 0) + 1;
      return;
    }

    const comandoFechamento = "#fecharnc";
    const houveFechamentoManual = mensagens.slice(1).some(msg => {
      const from = msg.getFrom().toLowerCase();
      const corpo = msg.getPlainBody().toLowerCase();
      const ehZara = remetentesNC.some(r => from.includes(r));
      return ehZara && corpo.includes(comandoFechamento);
    });

    if (houveFechamentoManual) {
      thread.addLabel(labelRespondido);
      thread.addLabel(labelCancelada);
      if (labels.includes("NC_Reenviada")) {
        thread.removeLabel(labelReenviado);
      }
      ignoradasCount++;
      ignoradasMotivo["NC encerrada manualmente"] = (ignoradasMotivo["NC encerrada manualmente"] || 0) + 1;
      return;
    }

    const mensagem = "⚠️ Reforço automático: Ainda não identificamos resposta a esta Não Conformidade.\n\nPor favor, verificar o histórico deste e-mail e responder diretamente por aqui conforme orientações.";
    thread.replyAll(mensagem);

    if (!labels.includes("NC_Reenviada")) {
      thread.addLabel(labelReenviado);
    }

    reenviadas.push(assunto);
  });

  Logger.log("Ignoradas (total: " + ignoradasCount + "):");
  for (const motivo in ignoradasMotivo) {
    Logger.log(`- ${motivo}: ${ignoradasMotivo[motivo]} thread(s)`);
  }

  if (reenviadas.length > 0) {
    Logger.log("Reenviadas:");
    reenviadas.forEach(r => Logger.log("- " + r));
  } else {
    Logger.log("Nenhum reenvio foi necessário nesta execução.");
  }

  Logger.log("Verificação concluída.");
}