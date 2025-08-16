function processarNC() {
  const EMAIL_AUTOR = "EMAIL.AUTORIZADO@PROVEDOR.com.br"; 
  const ASSUNTO_PADRAO = "NÃO CONFORMIDADE: NC - ";
  const PRAZO_HORAS = 48 // Verificação de 2 dias sem resposta.
  const DATA_MINIMA = new Date("2025-08-15T00:00:00"); 

  const labelReenviado = getOrCreateLabel("NC_Reenviada");
  const labelRespondido = getOrCreateLabel("NC_Respondida");
  const labelCancelado = getOrCreateLabel("NC_Cancelada");

  const query = `subject:"${ASSUNTO_PADRAO}" after:${formatDateForQuery(DATA_MINIMA)}`;
  const threads = GmailApp.search(query, 0, 500);

  let processadasCount = threads.length; // Total de threads encontradas
  let reenviadasCount = 0;
  let canceladasCount = 0;
  let respondidasCount = 0;
  let ignoradasMotivo = {};
  let reenviadasNCs = [];

  // Set para controlar threads já logadas como ignoradas
  const threadsIgnoradas = new Set();

  for (const thread of threads) {
    try {
      const mensagens = thread.getMessages();
      if (mensagens.length === 0) continue;

      const labels = thread.getLabels().map(l => l.getName());

      // Ignorar threads com NC_Respondida ou NC_Cancelada
      if (labels.includes(labelRespondido.getName()) || labels.includes(labelCancelado.getName())) {
        const motivo = "Já possui NC_Respondida ou NC_Cancelada";
        // Removido Logger.log para não exibir repetidamente
        if (labels.includes(labelRespondido.getName())) respondidasCount++;
        if (labels.includes(labelCancelado.getName())) canceladasCount++;
        ignoradasMotivo[motivo] = (ignoradasMotivo[motivo] || 0) + 1;
        continue;
      }

      // Verifica se o próprio autor enviou #fecharnc
      const autorFechouNC = mensagens.slice(1).some(msg => {
        const from = msg.getFrom().toLowerCase();
        const corpo = (msg.getPlainBody() || "").toLowerCase();
        return from.includes(EMAIL_AUTOR.toLowerCase()) && corpo.includes("#fecharnc");
      });
      if (autorFechouNC) {
        thread.addLabel(labelRespondido);
        thread.addLabel(labelCancelado);
        canceladasCount++;
        ignoradasMotivo["Encerrada pelo autor via #fecharnc"] = (ignoradasMotivo["Encerrada pelo autor via #fecharnc"] || 0) + 1;
        continue;
      }

      // Verifica se algum destinatário respondeu (não autor e não automática)
      const respostaValida = mensagens.slice(1).some(msg => {
        const from = msg.getFrom().toLowerCase();
        const corpo = (msg.getPlainBody() || "").toLowerCase();
        const ehAuto = corpo.includes("resposta automática") || corpo.includes("esta é uma resposta automática") ||
                       corpo.includes("auto-reply") || corpo.includes("out of office");
        return !from.includes(EMAIL_AUTOR.toLowerCase()) && !ehAuto;
      });
      if (respostaValida) {
        thread.addLabel(labelRespondido);
        respondidasCount++;
        ignoradasMotivo["Respondida por destinatário válido"] = (ignoradasMotivo["Respondida por destinatário válido"] || 0) + 1;
        continue;
      }

      // Verifica se passou o prazo para cobrança
      const ultimaMensagem = mensagens[mensagens.length - 1];
      const dataUltima = ultimaMensagem.getDate();
      const agora = new Date();
      const diffHoras = (agora - dataUltima) / 36e5;

      if (diffHoras < PRAZO_HORAS) {
        ignoradasMotivo["Ainda no prazo"] = (ignoradasMotivo["Ainda no prazo"] || 0) + 1;
        continue;
      }

      // Reenviar (responder todos usando emails originais da thread)
      const reenviada = responderTodosNaThread(thread, mensagens, EMAIL_AUTOR);
      if (reenviada) {
        thread.addLabel(labelReenviado);
        reenviadasCount++;
        reenviadasNCs.push(mensagens[0].getSubject());
      }

    } catch (erro) {
      Logger.log("Erro ao processar thread: " + erro);
    }
  }

  // Log final
  Logger.log(`Total processadas: ${processadasCount}`);
  Logger.log(`Total canceladas: ${canceladasCount}`);
  Logger.log(`Total respondidas: ${respondidasCount}`);
  Logger.log("NC reenviadas:");
  reenviadasNCs.forEach(nc => Logger.log(` - ${nc}`));
  Logger.log("Motivos ignorados:");
  for (const motivo in ignoradasMotivo) {
    Logger.log(` - ${motivo}: ${ignoradasMotivo[motivo]}`);
  }
}

// Atualizada: retorna true se realmente houve reenvio
function responderTodosNaThread(thread, mensagens, emailAutor) {
  let toOriginal = [];
  let ccOriginal = [];

  for (const msg of mensagens) {
    const toList = msg.getTo()?.split(',').map(e => e.trim()) || [];
    const ccList = msg.getCc()?.split(',').map(e => e.trim()) || [];
    toOriginal.push(...toList);
    ccOriginal.push(...ccList);
  }

  toOriginal = Array.from(new Set(toOriginal.filter(e => e.toLowerCase() !== emailAutor.toLowerCase())));
  ccOriginal = Array.from(new Set(ccOriginal.filter(e => e.toLowerCase() !== emailAutor.toLowerCase())));

  if (toOriginal.length === 0 && ccOriginal.length === 0) {
    Logger.log(`Thread "${mensagens[0].getSubject()}" não possui destinatários válidos para cobrança.`);
    return false;
  }

  const corpo = `⚠️ Atenção: esta Não Conformidade ainda não recebeu um retorno.\n` +
                `Por favor, responda a este email, para registrarmos na NC original.`;

  // Responder todos usando a primeira mensagem da thread
  mensagens[0].replyAll(corpo);
  return true;
}

function getOrCreateLabel(nome) {
  return GmailApp.getUserLabelByName(nome) || GmailApp.createLabel(nome);
}

function formatDateForQuery(data) {
  const ano = data.getFullYear();
  const mes = String(data.getMonth() + 1).padStart(2, "0");
  const dia = String(data.getDate()).padStart(2, "0");
  return `${ano}/${mes}/${dia}`;
}