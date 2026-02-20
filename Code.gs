const SHEET_CONFIG       = 'CONFIG';
const SHEET_ATENDIMENTOS = 'ATENDIMENTOS';
const SHEET_LOG          = 'LOG_CONVERSAS';
const SHEET_INBOX_TESTE  = 'INBOX_TESTE';
const SHEET_SETORES      = 'SETORES';
const SHEET_FAQ          = 'FAQ';
const SHEET_WEBHOOK_LOG  = 'WEBHOOK_LOG';

const ST_NOVO               = 'NOVO';
const ST_AGUARDANDO_DUVIDA   = 'AGUARDANDO_DUVIDA';
const ST_AGUARDANDO_CONFIRMA = 'AGUARDANDO_CONFIRMACAO_FAQ';
const ST_AGUARDANDO_SETOR    = 'AGUARDANDO_SETOR';
const ST_ABERTO_AGUARD_DP    = 'ABERTO – aguardando DP';
const ST_EM_ATENDIMENTO      = 'EM ATENDIMENTO';
const ST_ENCERRADO           = 'ENCERRADO';

function makeDedupKey_(telefone, mensagem) {
  const phone = String(telefone || '').replace(/\D/g, '');
  const msg = (mensagem || '').toString().trim().slice(0, 500);
  return Utilities.base64EncodeWebSafe(phone + '|' + msg);
}

function isDuplicateInbound_(telefone, mensagem) {
  const cache = CacheService.getScriptCache();
  const key = 'DEDUP_IN_' + makeDedupKey_(telefone, mensagem);

  const already = cache.get(key);
  if (already) return true;

  cache.put(key, '1', 45);
  return false;
}

function getConfig_(chave) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) throw new Error('Aba CONFIG não encontrada.');

  const values = sheet.getDataRange().getValues();
  const header = values[0];
  const chaveIndex = header.indexOf('CHAVE');
  const valorIndex = header.indexOf('VALOR');

  if (chaveIndex === -1 || valorIndex === -1) {
    throw new Error('Aba CONFIG precisa das colunas CHAVE e VALOR na linha 1.');
  }

  for (let i = 1; i < values.length; i++) {
    if ((values[i][chaveIndex] || '').toString().trim() === chave) {
      return values[i][valorIndex];
    }
  }
  return null;
}

function getZapiConfig_() {
  return {
    baseUrl: 'https://api.z-api.io',
    instanceId:  getConfig_('ZAPI_INSTANCE_ID'),
    token:       getConfig_('ZAPI_TOKEN'),
    clientToken: getConfig_('ZAPI_CLIENT_TOKEN')
  };
}

function getGroqConfig_() {
  return {
    apiKey: getConfig_('GROQ_API_KEY'),
    model: getConfig_('GROQ_MODEL') || 'llama3-8b-8192'
  };
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('BOT DP TESTE')
    .addItem('Processar mensagens de teste', 'processarInboxTeste')
    .addSeparator()
    .addItem('Abrir painel (URL WebApp nas implantações)', 'mostrarAlertaPainel')
    .addToUi();
}

function mostrarAlertaPainel() {
  SpreadsheetApp.getUi().alert(
    'Para abrir o painel HTML:\n' +
    '1) Vá em Implantar > Gerenciar implantações\n' +
    '2) Copie a URL do "Aplicativo da Web"\n' +
    '3) Abra essa URL no navegador.'
  );
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Painel BOT DP – Teste')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function sanitizeInboundText_(txt) {
  txt = (txt || '').toString().trim();
  if (!txt) return '';

  if (txt.startsWith('{') && txt.endsWith('}')) {
    try {
      const obj = JSON.parse(txt);
      if (obj && typeof obj === 'object') {
        if (typeof obj.message === 'string') return obj.message.trim();
        if (typeof obj.text === 'string') return obj.text.trim();
        if (typeof obj.body === 'string') return obj.body.trim();
      }
    } catch (e) {}
  }
  return txt;
}

function doPost(e) {
  const ss = SpreadsheetApp.getActive();
  const logSheet = ss.getSheetByName(SHEET_WEBHOOK_LOG);

  try {
    const rawBody = (e && e.postData && e.postData.contents) ? e.postData.contents : '';

    if (logSheet) logSheet.appendRow([new Date(), 'RAW', rawBody]);
    if (!rawBody) return ContentService.createTextOutput('NO_BODY');

    let body;
    try {
      body = JSON.parse(rawBody);
    } catch (errJson) {
      if (logSheet) logSheet.appendRow([new Date(), 'ERRO_JSON', errJson.toString()]);
      return ContentService.createTextOutput('ERRO_JSON');
    }

    const data = body.data || body;

    const fromMe =
      (data.message && data.message.fromMe === true) ||
      (typeof data.fromMe !== 'undefined' && data.fromMe === true);

    if (fromMe) return ContentService.createTextOutput('IGNORED_FROM_ME');

    let telefone =
      (data.phone) ||
      (data.chatId ? String(data.chatId).replace(/@c\.us$/i, '') : '') ||
      (data.from) ||
      (body.sender && body.sender.phone) ||
      '';

    let mensagem =
      (typeof data.message === 'string' && data.message) ||
      (data.message && (data.message.text || data.message.body || data.message.message)) ||
      data.text ||
      data.body ||
      (typeof body.message === 'string' && body.message) ||
      (body.message && (body.message.text || body.message.body || body.message.message)) ||
      '';

    if (mensagem && typeof mensagem === 'object') {
      if (typeof mensagem.text === 'string') mensagem = mensagem.text;
      else if (typeof mensagem.body === 'string') mensagem = mensagem.body;
      else if (typeof mensagem.message === 'string') mensagem = mensagem.message;
      else if (typeof mensagem.buttonText === 'string') mensagem = mensagem.buttonText;
      else mensagem = JSON.stringify(mensagem);
    }

    telefone = (telefone || '').toString().trim();
    mensagem = sanitizeInboundText_(mensagem);

    if (!telefone || !mensagem) {
      if (logSheet) logSheet.appendRow([new Date(), 'MISSING_FIELDS', telefone, mensagem]);
      return ContentService.createTextOutput('MISSING_FIELDS');
    }

    if (isDuplicateInbound_(telefone, mensagem)) {
      if (logSheet) logSheet.appendRow([new Date(), 'DUPLICADO_IGNORADO', telefone, mensagem]);
      return ContentService.createTextOutput('DUPLICATE_IGNORED');
    }

    processarMensagemGeral_(telefone, mensagem, null);

    if (logSheet) logSheet.appendRow([new Date(), 'OK', telefone, mensagem]);
    return ContentService.createTextOutput('OK');

  } catch (err) {
    if (logSheet) logSheet.appendRow([new Date(), 'ERRO_DOPOST', err.toString()]);
    return ContentService.createTextOutput('ERROR').setResponseCode(500);
  }
}

function processarInboxTeste() {
  const ss = SpreadsheetApp.getActive();
  const inbox = ss.getSheetByName(SHEET_INBOX_TESTE);
  if (!inbox) throw new Error('Aba INBOX_TESTE não encontrada.');

  const data = inbox.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rowIndex = i + 1;
    const status = (data[i][3] || '').toString().trim();

    if (status && status.toUpperCase() !== 'NOVO') continue;

    const telefone = (data[i][1] || '').toString().trim();
    const mensagem = sanitizeInboundText_(data[i][2]);

    if (!telefone || !mensagem) continue;

    inbox.getRange(rowIndex, 4).setValue('EM_PROCESSAMENTO');
    processarMensagemGeral_(telefone, mensagem, rowIndex);
  }
}

function getMensagemInicial_() {
  return 'Olá! Sou o assistente virtual do Departamento Pessoal da MSE Engenharia. Como posso te ajudar?';
}

function getMsgSetores_() {
  return (
    'A sua dúvida será melhor respondida por um atendente. Por favor, escolha o setor:\n\n' +
    '1 - Admissão\n' +
    '2 - Rescisão\n' +
    '3 - Folha\n' +
    '4 - Ponto'
  );
}

function processarMensagemGeral_(telefone, texto, rowInbox) {
  texto = sanitizeInboundText_(texto);

  const atendimento = getAtendimentoPorTelefone_(telefone);

  if (!atendimento) {
    const newId = criaAtendimento_(telefone);
    registraLog_(telefone, 'CLIENTE', texto, newId);

    const msgInicial = getMensagemInicial_();
    responderCanal_(telefone, rowInbox, msgInicial);
    registraLog_(telefone, 'BOT', msgInicial, newId);

    setStatusAtendimentoById_(newId, ST_AGUARDANDO_DUVIDA);
    return;
  }

  if (String(atendimento.status || '').toUpperCase().startsWith(ST_ENCERRADO)) {
    resetAtendimentoMesmoId_(atendimento.row);
    registraLog_(telefone, 'CLIENTE', texto, atendimento.id);

    const msgInicial = getMensagemInicial_();
    responderCanal_(telefone, rowInbox, msgInicial);
    registraLog_(telefone, 'BOT', msgInicial, atendimento.id);

    setStatusAtendimentoRow_(atendimento.row, ST_AGUARDANDO_DUVIDA);
    return;
  }

  registraLog_(telefone, 'CLIENTE', texto, atendimento.id);
  tratarMensagemFluxo_(atendimento, texto, rowInbox);
}

function tratarMensagemFluxo_(atendimento, texto, rowInbox) {
  const ss = SpreadsheetApp.getActive();
  const sheetAt = ss.getSheetByName(SHEET_ATENDIMENTOS);
  const row = atendimento.row;

  const setor = (sheetAt.getRange(row, 4).getValue() || '').toString().trim();
  let status  = (sheetAt.getRange(row, 6).getValue() || '').toString().trim();

  const now = new Date();
  sheetAt.getRange(row, 7).setValue(texto);
  sheetAt.getRange(row, 10).setValue(now);

  if (setor || status === ST_EM_ATENDIMENTO || status === ST_ABERTO_AGUARD_DP) {
    return;
  }

  if (!status) {
    status = ST_NOVO;
    sheetAt.getRange(row, 6).setValue(ST_NOVO);
  }

  if (status === ST_AGUARDANDO_SETOR) {
    const match = (texto || '').toString().match(/\d+/);
    const num = match ? parseInt(match[0], 10) : NaN;

    let setorEscolhido = '';
    switch (num) {
      case 1: setorEscolhido = 'ADMISSÃO'; break;
      case 2: setorEscolhido = 'RESCISÃO'; break;
      case 3: setorEscolhido = 'FOLHA'; break;
      case 4: setorEscolhido = 'PONTO'; break;
      default: setorEscolhido = '';
    }

    if (!setorEscolhido) {
      const msgInv = 'Opção inválida. Responda com 1 (Admissão), 2 (Rescisão), 3 (Folha) ou 4 (Ponto).';
      responderCanal_(atendimento.telefone, rowInbox, msgInv);
      registraLog_(atendimento.telefone, 'BOT', msgInv, atendimento.id);
      return;
    }

    sheetAt.getRange(row, 4).setValue(setorEscolhido);
    sheetAt.getRange(row, 6).setValue(ST_ABERTO_AGUARD_DP);
    sheetAt.getRange(row, 10).setValue(new Date());

    const msgOk = 'Perfeito! Um atendente do setor ' + setorEscolhido + ' irá assumir seu atendimento em breve.';
    responderCanal_(atendimento.telefone, rowInbox, msgOk);
    registraLog_(atendimento.telefone, 'BOT', msgOk, atendimento.id);
    return;
  }

  if (status === ST_AGUARDANDO_CONFIRMA) {
    const lower = (texto || '').toString().trim().toLowerCase();

    if (lower === 'não' || lower === 'nao' || lower.indexOf('não') >= 0 || lower.indexOf('nao') >= 0) {
      const msgFim = 'Atendimento finalizado.';
      responderCanal_(atendimento.telefone, rowInbox, msgFim);
      registraLog_(atendimento.telefone, 'BOT', msgFim, atendimento.id);
      sheetAt.getRange(row, 6).setValue(ST_ENCERRADO);
      return;
    }

    if (lower === 'sim' || lower.indexOf('sim') >= 0) {
      const msg = 'Qual sua dúvida?';
      responderCanal_(atendimento.telefone, rowInbox, msg);
      registraLog_(atendimento.telefone, 'BOT', msg, atendimento.id);
      sheetAt.getRange(row, 6).setValue(ST_AGUARDANDO_DUVIDA);
      return;
    }

    sheetAt.getRange(row, 6).setValue(ST_AGUARDANDO_DUVIDA);
  }

  if (status === ST_NOVO) {
    const msgInicial = getMensagemInicial_();
    responderCanal_(atendimento.telefone, rowInbox, msgInicial);
    registraLog_(atendimento.telefone, 'BOT', msgInicial, atendimento.id);
    sheetAt.getRange(row, 6).setValue(ST_AGUARDANDO_DUVIDA);
    return;
  }

  if (status === ST_AGUARDANDO_DUVIDA) {
    const faq = buscarFaq_(texto);
    if (faq && faq.resposta) {
      const respostaFaq =
        faq.resposta +
        '\n\nPosso te ajudar em algo mais? (responda "sim" ou "não")';

      responderCanal_(atendimento.telefone, rowInbox, respostaFaq);
      registraLog_(atendimento.telefone, 'BOT', respostaFaq, atendimento.id);
      sheetAt.getRange(row, 6).setValue(ST_AGUARDANDO_CONFIRMA);
      return;
    }

    const msgSetores = getMsgSetores_();
    responderCanal_(atendimento.telefone, rowInbox, msgSetores);
    registraLog_(atendimento.telefone, 'BOT', msgSetores, atendimento.id);
    sheetAt.getRange(row, 6).setValue(ST_AGUARDANDO_SETOR);
    return;
  }

  sheetAt.getRange(row, 6).setValue(ST_AGUARDANDO_DUVIDA);
}

function criaAtendimento_(telefone) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_ATENDIMENTOS);
  if (!sheet) throw new Error('Aba ATENDIMENTOS não encontrada.');

  const lastRow = sheet.getLastRow();

  let newId;
  if (lastRow >= 2) {
    const lastId = sheet.getRange(lastRow, 1).getValue();
    newId = (Number(lastId) || 0) + 1;
  } else {
    newId = 1;
  }

  const now = new Date();
  sheet.appendRow([
    newId,
    telefone,
    '',
    '',
    '',
    ST_NOVO,
    '',
    '',
    now,
    now
  ]);

  return newId;
}

function resetAtendimentoMesmoId_(row) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_ATENDIMENTOS);
  const now = new Date();

  sheet.getRange(row, 3).setValue('');
  sheet.getRange(row, 4).setValue('');
  sheet.getRange(row, 5).setValue('');
  sheet.getRange(row, 6).setValue(ST_NOVO);
  sheet.getRange(row, 7).setValue('');
  sheet.getRange(row, 8).setValue('');
  sheet.getRange(row, 9).setValue(now);
  sheet.getRange(row, 10).setValue(now);
}

function setStatusAtendimentoById_(idAtendimento, status) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_ATENDIMENTOS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(idAtendimento)) {
      sheet.getRange(i + 1, 6).setValue(status);
      sheet.getRange(i + 1, 10).setValue(new Date());
      return;
    }
  }
}

function setStatusAtendimentoRow_(row, status) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_ATENDIMENTOS);
  sheet.getRange(row, 6).setValue(status);
  sheet.getRange(row, 10).setValue(new Date());
}

function getAtendimentoPorTelefone_(telefone) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_ATENDIMENTOS);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const telBusca = (telefone || '').toString().trim();

  for (let i = data.length - 1; i >= 1; i--) {
    const tel = (data[i][1] || '').toString().trim();
    if (tel === telBusca) {
      return {
        row: i + 1,
        id: data[i][0],
        telefone: tel,
        setor: (data[i][3] || '').toString().trim(),
        responsavel: (data[i][4] || '').toString().trim(),
        status: (data[i][5] || '').toString().trim()
      };
    }
  }
  return null;
}

function registraLog_(telefone, origem, mensagem, idAtendimento) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet) throw new Error('Aba LOG_CONVERSAS não encontrada.');

  sheet.appendRow([
    new Date(),
    telefone,
    origem,
    mensagem,
    idAtendimento || ''
  ]);
}

function responderCanal_(telefone, rowInbox, mensagemBot) {
  const ss = SpreadsheetApp.getActive();
  if (rowInbox) {
    const inbox = ss.getSheetByName(SHEET_INBOX_TESTE);
    inbox.getRange(rowInbox, 5).setValue(mensagemBot);
    inbox.getRange(rowInbox, 4).setValue('RESPONDIDO');
  } else {
    enviarMensagemWhats_(telefone, mensagemBot);
  }
}

function buscarFaq_(mensagem) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_FAQ);
  if (!sheet) return null;

  const values = sheet.getDataRange().getValues();
  const header = values[0];
  const pergIndex = header.indexOf('PERGUNTA');
  const respIndex = header.indexOf('RESPOSTA');
  if (pergIndex === -1 || respIndex === -1) return null;

  const msgLower = (mensagem || '').toString().toLowerCase();

  for (let i = 1; i < values.length; i++) {
    const pergunta = String(values[i][pergIndex] || '').toLowerCase().trim();
    if (!pergunta) continue;

    if (msgLower.includes(pergunta)) {
      return { resposta: values[i][respIndex] };
    }
  }
  return null;
}

function responderComIAEstruturada_(mensagemUsuario) {
  const groq = getGroqConfig_();
  if (!groq.apiKey) return null;

  const systemPrompt =
    'Você é o assistente virtual do Departamento Pessoal (DP) de uma empresa brasileira (MSE Engenharia). ' +
    'Responda sempre em pt-BR. ' +
    'Se a pergunta exigir procedimento interno não conhecido, responda NAO_SEI. ' +
    'Saída OBRIGATÓRIA em JSON puro (sem markdown), no formato:\n' +
    '{ "answer": "...", "confidence": 0.0, "setor": "ADMISSÃO|RESCISÃO|FOLHA|PONTO|INDEFINIDO" }\n' +
    'confidence deve ser de 0 a 1. ' +
    'Se não tiver certeza, use confidence baixa (<=0.5) e answer=NAO_SEI.';

  const url = 'https://api.groq.com/openai/v1/chat/completions';

  const payload = {
    model: groq.model,
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: mensagemUsuario }
    ],
    temperature: 0.2
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + groq.apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    const content = data?.choices?.[0]?.message?.content || '';
    if (!content) return null;

    let obj;
    try { obj = JSON.parse(content); }
    catch (e) { return null; }

    const answer = (obj.answer || '').toString().trim();
    const confidence = Number(obj.confidence || 0);
    const setor = (obj.setor || 'INDEFINIDO').toString().trim().toUpperCase();

    return { answer, confidence, setor };
  } catch (e) {
    return null;
  }
}

function enviarMensagemWhats_(telefone, texto) {
  const cfg = getZapiConfig_();
  if (!cfg.instanceId || !cfg.token || !cfg.clientToken) return;

  const url = cfg.baseUrl + '/instances/' + cfg.instanceId + '/token/' + cfg.token + '/send-text';
  const phone = String(telefone || '').replace(/\D/g, '');

  const payload = { phone: phone, message: texto };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Client-Token': cfg.clientToken },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  UrlFetchApp.fetch(url, options);
}

function testeEnvioDiretoZapi() {
  const meuNumero = '554330310547';
  const texto = 'Teste direto da Z-API a partir do Apps Script.';
  enviarMensagemWhats_(meuNumero, texto);
}

function formatDate_(d) {
  if (!d || Object.prototype.toString.call(d) !== '[object Date]' || isNaN(d)) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

function listarAtendimentos(filtroSetor, filtroStatus) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_ATENDIMENTOS);
  const data = sheet.getDataRange().getValues();

  const result = [];
  filtroSetor = (filtroSetor || '').toString().trim();
  filtroStatus = (filtroStatus || '').toString().trim();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id          = row[0];
    const telefone    = row[1];
    const nome        = row[2];
    const setor       = row[3];
    const responsavel = row[4];
    const status      = row[5];
    const ultCli      = row[6];
    const ultAtend    = row[7];
    const dtAbert     = row[8];
    const dtUlt       = row[9];

    if (!id) continue;
    if (filtroSetor && setor !== filtroSetor) continue;
    if (filtroStatus && status !== filtroStatus) continue;

    result.push({
      id: id,
      telefone: telefone,
      nome: nome,
      setor: setor,
      responsavel: responsavel,
      status: status,
      ultimaMsgCliente: ultCli,
      ultimaMsgAtendente: ultAtend,
      dataAbertura: formatDate_(dtAbert),
      dataUltimaInteracao: formatDate_(dtUlt)
    });
  }

  result.sort((a, b) => (b.dataUltimaInteracao || '').localeCompare(a.dataUltimaInteracao || ''));
  return result;
}

function getSetores() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_SETORES);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const setores = [];

  for (let i = 1; i < data.length; i++) {
    const nome = (data[i][0] || '').toString().trim();
    if (nome && setores.indexOf(nome) === -1) setores.push(nome);
  }
  return setores;
}

function obterConversas(idAtendimento) {
  const ss = SpreadsheetApp.getActive();
  const sheetLog = ss.getSheetByName(SHEET_LOG);
  const data = sheetLog.getDataRange().getValues();
  const result = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dt       = row[0];
    const origem   = row[2];
    const mensagem = row[3];
    const idLog    = row[4];

    if (Number(idLog) === Number(idAtendimento)) {
      result.push({
        dataHora: formatDate_(dt),
        origem: origem,
        mensagem: mensagem
      });
    }
  }

  result.sort((a, b) => a.dataHora.localeCompare(b.dataHora));
  return result;
}

function enviarRespostaPainel(idAtendimento, mensagem, autor) {
  if (!mensagem) return;

  const ss = SpreadsheetApp.getActive();
  const sheetAt = ss.getSheetByName(SHEET_ATENDIMENTOS);
  const dataAt = sheetAt.getDataRange().getValues();

  let rowIndex = -1;
  let telefone = '';

  for (let i = 1; i < dataAt.length; i++) {
    const id = dataAt[i][0];
    if (Number(id) === Number(idAtendimento)) {
      rowIndex = i + 1;
      telefone = String(dataAt[i][1] || '');
      break;
    }
  }

  if (rowIndex === -1) throw new Error('Atendimento não encontrado para ID ' + idAtendimento);

  const now = new Date();
  const origem = autor ? ('ATENDENTE - ' + autor) : 'ATENDENTE';

  const textoTrim = (mensagem || '').toString().trim().toLowerCase();
  const novoStatus = (textoTrim === 'atendimento finalizado.') ? ST_ENCERRADO : ST_EM_ATENDIMENTO;

  sheetAt.getRange(rowIndex, 8).setValue(mensagem);
  sheetAt.getRange(rowIndex, 6).setValue(novoStatus);
  sheetAt.getRange(rowIndex, 10).setValue(now);

  const sheetLog = ss.getSheetByName(SHEET_LOG);
  sheetLog.appendRow([now, telefone, origem, mensagem, idAtendimento]);

  const textoWhats = autor ? (autor + ': ' + mensagem) : mensagem;
  enviarMensagemWhats_(telefone, textoWhats);

  return true;
}