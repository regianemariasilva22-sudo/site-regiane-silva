/**
 * Backend das Áreas de Membros (Programa + Check-up) — Regiane Silva
 * Lê e escreve numa Google Sheet, publicado como Web App.
 * Veja INSTRUCOES.md para o passo a passo de publicação.
 */

// ID da planilha (fica na URL dela, entre /d/ e /edit)
const SHEET_ID = '1-13n-7EhzF9b45OC--ijZBVQYMOJZr8PrWquUot1G7E';

// Client ID do OAuth do Google (Sign In With Google), criado no Google Cloud Console
const GOOGLE_CLIENT_ID = '288771217381-mt5g3dhdjhcoak6kphd1fhsarrkd44bc.apps.googleusercontent.com';

// E-mail da Regiane, para onde vão os avisos de novo cadastro/acesso liberado
const REGIANE_NOTIFICATION_EMAIL = 'regianemariasilva22@gmail.com';

// E-mails com acesso de administradora a qualquer área do site, sem precisar
// estar cadastrado nas planilhas de pacientes.
const ADMIN_EMAILS = ['divarebel.on@gmail.com', 'babadosdaaline@gmail.com', 'regianemariasilva22@gmail.com'];
function isAdmin(email) {
  return ADMIN_EMAILS.map(normEmail).indexOf(normEmail(email)) !== -1;
}

function getSheet(name) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
  if (!sheet) throw new Error('Aba "' + name + '" não encontrada na planilha.');
  return sheet;
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    })
    .filter(obj => Object.values(obj).some(v => v !== '' && v !== null));
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function normEmail(email) {
  return String(email || '').trim().toLowerCase();
}

/**
 * Verifica um id_token do Google Sign In (emitido pro nosso GOOGLE_CLIENT_ID)
 * e retorna o e-mail verificado. Lança erro se o token for inválido.
 */
function verifyGoogleToken(idToken) {
  if (!idToken) throw new Error('Token do Google ausente.');
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken), {
    muteHttpExceptions: true
  });
  const info = JSON.parse(res.getContentText());
  if (info.error) throw new Error('Token do Google inválido: ' + info.error);
  if (info.aud !== GOOGLE_CLIENT_ID) throw new Error('Token não pertence a este app.');
  if (info.email_verified !== 'true' && info.email_verified !== true) throw new Error('E-mail do Google não verificado.');
  return { email: normEmail(info.email), nome: info.name || info.email };
}

function notifyRegiane(assunto, corpo) {
  if (!REGIANE_NOTIFICATION_EMAIL || REGIANE_NOTIFICATION_EMAIL.indexOf('COLE_AQUI') !== -1) return;
  try {
    MailApp.sendEmail(REGIANE_NOTIFICATION_EMAIL, assunto, corpo);
  } catch (err) {
    // não deixa o fluxo principal quebrar se o e-mail falhar
  }
}

// ── ROTEAMENTO ──────────────────────────────────────────

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'login') return jsonResponse(actionLogin(e.parameter.email));
    if (action === 'dashboard') return jsonResponse(actionDashboard(e.parameter.email));
    if (action === 'comments') return jsonResponse(actionComments(e.parameter.postId));
    if (action === 'slots') return jsonResponse(actionSlots());
    return jsonResponse({ ok: false, error: 'Ação inválida: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'comment') return jsonResponse(actionAddComment(body));
    if (action === 'bookSlot') return jsonResponse(actionBookSlot(body));
    if (action === 'googleLoginPrograma') return jsonResponse(actionGoogleLoginPrograma(body));
    if (action === 'googleLoginCheckup') return jsonResponse(actionGoogleLoginCheckup(body));
    if (action === 'submitCheckup') return jsonResponse(actionSubmitCheckup(body));
    if (action === 'asaasWebhook') return jsonResponse(actionAsaasWebhook(body));
    return jsonResponse({ ok: false, error: 'Ação inválida: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

// ── PACIENTES ────────────────────────────────────────────

function findPatientRow(email) {
  const sheet = getSheet('Pacientes');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  for (let i = 1; i < data.length; i++) {
    if (normEmail(data[i][emailCol]) === normEmail(email)) {
      const obj = {};
      headers.forEach((h, idx) => obj[h] = data[i][idx]);
      obj._row = i + 1;
      return obj;
    }
  }
  return null;
}

function actionLogin(email) {
  const p = findPatientRow(email);
  if (!p) {
    return { ok: false, error: 'E-mail não encontrado. Verifique com a Regiane se seu cadastro já foi feito.' };
  }
  return { ok: true, nome: p.Nome, email: p.Email };
}

/**
 * Login da Área do Programa via "Continuar com o Google".
 * Verifica o token, confirma que o e-mail está cadastrado na aba Pacientes,
 * e libera o acesso (sem aceitar e-mail digitado à mão).
 */
function actionGoogleLoginPrograma(body) {
  const auth = verifyGoogleToken(body.idToken);
  if (isAdmin(auth.email)) {
    return { ok: true, nome: auth.nome, email: auth.email, admin: true };
  }
  const p = findPatientRow(auth.email);
  if (!p) {
    return { ok: false, error: 'Não encontramos seu cadastro no Programa com esta conta Google. Fale com a Regiane.' };
  }
  return { ok: true, nome: p.Nome, email: p.Email };
}

function daysSince(dateVal) {
  if (!dateVal) return 0;
  const start = new Date(dateVal);
  if (isNaN(start.getTime())) return 0;
  const now = new Date();
  return Math.max(0, Math.floor((now - start) / (1000 * 60 * 60 * 24)));
}

function actionDashboard(email) {
  let p = findPatientRow(email);
  if (!p && isAdmin(email)) {
    // administradora sem cadastro de paciente: mostra um painel de exemplo, sem erro.
    p = { Nome: 'Administradora', PontosTotal: 0, RetornosRealizados: 0, ReceitasSalvas: 0, ProgressoPercent: 0, DataInicio: '', ProximoRetornoData: '', ProximoRetornoHora: '', PlanoTexto: 'Acesso de administradora — sem plano individual.' };
  }
  if (!p) return { ok: false, error: 'Paciente não encontrada.' };

  const materiais = sheetToObjects(getSheet('Materiais')).filter(m => {
    const dest = normEmail(m.Email);
    return dest === normEmail(email) || dest === 'todos';
  });

  const pontosTotal = Number(p.PontosTotal) || 0;
  const totalInteracoes = sheetToObjects(getSheet('PontosLog')).filter(l => normEmail(l.Email) === normEmail(email)).length;

  return {
    ok: true,
    nome: p.Nome,
    diasAcompanhamento: daysSince(p.DataInicio),
    retornosRealizados: Number(p.RetornosRealizados) || 0,
    receitasSalvas: Number(p.ReceitasSalvas) || 0,
    progressoPercent: Number(p.ProgressoPercent) || 0,
    proximoRetornoData: p.ProximoRetornoData || '',
    proximoRetornoHora: p.ProximoRetornoHora || '',
    planoTexto: p.PlanoTexto || '',
    pontosTotal: pontosTotal,
    totalInteracoes: totalInteracoes,
    creditoDisponivel: Math.floor(pontosTotal / 100) * 10,
    faltamParaProximoCredito: 100 - (pontosTotal % 100),
    materiais: materiais.map(m => ({
      tipo: m.Tipo, titulo: m.Titulo, descricao: m.Descricao, link: m.Link
    }))
  };
}

// ── COMUNIDADE / COMENTÁRIOS ────────────────────────────

function actionComments(postId) {
  const all = sheetToObjects(getSheet('Comentarios'));
  const filtered = postId ? all.filter(c => String(c.PostId) === String(postId)) : all;
  return {
    ok: true,
    comments: filtered.map(c => ({ nome: c.Nome, texto: c.Texto, dataHora: c.DataHora }))
  };
}

function actionAddComment(body) {
  const email = normEmail(body.email);
  const p = findPatientRow(email);
  if (!p) return { ok: false, error: 'Paciente não encontrada.' };
  if (!body.texto || !String(body.texto).trim()) return { ok: false, error: 'Comentário vazio.' };

  const sheet = getSheet('Comentarios');
  const id = new Date().getTime();
  sheet.appendRow([id, body.postId, email, p.Nome, body.texto, new Date()]);

  addPoints(email, 'Comentário na comunidade', 3);

  return { ok: true, nome: p.Nome };
}

function addPoints(email, tipo, pontos) {
  getSheet('PontosLog').appendRow([new Date().getTime(), email, tipo, pontos, new Date()]);

  const pSheet = getSheet('Pacientes');
  const data = pSheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  const pontosCol = headers.indexOf('PontosTotal');
  for (let i = 1; i < data.length; i++) {
    if (normEmail(data[i][emailCol]) === normEmail(email)) {
      const atual = Number(data[i][pontosCol]) || 0;
      pSheet.getRange(i + 1, pontosCol + 1).setValue(atual + pontos);
      break;
    }
  }
}

// ── AGENDA ───────────────────────────────────────────────

function actionSlots() {
  const all = sheetToObjects(getSheet('Horarios'));
  const livres = all.filter(h => String(h.Disponivel).trim().toLowerCase() === 'sim');
  return {
    ok: true,
    slots: livres.map(h => ({ data: h.Data, hora: h.Hora }))
  };
}

function actionBookSlot(body) {
  const email = normEmail(body.email);
  const p = findPatientRow(email);
  if (!p) return { ok: false, error: 'Paciente não encontrada.' };
  if (!body.data || !body.hora) return { ok: false, error: 'Selecione data e hora.' };

  getSheet('Agendamentos').appendRow([email, body.data, body.hora, 'solicitado', new Date()]);

  // marca o horário como indisponível
  const hSheet = getSheet('Horarios');
  const data = hSheet.getDataRange().getValues();
  const headers = data[0];
  const dataCol = headers.indexOf('Data');
  const horaCol = headers.indexOf('Hora');
  const dispCol = headers.indexOf('Disponivel');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][dataCol]) === String(body.data) && String(data[i][horaCol]) === String(body.hora)) {
      hSheet.getRange(i + 1, dispCol + 1).setValue('Não');
      break;
    }
  }

  return { ok: true };
}

// ── CHECK-UP (acesso único por e-mail, liberado via pagamento) ──────

function findCheckupRow(email) {
  const sheet = getSheet('CheckupPacientes');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  for (let i = 1; i < data.length; i++) {
    if (normEmail(data[i][emailCol]) === normEmail(email)) {
      const obj = {};
      headers.forEach((h, idx) => obj[h] = data[i][idx]);
      obj._row = i + 1;
      return obj;
    }
  }
  return null;
}

/**
 * Login da Área do Check-up via "Continuar com o Google".
 * Só entra quem já teve o acesso liberado (por pagamento confirmado no Asaas,
 * ou cadastro manual da Regiane na planilha).
 */
function actionGoogleLoginCheckup(body) {
  const auth = verifyGoogleToken(body.idToken);
  if (isAdmin(auth.email)) {
    return { ok: true, nome: auth.nome, email: auth.email, admin: true, jaFezCheckup: false, respostasChecklist: null, respostasQuiz: null };
  }
  const c = findCheckupRow(auth.email);

  if (!c) {
    return { ok: false, error: 'Não encontramos seu Check-up. Se você já pagou, aguarde a liberação — pode levar alguns minutos.' };
  }
  if (String(c.Liberado).trim().toLowerCase() !== 'sim') {
    return { ok: false, error: 'Seu pagamento ainda não foi confirmado. Assim que for, seu acesso libera automaticamente.' };
  }

  return {
    ok: true,
    nome: c.Nome || auth.nome,
    email: c.Email,
    jaFezCheckup: String(c.JaFezCheckup).trim().toLowerCase() === 'sim',
    respostasChecklist: c.RespostasChecklist ? JSON.parse(c.RespostasChecklist) : null,
    respostasQuiz: c.RespostasQuiz ? JSON.parse(c.RespostasQuiz) : null
  };
}

/**
 * Salva as respostas do quiz/checklist do Check-up — só pode ser feito UMA vez
 * por e-mail. Depois disso, o login sempre retorna o resultado já salvo.
 */
function actionSubmitCheckup(body) {
  const auth = verifyGoogleToken(body.idToken);
  const c = findCheckupRow(auth.email);
  if (!c) return { ok: false, error: 'Check-up não encontrado para este e-mail.' };
  if (String(c.JaFezCheckup).trim().toLowerCase() === 'sim') {
    return { ok: false, error: 'Este e-mail já respondeu o Check-up. Cada Check-up pode ser feito apenas uma vez.' };
  }

  const sheet = getSheet('CheckupPacientes');
  const headers = sheet.getDataRange().getValues()[0];
  const row = c._row;
  sheet.getRange(row, headers.indexOf('JaFezCheckup') + 1).setValue('Sim');
  sheet.getRange(row, headers.indexOf('RespostasChecklist') + 1).setValue(JSON.stringify(body.respostasChecklist || {}));
  sheet.getRange(row, headers.indexOf('RespostasQuiz') + 1).setValue(JSON.stringify(body.respostasQuiz || {}));
  sheet.getRange(row, headers.indexOf('DataCheckup') + 1).setValue(new Date());

  return { ok: true };
}

/**
 * Recebe o webhook do Asaas quando um pagamento é confirmado, e libera o
 * acesso ao Check-up automaticamente pro e-mail usado na compra.
 *
 * IMPORTANTE: o payload padrão do Asaas traz o ID do cliente (payment.customer),
 * não o e-mail direto. Pra resolver o e-mail é preciso chamar a API do Asaas
 * (GET /customers/{id}) com a chave de API — isso ainda depende de vocês me
 * passarem a chave. Por enquanto, esta função aceita um e-mail já presente
 * no payload (caso configurem isso no Asaas) OU pode ser adaptada assim que
 * tivermos a chave de API.
 */
function actionAsaasWebhook(body) {
  const evento = body.event || '';
  const eventosConfirmados = ['PAYMENT_CONFIRMED', 'PAYMENT_RECEIVED'];
  if (eventosConfirmados.indexOf(evento) === -1) {
    return { ok: true, ignorado: true };
  }

  const payment = body.payment || {};
  const email = normEmail(payment.customerEmail || payment.email || body.email);
  const nome = payment.customerName || payment.name || '';

  if (!email) {
    return { ok: false, error: 'Webhook do Asaas sem e-mail do cliente. Configuração da API do Asaas ainda pendente.' };
  }

  const sheet = getSheet('CheckupPacientes');
  const existente = findCheckupRow(email);

  if (existente) {
    const headers = sheet.getDataRange().getValues()[0];
    sheet.getRange(existente._row, headers.indexOf('Liberado') + 1).setValue('Sim');
  } else {
    sheet.appendRow([email, nome, new Date(), 'Sim', 'Não', '', '', '']);
  }

  notifyRegiane(
    'Novo Check-up liberado — ' + email,
    'O pagamento de ' + (nome || email) + ' (' + email + ') foi confirmado no Asaas e o acesso ao Check-up foi liberado automaticamente.'
  );

  return { ok: true };
}

// ── CONFIGURAÇÃO INICIAL DA PLANILHA (rode uma vez, na mão) ──────
//
// No editor do Apps Script, selecione a função "setupSheetStructure" no
// menu de funções (ao lado do botão ▶ Executar) e clique em Executar.
// Isso cria as 7 abas, já com cabeçalho formatado nas cores da marca
// (marsala/rosé) e uma linha de exemplo em cada uma. Pode rodar de novo
// a qualquer momento — só recria os cabeçalhos, não apaga dados que já
// tiverem sido adicionados abaixo da linha de exemplo.

function setupSheetStructure() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const MARSALA = '#7A2A3B';
  const MARSALA_DEEP = '#551D29';
  const ROSE_MIST = '#F1DCDF';
  const INK = '#3B2024';
  const BG = '#FBF5F2';

  function buildSheet(name, headers, sampleRow) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);

    sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 2), Math.max(sheet.getMaxColumns(), headers.length))
      .setBackground(BG).setFontColor(INK).setFontFamily('Arial');

    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground(MARSALA)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(10)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(1, 34);
    sheet.setFrozenRows(1);

    if (sampleRow) {
      const r = sheet.getRange(2, 1, 1, sampleRow.length);
      r.setValues([sampleRow]).setBackground(ROSE_MIST).setFontColor(MARSALA_DEEP).setFontStyle('italic');
    }

    for (let c = 1; c <= headers.length; c++) sheet.autoResizeColumn(c);
    sheet.setTabColor(MARSALA);
    return sheet;
  }

  buildSheet('Pacientes',
    ['Email', 'Nome', 'DataInicio', 'RetornosRealizados', 'ReceitasSalvas', 'ProgressoPercent', 'ProximoRetornoData', 'ProximoRetornoHora', 'PlanoTexto', 'PontosTotal'],
    ['exemplo@paciente.com', 'Nome de Exemplo', new Date(), 0, 0, 0, '', '', 'Siga as orientações da última consulta.', 0]);

  buildSheet('Materiais',
    ['Id', 'Email', 'Tipo', 'Titulo', 'Descricao', 'Link'],
    [1, 'TODOS', 'PDF', 'Planner alimentar semanal', 'Vale para todas as pacientes do Programa', 'https://']);

  buildSheet('Comentarios', ['Id', 'PostId', 'Email', 'Nome', 'Texto', 'DataHora'], null);

  buildSheet('PontosLog', ['Id', 'Email', 'Tipo', 'Pontos', 'Data'], null);

  buildSheet('Horarios', ['Data', 'Hora', 'Disponivel'], ['28/07/2026', '15h00', 'Sim']);

  buildSheet('Agendamentos', ['Email', 'Data', 'Hora', 'Status', 'DataSolicitacao'], null);

  buildSheet('CheckupPacientes',
    ['Email', 'Nome', 'DataLiberacao', 'Liberado', 'JaFezCheckup', 'RespostasChecklist', 'RespostasQuiz', 'DataCheckup'],
    ['exemplo@checkup.com', 'Nome de Exemplo', new Date(), 'Sim', 'Não', '', '', '']);

  // remove a aba padrão em branco, se existir e não for a única
  ['Sheet1', 'Página1', 'Folha1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  // ordena as abas na ordem que faz mais sentido pro dia a dia da Regiane
  const ordem = ['Pacientes', 'CheckupPacientes', 'Materiais', 'Horarios', 'Agendamentos', 'Comentarios', 'PontosLog'];
  ordem.forEach((nome, i) => {
    const s = ss.getSheetByName(nome);
    if (s) ss.setActiveSheet(s);
    if (s) ss.moveActiveSheet(i + 1);
  });

  SpreadsheetApp.flush();
}
