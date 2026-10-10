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

// Todo e-mail direcionado às pacientes deve sair do endereço real da Regiane.
// O endereço precisa estar confirmado em "Enviar e-mail como" na conta que
// executa este Apps Script; caso contrário, o envio falha em vez de usar
// silenciosamente o e-mail da Aline como remetente.
const PATIENT_SENDER_EMAIL = REGIANE_NOTIFICATION_EMAIL;
const PATIENT_SENDER_NAME = 'Regiane Silva';

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

function normalizeInstagramProfile_(value) {
  let profile = String(value || '').trim();
  if (!profile) return '';
  profile = profile.replace(/^https?:\/\/(?:www\.)?instagram\.com\//i, '');
  profile = profile.replace(/^@+/, '').split(/[\/?#]/)[0].trim();
  if (!/^[A-Za-z0-9._]{1,30}$/.test(profile)) {
    throw new Error('Informe um @ do Instagram ou um link de perfil válido.');
  }
  return '@' + profile.toLowerCase();
}

function normalizeInstagramPostKey_(value) {
  const raw = String(value || '').trim();
  if (!raw) throw new Error('Informe o link ou identificador da publicação.');
  const urlMatch = raw.match(/instagram\.com\/(?:p|reel|tv)\/([A-Za-z0-9_-]+)/i);
  const key = urlMatch ? urlMatch[1] : raw.replace(/^@+/, '');
  if (!/^[A-Za-z0-9_-]{3,100}$/.test(key)) {
    throw new Error('Link ou identificador da publicação inválido.');
  }
  return key;
}

function ensurePatientInstagramHeader_() {
  const sheet = getSheet('Pacientes');
  ensureHeaders_(sheet, ['Instagram']);
  return sheet;
}

function ensureHeaders_(sheet, headers) {
  const lastColumn = sheet.getLastColumn();
  const current = lastColumn
    ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    : [];
  while (current.length && String(current[current.length - 1]).trim() === '') current.pop();
  headers.forEach(function(header) {
    if (current.indexOf(header) === -1) {
      sheet.getRange(1, current.length + 1).setValue(header);
      current.push(header);
    }
  });
}

function ensureGamificationSheets_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let comentarios = ss.getSheetByName('Comentarios');
  if (!comentarios) comentarios = ss.insertSheet('Comentarios');
  ensureHeaders_(comentarios, ['Id', 'PostId', 'Email', 'Nome', 'Texto', 'DataHora', 'AvisoEnviadoEm', 'StatusAviso', 'PontosConcedidos']);

  let pontos = ss.getSheetByName('PontosLog');
  if (!pontos) pontos = ss.insertSheet('PontosLog');
  ensureHeaders_(pontos, ['Id', 'Email', 'Tipo', 'Pontos', 'Data', 'ChaveUnica', 'SaldoDepois']);

  let fotos = ss.getSheetByName('Fotos');
  if (!fotos) fotos = ss.insertSheet('Fotos');
  ensureHeaders_(fotos, ['Id', 'Email', 'Tipo', 'MessageId', 'AttachmentIndex', 'NomeArquivo', 'MimeType', 'Hash', 'DataHora', 'Pontos', 'Status']);

  let creditos = ss.getSheetByName('Creditos');
  if (!creditos) creditos = ss.insertSheet('Creditos');
  ensureHeaders_(creditos, ['Id', 'Email', 'Tipo', 'Valor', 'Descricao', 'Data', 'AdminEmail']);

  [comentarios, pontos, fotos, creditos].forEach(function(sheet) {
    sheet.setFrozenRows(1);
  });
  return { comentarios: comentarios, pontos: pontos, fotos: fotos, creditos: creditos };
}

function isValidPublicLink_(value) {
  const link = String(value || '').trim();
  const match = link.match(/^https?:\/\/([^/?#\s]+)(?:[/?#]\S*)?$/i);
  return !!match && match[1].indexOf('.') > 0;
}

function sendPatientEmail_(to, subject, body) {
  const sender = normEmail(PATIENT_SENDER_EMAIL);
  const aliases = GmailApp.getAliases().map(normEmail);
  if (aliases.indexOf(sender) === -1) {
    throw new Error('O endereço da Regiane ainda não está autorizado como remetente no Gmail.');
  }
  GmailApp.sendEmail(normEmail(to), subject, body, {
    from: sender,
    name: PATIENT_SENDER_NAME,
    replyTo: sender
  });
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

/**
 * Verifica o token do Google e garante que quem está chamando é uma
 * administradora (Regiane ou Aline). Lança erro se não for.
 */
function assertAdmin(idToken) {
  const auth = verifyGoogleToken(idToken);
  if (!isAdmin(auth.email)) throw new Error('Acesso restrito à administradora.');
  return auth;
}

function patientEmailFromToken_(body) {
  const auth = verifyGoogleToken(body.idToken);
  const requested = normEmail(body.email || body.targetEmail);
  if (requested && requested !== auth.email) {
    if (!isAdmin(auth.email)) throw new Error('Você não pode acessar os dados de outra paciente.');
    if (!findPatientRow(requested)) throw new Error('Paciente não encontrada.');
    return requested;
  }
  if (!isAdmin(auth.email) && !findPatientRow(auth.email)) throw new Error('Paciente não encontrada.');
  return requested || auth.email;
}

function sendRegianeEmailStrict_(assunto, corpo) {
  if (!REGIANE_NOTIFICATION_EMAIL || REGIANE_NOTIFICATION_EMAIL.indexOf('COLE_AQUI') !== -1) {
    throw new Error('E-mail de notificação da Regiane não configurado.');
  }
  MailApp.sendEmail({
    to: REGIANE_NOTIFICATION_EMAIL,
    subject: assunto,
    body: corpo,
    name: 'Site Regiane Silva'
  });
}

function notifyRegiane(assunto, corpo) {
  try {
    sendRegianeEmailStrict_(assunto, corpo);
    return true;
  } catch (err) {
    // não deixa o fluxo principal quebrar se o e-mail falhar
    return false;
  }
}

// ── ROTEAMENTO ──────────────────────────────────────────

function doGet(e) {
  try {
    const action = e.parameter.action;
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
    if (action === 'patientDashboard') return jsonResponse(actionPatientDashboard(body));
    if (action === 'patientComments') return jsonResponse(actionComments(body));
    if (action === 'comment') return jsonResponse(actionAddComment(body));
    if (action === 'bookSlot') return jsonResponse(actionBookSlot(body));
    if (action === 'checkAccess') return jsonResponse(actionCheckAccess(body));
    if (action === 'googleLoginPrograma') return jsonResponse(actionGoogleLoginPrograma(body));
    if (action === 'googleLoginCheckup') return jsonResponse(actionGoogleLoginCheckup(body));
    if (action === 'checkupDashboard') return jsonResponse(actionCheckupDashboard(body));
    if (action === 'submitCheckup') return jsonResponse(actionSubmitCheckup(body));
    if (action === 'googleLoginConsultoria') return jsonResponse(actionGoogleLoginConsultoria(body));
    if (action === 'requestConsultoriaAccessCode') return jsonResponse(actionRequestConsultoriaAccessCode(body));
    if (action === 'verifyConsultoriaAccessCode') return jsonResponse(actionVerifyConsultoriaAccessCode(body));
    if (action === 'consultoriaDashboard') return jsonResponse(actionConsultoriaDashboard(body));
    if (action === 'submitConsultoriaAnamnese') return jsonResponse(actionSubmitConsultoriaAnamnese(body));
    // A integração automática com o Asaas ainda não possui validação
    // criptográfica configurada. Não aceite payloads públicos que poderiam
    // liberar o Check-up para qualquer e-mail. A liberação segue disponível
    // com autenticação pelo painel administrativo e pela planilha.
    if (action === 'asaasWebhook') return jsonResponse({
      ok: false,
      error: 'Integração automática do Asaas ainda não configurada. Libere o acesso pelo painel administrativo.'
    });
    if (action === 'saveRecipe') return jsonResponse(actionSaveRecipe(body));
    if (action === 'uploadFoto') return jsonResponse(actionUploadFoto(body));
    if (action === 'patientPhotos') return jsonResponse(actionPatientPhotos(body));
    if (action === 'patientPhotoData') return jsonResponse(actionPatientPhotoData(body));
    if (action === 'saveInstagramProfile') return jsonResponse(actionSaveInstagramProfile(body));
    if (action === 'adminListPatients') return jsonResponse(actionAdminListPatients(body));
    if (action === 'adminAddPatient') return jsonResponse(actionAdminAddPatient(body));
    if (action === 'adminRemovePatient') return jsonResponse(actionAdminRemovePatient(body));
    if (action === 'adminSavePlan') return jsonResponse(actionAdminSavePlan(body));
    if (action === 'adminUploadPlanPdf') return jsonResponse(actionAdminUploadPlanPdf(body));
    if (action === 'adminAddMaterial') return jsonResponse(actionAdminAddMaterial(body));
    if (action === 'adminUseCredit') return jsonResponse(actionAdminUseCredit(body));
    if (action === 'adminRegisterInstagramEngagement') return jsonResponse(actionAdminRegisterInstagramEngagement(body));
    if (action === 'adminListRoutines') return jsonResponse(actionAdminListRoutines(body));
    if (action === 'adminSaveRoutine') return jsonResponse(actionAdminSaveRoutine(body));
    if (action === 'adminDeleteRoutine') return jsonResponse(actionAdminDeleteRoutine(body));
    if (action === 'patientRoutineData') return jsonResponse(actionPatientRoutineData(body));
    if (action === 'patientCompleteRoutine') return jsonResponse(actionPatientCompleteRoutine(body));
    if (action === 'patientReadNotification') return jsonResponse(actionPatientReadNotification(body));
    if (action === 'adminListPendingBookings') return jsonResponse(actionAdminListPendingBookings(body));
    if (action === 'adminConfirmBooking') return jsonResponse(actionAdminConfirmBooking(body));
    if (action === 'adminRejectBooking') return jsonResponse(actionAdminRejectBooking(body));
    if (action === 'adminListCheckupPatients') return jsonResponse(actionAdminListCheckupPatients(body));
    if (action === 'adminAddCheckupPatient') return jsonResponse(actionAdminAddCheckupPatient(body));
    if (action === 'adminSetCheckupAccess') return jsonResponse(actionAdminSetCheckupAccess(body));
    if (action === 'adminListConsultoriaPatients') return jsonResponse(actionAdminListConsultoriaPatients(body));
    if (action === 'adminAddConsultoriaPatient') return jsonResponse(actionAdminAddConsultoriaPatient(body));
    if (action === 'adminSetConsultoriaAccess') return jsonResponse(actionAdminSetConsultoriaAccess(body));
    if (action === 'bioLead') return jsonResponse(actionBioLead(body));
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
  ensurePatientInstagramHeader_();
  let p = findPatientRow(email);
  if (!p && isAdmin(email)) {
    // administradora sem cadastro de paciente: mostra um painel de exemplo, sem erro.
    p = { Nome: 'Administradora', PontosTotal: 0, RetornosRealizados: 0, ReceitasSalvas: 0, ProgressoPercent: 0, DataInicio: '', ProximoRetornoData: '', ProximoRetornoHora: '', PlanoTexto: 'Acesso de administradora — sem plano individual.' };
  }
  if (!p) return { ok: false, error: 'Paciente não encontrada.' };

  const materiais = sheetToObjects(getSheet('Materiais')).filter(m => {
    const dest = normEmail(m.Email);
    return (dest === normEmail(email) || dest === 'todos') && isValidPublicLink_(m.Link);
  });

  ensureGamificationSheets_();
  const pontosTotal = Number(p.PontosTotal) || 0;
  const pontosLog = sheetToObjects(getSheet('PontosLog')).filter(function(l) { return normEmail(l.Email) === normEmail(email); });
  const totalInteracoes = pontosLog.filter(function(l) {
    const tipo = String(l.Tipo || '').toLowerCase();
    return tipo.indexOf('comentário') === 0 || tipo.indexOf('foto') === 0 || tipo.indexOf('instagram') === 0;
  }).length;
  const credito = creditSummary_(email, pontosTotal);

  // próxima solicitação/consulta desta paciente (a mais próxima no futuro)
  const meusAgendamentos = sheetToObjects(getSheet('Agendamentos'))
    .filter(a => normEmail(a.Email) === normEmail(email) && a.IsoInicio && new Date(a.IsoInicio) > new Date())
    .sort((a, b) => new Date(a.IsoInicio) - new Date(b.IsoInicio));
  const proximoAgendamento = meusAgendamentos[0] || null;

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
    planoPdfUrl: p.PlanoPdfUrl || '',
    agendamentoStatus: proximoAgendamento ? String(proximoAgendamento.Status).trim().toLowerCase() : null,
    agendamentoData: proximoAgendamento ? proximoAgendamento.Data : '',
    agendamentoHora: proximoAgendamento ? proximoAgendamento.Hora : '',
    pontosTotal: pontosTotal,
    instagram: p.Instagram || '',
    totalInteracoes: totalInteracoes,
    pontosHistorico: pontosLog.sort(function(a, b) { return new Date(b.Data) - new Date(a.Data); }).slice(0, 30).map(function(l) {
      return { tipo: l.Tipo || 'Pontos', pontos: Number(l.Pontos) || 0, data: l.Data || '' };
    }),
    creditoGerado: credito.gerado,
    creditoUsado: credito.usado,
    creditoDisponivel: credito.disponivel,
    creditoHistorico: credito.historico,
    faltamParaProximoCredito: 100 - (pontosTotal % 100),
    materiais: materiais.map(m => ({
      tipo: m.Tipo, titulo: m.Titulo, descricao: m.Descricao, link: m.Link, area: m.Area || 'materiais'
    }))
  };
}

function actionPatientDashboard(body) {
  const email = patientEmailFromToken_(body);
  return actionDashboard(email);
}

function actionSaveInstagramProfile(body) {
  const email = patientEmailFromToken_(body);
  const instagram = normalizeInstagramProfile_(body.instagram);
  const sheet = ensurePatientInstagramHeader_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  const instagramCol = headers.indexOf('Instagram');
  for (let i = 1; i < data.length; i++) {
    if (normEmail(data[i][emailCol]) === normEmail(email)) {
      sheet.getRange(i + 1, instagramCol + 1).setValue(instagram);
      return { ok: true, instagram: instagram };
    }
  }
  return { ok: false, error: 'Paciente não encontrada.' };
}

/**
 * Salva o "salvar receita" de uma paciente — idempotente (a mesma receita
 * não conta pontos/contador duas vezes pra mesma pessoa) e atualiza o
 * contador ReceitasSalvas automaticamente, sem a Regiane precisar mexer.
 */
function actionSaveRecipe(body) {
  const email = patientEmailFromToken_(body);
  const recipeId = String(body.recipeId || '').trim();
  if (!recipeId) return { ok: false, error: 'Receita inválida.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
  const sheet = getSheet('Pacientes');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  const receitasCol = headers.indexOf('ReceitasSalvas');
  const idsCol = headers.indexOf('ReceitasSalvasIds');

  for (let i = 1; i < data.length; i++) {
    if (normEmail(data[i][emailCol]) === email) {
      const idsAtuais = String(data[i][idsCol] || '').split(',').map(s => s.trim()).filter(Boolean);
      if (idsAtuais.indexOf(recipeId) !== -1) {
        return { ok: true, jaSalva: true, receitasSalvas: Number(data[i][receitasCol]) || 0 };
      }
      idsAtuais.push(recipeId);
      const novoTotal = (Number(data[i][receitasCol]) || 0) + 1;
      sheet.getRange(i + 1, receitasCol + 1).setValue(novoTotal);
      sheet.getRange(i + 1, idsCol + 1).setValue(idsAtuais.join(','));
      addPointsUnlocked_(email, 'Receita salva: ' + (body.recipeTitle || recipeId), 2, 'receita:' + email + ':' + recipeId);
      return { ok: true, jaSalva: false, receitasSalvas: novoTotal };
    }
  }
  return { ok: false, error: 'Paciente não encontrada.' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * A paciente envia uma foto de resultado ou do prato — vai direto por e-mail
 * pra Regiane (anexada) e conta pontos automaticamente (5 pontos), sem ela
 * precisar mexer em nada. Usa MailApp em vez de Drive de propósito: evita
 * depender da autorização de Drive, que trava a implantação existente.
 */
function actionUploadFoto(body) {
  const email = patientEmailFromToken_(body);
  const p = findPatientRow(email);
  if (!p) return { ok: false, error: 'Paciente não encontrada.' };
  if (!body.fileBase64) return { ok: false, error: 'Nenhuma foto enviada.' };

  const mimeType = String(body.mimeType || 'image/jpeg').toLowerCase();
  if (mimeType.indexOf('image/') !== 0) return { ok: false, error: 'Envie somente uma imagem.' };
  const bytes = Utilities.base64Decode(body.fileBase64);
  if (bytes.length > 8 * 1024 * 1024) return { ok: false, error: 'A imagem deve ter no máximo 8 MB.' };
  const blob = Utilities.newBlob(bytes, mimeType, body.fileName || 'foto.jpg');

  const tipo = body.tipo === 'prato' ? 'Foto do prato' : 'Foto de resultado';
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes)
    .map(function(b) { return ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2); }).join('');
  const sheets = ensureGamificationSheets_();
  const existente = sheetToObjects(sheets.fotos).find(function(f) { return normEmail(f.Email) === email && String(f.Hash) === hash; });
  if (existente) return { ok: true, duplicada: true, pontosAdicionados: 0, error: 'Esta foto já foi enviada anteriormente.' };

  const fotosHoje = sheetToObjects(sheets.fotos).filter(function(f) {
    return normEmail(f.Email) === email && sameLocalDay_(f.DataHora, new Date());
  }).length;
  if (fotosHoje >= 5) return { ok: false, error: 'Limite diário de 5 fotos atingido. Tente novamente amanhã.' };

  const sender = normEmail(PATIENT_SENDER_EMAIL);
  if (GmailApp.getAliases().map(normEmail).indexOf(sender) === -1) throw new Error('O endereço da Regiane não está autorizado como remetente no Gmail.');
  const sent = GmailApp.createDraft(
    REGIANE_NOTIFICATION_EMAIL,
    'Nova ' + tipo.toLowerCase() + ' — ' + (p.Nome || email),
    (p.Nome || email) + ' (' + email + ') enviou uma ' + tipo.toLowerCase() + '. A foto também está disponível no painel da paciente.',
    { from: sender, name: PATIENT_SENDER_NAME, replyTo: sender, attachments: [blob] }
  ).send();

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const id = 'foto-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
    sheets.fotos.appendRow([id, email, tipo, sent.getId(), 0, blob.getName(), mimeType, hash, new Date(), 5, 'Enviada']);
    addPointsUnlocked_(email, tipo, 5, 'foto:' + email + ':' + hash);
    return { ok: true, id: id, duplicada: false, pontosAdicionados: 5 };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Checagem autenticada usada pelo site pra saber, de tempos em tempos
 * enquanto a página está aberta, se aquele e-mail ainda está cadastrado —
 * se a Regiane excluir a linha da planilha, o site desloga sozinho.
 */
function actionCheckAccess(body) {
  const auth = verifyGoogleToken(body.idToken);
  if (isAdmin(auth.email)) return { ok: true };
  if (body.area === 'checkup') {
    const email = auth.email;
    const c = findCheckupRow(email);
    if (!c || String(c.Liberado).trim().toLowerCase() !== 'sim') return { ok: false };
    return { ok: true };
  }
  return { ok: !!findPatientRow(auth.email) };
}

// ── PAINEL DA ADMINISTRADORA ─────────────────────────────

/**
 * Lista todas as pacientes do Programa pra administradora — usado no
 * "Painel da Regiane" dentro do próprio site, pra ela ver quem já tem
 * plano alimentar cadastrado e quem ainda falta.
 */
function actionAdminListPatients(body) {
  assertAdmin(body.idToken);
  ensurePatientInstagramHeader_();
  ensureGamificationSheets_();
  const pacientes = sheetToObjects(getSheet('Pacientes'));
  return {
    ok: true,
    pacientes: pacientes.map(function(p) {
      const pontosTotal = Number(p.PontosTotal) || 0;
      const credito = creditSummary_(p.Email, pontosTotal);
      const fotos = sheetToObjects(getSheet('Fotos')).filter(function(f) { return normEmail(f.Email) === normEmail(p.Email); }).length;
      return {
        email: p.Email,
        nome: p.Nome,
        instagram: p.Instagram || '',
        diasAcompanhamento: daysSince(p.DataInicio),
        retornosRealizados: Number(p.RetornosRealizados) || 0,
        receitasSalvas: Number(p.ReceitasSalvas) || 0,
        pontosTotal: pontosTotal,
        creditoGerado: credito.gerado,
        creditoUsado: credito.usado,
        creditoDisponivel: credito.disponivel,
        creditoHistorico: credito.historico,
        fotosTotal: fotos,
        progressoPercent: Number(p.ProgressoPercent) || 0,
        temPlano: !!(p.PlanoTexto && String(p.PlanoTexto).trim()),
        planoTexto: p.PlanoTexto || '',
        planoPdfUrl: p.PlanoPdfUrl || ''
      };
    })
  };
}

function creditSummary_(email, pontosTotal) {
  ensureGamificationSheets_();
  const historico = sheetToObjects(getSheet('Creditos')).filter(function(c) { return normEmail(c.Email) === normEmail(email); })
    .sort(function(a, b) { return new Date(b.Data) - new Date(a.Data); });
  let usado = 0;
  historico.forEach(function(c) {
    const valor = Math.abs(Number(c.Valor) || 0);
    const tipo = String(c.Tipo || 'Uso').trim().toLowerCase();
    if (tipo === 'estorno') usado -= valor;
    else if (tipo === 'uso') usado += valor;
  });
  usado = Math.max(0, usado);
  const gerado = Math.floor((Number(pontosTotal) || 0) / 100) * 10;
  return {
    gerado: gerado,
    usado: usado,
    disponivel: Math.max(0, gerado - usado),
    historico: historico.slice(0, 20).map(function(c) {
      return { id: String(c.Id), tipo: c.Tipo || 'Uso', valor: Number(c.Valor) || 0, descricao: c.Descricao || '', data: c.Data || '' };
    })
  };
}

function actionAdminUseCredit(body) {
  const admin = assertAdmin(body.idToken);
  const email = normEmail(body.email);
  const paciente = findPatientRow(email);
  if (!paciente) return { ok: false, error: 'Paciente não encontrada.' };
  const valor = Number(String(body.valor || '').replace(',', '.'));
  if (!isFinite(valor) || valor <= 0) return { ok: false, error: 'Informe um valor de crédito maior que zero.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const resumo = creditSummary_(email, Number(paciente.PontosTotal) || 0);
    if (valor > resumo.disponivel) return { ok: false, error: 'Crédito insuficiente. Disponível: R$ ' + resumo.disponivel.toFixed(2).replace('.', ',') };
    const sheet = ensureGamificationSheets_().creditos;
    const id = 'credito-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
    sheet.appendRow([id, email, 'Uso', valor, String(body.descricao || 'Crédito utilizado').trim(), new Date(), admin.email]);
    const atualizado = creditSummary_(email, Number(paciente.PontosTotal) || 0);
    return { ok: true, creditoGerado: atualizado.gerado, creditoUsado: atualizado.usado, creditoDisponivel: atualizado.disponivel };
  } finally {
    lock.releaseLock();
  }
}

function actionAdminRegisterInstagramEngagement(body) {
  const admin = assertAdmin(body.idToken);
  const email = normEmail(body.email);
  ensurePatientInstagramHeader_();
  const paciente = findPatientRow(email);
  if (!paciente) return { ok: false, error: 'Paciente não encontrada.' };
  const instagram = normalizeInstagramProfile_(paciente.Instagram);
  if (!instagram) return { ok: false, error: 'A paciente ainda não cadastrou o Instagram.' };
  const postKey = normalizeInstagramPostKey_(body.postKey);
  const liked = body.liked === true || String(body.liked).toLowerCase() === 'true';
  const commented = body.commented === true || String(body.commented).toLowerCase() === 'true';
  if (!liked && !commented) return { ok: false, error: 'Marque curtida e/ou comentário.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const awarded = [];
    const duplicate = [];
    if (liked) {
      const result = addPointsUnlocked_(email, 'Instagram — curtida', 5, 'instagram:curtida:' + postKey + ':' + email);
      (result.added ? awarded : duplicate).push('curtida');
    }
    if (commented) {
      const result = addPointsUnlocked_(email, 'Instagram — comentário', 5, 'instagram:comentario:' + postKey + ':' + email);
      (result.added ? awarded : duplicate).push('comentário');
    }
    const atualizado = findPatientRow(email);
    return {
      ok: true,
      instagram: instagram,
      postKey: postKey,
      pontosAdicionados: awarded.length * 5,
      registrados: awarded,
      duplicados: duplicate,
      pontosTotal: Number(atualizado && atualizado.PontosTotal) || 0,
      registradoPor: admin.email
    };
  } finally {
    lock.releaseLock();
  }
}

function actionAdminAddPatient(body) {
  assertAdmin(body.idToken);
  const email = normEmail(body.email);
  const nome = String(body.nome || '').trim();
  const instagram = normalizeInstagramProfile_(body.instagram);
  if (!email || email.indexOf('@') === -1 || !nome) {
    return { ok: false, error: 'Informe o nome e um e-mail válido.' };
  }

  const sheet = ensurePatientInstagramHeader_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  const nomeCol = headers.indexOf('Nome');
  const instagramCol = headers.indexOf('Instagram');
  for (let i = 1; i < data.length; i++) {
    if (normEmail(data[i][emailCol]) === email) {
      sheet.getRange(i + 1, nomeCol + 1).setValue(nome);
      if (instagram) sheet.getRange(i + 1, instagramCol + 1).setValue(instagram);
      return { ok: true, created: false };
    }
  }

  const row = headers.map(function() { return ''; });
  row[emailCol] = email;
  row[nomeCol] = nome;
  row[headers.indexOf('DataInicio')] = new Date();
  row[headers.indexOf('RetornosRealizados')] = 0;
  row[headers.indexOf('ReceitasSalvas')] = 0;
  row[headers.indexOf('ProgressoPercent')] = 0;
  row[headers.indexOf('PontosTotal')] = 0;
  row[instagramCol] = instagram;
  sheet.appendRow(row);
  return { ok: true, created: true };
}

function actionAdminRemovePatient(body) {
  assertAdmin(body.idToken);
  const email = normEmail(body.email);
  const sheet = getSheet('Pacientes');
  const data = sheet.getDataRange().getValues();
  const emailCol = data[0].indexOf('Email');
  for (let i = 1; i < data.length; i++) {
    if (normEmail(data[i][emailCol]) === email) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Paciente não encontrada.' };
}

/**
 * A administradora adiciona/atualiza o plano alimentar de uma paciente
 * direto pelo site — a planilha (coluna PlanoTexto) atualiza sozinha,
 * sem ela precisar editar a planilha na mão.
 */
function actionAdminSavePlan(body) {
  assertAdmin(body.idToken);
  const email = normEmail(body.email);
  const sheet = getSheet('Pacientes');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  const planoCol = headers.indexOf('PlanoTexto');

  for (let i = 1; i < data.length; i++) {
    if (normEmail(data[i][emailCol]) === email) {
      sheet.getRange(i + 1, planoCol + 1).setValue(body.planoTexto || '');
      createImmediateNotification_(email, 'Seu plano foi atualizado', 'A Regiane publicou uma nova orientação no seu plano. Acesse a área de membros para conferir.', 'plano');
      return { ok: true };
    }
  }
  return { ok: false, error: 'Paciente não encontrada.' };
}

/**
 * A administradora envia o plano alimentar em PDF direto pelo site.
 * O arquivo vai pro Google Drive dela e o link fica salvo na planilha,
 * na coluna PlanoPdfUrl — a paciente vê um botão de baixar no Meu Plano.
 */
function actionAdminUploadPlanPdf(body) {
  assertAdmin(body.idToken);
  const email = normEmail(body.email);
  const sheet = getSheet('Pacientes');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  const pdfCol = headers.indexOf('PlanoPdfUrl');

  for (let i = 1; i < data.length; i++) {
    if (normEmail(data[i][emailCol]) === email) {
      const bytes = Utilities.base64Decode(body.fileBase64);
      const blob = Utilities.newBlob(bytes, 'application/pdf', body.fileName || 'plano.pdf');
      const file = DriveApp.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const url = file.getUrl();
      sheet.getRange(i + 1, pdfCol + 1).setValue(url);
      return { ok: true, url: url };
    }
  }
  return { ok: false, error: 'Paciente não encontrada.' };
}

/**
 * A administradora adiciona um material (vídeo, PDF, link etc.) direto pelo
 * site — vira uma linha na aba Materiais. Se "email" vier vazio ou como
 * "TODOS", o material aparece pra todas as pacientes do Programa; senão,
 * só pra quem tem aquele e-mail cadastrado.
 */
function actionAdminAddMaterial(body) {
  assertAdmin(body.idToken);
  const link = String(body.link || '').trim();
  const titulo = String(body.titulo || '').trim();
  if (!link || !titulo) return { ok: false, error: 'Preencha pelo menos o título e o link.' };
  if (!isValidPublicLink_(link)) return { ok: false, error: 'Informe um link completo e válido, começando com https://.' };

  const areasValidas = ['aulas', 'rotulo', 'materiais', 'consultoria'];
  const area = areasValidas.indexOf(body.area) !== -1 ? body.area : 'materiais';
  const tipoPadrao = area === 'materiais' ? 'Material' : 'Vídeo';

  const destino = normEmail(body.email) || 'todos';
  const sheet = getSheet('Materiais');
  const id = new Date().getTime();
  sheet.appendRow([id, destino === 'todos' ? 'TODOS' : destino, body.tipo || tipoPadrao, titulo, body.descricao || '', link, area]);

  return { ok: true };
}

// ── ROTINAS E NOTIFICAÇÕES INDIVIDUAIS ──────────────────

const NOTIFICATION_TIMEZONE = 'America/Sao_Paulo';

function ensureNotificationSheets_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let rotinas = ss.getSheetByName('Rotinas');
  if (!rotinas) {
    rotinas = ss.insertSheet('Rotinas');
    rotinas.appendRow(['Id', 'Email', 'Titulo', 'Categoria', 'MetaValor', 'MetaUnidade', 'DiasSemana', 'Horarios', 'Mensagem', 'Ativo', 'DataInicio', 'DataFim', 'CriadoEm', 'AtualizadoEm']);
    rotinas.setFrozenRows(1);
  }
  let notificacoes = ss.getSheetByName('Notificacoes');
  if (!notificacoes) {
    notificacoes = ss.insertSheet('Notificacoes');
    notificacoes.appendRow(['Id', 'RotinaId', 'Email', 'Titulo', 'Mensagem', 'AgendadaPara', 'EnviadaEm', 'LidaEm', 'ConcluidaEm', 'Status']);
    notificacoes.setFrozenRows(1);
  }
  return { rotinas: rotinas, notificacoes: notificacoes };
}

function routineObjects_() {
  return sheetToObjects(ensureNotificationSheets_().rotinas);
}

function notificationObjects_() {
  return sheetToObjects(ensureNotificationSheets_().notificacoes);
}

function normalizeTimes_(value) {
  // A planilha converte horários únicos em Date; preserve o horário exibido.
  if (value instanceof Date && !isNaN(value.getTime())) {
    value = Utilities.formatDate(value, NOTIFICATION_TIMEZONE, 'HH:mm');
  }
  const times = String(value || '').split(/[;,\s]+/).map(function(v) { return v.trim(); }).filter(Boolean);
  const valid = [];
  times.forEach(function(v) {
    const m = v.match(/^(\d{1,2}):(\d{2})$/);
    if (!m) return;
    const h = Number(m[1]);
    const min = Number(m[2]);
    if (h >= 0 && h <= 23 && min >= 0 && min <= 59) valid.push(('0' + h).slice(-2) + ':' + ('0' + min).slice(-2));
  });
  return Array.from(new Set(valid)).sort().join(',');
}

function notificationDateKey_(value) {
  if (!value) return '';
  const text = String(value).trim();
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return iso[1] + iso[2] + iso[3];
  const date = value instanceof Date ? value : new Date(value);
  return isNaN(date.getTime()) ? '' : Utilities.formatDate(date, NOTIFICATION_TIMEZONE, 'yyyyMMdd');
}

function isRoutineActiveToday_(routine) {
  if (String(routine.Ativo).trim().toLowerCase() === 'não') return false;
  const todayKey = Utilities.formatDate(new Date(), NOTIFICATION_TIMEZONE, 'yyyyMMdd');
  const startKey = notificationDateKey_(routine.DataInicio);
  const endKey = notificationDateKey_(routine.DataFim);
  return !(startKey && todayKey < startKey) && !(endKey && todayKey > endKey);
}

function actionAdminListRoutines(body) {
  assertAdmin(body.idToken);
  const email = normEmail(body.email);
  return {
    ok: true,
    rotinas: routineObjects_().filter(function(r) { return !email || normEmail(r.Email) === email; }).map(function(r) {
      return {
        id: String(r.Id), email: normEmail(r.Email), titulo: r.Titulo || '', categoria: r.Categoria || '',
        metaValor: r.MetaValor || '', metaUnidade: r.MetaUnidade || '', diasSemana: r.DiasSemana || '',
        horarios: r.Horarios || '', mensagem: r.Mensagem || '', ativo: String(r.Ativo).toLowerCase() !== 'não',
        dataInicio: r.DataInicio || '', dataFim: r.DataFim || ''
      };
    })
  };
}

function actionAdminSaveRoutine(body) {
  assertAdmin(body.idToken);
  const email = normEmail(body.email);
  const paciente = findPatientRow(email);
  if (!paciente) return { ok: false, error: 'Paciente não encontrada na aba Pacientes.' };
  const titulo = String(body.titulo || '').trim();
  if (!titulo) return { ok: false, error: 'Informe o título da orientação.' };
  const horarios = normalizeTimes_(body.horarios);
  if (!horarios) return { ok: false, error: 'Informe pelo menos um horário no formato HH:MM.' };
  const sheets = ensureNotificationSheets_();
  const sheet = sheets.rotinas;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const id = String(body.id || new Date().getTime());
  const values = {
    Id: id, Email: email, Titulo: titulo, Categoria: body.categoria || 'Rotina',
    MetaValor: body.metaValor || '', MetaUnidade: body.metaUnidade || '',
    DiasSemana: body.diasSemana || 'SEG,TER,QUA,QUI,SEX,SAB,DOM', Horarios: horarios,
    Mensagem: body.mensagem || ('Está na hora de: ' + titulo), Ativo: body.ativo === false ? 'Não' : 'Sim',
    DataInicio: body.dataInicio || '', DataFim: body.dataFim || '', AtualizadoEm: new Date()
  };
  let updated = false;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][headers.indexOf('Id')]) === id) {
      headers.forEach(function(h, idx) { if (values[h] !== undefined) sheet.getRange(i + 1, idx + 1).setValue(values[h]); });
      updated = true;
      break;
    }
  }
  if (!updated) {
    values.CriadoEm = new Date();
    sheet.appendRow(headers.map(function(h) { return values[h] !== undefined ? values[h] : ''; }));
  }
  const triggerConfigured = ensureNotificationTrigger_();
  createImmediateNotification_(email, 'Nova orientação disponível', 'A Regiane adicionou "' + titulo + '" à sua rotina. Acesse a área de membros para ver os horários e detalhes.', 'rotina-' + id);
  return { ok: true, id: id, triggerConfigured: triggerConfigured };
}

function actionAdminDeleteRoutine(body) {
  assertAdmin(body.idToken);
  const id = String(body.id || '');
  const sheet = ensureNotificationSheets_().rotinas;
  const data = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('Id');
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idCol]) === id) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Orientação não encontrada.' };
}

function actionPatientRoutineData(body) {
  const email = patientEmailFromToken_(body);
  const routines = routineObjects_().filter(function(r) {
    return normEmail(r.Email) === email && isRoutineActiveToday_(r);
  }).map(function(r) {
    return {
      id: String(r.Id), titulo: r.Titulo || '', categoria: r.Categoria || '', metaValor: r.MetaValor || '',
      metaUnidade: r.MetaUnidade || '', diasSemana: r.DiasSemana || '', horarios: r.Horarios || '', mensagem: r.Mensagem || ''
    };
  });
  const notifications = notificationObjects_().filter(function(n) { return normEmail(n.Email) === email; })
    .sort(function(a, b) { return new Date(b.AgendadaPara) - new Date(a.AgendadaPara); }).slice(0, 50)
    .map(function(n) {
      return { id: String(n.Id), rotinaId: String(n.RotinaId || ''), titulo: n.Titulo || '', mensagem: n.Mensagem || '',
        agendadaPara: n.AgendadaPara || '', lida: !!n.LidaEm, concluida: !!n.ConcluidaEm, status: n.Status || '' };
    });
  return { ok: true, email: email, rotinas: routines, notificacoes: notifications, naoLidas: notifications.filter(function(n) { return !n.lida; }).length };
}

function actionPatientCompleteRoutine(body) {
  const email = patientEmailFromToken_(body);
  const routineId = String(body.routineId || '');
  const routine = routineObjects_().find(function(r) { return String(r.Id) === routineId && normEmail(r.Email) === email; });
  if (!routine) return { ok: false, error: 'Orientação não encontrada.' };
  const sheets = ensureNotificationSheets_();
  const id = 'feito-' + routineId + '-' + Utilities.formatDate(new Date(), NOTIFICATION_TIMEZONE, 'yyyyMMdd');
  const data = sheets.notificacoes.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('Id');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === id) {
      sheets.notificacoes.getRange(i + 1, headers.indexOf('ConcluidaEm') + 1).setValue(new Date());
      sheets.notificacoes.getRange(i + 1, headers.indexOf('Status') + 1).setValue('Concluída');
      return { ok: true };
    }
  }
  sheets.notificacoes.appendRow([id, routineId, email, routine.Titulo, body.observacao || 'Atividade marcada como realizada pela paciente.', new Date(), '', new Date(), new Date(), 'Concluída']);
  return { ok: true };
}

function actionPatientReadNotification(body) {
  const email = patientEmailFromToken_(body);
  const id = String(body.id || '');
  const sheet = ensureNotificationSheets_().notificacoes;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][headers.indexOf('Id')]) === id && normEmail(data[i][headers.indexOf('Email')]) === email) {
      sheet.getRange(i + 1, headers.indexOf('LidaEm') + 1).setValue(new Date());
      return { ok: true };
    }
  }
  return { ok: false, error: 'Notificação não encontrada.' };
}

function createImmediateNotification_(email, titulo, mensagem, routineId) {
  const sheets = ensureNotificationSheets_();
  const id = 'imediata-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
  const now = new Date();
  sheets.notificacoes.appendRow([id, routineId || '', normEmail(email), titulo, mensagem, now, now, '', '', 'Enviada']);
  const notificationRow = sheets.notificacoes.getLastRow();
  try {
    sendPatientEmail_(
      email,
      titulo + ' — Regiane Silva',
      mensagem + '\n\nAcesse sua área de membros: https://regianemariasilva22-sudo.github.io/site-regiane-silva/membros/programa/login.html'
    );
  } catch (err) {
    sheets.notificacoes.getRange(notificationRow, 10).setValue('Erro no e-mail: ' + String(err));
  }
  return id;
}

function ensureNotificationTrigger_() {
  try {
    const exists = ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === 'processScheduledNotifications'; });
    if (!exists) ScriptApp.newTrigger('processScheduledNotifications').timeBased().everyMinutes(5).create();
    return true;
  } catch (err) {
    return false;
  }
}

function processScheduledNotifications() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    const now = new Date();
    const dateKey = Utilities.formatDate(now, NOTIFICATION_TIMEZONE, 'yyyyMMdd');
    const dayCode = ['DOM', 'SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SAB'][Number(Utilities.formatDate(now, NOTIFICATION_TIMEZONE, 'u')) % 7];
    const currentMinutes = Number(Utilities.formatDate(now, NOTIFICATION_TIMEZONE, 'H')) * 60 + Number(Utilities.formatDate(now, NOTIFICATION_TIMEZONE, 'm'));
    const sentIds = {};
    notificationObjects_().forEach(function(n) { sentIds[String(n.Id)] = true; });
    routineObjects_().forEach(function(r) {
      if (!isRoutineActiveToday_(r)) return;
      const days = String(r.DiasSemana || '').split(',').map(function(v) { return v.trim().toUpperCase(); });
      if (days.indexOf(dayCode) === -1) return;
      normalizeTimes_(r.Horarios).split(',').filter(Boolean).forEach(function(time) {
        const parts = time.split(':');
        const scheduledMinutes = Number(parts[0]) * 60 + Number(parts[1]);
        const delta = currentMinutes - scheduledMinutes;
        if (delta < 0 || delta > 9) return;
        const id = 'rotina-' + r.Id + '-' + dateKey + '-' + time.replace(':', '');
        if (sentIds[id]) return;
        const when = new Date(now.getTime() - delta * 60000);
        const sheets = ensureNotificationSheets_();
        let status = 'Enviada';
        try {
          sendPatientEmail_(
            r.Email,
            (r.Titulo || 'Lembrete') + ' — Regiane Silva',
            (r.Mensagem || ('Está na hora de: ' + r.Titulo)) + '\n\nConfira sua rotina na área de membros: https://regianemariasilva22-sudo.github.io/site-regiane-silva/membros/programa/login.html'
          );
        } catch (err) { status = 'Erro no e-mail: ' + String(err); }
        sheets.notificacoes.appendRow([id, r.Id, normEmail(r.Email), r.Titulo, r.Mensagem, when, status === 'Enviada' ? now : '', '', '', status]);
        sentIds[id] = true;
      });
    });
    processPendingBioLeadNotifications_();
  } finally {
    lock.releaseLock();
  }
}

/** Teste controlado solicitado pela Aline. Cria uma paciente simbólica,
 * envia um lembrete real e limita a rotina ao dia do teste para não gerar spam. */
function testarNotificacoesBabados() {
  const email = 'babadosdaaline@gmail.com';
  const patientSheet = getSheet('Pacientes');
  if (!findPatientRow(email)) {
    const headers = patientSheet.getRange(1, 1, 1, patientSheet.getLastColumn()).getValues()[0];
    const values = { Email: email, Nome: 'Maria — Teste de Notificações', DataInicio: new Date(), PlanoTexto: 'Paciente simbólica para validar as notificações.', PontosTotal: 0 };
    patientSheet.appendRow(headers.map(function(h) { return values[h] !== undefined ? values[h] : ''; }));
  }
  const sheets = ensureNotificationSheets_();
  const id = 'teste-agua-babados';
  const existing = routineObjects_().find(function(r) { return String(r.Id) === id; });
  if (existing) {
    const data = sheets.rotinas.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('Id');
    for (let i = 1; i < data.length; i++) if (String(data[i][idCol]) === id) sheets.rotinas.deleteRow(i + 1);
  }
  const today = Utilities.formatDate(new Date(), NOTIFICATION_TIMEZONE, 'yyyy-MM-dd');
  sheets.rotinas.appendRow([id, email, 'Meta de água', 'Água', 2, 'litros por dia', 'SEG,TER,QUA,QUI,SEX,SAB,DOM', '08:00,10:00,12:00,14:00,16:00,18:00,20:00', 'Maria, está na hora de tomar sua água. Sua meta diária é de 2 litros.', 'Sim', today, today, new Date(), new Date()]);
  const notificationId = createImmediateNotification_(email, 'Meta de água — lembrete de teste', 'Maria, está na hora de tomar sua água. Sua meta diária é de 2 litros. Este é o teste solicitado pela Aline.', id);
  const triggerConfigured = ensureNotificationTrigger_();
  const notification = notificationObjects_().find(function(n) { return String(n.Id) === notificationId; });
  return {
    ok: !!notification && String(notification.Status) === 'Enviada',
    email: email,
    patient: 'Maria — Teste de Notificações',
    routineId: id,
    notificationId: notificationId,
    notificationStatus: notification ? notification.Status : 'Não encontrada',
    triggerConfigured: triggerConfigured,
    routineActiveToday: true,
    routineEndsToday: true
  };
}

// ── COMUNIDADE / COMENTÁRIOS / FOTOS / PONTOS ──────────

function sameLocalDay_(a, b) {
  const da = a instanceof Date ? a : new Date(a);
  const db = b instanceof Date ? b : new Date(b);
  if (isNaN(da.getTime()) || isNaN(db.getTime())) return false;
  return Utilities.formatDate(da, NOTIFICATION_TIMEZONE, 'yyyyMMdd') === Utilities.formatDate(db, NOTIFICATION_TIMEZONE, 'yyyyMMdd');
}

function actionComments(body) {
  patientEmailFromToken_(body);
  const postId = body.postId;
  const all = sheetToObjects(ensureGamificationSheets_().comentarios);
  const filtered = postId ? all.filter(c => String(c.PostId) === String(postId)) : all;
  return {
    ok: true,
    comments: filtered.map(c => ({ id: String(c.Id), nome: c.Nome, texto: c.Texto, dataHora: c.DataHora }))
  };
}

function actionCommentsLegacy_(postId) {
  const all = sheetToObjects(ensureGamificationSheets_().comentarios);
  const filtered = postId ? all.filter(function(c) { return String(c.PostId) === String(postId); }) : all;
  return { ok: true, comments: filtered.map(function(c) { return { nome: c.Nome, texto: c.Texto, dataHora: c.DataHora }; }) };
}

function actionAddComment(body) {
  const email = patientEmailFromToken_(body);
  const p = findPatientRow(email);
  if (!p) return { ok: false, error: 'Paciente não encontrada.' };
  const texto = String(body.texto || '').trim();
  if (!texto) return { ok: false, error: 'Comentário vazio.' };
  if (texto.length > 1000) return { ok: false, error: 'O comentário deve ter no máximo 1.000 caracteres.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let id;
  let pontosConcedidos = 0;
  let row;
  try {
    const sheets = ensureGamificationSheets_();
    const anteriores = sheetToObjects(sheets.comentarios).filter(function(c) { return normEmail(c.Email) === email; });
    const duplicadoRecente = anteriores.some(function(c) {
      const quando = new Date(c.DataHora);
      return String(c.PostId) === String(body.postId) && String(c.Texto || '').trim().toLowerCase() === texto.toLowerCase() &&
        !isNaN(quando.getTime()) && (new Date().getTime() - quando.getTime()) < 24 * 60 * 60 * 1000;
    });
    const comentariosHoje = anteriores.filter(function(c) { return sameLocalDay_(c.DataHora, new Date()) && Number(c.PontosConcedidos) > 0; }).length;

    id = 'comentario-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
    pontosConcedidos = (!duplicadoRecente && comentariosHoje < 10) ? 3 : 0;
    sheets.comentarios.appendRow([id, body.postId, email, p.Nome, texto, new Date(), '', 'Pendente', pontosConcedidos]);
    row = sheets.comentarios.getLastRow();
    if (pontosConcedidos) addPointsUnlocked_(email, 'Comentário na comunidade', pontosConcedidos, id);
  } finally {
    lock.releaseLock();
  }

  const enviado = notifyRegiane(
    'Novo comentário ou dúvida — ' + (p.Nome || email),
    (p.Nome || email) + ' (' + email + ') escreveu na comunidade:\n\n' + texto + '\n\nAcesse o painel da área de membros para acompanhar.'
  );
  const comentarios = ensureGamificationSheets_().comentarios;
  comentarios.getRange(row, 7).setValue(enviado ? new Date() : '');
  comentarios.getRange(row, 8).setValue(enviado ? 'Enviado' : 'Erro no envio');

  return { ok: true, id: id, nome: p.Nome, pontosAdicionados: pontosConcedidos, avisoRegianeEnviado: enviado };
}

function addPointsUnlocked_(email, tipo, pontos, chaveUnica) {
  const sheets = ensureGamificationSheets_();
  const key = String(chaveUnica || '').trim();
  if (key) {
    const already = sheetToObjects(sheets.pontos).some(function(l) { return String(l.ChaveUnica || '') === key; });
    if (already) return { added: false, total: null };
  }
  const pSheet = getSheet('Pacientes');
  const data = pSheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  const pontosCol = headers.indexOf('PontosTotal');
  for (let i = 1; i < data.length; i++) {
    if (normEmail(data[i][emailCol]) === normEmail(email)) {
      const atual = Number(data[i][pontosCol]) || 0;
      const novoTotal = atual + Number(pontos || 0);
      pSheet.getRange(i + 1, pontosCol + 1).setValue(novoTotal);
      sheets.pontos.appendRow([new Date().getTime(), normEmail(email), tipo, Number(pontos) || 0, new Date(), key, novoTotal]);
      return { added: true, total: novoTotal };
    }
  }
  return { added: false, total: null };
}

function addPoints(email, tipo, pontos, chaveUnica) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return addPointsUnlocked_(email, tipo, pontos, chaveUnica);
  } finally {
    lock.releaseLock();
  }
}

function actionPatientPhotos(body) {
  const email = patientEmailFromToken_(body);
  const fotos = sheetToObjects(ensureGamificationSheets_().fotos)
    .filter(function(f) { return normEmail(f.Email) === email; })
    .sort(function(a, b) { return new Date(b.DataHora) - new Date(a.DataHora); })
    .slice(0, 50)
    .map(function(f) {
      return { id: String(f.Id), tipo: f.Tipo || 'Foto', nomeArquivo: f.NomeArquivo || 'foto', mimeType: f.MimeType || 'image/jpeg', dataHora: f.DataHora || '', pontos: Number(f.Pontos) || 0 };
    });
  return { ok: true, email: email, fotos: fotos };
}

function actionPatientPhotoData(body) {
  const email = patientEmailFromToken_(body);
  const foto = sheetToObjects(ensureGamificationSheets_().fotos).find(function(f) {
    return String(f.Id) === String(body.id) && normEmail(f.Email) === email;
  });
  if (!foto) return { ok: false, error: 'Foto não encontrada.' };
  const message = GmailApp.getMessageById(String(foto.MessageId || ''));
  if (!message) return { ok: false, error: 'O arquivo da foto não está mais disponível.' };
  const attachments = message.getAttachments({ includeInlineImages: false, includeAttachments: true });
  const index = Number(foto.AttachmentIndex) || 0;
  if (!attachments[index]) return { ok: false, error: 'O anexo da foto não foi encontrado.' };
  const blob = attachments[index];
  return { ok: true, id: String(foto.Id), mimeType: blob.getContentType(), fileBase64: Utilities.base64Encode(blob.getBytes()) };
}

// ── AGENDA (conectada de verdade ao Google Agenda da Regiane) ────
//
// A Regiane precisa compartilhar o Google Agenda dela com a conta que
// publica o backend (quem "Executa como" no deploy — veja o topo do
// arquivo), com permissão de "Fazer alterações em eventos". Sem isso,
// esta seção não consegue ler os horários livres nem criar a consulta.

const REGIANE_CALENDAR_ID = REGIANE_NOTIFICATION_EMAIL; // e-mail do Google Agenda da Regiane
const AGENDA_DIAS_A_FRENTE = 14;   // até quantos dias no futuro mostrar horários
const AGENDA_HORA_INICIO = 9;      // agenda abre às 09:00
const AGENDA_HORA_FIM = 18;        // último horário considerado antes das 18:00
const AGENDA_DURACAO_MIN = 60;     // duração de cada consulta, em minutos
const FUSO_AGENDA = 'GMT-3';       // horário de Brasília

function getRegianeCalendar() {
  const cal = CalendarApp.getCalendarById(REGIANE_CALENDAR_ID);
  if (!cal) {
    throw new Error('Não foi possível acessar o Google Agenda da Regiane. Ela precisa compartilhar a agenda (com permissão de fazer alterações) com quem publicou o site.');
  }
  return cal;
}

/**
 * Gera os horários livres olhando de verdade o Google Agenda da Regiane:
 * dias úteis, dentro do horário comercial configurado acima, excluindo
 * qualquer horário que já tenha um evento (compromisso) na agenda dela.
 */
function actionSlots() {
  const cal = getRegianeCalendar();
  const agora = new Date();
  const fim = new Date();
  fim.setDate(fim.getDate() + AGENDA_DIAS_A_FRENTE);
  const eventos = cal.getEvents(agora, fim);

  const slots = [];
  for (let d = 0; d < AGENDA_DIAS_A_FRENTE; d++) {
    const dia = new Date();
    dia.setDate(dia.getDate() + d);
    const diaSemana = dia.getDay();
    if (diaSemana === 0 || diaSemana === 6) continue; // pula sábado e domingo

    for (let h = AGENDA_HORA_INICIO; h < AGENDA_HORA_FIM; h++) {
      const inicio = new Date(dia.getFullYear(), dia.getMonth(), dia.getDate(), h, 0, 0);
      if (inicio <= agora) continue;
      const termino = new Date(inicio.getTime() + AGENDA_DURACAO_MIN * 60000);

      const ocupado = eventos.some(ev => ev.getStartTime() < termino && ev.getEndTime() > inicio);
      if (!ocupado) {
        slots.push({
          data: Utilities.formatDate(inicio, FUSO_AGENDA, 'dd/MM/yyyy'),
          hora: Utilities.formatDate(inicio, FUSO_AGENDA, 'HH:mm'),
          iso: inicio.toISOString()
        });
      }
    }
  }

  return { ok: true, slots: slots };
}

// Cores de destaque na aba Agendamentos, pra Regiane bater o olho e ver
// na hora quem está pedindo horário e quem já foi confirmado.
const COR_SOLICITADO = '#FFF3CD';
const COR_CONFIRMADO = '#D4EDDA';

/**
 * A paciente SOLICITA um horário — isso não cria o evento na agenda
 * ainda. Fica marcado como "solicitado" (destacado em amarelo na aba
 * Agendamentos) e a Regiane recebe um aviso por e-mail. Só quando ela
 * confirma pelo Painel (actionAdminConfirmBooking) é que o evento entra
 * de verdade no Google Agenda dela.
 */
function actionBookSlot(body) {
  const email = patientEmailFromToken_(body);
  let nome;
  const p = findPatientRow(email);
  if (p) {
    nome = p.Nome;
  } else if (isAdmin(email)) {
    nome = 'Administradora (teste)';
  } else {
    return { ok: false, error: 'Paciente não encontrada.' };
  }
  if (!body.iso) return { ok: false, error: 'Selecione um horário da lista.' };

  const inicio = new Date(body.iso);
  if (isNaN(inicio.getTime())) return { ok: false, error: 'Horário inválido.' };
  const termino = new Date(inicio.getTime() + AGENDA_DURACAO_MIN * 60000);

  const cal = getRegianeCalendar();
  const jaOcupado = cal.getEvents(inicio, termino).length > 0;
  if (jaOcupado) return { ok: false, error: 'Esse horário acabou de ser ocupado. Escolha outro.' };

  const sheet = getSheet('Agendamentos');
  const linha = [email, Utilities.formatDate(inicio, FUSO_AGENDA, 'dd/MM/yyyy'), Utilities.formatDate(inicio, FUSO_AGENDA, 'HH:mm'), 'solicitado', new Date(), nome, inicio.toISOString()];
  sheet.appendRow(linha);
  sheet.getRange(sheet.getLastRow(), 1, 1, linha.length).setBackground(COR_SOLICITADO);

  notifyRegiane(
    'Nova solicitação de agendamento — ' + nome,
    nome + ' (' + email + ') pediu consulta para ' + Utilities.formatDate(inicio, FUSO_AGENDA, 'dd/MM/yyyy \'às\' HH:mm') + '. Entre no Painel da Regiane no site (ou na aba Agendamentos da planilha) pra confirmar ou não.'
  );

  return { ok: true };
}

/**
 * Lista as solicitações de agendamento ainda pendentes, pro "Painel da
 * Regiane" mostrar em destaque — ela confirma ou recusa por lá.
 */
function actionAdminListPendingBookings(body) {
  assertAdmin(body.idToken);
  const pendentes = sheetToObjects(getSheet('Agendamentos')).filter(a => String(a.Status).trim().toLowerCase() === 'solicitado');
  return {
    ok: true,
    pendentes: pendentes.map(a => ({
      email: a.Email, nome: a.Nome || a.Email, data: a.Data, hora: a.Hora, iso: a.IsoInicio
    }))
  };
}

/**
 * A Regiane confirma uma solicitação — só agora o evento é criado de
 * verdade no Google Agenda dela (com a paciente como convidada), e a
 * linha na aba Agendamentos vira "confirmado" (destaque verde).
 */
function actionAdminConfirmBooking(body) {
  assertAdmin(body.idToken);
  const email = normEmail(body.email);
  const iso = body.iso;
  const sheet = getSheet('Agendamentos');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  const statusCol = headers.indexOf('Status');
  const isoCol = headers.indexOf('IsoInicio');
  const nomeCol = headers.indexOf('Nome');

  for (let i = 1; i < data.length; i++) {
    if (normEmail(data[i][emailCol]) === email && String(data[i][isoCol]) === String(iso) && String(data[i][statusCol]).trim().toLowerCase() === 'solicitado') {
      const inicio = new Date(iso);
      const termino = new Date(inicio.getTime() + AGENDA_DURACAO_MIN * 60000);
      const cal = getRegianeCalendar();
      if (cal.getEvents(inicio, termino).length > 0) {
        return { ok: false, error: 'Esse horário já está ocupado na sua agenda — não dá pra confirmar.' };
      }

      const nome = data[i][nomeCol] || email;
      cal.createEvent('Consulta - ' + nome, inicio, termino, {
        guests: email,
        description: 'Agendado pelo site. Paciente: ' + nome + ' (' + email + ')'
      });

      sheet.getRange(i + 1, statusCol + 1).setValue('confirmado');
      sheet.getRange(i + 1, 1, 1, headers.length).setBackground(COR_CONFIRMADO);

      return { ok: true };
    }
  }
  return { ok: false, error: 'Solicitação não encontrada (talvez já tenha sido confirmada).' };
}

/**
 * A Regiane recusa uma solicitação — some da lista de pendentes, sem
 * criar nada na agenda. A paciente pode tentar outro horário.
 */
function actionAdminRejectBooking(body) {
  assertAdmin(body.idToken);
  const email = normEmail(body.email);
  const iso = body.iso;
  const sheet = getSheet('Agendamentos');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  const statusCol = headers.indexOf('Status');
  const isoCol = headers.indexOf('IsoInicio');

  for (let i = 1; i < data.length; i++) {
    if (normEmail(data[i][emailCol]) === email && String(data[i][isoCol]) === String(iso) && String(data[i][statusCol]).trim().toLowerCase() === 'solicitado') {
      sheet.getRange(i + 1, statusCol + 1).setValue('recusado');
      sheet.getRange(i + 1, 1, 1, headers.length).setBackground(null);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Solicitação não encontrada.' };
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
 * Devolve o status do Check-up de um e-mail (se já respondeu e quais foram
 * as respostas) — usado tanto pela própria paciente reabrindo a página
 * quanto pela Regiane no "ver como paciente" do Painel dela.
 */
function actionCheckupDashboard(body) {
  const auth = verifyGoogleToken(body.idToken);
  const requested = normEmail(body.email || body.targetEmail);
  let email = auth.email;
  if (requested && requested !== auth.email) {
    if (!isAdmin(auth.email)) throw new Error('Você não pode acessar os dados de outra paciente.');
    email = requested;
  }
  const c = findCheckupRow(email);
  if (!c) return { ok: false, error: 'Check-up não encontrado.' };
  return {
    ok: true,
    nome: c.Nome || '',
    jaFezCheckup: String(c.JaFezCheckup).trim().toLowerCase() === 'sim',
    respostasChecklist: c.RespostasChecklist ? JSON.parse(c.RespostasChecklist) : null,
    respostasQuiz: c.RespostasQuiz ? JSON.parse(c.RespostasQuiz) : null,
    dataCheckup: c.DataCheckup || ''
  };
}

/**
 * Monta um texto legível com as respostas do checklist/quiz, pra mandar
 * por e-mail pra paciente e pra Regiane.
 */
function formatCheckupRespostas(checklist, quiz) {
  let txt = '';
  if (checklist && checklist.length) {
    txt += 'Sinais marcados no checklist:\n';
    checklist.forEach(item => { txt += '- ' + item + '\n'; });
    txt += '\n';
  }
  if (quiz && Object.keys(quiz).length) {
    txt += 'Respostas do quiz:\n';
    Object.keys(quiz).forEach(pergunta => { txt += '- ' + pergunta + ': ' + quiz[pergunta] + '\n'; });
  }
  return txt || '(sem respostas registradas)';
}

/**
 * Salva as respostas do quiz/checklist do Check-up — só pode ser feito UMA vez
 * por e-mail. Depois disso, o login sempre retorna o resultado já salvo.
 * Manda uma cópia das respostas por e-mail pra paciente e pra Regiane.
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
  sheet.getRange(row, headers.indexOf('RespostasChecklist') + 1).setValue(JSON.stringify(body.respostasChecklist || []));
  sheet.getRange(row, headers.indexOf('RespostasQuiz') + 1).setValue(JSON.stringify(body.respostasQuiz || {}));
  sheet.getRange(row, headers.indexOf('DataCheckup') + 1).setValue(new Date());

  const nome = c.Nome || auth.nome || auth.email;
  const resumo = formatCheckupRespostas(body.respostasChecklist, body.respostasQuiz);

  try {
    sendPatientEmail_(auth.email, 'Suas respostas do Check-up Alimentar — Regiane Silva',
      'Oi, ' + nome + '! Aqui está uma cópia das suas respostas no Check-up Alimentar Funcional:\n\n' + resumo);
  } catch (err) {
    // não deixa o fluxo principal quebrar se o e-mail falhar
  }

  notifyRegiane(
    'Check-up respondido — ' + nome,
    nome + ' (' + auth.email + ') acabou de responder o Check-up. Respostas:\n\n' + resumo
  );

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

/**
 * Lista todas as pacientes do Check-up pra administradora — usado no
 * Painel da Regiane da Área do Check-up, incluindo as respostas de quem
 * já respondeu, pra ela ter controle sem precisar abrir a planilha.
 */
function actionAdminListCheckupPatients(body) {
  assertAdmin(body.idToken);
  const pacientes = sheetToObjects(getSheet('CheckupPacientes'));
  return {
    ok: true,
    pacientes: pacientes.map(c => ({
      email: c.Email,
      nome: c.Nome,
      liberado: String(c.Liberado).trim().toLowerCase() === 'sim',
      jaFezCheckup: String(c.JaFezCheckup).trim().toLowerCase() === 'sim',
      dataCheckup: c.DataCheckup || '',
      respostasChecklist: c.RespostasChecklist ? JSON.parse(c.RespostasChecklist) : null,
      respostasQuiz: c.RespostasQuiz ? JSON.parse(c.RespostasQuiz) : null
    }))
  };
}

function actionAdminAddCheckupPatient(body) {
  assertAdmin(body.idToken);
  const email = normEmail(body.email);
  const nome = String(body.nome || '').trim();
  if (!email || email.indexOf('@') === -1 || !nome) {
    return { ok: false, error: 'Informe o nome e um e-mail válido.' };
  }

  const sheet = getSheet('CheckupPacientes');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  const nomeCol = headers.indexOf('Nome');
  const liberadoCol = headers.indexOf('Liberado');
  const dataCol = headers.indexOf('DataLiberacao');
  for (let i = 1; i < data.length; i++) {
    if (normEmail(data[i][emailCol]) === email) {
      sheet.getRange(i + 1, nomeCol + 1).setValue(nome);
      sheet.getRange(i + 1, liberadoCol + 1).setValue('Sim');
      sheet.getRange(i + 1, dataCol + 1).setValue(new Date());
      return { ok: true, created: false };
    }
  }

  sheet.appendRow([email, nome, new Date(), 'Sim', 'Não', '', '', '']);
  return { ok: true, created: true };
}

function actionAdminSetCheckupAccess(body) {
  assertAdmin(body.idToken);
  const email = normEmail(body.email);
  const liberado = body.liberado ? 'Sim' : 'Não';
  const sheet = getSheet('CheckupPacientes');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  const liberadoCol = headers.indexOf('Liberado');
  for (let i = 1; i < data.length; i++) {
    if (normEmail(data[i][emailCol]) === email) {
      sheet.getRange(i + 1, liberadoCol + 1).setValue(liberado);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Paciente do check-up não encontrada.' };
}

// ── CONSULTORIA (anamnese simplificada pré-atendimento) ─────

function ensureConsultoriaSheets_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let pacientes = ss.getSheetByName('ConsultoriaPacientes');
  if (!pacientes) pacientes = ss.insertSheet('ConsultoriaPacientes');
  ensureHeaders_(pacientes, ['Email', 'Nome', 'DataLiberacao', 'Liberado', 'RespostasAnamnese', 'DataResposta']);
  pacientes.setFrozenRows(1);

  let respostas = ss.getSheetByName('ConsultoriaRespostas');
  if (!respostas) respostas = ss.insertSheet('ConsultoriaRespostas');
  ensureHeaders_(respostas, ['Id', 'Email', 'Nome', 'DataResposta', 'Modalidade', 'Objetivo', 'MedicamentosSuplementos', 'PatologiasCondicoes', 'Intestino', 'AlergiasIntolerancias', 'Preferencias', 'CafeDaManha', 'Almoco', 'Lanches', 'Jantar', 'OutrasInformacoes']);
  respostas.setFrozenRows(1);
  return { pacientes: pacientes, respostas: respostas };
}

function setupConsultoriaSheets() {
  ensureConsultoriaSheets_();
  SpreadsheetApp.flush();
  return { ok: true };
}

function findConsultoriaRow(email) {
  const sheet = ensureConsultoriaSheets_().pacientes;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  for (let i = 1; i < data.length; i++) {
    if (normEmail(data[i][emailCol]) === normEmail(email)) {
      const obj = {};
      headers.forEach(function(h, idx) { obj[h] = data[i][idx]; });
      obj._row = i + 1;
      return obj;
    }
  }
  return null;
}

function consultoriaIdentityFromAuth_(body) {
  if (body.idToken) return verifyGoogleToken(body.idToken);
  const accessToken = String(body.accessToken || '').trim();
  if (!accessToken) throw new Error('Sessão de acesso ausente. Entre novamente.');
  const cached = CacheService.getScriptCache().get('consultoria-session:' + accessToken);
  if (!cached) throw new Error('Sua sessão expirou. Solicite um novo código de acesso.');
  const session = JSON.parse(cached);
  return { email: normEmail(session.email), nome: session.nome || session.email, emailCode: true };
}

function consultoriaEmailFromToken_(body) {
  const auth = consultoriaIdentityFromAuth_(body);
  const requested = normEmail(body.email || body.targetEmail);
  let email = auth.email;
  if (requested && requested !== auth.email) {
    if (!isAdmin(auth.email)) throw new Error('Você não pode acessar os dados de outra paciente.');
    email = requested;
  }
  if (!isAdmin(auth.email) && !findConsultoriaRow(email)) throw new Error('Paciente da consultoria não encontrada.');
  return { email: email, auth: auth };
}

function actionRequestConsultoriaAccessCode(body) {
  const email = normEmail(body.email);
  if (!email || email.indexOf('@') === -1) return { ok: false, error: 'Informe um e-mail válido.' };
  const paciente = findConsultoriaRow(email);
  if (!paciente || String(paciente.Liberado).trim().toLowerCase() !== 'sim') {
    return { ok: false, error: 'Este e-mail não possui acesso ativo à Consultoria.' };
  }
  const cache = CacheService.getScriptCache();
  if (cache.get('consultoria-otp-rate:' + email)) {
    return { ok: false, error: 'Aguarde um minuto antes de solicitar outro código.' };
  }
  const code = String(Math.floor(100000 + Math.random() * 900000));
  cache.put('consultoria-otp:' + email, JSON.stringify({ code: code, attempts: 0 }), 600);
  cache.put('consultoria-otp-rate:' + email, '1', 60);
  sendPatientEmail_(email, 'Seu código de acesso - Consultoria Regiane Silva',
    'Oi, ' + (paciente.Nome || '') + '!\n\nSeu código de acesso à Área da Consultoria é: ' + code + '\n\nEle vale por 10 minutos e pode ser usado uma vez. Se você não solicitou este código, ignore esta mensagem.\n\nRegiane Silva');
  return { ok: true, email: email };
}

function actionVerifyConsultoriaAccessCode(body) {
  const email = normEmail(body.email);
  const code = String(body.code || '').replace(/\D/g, '');
  if (!email || code.length !== 6) return { ok: false, error: 'Informe o e-mail e o código de seis dígitos.' };
  const paciente = findConsultoriaRow(email);
  if (!paciente || String(paciente.Liberado).trim().toLowerCase() !== 'sim') return { ok: false, error: 'Acesso não autorizado.' };
  const cache = CacheService.getScriptCache();
  const key = 'consultoria-otp:' + email;
  const raw = cache.get(key);
  if (!raw) return { ok: false, error: 'Código expirado. Solicite um novo código.' };
  const record = JSON.parse(raw);
  record.attempts = Number(record.attempts || 0) + 1;
  if (record.attempts > 5) {
    cache.remove(key);
    return { ok: false, error: 'Muitas tentativas. Solicite um novo código.' };
  }
  if (String(record.code) !== code) {
    cache.put(key, JSON.stringify(record), 600);
    return { ok: false, error: 'Código incorreto.' };
  }
  cache.remove(key);
  const accessToken = Utilities.getUuid() + Utilities.getUuid();
  cache.put('consultoria-session:' + accessToken, JSON.stringify({ email: email, nome: paciente.Nome || email }), 21600);
  return { ok: true, email: email, nome: paciente.Nome || email, accessToken: accessToken, expiresIn: 21600 };
}

function actionGoogleLoginConsultoria(body) {
  const auth = verifyGoogleToken(body.idToken);
  if (isAdmin(auth.email)) return { ok: true, nome: auth.nome, email: auth.email, admin: true };
  const paciente = findConsultoriaRow(auth.email);
  if (!paciente || String(paciente.Liberado).trim().toLowerCase() !== 'sim') {
    return { ok: false, error: 'Não encontramos uma consultoria ativa para esta conta Google. Fale com a Regiane.' };
  }
  return { ok: true, nome: paciente.Nome || auth.nome, email: paciente.Email, admin: false };
}

function actionConsultoriaDashboard(body) {
  const identity = consultoriaEmailFromToken_(body);
  const paciente = findConsultoriaRow(identity.email);
  const materiais = sheetToObjects(getSheet('Materiais')).filter(function(m) {
    const area = String(m.Area || '').trim().toLowerCase();
    const destino = normEmail(m.Email);
    return area === 'consultoria' &&
      (destino === 'consultoria' || destino === 'todos' || destino === identity.email) &&
      isValidPublicLink_(m.Link);
  }).map(function(m) {
    return { id: m.Id, titulo: m.Titulo || 'Vídeo da Regiane', descricao: m.Descricao || '', link: m.Link, tipo: m.Tipo || 'Vídeo' };
  });
  if (!paciente) {
    if (isAdmin(identity.auth.email)) return { ok: true, nome: 'Administradora', admin: true, respondeu: false, respostas: null, materiais: materiais };
    return { ok: false, error: 'Paciente da consultoria não encontrada.' };
  }
  if (!isAdmin(identity.auth.email) && String(paciente.Liberado).trim().toLowerCase() !== 'sim') {
    return { ok: false, error: 'Seu acesso à consultoria não está liberado.' };
  }
  let respostas = null;
  try { respostas = paciente.RespostasAnamnese ? JSON.parse(paciente.RespostasAnamnese) : null; } catch (err) { respostas = null; }
  return {
    ok: true, nome: paciente.Nome || '', email: paciente.Email,
    admin: isAdmin(identity.auth.email),
    liberado: String(paciente.Liberado).trim().toLowerCase() === 'sim',
    respondeu: !!paciente.DataResposta, dataResposta: paciente.DataResposta || '', respostas: respostas, materiais: materiais
  };
}

function formatConsultoriaRespostas_(respostas) {
  const labels = {
    modalidade: 'Modalidade da consultoria', objetivo: 'Objetivo principal',
    medicamentos: 'Medicamentos e suplementos', patologias: 'Diagnósticos, patologias ou condições de saúde',
    intestino: 'Funcionamento do intestino', alergias: 'Alergias e intolerâncias alimentares',
    preferencias: 'Preferências e alimentos que não consome', cafe: 'Café da manhã atual',
    almoco: 'Almoço atual', lanche: 'Lanches atuais', jantar: 'Jantar atual',
    outros: 'Outras informações importantes'
  };
  return Object.keys(labels).map(function(key) {
    return labels[key] + ': ' + String(respostas[key] || 'Não informado');
  }).join('\n');
}

function actionSubmitConsultoriaAnamnese(body) {
  const identity = consultoriaEmailFromToken_(body);
  if (isAdmin(identity.auth.email) && identity.email === identity.auth.email) {
    return { ok: false, error: 'Selecione uma paciente para responder como paciente.' };
  }
  const paciente = findConsultoriaRow(identity.email);
  if (!paciente || String(paciente.Liberado).trim().toLowerCase() !== 'sim') return { ok: false, error: 'Acesso à consultoria não liberado.' };
  const respostas = body.respostas || {};
  if (!String(respostas.objetivo || '').trim()) return { ok: false, error: 'Informe seu objetivo principal.' };
  if (!String(respostas.intestino || '').trim()) return { ok: false, error: 'Conte como está o funcionamento do seu intestino.' };
  if (!String(respostas.cafe || '').trim() || !String(respostas.almoco || '').trim() || !String(respostas.jantar || '').trim()) {
    return { ok: false, error: 'Preencha café da manhã, almoço e jantar atuais.' };
  }

  const now = new Date();
  const sheet = getSheet('ConsultoriaPacientes');
  const headers = sheet.getDataRange().getValues()[0];
  sheet.getRange(paciente._row, headers.indexOf('RespostasAnamnese') + 1).setValue(JSON.stringify(respostas));
  sheet.getRange(paciente._row, headers.indexOf('DataResposta') + 1).setValue(now);

  const history = getSheet('ConsultoriaRespostas');
  history.appendRow([
    'consultoria-' + now.getTime(), identity.email, paciente.Nome || identity.auth.nome || '', now,
    sheetLiteral_(respostas.modalidade), sheetLiteral_(respostas.objetivo), sheetLiteral_(respostas.medicamentos),
    sheetLiteral_(respostas.patologias), sheetLiteral_(respostas.intestino), sheetLiteral_(respostas.alergias),
    sheetLiteral_(respostas.preferencias), sheetLiteral_(respostas.cafe), sheetLiteral_(respostas.almoco),
    sheetLiteral_(respostas.lanche), sheetLiteral_(respostas.jantar), sheetLiteral_(respostas.outros)
  ]);

  const resumo = formatConsultoriaRespostas_(respostas);
  let emailPacienteEnviado = false;
  try {
    sendPatientEmail_(identity.email, 'Suas respostas para a Consultoria Nutricional - Regiane Silva',
      'Oi, ' + (paciente.Nome || identity.auth.nome || '') + '!\n\nRecebemos suas informações para preparar a consultoria. Confira a cópia abaixo:\n\n' + resumo + '\n\nAté a nossa consultoria!\nRegiane Silva');
    emailPacienteEnviado = true;
  } catch (err) {}
  const emailRegianeEnviado = notifyRegiane(
    'Anamnese da Consultoria respondida - ' + (paciente.Nome || identity.email),
    (paciente.Nome || identity.email) + ' (' + identity.email + ') respondeu à preparação da consultoria.\n\n' + resumo
  );
  return { ok: true, dataResposta: now, emailPacienteEnviado: emailPacienteEnviado, emailRegianeEnviado: emailRegianeEnviado };
}

function actionAdminListConsultoriaPatients(body) {
  assertAdmin(body.idToken);
  return {
    ok: true,
    pacientes: sheetToObjects(getSheet('ConsultoriaPacientes')).map(function(p) {
      let respostas = null;
      try { respostas = p.RespostasAnamnese ? JSON.parse(p.RespostasAnamnese) : null; } catch (err) { respostas = null; }
      return { email: p.Email, nome: p.Nome, liberado: String(p.Liberado).trim().toLowerCase() === 'sim', dataResposta: p.DataResposta || '', respostas: respostas };
    })
  };
}

function actionAdminAddConsultoriaPatient(body) {
  assertAdmin(body.idToken);
  const email = normEmail(body.email);
  const nome = String(body.nome || '').trim();
  if (!email || email.indexOf('@') === -1 || !nome) return { ok: false, error: 'Informe nome e e-mail válido.' };
  const sheet = getSheet('ConsultoriaPacientes');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  for (let i = 1; i < data.length; i++) {
    if (normEmail(data[i][emailCol]) === email) {
      sheet.getRange(i + 1, headers.indexOf('Nome') + 1).setValue(nome);
      sheet.getRange(i + 1, headers.indexOf('Liberado') + 1).setValue('Sim');
      sheet.getRange(i + 1, headers.indexOf('DataLiberacao') + 1).setValue(new Date());
      return { ok: true, created: false };
    }
  }
  sheet.appendRow([email, nome, new Date(), 'Sim', '', '']);
  return { ok: true, created: true };
}

function actionAdminSetConsultoriaAccess(body) {
  assertAdmin(body.idToken);
  const email = normEmail(body.email);
  const sheet = getSheet('ConsultoriaPacientes');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  for (let i = 1; i < data.length; i++) {
    if (normEmail(data[i][emailCol]) === email) {
      sheet.getRange(i + 1, headers.indexOf('Liberado') + 1).setValue(body.liberado ? 'Sim' : 'Não');
      return { ok: true };
    }
  }
  return { ok: false, error: 'Paciente da consultoria não encontrada.' };
}

// ── LINK NA BIO (biolink.html) ───────────────────────────

/** Mantém dados enviados pelo público como texto literal na planilha.
 * Evita que telefones com +55 ou textos iniciados por =, -, + ou @ virem fórmulas.
 */
function sheetLiteral_(value) {
  const text = String(value || '');
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function buildBioLeadNotification_(type, lead) {
  if (type === 'newsletter') {
    return {
      subject: '[NOVO LEAD] Newsletter do link da bio — ' + (lead.Nome || lead.Email || 'novo contato'),
      body: 'Um novo contato entrou pela newsletter do link da bio.\n\n' +
        'Nome: ' + (lead.Nome || '-') + '\n' +
        'E-mail: ' + (lead.Email || '-') + '\n\n' +
        'Entre em contato com esse lead assim que possível.\n\n' +
        'Planilha: https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit#gid=585346695'
    };
  }

  return {
    subject: '[NOVO LEAD] Quiz do link da bio — ' + (lead.Nome || lead.Telefone || 'novo contato'),
    body: 'Uma nova pessoa respondeu ao quiz do link da bio.\n\n' +
      'Nome: ' + (lead.Nome || '-') + '\n' +
      'Telefone: ' + (lead.Telefone || '-') + '\n' +
      'Momento: ' + (lead.Momento || '-') + '\n' +
      'Sintoma: ' + (lead.Sintoma || '-') + '\n' +
      'Mensagem: ' + (lead.Mensagem || '-') + '\n' +
      'Recomendação sugerida: ' + (lead.RecomendacaoSugerida || '-') + '\n\n' +
      'Entre em contato com esse lead assim que possível.\n\n' +
      'Planilha: https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit#gid=818925408'
  };
}

function trySendBioLeadNotification_(sheet, row, notification, sentCol, statusCol, attemptsCol) {
  const attempts = Number(sheet.getRange(row, attemptsCol).getValue() || 0) + 1;
  sheet.getRange(row, attemptsCol).setValue(attempts);
  try {
    sendRegianeEmailStrict_(notification.subject, notification.body);
    sheet.getRange(row, sentCol, 1, 2).setValues([[new Date(), 'Enviado']]);
    return { ok: true, attempts: attempts };
  } catch (err) {
    const message = String(err && err.message ? err.message : err).slice(0, 300);
    sheet.getRange(row, statusCol).setValue('Erro no envio: ' + message);
    ensureNotificationTrigger_();
    return { ok: false, attempts: attempts, error: message };
  }
}

function processPendingBioLeadNotifications_() {
  [
    { name: 'BioLeadsQuiz', type: 'quiz' },
    { name: 'BioNewsletter', type: 'newsletter' }
  ].forEach(function(config) {
    const sheet = getSheet(config.name);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    const headers = data[0];
    const sentCol = headers.indexOf('AvisoEnviadoEm') + 1;
    const statusCol = headers.indexOf('StatusAviso') + 1;
    const attemptsCol = headers.indexOf('TentativasAviso') + 1;
    if (!sentCol || !statusCol || !attemptsCol) return;

    data.slice(1).forEach(function(values, index) {
      const status = String(values[statusCol - 1] || '');
      const attempts = Number(values[attemptsCol - 1] || 0);
      if ((status !== 'Pendente' && status.indexOf('Erro no envio:') !== 0) || attempts >= 5) return;
      const lead = {};
      headers.forEach(function(header, col) { lead[header] = values[col]; });
      const notification = buildBioLeadNotification_(config.type, lead);
      trySendBioLeadNotification_(sheet, index + 2, notification, sentCol, statusCol, attemptsCol);
    });
  });
}

/**
 * Recebe os leads do link na bio e separa em duas abas: quem respondeu o
 * quiz de diagnóstico vai para BioLeadsQuiz, quem só deixou o e-mail na
 * newsletter vai para BioNewsletter. Nos dois casos avisa a Regiane.
 */
function actionBioLead(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (body.tag === 'newsletter') {
      const sheet = getSheet('BioNewsletter');
      sheet.appendRow([new Date(), sheetLiteral_(body.nome), sheetLiteral_(body.email), '', 'Pendente', 0]);
      const row = sheet.getLastRow();
      const notification = buildBioLeadNotification_('newsletter', { Nome: body.nome, Email: body.email });
      const delivery = trySendBioLeadNotification_(sheet, row, notification, 4, 5, 6);
      return { ok: true, avisoEmailEnviado: delivery.ok };
    }

    const sheet = getSheet('BioLeadsQuiz');
    sheet.appendRow([
      new Date(),
      sheetLiteral_(body.nome),
      sheetLiteral_(body.telefone),
      sheetLiteral_(body.pergunta1),
      sheetLiteral_(body.pergunta2),
      sheetLiteral_(body.textoLivre),
      sheetLiteral_(body.cursoSugerido),
      '',
      'Pendente',
      0
    ]);
    const row = sheet.getLastRow();
    const notification = buildBioLeadNotification_('quiz', {
      Nome: body.nome,
      Telefone: body.telefone,
      Momento: body.pergunta1,
      Sintoma: body.pergunta2,
      Mensagem: body.textoLivre,
      RecomendacaoSugerida: body.cursoSugerido
    });
    const delivery = trySendBioLeadNotification_(sheet, row, notification, 8, 9, 10);
    return { ok: true, avisoEmailEnviado: delivery.ok };
  } finally {
    lock.releaseLock();
  }
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
    ['Email', 'Nome', 'DataInicio', 'RetornosRealizados', 'ReceitasSalvas', 'ProgressoPercent', 'ProximoRetornoData', 'ProximoRetornoHora', 'PlanoTexto', 'PontosTotal', 'ReceitasSalvasIds', 'PlanoPdfUrl', 'Instagram'],
    ['exemplo@paciente.com', 'Nome de Exemplo', new Date(), 0, 0, 0, '', '', 'Siga as orientações da última consulta.', 0, '', '', '@perfil.exemplo']);

  buildSheet('Materiais',
    ['Id', 'Email', 'Tipo', 'Titulo', 'Descricao', 'Link', 'Area'],
    [1, 'TODOS', 'PDF', 'Planner alimentar semanal', 'Vale para todas as pacientes do Programa', 'https://', 'materiais']);

  buildSheet('Comentarios', ['Id', 'PostId', 'Email', 'Nome', 'Texto', 'DataHora', 'AvisoEnviadoEm', 'StatusAviso', 'PontosConcedidos'], null);

  buildSheet('PontosLog', ['Id', 'Email', 'Tipo', 'Pontos', 'Data', 'ChaveUnica', 'SaldoDepois'], null);

  buildSheet('Fotos', ['Id', 'Email', 'Tipo', 'MessageId', 'AttachmentIndex', 'NomeArquivo', 'MimeType', 'Hash', 'DataHora', 'Pontos', 'Status'], null);

  buildSheet('Creditos', ['Id', 'Email', 'Tipo', 'Valor', 'Descricao', 'Data', 'AdminEmail'], null);

  buildSheet('Agendamentos', ['Email', 'Data', 'Hora', 'Status', 'DataSolicitacao', 'Nome', 'IsoInicio'], null);

  buildSheet('Rotinas',
    ['Id', 'Email', 'Titulo', 'Categoria', 'MetaValor', 'MetaUnidade', 'DiasSemana', 'Horarios', 'Mensagem', 'Ativo', 'DataInicio', 'DataFim', 'CriadoEm', 'AtualizadoEm'],
    null);

  buildSheet('Notificacoes',
    ['Id', 'RotinaId', 'Email', 'Titulo', 'Mensagem', 'AgendadaPara', 'EnviadaEm', 'LidaEm', 'ConcluidaEm', 'Status'],
    null);

  buildSheet('CheckupPacientes',
    ['Email', 'Nome', 'DataLiberacao', 'Liberado', 'JaFezCheckup', 'RespostasChecklist', 'RespostasQuiz', 'DataCheckup'],
    ['exemplo@checkup.com', 'Nome de Exemplo', new Date(), 'Sim', 'Não', '', '', '']);

  buildSheet('ConsultoriaPacientes',
    ['Email', 'Nome', 'DataLiberacao', 'Liberado', 'RespostasAnamnese', 'DataResposta'],
    null);

  buildSheet('ConsultoriaRespostas',
    ['Id', 'Email', 'Nome', 'DataResposta', 'Modalidade', 'Objetivo', 'MedicamentosSuplementos', 'PatologiasCondicoes', 'Intestino', 'AlergiasIntolerancias', 'Preferencias', 'CafeDaManha', 'Almoco', 'Lanches', 'Jantar', 'OutrasInformacoes'],
    null);

  buildSheet('BioLeadsQuiz',
    ['Data', 'Nome', 'Telefone', 'Momento', 'Sintoma', 'Mensagem', 'RecomendacaoSugerida', 'AvisoEnviadoEm', 'StatusAviso', 'TentativasAviso'],
    null);

  buildSheet('BioNewsletter',
    ['Data', 'Nome', 'Email', 'AvisoEnviadoEm', 'StatusAviso', 'TentativasAviso'],
    null);

  // remove a aba padrão em branco, se existir e não for a única
  ['Sheet1', 'Página1', 'Folha1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  // ordena as abas na ordem que faz mais sentido pro dia a dia da Regiane
  const ordem = ['Pacientes', 'CheckupPacientes', 'ConsultoriaPacientes', 'ConsultoriaRespostas', 'Rotinas', 'Notificacoes', 'Materiais', 'Agendamentos', 'Comentarios', 'Fotos', 'PontosLog', 'Creditos', 'BioLeadsQuiz', 'BioNewsletter'];
  ordem.forEach((nome, i) => {
    const s = ss.getSheetByName(nome);
    if (s) ss.setActiveSheet(s);
    if (s) ss.moveActiveSheet(i + 1);
  });

  SpreadsheetApp.flush();
}
