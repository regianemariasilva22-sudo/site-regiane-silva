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

/**
 * Verifica o token do Google e garante que quem está chamando é uma
 * administradora (Regiane ou Aline). Lança erro se não for.
 */
function assertAdmin(idToken) {
  const auth = verifyGoogleToken(idToken);
  if (!isAdmin(auth.email)) throw new Error('Acesso restrito à administradora.');
  return auth;
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
    if (action === 'checkAccess') return jsonResponse(actionCheckAccess(e.parameter.email, e.parameter.area));
    if (action === 'checkupDashboard') return jsonResponse(actionCheckupDashboard(e.parameter.email));
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
    if (action === 'saveRecipe') return jsonResponse(actionSaveRecipe(body));
    if (action === 'uploadFoto') return jsonResponse(actionUploadFoto(body));
    if (action === 'adminListPatients') return jsonResponse(actionAdminListPatients(body));
    if (action === 'adminSavePlan') return jsonResponse(actionAdminSavePlan(body));
    if (action === 'adminUploadPlanPdf') return jsonResponse(actionAdminUploadPlanPdf(body));
    if (action === 'adminAddMaterial') return jsonResponse(actionAdminAddMaterial(body));
    if (action === 'adminListPendingBookings') return jsonResponse(actionAdminListPendingBookings(body));
    if (action === 'adminConfirmBooking') return jsonResponse(actionAdminConfirmBooking(body));
    if (action === 'adminRejectBooking') return jsonResponse(actionAdminRejectBooking(body));
    if (action === 'adminListCheckupPatients') return jsonResponse(actionAdminListCheckupPatients(body));
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
    totalInteracoes: totalInteracoes,
    creditoDisponivel: Math.floor(pontosTotal / 100) * 10,
    faltamParaProximoCredito: 100 - (pontosTotal % 100),
    materiais: materiais.map(m => ({
      tipo: m.Tipo, titulo: m.Titulo, descricao: m.Descricao, link: m.Link, area: m.Area || 'materiais'
    }))
  };
}

/**
 * Salva o "salvar receita" de uma paciente — idempotente (a mesma receita
 * não conta pontos/contador duas vezes pra mesma pessoa) e atualiza o
 * contador ReceitasSalvas automaticamente, sem a Regiane precisar mexer.
 */
function actionSaveRecipe(body) {
  const email = normEmail(body.email);
  const recipeId = String(body.recipeId || '').trim();
  if (!recipeId) return { ok: false, error: 'Receita inválida.' };

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
      addPoints(email, 'Receita salva: ' + (body.recipeTitle || recipeId), 2);
      return { ok: true, jaSalva: false, receitasSalvas: novoTotal };
    }
  }
  return { ok: false, error: 'Paciente não encontrada.' };
}

/**
 * A paciente envia uma foto de resultado ou do prato — vai pro Google Drive
 * da Regiane e conta pontos automaticamente (5 pontos), sem ela precisar
 * mexer em nada. A Regiane recebe um aviso por e-mail com o link da foto.
 */
function actionUploadFoto(body) {
  const email = normEmail(body.email);
  const p = findPatientRow(email);
  if (!p) return { ok: false, error: 'Paciente não encontrada.' };
  if (!body.fileBase64) return { ok: false, error: 'Nenhuma foto enviada.' };

  const bytes = Utilities.base64Decode(body.fileBase64);
  const blob = Utilities.newBlob(bytes, body.mimeType || 'image/jpeg', body.fileName || 'foto.jpg');
  const file = DriveApp.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = file.getUrl();

  const tipo = body.tipo === 'prato' ? 'Foto do prato' : 'Foto de resultado';
  addPoints(email, tipo, 5);

  notifyRegiane(
    'Nova ' + tipo.toLowerCase() + ' — ' + (p.Nome || email),
    (p.Nome || email) + ' enviou uma ' + tipo.toLowerCase() + '.\n' + url
  );

  return { ok: true, url: url };
}

/**
 * Checagem leve (sem token) usada pelo site pra saber, de tempos em tempos
 * enquanto a página está aberta, se aquele e-mail ainda está cadastrado —
 * se a Regiane excluir a linha da planilha, o site desloga sozinho.
 */
function actionCheckAccess(email, area) {
  if (isAdmin(email)) return { ok: true };
  if (area === 'checkup') {
    const c = findCheckupRow(email);
    if (!c || String(c.Liberado).trim().toLowerCase() !== 'sim') return { ok: false };
    return { ok: true };
  }
  return { ok: !!findPatientRow(email) };
}

// ── PAINEL DA ADMINISTRADORA ─────────────────────────────

/**
 * Lista todas as pacientes do Programa pra administradora — usado no
 * "Painel da Regiane" dentro do próprio site, pra ela ver quem já tem
 * plano alimentar cadastrado e quem ainda falta.
 */
function actionAdminListPatients(body) {
  assertAdmin(body.idToken);
  const pacientes = sheetToObjects(getSheet('Pacientes'));
  return {
    ok: true,
    pacientes: pacientes.map(p => ({
      email: p.Email,
      nome: p.Nome,
      diasAcompanhamento: daysSince(p.DataInicio),
      retornosRealizados: Number(p.RetornosRealizados) || 0,
      receitasSalvas: Number(p.ReceitasSalvas) || 0,
      pontosTotal: Number(p.PontosTotal) || 0,
      progressoPercent: Number(p.ProgressoPercent) || 0,
      temPlano: !!(p.PlanoTexto && String(p.PlanoTexto).trim()),
      planoTexto: p.PlanoTexto || '',
      planoPdfUrl: p.PlanoPdfUrl || ''
    }))
  };
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

  const areasValidas = ['aulas', 'rotulo', 'materiais'];
  const area = areasValidas.indexOf(body.area) !== -1 ? body.area : 'materiais';
  const tipoPadrao = area === 'materiais' ? 'Material' : 'Vídeo';

  const destino = normEmail(body.email) || 'todos';
  const sheet = getSheet('Materiais');
  const id = new Date().getTime();
  sheet.appendRow([id, destino === 'todos' ? 'TODOS' : destino, body.tipo || tipoPadrao, titulo, body.descricao || '', link, area]);

  return { ok: true };
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
  const email = normEmail(body.email);
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
function actionCheckupDashboard(email) {
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
    MailApp.sendEmail(auth.email, 'Suas respostas do Check-up Alimentar — Regiane Silva',
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

// ── LINK NA BIO (biolink.html) ───────────────────────────

/**
 * Recebe os leads do link na bio e separa em duas abas: quem respondeu o
 * quiz de diagnóstico vai para BioLeadsQuiz, quem só deixou o e-mail na
 * newsletter vai para BioNewsletter. Nos dois casos avisa a Regiane.
 */
function actionBioLead(body) {
  if (body.tag === 'newsletter') {
    const sheet = getSheet('BioNewsletter');
    sheet.appendRow([new Date(), body.nome || '', body.email || '']);
    notifyRegiane('Novo inscrito na newsletter do link na bio', 'Nome: ' + (body.nome || '-') + '\nE-mail: ' + (body.email || '-'));
    return { ok: true };
  }

  const sheet = getSheet('BioLeadsQuiz');
  sheet.appendRow([
    new Date(),
    body.nome || '',
    body.telefone || '',
    body.pergunta1 || '',
    body.pergunta2 || '',
    body.textoLivre || '',
    body.cursoSugerido || ''
  ]);

  const quem = body.nome || body.telefone || 'alguém';
  let corpo = 'Novo diagnóstico respondido no link na bio.\n\nNome: ' + (body.nome || '-') +
    '\nTelefone: ' + (body.telefone || '-');
  if (body.pergunta1) corpo += '\nMomento: ' + body.pergunta1;
  if (body.pergunta2) corpo += '\nSintoma: ' + body.pergunta2;
  if (body.textoLivre) corpo += '\nMensagem: ' + body.textoLivre;
  if (body.cursoSugerido) corpo += '\nRecomendação sugerida: ' + body.cursoSugerido;

  notifyRegiane('Novo lead do link na bio — ' + quem, corpo);

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
    ['Email', 'Nome', 'DataInicio', 'RetornosRealizados', 'ReceitasSalvas', 'ProgressoPercent', 'ProximoRetornoData', 'ProximoRetornoHora', 'PlanoTexto', 'PontosTotal', 'ReceitasSalvasIds', 'PlanoPdfUrl'],
    ['exemplo@paciente.com', 'Nome de Exemplo', new Date(), 0, 0, 0, '', '', 'Siga as orientações da última consulta.', 0, '', '']);

  buildSheet('Materiais',
    ['Id', 'Email', 'Tipo', 'Titulo', 'Descricao', 'Link', 'Area'],
    [1, 'TODOS', 'PDF', 'Planner alimentar semanal', 'Vale para todas as pacientes do Programa', 'https://', 'materiais']);

  buildSheet('Comentarios', ['Id', 'PostId', 'Email', 'Nome', 'Texto', 'DataHora'], null);

  buildSheet('PontosLog', ['Id', 'Email', 'Tipo', 'Pontos', 'Data'], null);

  buildSheet('Agendamentos', ['Email', 'Data', 'Hora', 'Status', 'DataSolicitacao', 'Nome', 'IsoInicio'], null);

  buildSheet('CheckupPacientes',
    ['Email', 'Nome', 'DataLiberacao', 'Liberado', 'JaFezCheckup', 'RespostasChecklist', 'RespostasQuiz', 'DataCheckup'],
    ['exemplo@checkup.com', 'Nome de Exemplo', new Date(), 'Sim', 'Não', '', '', '']);

  buildSheet('BioLeadsQuiz',
    ['Data', 'Nome', 'Telefone', 'Momento', 'Sintoma', 'Mensagem', 'RecomendacaoSugerida'],
    null);

  buildSheet('BioNewsletter',
    ['Data', 'Nome', 'Email'],
    null);

  // remove a aba padrão em branco, se existir e não for a única
  ['Sheet1', 'Página1', 'Folha1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  // ordena as abas na ordem que faz mais sentido pro dia a dia da Regiane
  const ordem = ['Pacientes', 'CheckupPacientes', 'Materiais', 'Agendamentos', 'Comentarios', 'PontosLog', 'BioLeadsQuiz', 'BioNewsletter'];
  ordem.forEach((nome, i) => {
    const s = ss.getSheetByName(nome);
    if (s) ss.setActiveSheet(s);
    if (s) ss.moveActiveSheet(i + 1);
  });

  SpreadsheetApp.flush();
}
