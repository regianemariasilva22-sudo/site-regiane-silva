/**
 * Backend da Área do Programa — Regiane Silva
 * Lê e escreve numa Google Sheet, publicado como Web App.
 * Veja INSTRUCOES.md para o passo a passo de publicação.
 */

// Cole aqui o ID da planilha (fica na URL dela, entre /d/ e /edit)
const SHEET_ID = 'COLE_AQUI_O_ID_DA_PLANILHA';

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

function daysSince(dateVal) {
  if (!dateVal) return 0;
  const start = new Date(dateVal);
  if (isNaN(start.getTime())) return 0;
  const now = new Date();
  return Math.max(0, Math.floor((now - start) / (1000 * 60 * 60 * 24)));
}

function actionDashboard(email) {
  const p = findPatientRow(email);
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
