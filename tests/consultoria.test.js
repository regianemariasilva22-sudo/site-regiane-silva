const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

class Range {
  constructor(sheet, row, col, rows = 1, cols = 1) { Object.assign(this, { sheet, row, col, rows, cols }); }
  getValues() { return Array.from({ length: this.rows }, (_, r) => Array.from({ length: this.cols }, (_, c) => this.sheet.cell(this.row + r, this.col + c))); }
  getValue() { return this.sheet.cell(this.row, this.col); }
  setValue(value) { this.sheet.setCell(this.row, this.col, value); return this; }
  setValues(values) { values.forEach((row, r) => row.forEach((value, c) => this.sheet.setCell(this.row + r, this.col + c, value))); return this; }
  setBackground() { return this; } setFontColor() { return this; } setFontFamily() { return this; }
  setFontWeight() { return this; } setFontSize() { return this; } setVerticalAlignment() { return this; }
  setHorizontalAlignment() { return this; } setFontStyle() { return this; }
}
class Sheet {
  constructor(name) { this.name = name; this.rows = []; }
  cell(row, col) { return (this.rows[row - 1] || [])[col - 1] ?? ''; }
  setCell(row, col, value) { while (this.rows.length < row) this.rows.push([]); while (this.rows[row - 1].length < col) this.rows[row - 1].push(''); this.rows[row - 1][col - 1] = value; }
  getDataRange() { return new Range(this, 1, 1, Math.max(this.rows.length, 1), Math.max(this.getLastColumn(), 1)); }
  getRange(row, col, rows, cols) { return new Range(this, row, col, rows, cols); }
  getLastColumn() { return this.rows.reduce((max, row) => Math.max(max, row.length), 0); }
  getLastRow() { return this.rows.length; }
  appendRow(row) { this.rows.push(row.slice()); }
  setFrozenRows() {}
}
class Spreadsheet {
  constructor() { this.sheets = {}; }
  getSheetByName(name) { return this.sheets[name] || null; }
  insertSheet(name) { return this.sheets[name] = new Sheet(name); }
}

const spreadsheet = new Spreadsheet();
const patientEmails = [];
const regianeEmails = [];
const cacheValues = new Map();
const context = {
  console,
  SpreadsheetApp: { openById: () => spreadsheet, flush: () => {} },
  GmailApp: {
    getAliases: () => ['regianemariasilva22@gmail.com'],
    sendEmail: (to, subject, body, options) => patientEmails.push({ to, subject, body, options })
  },
  MailApp: { sendEmail: message => regianeEmails.push(message) },
  UrlFetchApp: { fetch: () => ({ getContentText: () => JSON.stringify({ error: 'invalid_token' }) }) },
  Utilities: { getUuid: () => 'uuid-' + Math.random().toString(16).slice(2) },
  CacheService: { getScriptCache: () => ({
    get: key => cacheValues.get(key) || null,
    put: (key, value) => cacheValues.set(key, value),
    remove: key => cacheValues.delete(key)
  }) },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  ScriptApp: {}, CalendarApp: {}, DriveApp: {}, ContentService: {}, Session: {},
  Date, JSON, Object, Array, String, Number, Math, RegExp, Error, isFinite
};
vm.createContext(context);
vm.runInContext(fs.readFileSync('google-apps-script/Code.gs', 'utf8'), context);

const tokens = {
  admin: { email: 'regianemariasilva22@gmail.com', nome: 'Regiane' },
  patient: { email: 'paciente.teste@example.com', nome: 'Paciente Teste' }
};
context.verifyGoogleToken = token => {
  if (!tokens[token]) throw new Error('Token inválido');
  return tokens[token];
};

assert.strictEqual(context.setupConsultoriaSheets().ok, true);
assert.ok(spreadsheet.getSheetByName('ConsultoriaPacientes'));
assert.ok(spreadsheet.getSheetByName('ConsultoriaRespostas'));
const materiais = spreadsheet.insertSheet('Materiais');
materiais.appendRow(['Id', 'Email', 'Tipo', 'Titulo', 'Descricao', 'Link', 'Area']);

const added = context.actionAdminAddConsultoriaPatient({ idToken: 'admin', nome: 'Paciente Teste', email: 'paciente.teste@example.com' });
assert.strictEqual(added.ok, true);
assert.strictEqual(context.actionGoogleLoginConsultoria({ idToken: 'patient' }).ok, true);

const requestedCode = context.actionRequestConsultoriaAccessCode({ email: 'paciente.teste@example.com' });
assert.strictEqual(requestedCode.ok, true);
const deliveredCode = patientEmails[0].body.match(/\b(\d{6})\b/)[1];
assert.strictEqual(context.actionVerifyConsultoriaAccessCode({ email: 'paciente.teste@example.com', code: '000000' }).ok, false);
const verifiedCode = context.actionVerifyConsultoriaAccessCode({ email: 'paciente.teste@example.com', code: deliveredCode });
assert.strictEqual(verifiedCode.ok, true);
assert.ok(verifiedCode.accessToken);
assert.strictEqual(context.actionConsultoriaDashboard({ accessToken: verifiedCode.accessToken }).ok, true);

const respostas = {
  modalidade: 'Online', objetivo: 'Melhorar a alimentação', medicamentos: 'Não uso', patologias: 'Não possuo',
  intestino: 'Funciona diariamente', alergias: 'Não possuo', preferencias: 'Prefiro arroz e feijão',
  cafe: 'Café e pão', almoco: 'Arroz, feijão e proteína', lanche: 'Fruta', jantar: 'Sopa', outros: ''
};
const submitted = context.actionSubmitConsultoriaAnamnese({ idToken: 'patient', respostas });
assert.strictEqual(submitted.ok, true);
assert.strictEqual(submitted.emailPacienteEnviado, true);
assert.strictEqual(submitted.emailRegianeEnviado, true);
assert.strictEqual(patientEmails[1].options.from, 'regianemariasilva22@gmail.com');
assert.strictEqual(patientEmails[1].to, 'paciente.teste@example.com');
assert.strictEqual(regianeEmails[0].to, 'regianemariasilva22@gmail.com');
assert.strictEqual(spreadsheet.getSheetByName('ConsultoriaRespostas').rows.length, 2);

const disabled = context.actionAdminSetConsultoriaAccess({ idToken: 'admin', email: 'paciente.teste@example.com', liberado: false });
assert.strictEqual(disabled.ok, true);
assert.strictEqual(context.actionGoogleLoginConsultoria({ idToken: 'patient' }).ok, false);
const adminList = context.actionAdminListConsultoriaPatients({ idToken: 'admin' });
assert.strictEqual(adminList.pacientes[0].respostas.objetivo, respostas.objetivo);
assert.strictEqual(adminList.pacientes[0].liberado, false);

const pacientes = spreadsheet.insertSheet('Pacientes');
pacientes.appendRow(['Email', 'Nome', 'DataInicio', 'RetornosRealizados', 'ReceitasSalvas', 'ProgressoPercent', 'PontosTotal', 'Instagram']);
const checkups = spreadsheet.insertSheet('CheckupPacientes');
checkups.appendRow(['Email', 'Nome', 'DataLiberacao', 'Liberado', 'JaFezCheckup', 'DataCheckup', 'RespostasChecklist', 'RespostasQuiz']);

const financialCases = [
  { id: 'fin-consultoria', email: 'consultoria@example.com', service: 'Consultoria', sheet: 'ConsultoriaPacientes' },
  { id: 'fin-checkup', email: 'checkup@example.com', service: 'Check-up', sheet: 'CheckupPacientes' },
  { id: 'fin-dupla', email: 'dupla@example.com', service: 'Consulta Dupla', sheet: 'Pacientes' },
  { id: 'fin-programa', email: 'programa@example.com', service: 'Programa Reequilíbrio', sheet: 'Pacientes' },
  { id: 'fin-acompanhamento', email: 'acompanhamento@example.com', service: 'Acompanhamento de 5 meses', sheet: 'Pacientes' }
];
financialCases.forEach(item => {
  const result = context.actionFinanceiroSavePatient({
    idToken: 'admin',
    patient: { id: item.id, email: item.email, name: 'Paciente ' + item.id, service: item.service, value: 100 }
  });
  assert.strictEqual(result.ok, true);
  const target = spreadsheet.getSheetByName(item.sheet);
  const emailCol = target.rows[0].indexOf('Email');
  assert.ok(target.rows.slice(1).some(row => row[emailCol] === item.email), item.service + ' não liberou ' + item.sheet);
});
assert.strictEqual(spreadsheet.getSheetByName('FinanceiroPacientes').rows.length, financialCases.length + 1);

console.log('Consultoria e financeiro: cadastro, áreas corretas, login, respostas, e-mails e histórico OK');
