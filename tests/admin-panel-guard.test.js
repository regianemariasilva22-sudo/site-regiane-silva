const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

function extractShowTab(source) {
  const start = source.indexOf('function showTab(name)');
  assert.notStrictEqual(start, -1, 'dashboard must define showTab(name)');
  const open = source.indexOf('{', start);
  let depth = 0;
  for (let i = open; i < source.length; i++) {
    if (source[i] === '{') depth++;
    if (source[i] === '}' && --depth === 0) return source.slice(start, i + 1);
  }
  throw new Error('showTab(name) has unbalanced braces');
}

function exerciseDashboard(file, isAdminUser) {
  const adminTab = { style: { display: 'none' } };
  const homeTab = { style: { display: 'none' } };
  let adminLoads = 0;
  const links = [{ classList: { remove() {}, add() {} } }];
  const context = {
    isAdminUser,
    event: {},
    window: { scrollTo() {}, __rebindCursor() {} },
    document: {
      querySelectorAll(selector) { return selector.startsWith('[id^="tab-"') ? [homeTab, adminTab] : links; },
      getElementById(id) { return id === 'tab-admin' ? adminTab : id === 'tab-inicio' ? homeTab : null; }
    },
    loadAdmin() { adminLoads++; },
    loadAdminBookings() { adminLoads++; },
    loadAdminPatients() { adminLoads++; },
    loadAdminCheckupPatients() { adminLoads++; }
  };
  vm.runInNewContext(`${extractShowTab(fs.readFileSync(file, 'utf8'))}; showTab('admin');`, context);
  return { adminVisible: adminTab.style.display === 'block', homeVisible: homeTab.style.display === 'block', adminLoads };
}

const dashboards = [
  'membros/programa/dashboard.html',
  'membros/checkup/dashboard.html',
  'membros/consultoria/dashboard.html'
];

for (const file of dashboards) {
  const patientView = exerciseDashboard(file, false);
  assert.deepStrictEqual(patientView, { adminVisible: false, homeVisible: true, adminLoads: 0 }, `${file}: patient must not see or load admin panel`);
  const adminView = exerciseDashboard(file, true);
  assert.strictEqual(adminView.adminVisible, true, `${file}: administrator must retain panel access`);
}

const program = fs.readFileSync('membros/programa/dashboard.html', 'utf8');
const backend = fs.readFileSync('google-apps-script/Code.gs', 'utf8');
assert.match(program, /Escolha um horário para enviar sua solicitação\. Ele ficará pendente até a Regiane confirmar/);
assert.match(program, /setInterval\(\(\)=>\{ if\(pendingAppointmentStatusRefresh\) loadDashboardData\(\); \}, 60000\)/);
assert.match(backend, /\['solicitado', 'confirmado'\]\.indexOf\(String\(a\.Status\)\.trim\(\)\.toLowerCase\(\)\)/);

console.log('Aba administrativa bloqueada para pacientes nos três painéis; agenda informa pendência e atualiza o status.');
