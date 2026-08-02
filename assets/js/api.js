/* ============================================================
   REGIANE SILVA — cliente da API (Google Apps Script)
   Veja google-apps-script/INSTRUCOES.md para publicar o backend.
   ============================================================ */

// Depois de publicar o Apps Script como Web App, cole a URL aqui:
const API_URL = 'https://script.google.com/macros/s/AKfycbykDiX4gWybvlRh2s-iZCb_q3UH0bE2HVuG23eYsvu-nyOEHCTyF4oquAZf1BzYqdGN2Q/exec';

// Client ID do OAuth do Google (Sign In With Google) — já criado no Google Cloud Console
const GOOGLE_CLIENT_ID = '288771217381-mt5g3dhdjhcoak6kphd1fhsarrkd44bc.apps.googleusercontent.com';

function apiConfigured() {
  return typeof API_URL === 'string' && API_URL.indexOf('COLE_AQUI') === -1;
}

async function apiGet(action, params) {
  params = params || {};
  if (!apiConfigured()) return { ok: false, error: 'API ainda não configurada.' };
  const qs = new URLSearchParams(Object.assign({ action: action }, params)).toString();
  try {
    const res = await fetch(API_URL + '?' + qs);
    return await res.json();
  } catch (err) {
    return { ok: false, error: 'Não foi possível conectar ao servidor. Tente novamente em instantes.' };
  }
}

async function apiPost(action, data) {
  data = data || {};
  if (!apiConfigured()) return { ok: false, error: 'API ainda não configurada.' };
  try {
    const res = await fetch(API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(Object.assign({ action: action }, data))
    });
    return await res.json();
  } catch (err) {
    return { ok: false, error: 'Não foi possível conectar ao servidor. Tente novamente em instantes.' };
  }
}
