const SHEET_ID = '1UHpshK0flhE0Bq57h0FxFPIrPom7KSRpkRYE_wQ-0QU';

// ── QR CODE DO DIA ─────────────────────────────────────────
const SECRET = 'HA-CARNAXIDE-2024';

function getQRCodeDoDia() {
  const hoje = Utilities.formatDate(new Date(), 'Europe/Lisbon', 'yyyyMMdd');
  const raw = hoje + SECRET;
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw);
  return bytes.slice(0, 4).map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('').toUpperCase();
}

// ── ROUTING ────────────────────────────────────────────────
function doGet(e) {
  const page = (e.parameter.page || 'client');
  if (page === 'staff') return serveStaff();
  return serveClient();
}

function doPost(e) {
  const action = e.parameter.action || '';
  let result;
  if (action === 'registar')  result = registarCliente(e);
  else if (action === 'scan') result = registarScan(e);
  else if (action === 'cliente') result = getCliente(e);
  else if (action === 'novidades') result = getNovidades();
  else result = { ok: false, erro: 'Accao desconhecida' };

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── SERVE CLIENT APP ───────────────────────────────────────
function serveClient() {
  const url = 'https://raw.githubusercontent.com/davidpinheiro-7777777/ha-fidelizacao/main/client.html?nc=' + Date.now();
  const html = UrlFetchApp.fetch(url).getContentText();
  return HtmlService.createHtmlOutput(html)
    .setTitle('Ha - Clube de Fidelizacao')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── SERVE STAFF APP ────────────────────────────────────────
function serveStaff() {
  const url = 'https://raw.githubusercontent.com/davidpinheiro-7777777/ha-fidelizacao/main/staff.html?nc=' + Date.now();
  const html = UrlFetchApp.fetch(url).getContentText();
  return HtmlService.createHtmlOutput(html)
    .setTitle('Ha - Staff')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── REGISTAR CLIENTE ───────────────────────────────────────
function registarCliente(e) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Clientes');
  const nome     = e.parameter.nome || '';
  const apelido  = e.parameter.apelido || '';
  const telefone = e.parameter.telefone || '';
  const email    = e.parameter.email || '';
  const nasc     = e.parameter.nascimento || '';

  if (!nome || !email) return { ok: false, erro: 'Dados incompletos' };

  // Verifica se ja existe pelo email
  const dados = sheet.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][4]).toLowerCase() === email.toLowerCase()) {
      return {
        ok: true, jaExiste: true,
        id: dados[i][0], nome: dados[i][1], apelido: dados[i][2],
        visitas: parseInt(dados[i][7]) || 0
      };
    }
  }

  // Cria novo cliente
  const id = 'C' + Date.now();
  const agora = Utilities.formatDate(new Date(), 'Europe/Lisbon', 'dd/MM/yyyy HH:mm');
  sheet.appendRow([id, nome, apelido, telefone, email, nasc, agora, 0, '']);
  return { ok: true, jaExiste: false, id, nome, apelido, visitas: 0 };
}

// ── GET CLIENTE ────────────────────────────────────────────
function getCliente(e) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Clientes');
  const id = e.parameter.id || '';
  const email = e.parameter.email || '';
  const dados = sheet.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0] === id || String(dados[i][4]).toLowerCase() === email.toLowerCase()) {
      return {
        ok: true, id: dados[i][0], nome: dados[i][1], apelido: dados[i][2],
        visitas: parseInt(dados[i][7]) || 0, ultimaVisita: dados[i][8] || ''
      };
    }
  }
  return { ok: false, erro: 'Cliente nao encontrado' };
}

// ── REGISTAR SCAN ──────────────────────────────────────────
function registarScan(e) {
  const qr       = e.parameter.qr || '';
  const clientId = e.parameter.id || '';

  // Valida QR code do dia
  if (qr.toUpperCase() !== getQRCodeDoDia()) {
    return { ok: false, erro: 'QR code invalido ou expirado' };
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheetC = ss.getSheetByName('Clientes');
  const sheetV = ss.getSheetByName('Visitas');
  const dados = sheetC.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0] === clientId) {
      const visitas = (parseInt(dados[i][7]) || 0) + 1;
      const agora = Utilities.formatDate(new Date(), 'Europe/Lisbon', 'dd/MM/yyyy HH:mm');
      const refeicaoGratis = visitas >= 10;
      const novasVisitas = refeicaoGratis ? 0 : visitas;

      sheetC.getRange(i + 1, 8).setValue(novasVisitas);
      sheetC.getRange(i + 1, 9).setValue(agora);
      sheetV.appendRow([clientId, agora, qr]);

      return {
        ok: true, visitas: novasVisitas, refeicaoGratis,
        mensagem: refeicaoGratis
          ? 'Parabens! Refeicao gratis desbloqueada!'
          : 'Refeicao registada! Tens agora ' + novasVisitas + ' de 10.'
      };
    }
  }
  return { ok: false, erro: 'Cliente nao encontrado' };
}

// ── NOVIDADES ──────────────────────────────────────────────
function getNovidades() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Novidades');
  const dados = sheet.getDataRange().getValues();
  const novidades = [];
  for (let i = 1; i < dados.length; i++) {
    const activo = String(dados[i][4]).toLowerCase();
    if (activo === 'sim' || activo === 'true' || activo === '1') {
      novidades.push({
        titulo: dados[i][0], descricao: dados[i][1],
        data: dados[i][2], foto: dados[i][3]
      });
    }
  }
  return { ok: true, novidades };
}
