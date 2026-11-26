/**
 * Recebe os cadastros da página da Black Friday e grava na aba "Leads".
 * Cole este código em: Planilha > Extensões > Apps Script (veja LEIA-ME.txt).
 */
const ABA = 'Leads';
const CABECALHO = ['Data/hora', 'Nome', 'WhatsApp', 'E-mail', 'utm_source', 'utm_medium', 'utm_campaign', 'utm_content', 'utm_term', 'Página'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(ABA);
    if (!sh) {
      sh = ss.insertSheet(ABA);
      sh.appendRow(CABECALHO);
      sh.setFrozenRows(1);
    }
    const p = (e && e.parameter) || {};
    // Grava tudo como texto e impede fórmulas maliciosas (=, +, -, @)
    const txt = v => {
      v = String(v || '').trim().slice(0, 200);
      return /^[=+\-@]/.test(v) ? "'" + v : v;
    };
    sh.appendRow([
      new Date(),
      txt(p.nome),
      "'" + String(p.whatsapp || '').replace(/\D/g, ''),
      txt(p.email),
      txt(p.utm_source), txt(p.utm_medium), txt(p.utm_campaign), txt(p.utm_content), txt(p.utm_term),
      txt(p.pagina)
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('Planilha de leads da Black Friday: funcionando.');
}
