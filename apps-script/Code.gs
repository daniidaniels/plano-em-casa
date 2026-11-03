// Plano em Casa · sincronização entre os celulares
// 1) Cole este código em Extensões › Apps Script da planilha
// 2) Troque a CHAVE abaixo por uma senha só de vocês
// 3) Implantar › Nova implantação › App da Web
//    Executar como: Eu · Quem pode acessar: Qualquer pessoa
const CHAVE = 'troque-esta-chave';

function aba_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('docs');
  if (!sh) { sh = ss.insertSheet('docs'); sh.appendRow(['path', 'json', 't']); }
  return sh;
}
function resp_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) {
  if ((e.parameter.k || '') !== CHAVE) return resp_({ok: false, error: 'chave inválida'});
  const since = Number(e.parameter.since || 0);
  const rows = aba_().getDataRange().getValues().slice(1);
  const docs = rows.filter(r => Number(r[2]) > since).map(r => ({path: r[0], data: JSON.parse(r[1]), t: Number(r[2])}));
  return resp_({ok: true, now: Date.now(), docs});
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents || '{}');
  if ((b.k || '') !== CHAVE) return resp_({ok: false, error: 'chave inválida'});
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = aba_(), vals = sh.getDataRange().getValues(), idx = {};
    vals.forEach((r, i) => { if (i) idx[r[0]] = i + 1; });
    (b.docs || []).forEach(d => {
      const row = [d.path, JSON.stringify(d.data), Number(d.t) || Date.now()];
      const at = idx[d.path];
      if (at) { if (Number(sh.getRange(at, 3).getValue()) <= row[2]) sh.getRange(at, 1, 1, 3).setValues([row]); }
      else { sh.appendRow(row); idx[d.path] = sh.getLastRow(); }
    });
    return resp_({ok: true});
  } finally { lock.releaseLock(); }
}