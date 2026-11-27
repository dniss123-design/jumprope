const SHEET = '순위';
const MAX_SCORE = 200000;
const PHOTO_RE = /^data:image\/jpeg;base64,[A-Za-z0-9+/=]+$/;

function doGet(e) {
  const p = (e && e.parameter) || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET);
    if (!sh) {
      sh = ss.insertSheet(SHEET);
      sh.appendRow(['기기ID', '이름', '점수', '신발', '도달', '날짜', '사진']);
    }
    if (sh.getRange(1, 7).getValue() !== '사진') sh.getRange(1, 7).setValue('사진');
    let data = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues() : [];
    const id = String(p.id || '').slice(0, 40);

    let photo = null;
    if (p.p === 'none') photo = '';
    else if (typeof p.p === 'string' && p.p.length < 16000 && PHOTO_RE.test(p.p)) photo = p.p;

    if (p.a === 'save') {
      const n = String(p.n || '').trim().slice(0, 8);
      const s = Math.floor(Number(p.s) || 0);
      if (id && n && s >= 0 && s <= MAX_SCORE) {
        const i = data.findIndex(r => r[0] === id);
        if (i < 0) {
          const row = [id, n, s, Number(p.sh) || 0, String(p.st || '').slice(0, 12), new Date(), photo || ''];
          sh.appendRow(row); data.push(row);
        } else if (s > Number(data[i][2])) {
          const row = [id, n, s, Number(p.sh) || 0, String(p.st || '').slice(0, 12), new Date(), photo === null ? data[i][6] : photo];
          sh.getRange(i + 2, 1, 1, 7).setValues([row]); data[i] = row;
        } else {
          sh.getRange(i + 2, 2).setValue(n); data[i][1] = n;
          if (photo !== null) { sh.getRange(i + 2, 7).setValue(photo); data[i][6] = photo; }
        }
      }
    }

    data.sort((a, b) => Number(b[2]) - Number(a[2]));
    const top = data.slice(0, 10).map(r => ({ n: String(r[1]), s: Number(r[2]), sh: Number(r[3]), st: String(r[4]), p: String(r[6] || ''), me: !!id && r[0] === id }));
    const idx = id ? data.findIndex(r => r[0] === id) : -1;
    const out = {
      top: top,
      total: data.length,
      rank: idx + 1,
      me: idx >= 0 ? { n: String(data[idx][1]), s: Number(data[idx][2]) } : null
    };
    return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
