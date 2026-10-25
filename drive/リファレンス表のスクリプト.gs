/**
 * リファレンス表(Google スプレッドシート)に「サイト」メニューを足すスクリプト。
 * 入れ方:表を開く → 拡張機能 → Apps Script → このファイルの中身を全部貼り付けて保存 → 表を開き直す。
 * 初回だけ「承認が必要です」と出るので、自分のアカウントで許可する。
 *
 * サイト → 画像と作品を取り込む
 *   - references フォルダ(サブフォルダも)にあって表にない画像を、下に1行ずつ足す(見た目・画像・題名・ID)
 *   - 画像の名前を変えていたら、表の「画像」も新しい名前にする
 *   - 「作品一覧」を works フォルダの今のフォルダ名に、「タグ一覧」を表で使っているタグにそろえる
 * 表はホームページのフォルダ(works / references / texts のある所)か、references の中に置く。
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('サイト')
    .addItem('画像と作品を取り込む', 'importAll')
    .addToUi();
}

function importAll() {
  const ss = SpreadsheetApp.getActive();
  const home = findHome_(ss);
  if (!home) { SpreadsheetApp.getUi().alert('works と references のあるフォルダが見つかりません。表をホームページのフォルダに移してください。'); return; }

  const sh = ss.getSheetByName('リファレンス') || ss.getSheets()[0];
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const col = name => head.findIndex(h => h.replace(/\s/g, '') === name) + 1;
  const cImg = col('画像'), cTitle = col('題名'), cThumb = col('見た目'), cId = col('ID');
  if (!cImg || !cId) { SpreadsheetApp.getUi().alert('1行目に「画像」と「ID」の列が必要です。'); return; }

  // references の画像(サブフォルダも)
  const refs = home.getFoldersByName('references');
  const files = [];
  if (refs.hasNext()) collect_(refs.next(), files);

  const last = Math.max(sh.getLastRow(), 1);
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, head.length).getValues() : [];
  const norm = s => String(s).normalize('NFKC').toLowerCase().replace(/\.[a-z0-9]{2,5}$/, '').replace(/^\d+[\s_.\-]*/, '').replace(/[\s_\-—–・.,]+/g, ' ').trim();
  let added = 0, renamed = 0;
  files.forEach(f => {
    const id = f.getId();
    let r = rows.findIndex(row => row[cId - 1] === id);
    if (r < 0) r = rows.findIndex(row => !row[cId - 1] && row[cImg - 1] && (String(row[cImg - 1]) === f.getName() || norm(row[cImg - 1]) === norm(f.getName())));
    if (r >= 0) {
      const line = r + 2;
      if (!rows[r][cId - 1]) sh.getRange(line, cId).setValue(id);
      if (String(rows[r][cImg - 1]) !== f.getName()) { sh.getRange(line, cImg).setValue(f.getName()); renamed++; }
      if (cThumb && !String(sh.getRange(line, cThumb).getFormula())) sh.getRange(line, cThumb).setFormula(thumb_(id));
      return;
    }
    const row = new Array(head.length).fill('');
    row[cImg - 1] = f.getName();
    if (cTitle) row[cTitle - 1] = f.getName().replace(/\.[^.]+$/, '').replace(/^\d+[\s_.\-]*/, '');
    row[cId - 1] = id;
    sh.appendRow(row);
    if (cThumb) sh.getRange(sh.getLastRow(), cThumb).setFormula(thumb_(id));
    added++;
  });
  if (cThumb && sh.getLastRow() > 1) sh.setRowHeights(2, sh.getLastRow() - 1, 84);

  // 作品一覧
  const works = home.getFoldersByName('works');
  const names = [];
  if (works.hasNext()) { const it = works.next().getFolders(); while (it.hasNext()) names.push(it.next().getName()); }
  names.sort((a, b) => b.localeCompare(a, 'ja', { numeric: true }));
  fillList_(ss, '作品一覧', '作品', names);

  // タグ一覧:今の一覧+表で使われているタグ
  const cTags = col('タグ');
  if (cTags) {
    const tl = ss.getSheetByName('タグ一覧');
    const have = tl && tl.getLastRow() > 1 ? tl.getRange(2, 1, tl.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
    const used = sh.getLastRow() > 1 ? sh.getRange(2, cTags, sh.getLastRow() - 1, 1).getValues().flatMap(r => String(r[0]).split(/[,、]/)) : [];
    const tags = [...new Set([...have, ...used].map(s => s.trim()).filter(Boolean))].sort();
    fillList_(ss, 'タグ一覧', 'タグ', tags);
  }
  ss.toast(`新しい画像 ${added} 件を足しました` + (renamed ? `(名前の変わった画像 ${renamed} 件)` : '') + `。作品 ${names.length} 件。`, 'サイト', 8);
}

function findHome_(ss) {
  const parents = DriveApp.getFileById(ss.getId()).getParents();
  while (parents.hasNext()) {
    const p = parents.next();
    if (p.getFoldersByName('works').hasNext()) return p;
    if (p.getName() === 'references') { const up = p.getParents(); if (up.hasNext()) return up.next(); }
  }
  return null;
}
function collect_(folder, out) {
  const fs = folder.getFiles();
  const list = [];
  while (fs.hasNext()) { const f = fs.next(); if (/^image\/|\/pdf$/.test(f.getMimeType())) list.push(f); }
  list.sort((a, b) => a.getName().localeCompare(b.getName(), 'ja', { numeric: true }));
  out.push(...list);
  const ds = folder.getFolders();
  const subs = [];
  while (ds.hasNext()) subs.push(ds.next());
  subs.sort((a, b) => a.getName().localeCompare(b.getName(), 'ja', { numeric: true })).forEach(d => collect_(d, out));
}
function thumb_(id) { return '=IMAGE("https://drive.google.com/thumbnail?id=' + id + '&sz=w400")'; }
function fillList_(ss, name, title, values) {
  const s = ss.getSheetByName(name) || ss.insertSheet(name);
  s.getRange(2, 1, Math.max(s.getMaxRows() - 1, 1), 1).clearContent();
  s.getRange(1, 1).setValue(title);
  if (values.length) s.getRange(2, 1, values.length, 1).setValues(values.map(v => [v]));
}
