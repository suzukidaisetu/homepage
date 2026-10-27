/**
 * リファレンス表(Google スプレッドシート)に「サイト」メニューを足すスクリプト。
 * 入れ方:表を開く → 拡張機能 → Apps Script → このファイルの中身を全部貼り付けて保存 → 表を開き直す。
 * 初回だけ「承認が必要です」と出るので、自分のアカウントで許可する。
 *
 * サイト → 画像と作品を取り込む
 *   - references フォルダ(サブフォルダも)にあって表にない画像を、下に1行ずつ足す(見た目・画像・題名・ID)
 *   - 画像の名前を変えていたら、表の「画像」も新しい名前にする
 *   - 「作品一覧」を works フォルダの今のフォルダ名に、「タグ一覧」を表で使っているタグにそろえる
 * サイト → 英語を今すぐ入れる / 英語を毎時自動で入れる / 自動の英訳を止める
 *   - 日本語だけ書いてあって英語が空いている所を Google 翻訳で埋める(詳しくは下の「英語の自動翻訳」)
 * 表はホームページのフォルダ(works / references / texts のある所)か、references の中に置く。
 * スクリプトを更新したときは、中身を全部消してから貼り直して保存する。
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('サイト')
    .addItem('画像と作品を取り込む', 'importAll')
    .addSeparator()
    .addItem('英語を今すぐ入れる', 'translateAll')
    .addItem('英語を毎時自動で入れる', 'autoTranslateOn')
    .addItem('自動の英訳を止める', 'autoTranslateOff')
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

/* ================================================================
 * 英語の自動翻訳
 * 日本語だけ書いてあって英語が空いている所を、Google 翻訳で埋める。
 *   - リファレンス表の「一言」→「一言(英)」
 *   - texts の Google ドキュメント(「English」から下を作る)
 *   - works の各フォルダの info と about(「--- en」から下を作る)
 * 自分で書いた英語には触らない。自動で入れた英語は、日本語を直したときだけ作り直す
 * (自動で入れた英語を自分で直すと、それ以降は触らない)。
 * 印として、自動で入れた所には「自動翻訳: ...」の行(表ではセルのメモ)が付く。サイトには出ない。
 * 訳し方を決めたい言葉は、「訳語」シートに 日本語|英語 で書いておくと、その英語で訳す。
 * ================================================================ */

const AUTO_ = '自動翻訳';
const EN_MARK_ = /^(?:-{2,}\s*)?(en|english|英語)$/i;
const JA_MARK_ = /^(?:-{2,}\s*)?(ja|日本語)$/i;
const AUTO_LINE_ = /^自動翻訳\s*[:：]\s*([0-9a-f]+)\.([0-9a-f]+)\s*$/;

function hash_(s) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(s).replace(/\s+/g, ' ').trim(), Utilities.Charset.UTF_8);
  return b.slice(0, 5).map(x => ((x + 256) % 256).toString(16).padStart(2, '0')).join('');
}

/* the glossary sheet 「訳語」: Japanese words put in English before translating, so they come out as you want */
function glossary_(ss) {
  const sh = ss && ss.getSheetByName('訳語');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().filter(r => r[0] && r[1]).map(r => [String(r[0]), String(r[1])]).sort((a, b) => b[0].length - a[0].length);
}
function translator_(ss) {
  const g = glossary_(ss);
  const cache = {};
  return s => {
    s = String(s || '');
    if (!s.trim() || !/[぀-ヿ㐀-鿿]/.test(s)) return s;   // nothing Japanese: keep as it is
    if (cache[s] != null) return cache[s];
    let src = s;
    g.forEach(([ja, en]) => { src = src.split(ja).join(' ' + en + ' '); });
    let out = LanguageApp.translate(src, 'ja', 'en');
    out = out.replace(/\s{2,}/g, ' ').replace(/\s+([,.;:!?)])/g, '$1').trim();
    return (cache[s] = out);
  };
}

function translateAll() {
  const ss = SpreadsheetApp.getActive();
  const home = findHome_(ss);
  const tr = translator_(ss);
  const done = { sheet: 0, texts: 0, works: 0, about: 0 };
  done.sheet = translateSheet_(ss, tr);
  if (home) {
    const texts = home.getFoldersByName('texts');
    if (texts.hasNext()) {
      const it = texts.next().getFiles();
      while (it.hasNext()) { const f = it.next(); if (translateFile_(f, tr, 'English')) done.texts++; }
    }
    const works = home.getFoldersByName('works');
    if (works.hasNext()) {
      const wf = works.next().getFolders();
      while (wf.hasNext()) {
        const fs = wf.next().getFiles();
        while (fs.hasNext()) { const f = fs.next(); if (/^info(\.txt)?$/i.test(f.getName()) && translateFile_(f, tr, '--- en')) done.works++; }
      }
    }
    const top = home.getFiles();
    while (top.hasNext()) { const f = top.next(); if (/^about(\.txt)?$/i.test(f.getName()) && translateFile_(f, tr, '--- en')) done.about++; }
  }
  const msg = `英語を入れました:リファレンス ${done.sheet} 件、テキスト ${done.texts} 本、作品 ${done.works} 件、about ${done.about}`;
  try { ss.toast(msg, 'サイト', 8); } catch (e) {}
  console.log(msg);
}

/* the reference sheet: 一言 → 一言(英) */
function translateSheet_(ss, tr) {
  const sh = ss.getSheetByName('リファレンス') || ss.getSheets()[0];
  if (sh.getLastRow() < 2) return 0;
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).normalize('NFKC').replace(/[\s()()]/g, ''));
  const cJa = head.indexOf('一言') + 1, cEn = head.findIndex(h => /^(一言英|一言英語|noteen)$/i.test(h)) + 1;
  if (!cJa || !cEn) return 0;
  const n = sh.getLastRow() - 1;
  const ja = sh.getRange(2, cJa, n, 1).getValues(), enR = sh.getRange(2, cEn, n, 1), en = enR.getValues(), notes = enR.getNotes();
  let count = 0;
  for (let i = 0; i < n; i++) {
    const j = String(ja[i][0] || '').trim(), e = String(en[i][0] || '').trim();
    const m = String(notes[i][0] || '').match(AUTO_LINE_);
    if (!j) { if (m && e && hash_(e) === m[2]) { en[i][0] = ''; notes[i][0] = ''; count++; } continue; }   // Japanese removed: remove its automatic English too
    if (e && (!m || hash_(e) !== m[2] || hash_(j) === m[1])) continue;   // your own English, or already up to date
    const out = j.split(/\n\s*\n/).map(p => tr(p.replace(/\n/g, ' '))).join('\n\n');
    en[i][0] = out; notes[i][0] = `${AUTO_}: ${hash_(j)}.${hash_(out)}`; count++;
  }
  if (count) { enR.setValues(en); enR.setNotes(notes); }
  return count;
}

function translateFile_(f, tr, mark) {
  const type = f.getMimeType();
  try {
    if (type === MimeType.GOOGLE_DOCS) return translateDoc_(DocumentApp.openById(f.getId()), tr, mark);
    if (type === MimeType.PLAIN_TEXT || /\.txt$/i.test(f.getName())) {
      const src = f.getBlob().getDataAsString('UTF-8');
      const out = translatePlain_(src, tr);
      if (out != null && out !== src) { f.setContent(out); return true; }
    }
  } catch (e) { console.warn(f.getName() + ': ' + e.message); }
  return false;
}

/* a line of Japanese turned into its English line: headings keep "# ", "kind:" style lines keep their key */
function trLine_(line, tr) {
  const t = line.trim();
  if (!t) return '';
  if (/^#\s/.test(t)) return '# ' + tr(t.replace(/^#\s*/, ''));
  const kv = t.match(/^([^:：]{1,24}?)\s*[:：]\s*(.*)$/);
  if (kv && !/^https?$/i.test(kv[1])) {
    const k = kv[1].trim(), v = kv[2].trim();
    const key = /^(kind|種類)$/i.test(k) ? 'kind' : /^(role|肩書き?)$/i.test(k) ? 'role' : /^(title|題名)$/i.test(k) ? 'title' : tr(k);
    const val = !v || /^https?:\/\//.test(v) ? v : key === 'role' ? v.split('/').map(x => tr(x.trim())).join(' / ') : tr(v);
    return key + ': ' + val;
  }
  return tr(t);
}

/* .txt (info, about, old texts): "--- ja" … "--- en" */
function translatePlain_(src, tr) {
  const text = String(src).replace(/\r\n?/g, '\n');
  const ja = text.match(/^---\s*ja\s*$/m);
  if (!ja) return null;
  const enAt = text.search(/^---\s*en\s*$/m);
  const jaBody = text.slice(ja.index + ja[0].length, enAt < 0 ? undefined : enAt).replace(/^\n+|\s+$/g, '');
  if (!jaBody.trim()) return null;
  const noTitle = t => t.split('\n').filter(l => !/^#\s/.test(l.trim())).join('\n');   // the English title is yours to rewrite
  let keepTitle = null;
  if (enAt >= 0) {
    const lines = text.slice(enAt).split('\n').slice(1);
    const autoIdx = lines.findIndex(l => AUTO_LINE_.test(l.trim()));
    const enBody = lines.filter((l, k) => k !== autoIdx).join('\n').trim();
    if (enBody) {
      if (autoIdx < 0) return null;                          // your own English
      const m = lines[autoIdx].trim().match(AUTO_LINE_);
      if (hash_(noTitle(enBody)) !== m[2] || hash_(jaBody) === m[1]) return null;   // edited by you, or up to date
      keepTitle = lines.find(l => /^#\s/.test(l.trim())) || null;
    }
  }
  let out = jaBody.split('\n').map(l => trLine_(l, tr)).join('\n').replace(/\n{3,}/g, '\n\n').trim();
  if (keepTitle) out = out.replace(/^#\s.*$/m, keepTitle.trim());
  const before = (enAt < 0 ? text : text.slice(0, enAt)).replace(/\s+$/, '');
  return `${before}\n\n--- en\n${out}\n\n${AUTO_}: ${hash_(jaBody)}.${hash_(noTitle(out))}\n`;
}

/* a Google Doc: paragraphs up to "English" (or "--- en") are the Japanese; the English goes below */
function translateDoc_(doc, tr, mark) {
  const body = doc.getBody();
  const paras = [];
  for (let i = 0; i < body.getNumChildren(); i++) {
    const c = body.getChild(i), t = c.getType();
    if (t === DocumentApp.ElementType.PARAGRAPH || t === DocumentApp.ElementType.LIST_ITEM) paras.push({ i, el: c, text: c.asText().getText().trim(), heading: c.getType() === DocumentApp.ElementType.PARAGRAPH ? c.asParagraph().getHeading() : DocumentApp.ParagraphHeading.NORMAL });
  }
  const enAt = paras.findIndex(p => EN_MARK_.test(p.text));
  const jaAt = paras.findIndex(p => JA_MARK_.test(p.text));
  const jaParas = paras.filter((p, k) => k > jaAt && (enAt < 0 || k < enAt));
  const jaText = jaParas.map(p => p.text).join('\n');
  if (!jaText.trim()) return false;
  const H = DocumentApp.ParagraphHeading;
  const isTitle = p => p.heading === H.TITLE || p.heading === H.SUBTITLE || String(p.heading).indexOf('HEADING') === 0 || /^#\s/.test(p.text);
  let keepTitle = null;
  if (enAt >= 0) {
    const rest = paras.slice(enAt + 1);
    const auto = rest.find(p => AUTO_LINE_.test(p.text));
    const tPara = rest.find(p => p !== auto && p.text && isTitle(p));   // the English title is yours to rewrite
    const enText = rest.filter(p => p !== auto && p !== tPara).map(p => p.text).join('\n');
    if (rest.some(p => p !== auto && p.text)) {
      if (!auto) return false;
      const m = auto.text.match(AUTO_LINE_);
      if (hash_(enText) !== m[2] || hash_(jaText) === m[1]) return false;
      keepTitle = tPara ? tPara.text : null;
    }
    // take away the old English (a document must keep a last paragraph, so one empty paragraph stays and becomes the new heading)
    body.appendParagraph('');
    for (let i = body.getNumChildren() - 2; i >= paras[enAt].i; i--) body.removeChild(body.getChild(i));
  }
  const top = enAt >= 0 ? body.getChild(body.getNumChildren() - 1).asParagraph() : body.appendParagraph('');
  top.setText(mark); top.setHeading(mark === 'English' ? H.HEADING1 : H.NORMAL);
  const HEAD = /^(日付|date|作品|works|リファレンス|refs|下書き|draft)\s*[:：]/i;
  const lines = [];
  let titled = false;
  jaParas.forEach(p => {
    if (!p.text || HEAD.test(p.text)) return;
    const head = isTitle(p);
    let line = trLine_(p.text, tr);
    if (head && !titled) { titled = true; if (keepTitle) line = keepTitle; } else lines.push(line);
    body.appendParagraph(line).setHeading(head && !/^#\s/.test(p.text) ? p.heading : H.NORMAL);
  });
  const out = lines.join('\n');
  const last = body.appendParagraph(`${AUTO_}: ${hash_(jaText)}.${hash_(out)}`);
  last.setHeading(H.NORMAL); last.editAsText().setFontSize(8).setForegroundColor('#999999');
  doc.saveAndClose();
  return true;
}

/* every hour, a little before the site is rebuilt */
function autoTranslateOn() {
  autoTranslateOff(true);
  ScriptApp.newTrigger('translateAll').timeBased().everyHours(1).create();
  SpreadsheetApp.getActive().toast('これから毎時、空いている英語を自動で入れます。', 'サイト', 8);
}
function autoTranslateOff(quiet) {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'translateAll').forEach(t => ScriptApp.deleteTrigger(t));
  if (quiet !== true) SpreadsheetApp.getActive().toast('自動の英訳を止めました。', 'サイト', 8);
}
