// Builds the site into dist/ from a content folder.
// Source: a public Google Drive folder (DRIVE_FOLDER_ID + GOOGLE_API_KEY), or a local folder (CONTENT_DIR, default ./content).
import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import sharp from 'sharp';
import os from 'node:os';
import { execFileSync } from 'node:child_process';
import { readXlsx, readDocx, docxToPlain, blockText } from './office.mjs';

const ROOT = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
const DIST = path.join(ROOT, 'dist');
const CACHE = path.join(ROOT, '.cache', 'img');
const MAX = 1800;            // longest side of published images, px
const QUALITY = 82;
// smaller copies for thumbnails and mid-size sheets: img/s/<key>.jpg and img/m/<key>.jpg
const SIZES = { s: 480, m: 1200 };
const IMAGE_EXT = /\.(jpe?g|png|webp|gif|tiff?|avif|heic|heif|pdf)$/i;
const SKIP_EXT = /\.(txt|md|docx|xlsx|ds_store|ini|db)$/i;

const { DRIVE_FOLDER_ID, GOOGLE_API_KEY } = process.env;
const CONTENT_DIR = path.resolve(ROOT, process.env.CONTENT_DIR || 'content');

/* ---------- sources: both return a tree of {name, kind:'dir'|'file', children?, md5?, read()} ---------- */

async function localTree(dir) {
  const entries = await fs.readdir(dir, { withFileTypes: true });
  const out = [];
  for (const e of entries) {
    if (e.name.startsWith('.')) continue;
    const p = path.join(dir, e.name);
    if (e.isDirectory()) out.push({ name: e.name, kind: 'dir', children: await localTree(p) });
    else out.push({ name: e.name, kind: 'file', read: () => fs.readFile(p) });
  }
  return out;
}

const DRIVE = 'https://www.googleapis.com/drive/v3/files';
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
async function driveJSON(url) {
  for (let attempt = 0; attempt < 4; attempt++) {
    const r = await fetch(url);
    if (r.ok) return r.json();
    if (r.status === 429 || r.status >= 500) { await new Promise(s => setTimeout(s, 1500 * (attempt + 1))); continue; }
    throw new Error(`Drive request failed (${r.status}): ${await r.text()}`);
  }
  throw new Error('Drive request kept failing');
}
async function driveTree(folderId) {
  const out = [];
  let pageToken = '';
  do {
    const q = encodeURIComponent(`'${folderId}' in parents and trashed = false`);
    const url = `${DRIVE}?q=${q}&fields=nextPageToken,files(id,name,mimeType,md5Checksum,shortcutDetails)&pageSize=1000&key=${GOOGLE_API_KEY}` + (pageToken ? `&pageToken=${pageToken}` : '');
    const j = await driveJSON(url);
    for (let f of j.files) {
      // follow shortcuts to the file or folder they point at
      if (f.mimeType === 'application/vnd.google-apps.shortcut' && f.shortcutDetails) {
        f = { ...f, id: f.shortcutDetails.targetId, mimeType: f.shortcutDetails.targetMimeType, md5Checksum: undefined };
      }
      if (f.mimeType === 'application/vnd.google-apps.folder') {
        out.push({ name: f.name, kind: 'dir', id: f.id, children: await driveTree(f.id) });
      } else if (f.mimeType === 'application/vnd.google-apps.document') {
        // a Google Doc is read as .docx, so its links, headings and pasted images come through
        const id = f.id;
        out.push({ name: f.name.replace(/\.(txt|docx?)$/i, '') + '.docx', kind: 'file', id, read: () => driveDownload(`${DRIVE}/${id}/export?mimeType=${encodeURIComponent(DOCX)}&key=${GOOGLE_API_KEY}`) });
      } else if (f.mimeType === 'application/vnd.google-apps.spreadsheet') {
        const id = f.id;
        out.push({ name: f.name.replace(/\.xlsx?$/i, '') + '.xlsx', kind: 'file', id, read: () => driveDownload(`${DRIVE}/${id}/export?mimeType=${encodeURIComponent(XLSX)}&key=${GOOGLE_API_KEY}`) });
      } else if (f.mimeType.startsWith('application/vnd.google-apps.')) {
        out.push({ name: f.name + ' (Google ' + f.mimeType.split('.').pop() + ')', kind: 'file', read: async () => Buffer.alloc(0) });
      } else {
        out.push({ name: f.name, kind: 'file', id: f.id, md5: f.md5Checksum, read: () => driveDownload(`${DRIVE}/${f.id}?alt=media&key=${GOOGLE_API_KEY}`) });
      }
    }
    pageToken = j.nextPageToken || '';
  } while (pageToken);
  return out;
}
async function driveDownload(url) {
  for (let attempt = 0; attempt < 4; attempt++) {
    const r = await fetch(url);
    if (r.ok) return Buffer.from(await r.arrayBuffer());
    if (r.status === 429 || r.status >= 500) { await new Promise(s => setTimeout(s, 1500 * (attempt + 1))); continue; }
    throw new Error(`Drive download failed (${r.status})`);
  }
  throw new Error('Drive download kept failing');
}

/* ---------- helpers ---------- */

const byName = (a, b) => a.name.localeCompare(b.name, 'ja', { numeric: true });
const dir = (tree, name) => tree.find(e => e.kind === 'dir' && e.name.toLowerCase() === name);
const file = (tree, name) => tree.find(e => e.kind === 'file' && e.name.toLowerCase() === name);
// "info" / "about" may be a .txt, a .docx or a Google Doc
const textFile = (tree, base) => file(tree, base + '.docx') || file(tree, base + '.txt');
const norm = s => String(s ?? '').normalize('NFKC').toLowerCase().replace(/\.(jpe?g|png|webp|gif|tiff?|avif|heic|heif|pdf|txt|docx|xlsx)$/i, '').replace(/[\s_\-—–・.,]+/g, ' ').trim();
const unesc = s => String(s ?? '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&');
// links in a Google Doc export may come wrapped as https://www.google.com/url?q=<link>
const unwrap = u => { const m = String(u || '').match(/^https?:\/\/(?:www\.)?google\.[a-z.]+\/url\?(?:.*&)?q=([^&]+)/); return m ? decodeURIComponent(m[1]) : u; };
const driveId = url => (String(url).match(/\/d\/([\w-]{20,})/) || String(url).match(/[?&]id=([\w-]{20,})/) || [])[1] || '';
const esc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const stripNo = s => s.replace(/\.[^.]+$/, '').replace(/^\d+[\s_-]*/, '').trim();
const leadNo = s => (s.match(/^(\d+)/) || [])[1] || '';
// a text file, or a .docx / Google Doc read as plain text
const text = async f => !f ? '' : /\.docx$/i.test(f.name) ? docxToPlain(await f.read()) : ((await f.read()).toString('utf8').replace(/^﻿/, '').replace(/\r\n?/g, '\n'));

// "key: value" lines at the top, then optional "--- ja" / "--- en" sections.
// Inside a section: "key: value" lines, a "# heading", and paragraphs separated by blank lines.
function parseDoc(src, withRows = false) {
  const parts = src.split(/^---\s*(ja|en)\s*$/m);
  const head = kv(parts[0]);
  const lang = {};
  for (let i = 1; i < parts.length; i += 2) lang[parts[i]] = section(parts[i + 1], withRows);
  return { head, lang };
}
function kv(block) {
  const o = {};
  for (const line of block.split('\n')) {
    if (/^\s*#/.test(line)) continue;
    const m = line.match(/^\s*([^:]+?)\s*:\s*(.*)$/);
    if (m) o[m[1].trim().toLowerCase()] = m[2].trim();
  }
  return o;
}
const AUTO_TR = /^自動翻訳\s*[:：]/;
function section(block, withRows = false) {
  const lines = block.split('\n');
  const meta = {}, rows = [];
  let title = '', body = [], para = [];
  const flush = () => { if (para.length) { body.push(para.join(' ').trim()); para = []; } };
  for (const raw of lines) {
    const line = raw.trim();
    if (!line) { flush(); continue; }
    if (AUTO_TR.test(line)) continue;   // the mark left by the automatic translation
    if (line.startsWith('# ') && !title) { flush(); title = line.slice(2).trim(); continue; }
    const m = line.match(/^([^:：]{1,24})\s*[:：]\s*(.*)$/);
    if (m && /^(kind|role)$/i.test(m[1].trim())) { meta[m[1].trim().toLowerCase()] = m[2].trim(); continue; }
    if (withRows && m && !/https?$/i.test(m[1]) && line.length < 140 && body.length) { flush(); rows.push([m[1].trim(), m[2].trim()]); continue; }
    para.push(line);
  }
  flush();
  return { meta, title, body, rows };
}
const list = s => String(s || '').split(/[,、，;\n]/).map(x => x.trim()).filter(Boolean);
// A reference's own words: "note:" (or 一言: / メモ:) for Japanese, "note en:" for English.
// The note may run over several lines until the next known key; blank lines start a new paragraph.
const REF_KEYS = /^\s*(title|kind|tags|works|spotify|apple music|youtube|link|url|note|note ja|note en|一言|メモ)\s*[:：]\s*(.*)$/i;
function refNote(lines) {
  const out = { ja: [], en: [] };
  let cur = null, para = [];
  const flush = () => { if (cur && para.length) { out[cur].push(esc(para.reduce((t, l) => t + (/[\x21-\x7E]$/.test(t) && /^[\x21-\x7E]/.test(l) ? ' ' : '') + l, ''))); } para = []; };
  for (const raw of lines) {
    if (/^\s*#/.test(raw)) continue;
    const m = raw.match(REF_KEYS);
    if (m) {
      flush();
      const k = m[1].toLowerCase();
      cur = k === 'note en' ? 'en' : /^(note|note ja|一言|メモ)$/.test(k) ? 'ja' : null;
      if (cur && m[2].trim()) para.push(m[2].trim());
      continue;
    }
    if (!cur) continue;
    if (!raw.trim()) { flush(); continue; }
    para.push(raw.trim());
  }
  flush();
  return out.ja.length || out.en.length ? out : '';
}

// A .docx / Google Doc text: "key: value" lines at the top (日付 / date, 種類 / kind, 作品 / works, リファレンス / refs, 下書き / draft),
// the first heading (or Title) is the title, "English" / "日本語" (or "--- en" / "--- ja") on a line of its own starts the other language.
const HEAD_KEYS = { date: 'date', 日付: 'date', works: 'works', 作品: 'works', refs: 'refs', リファレンス: 'refs', draft: 'draft', 下書き: 'draft' };
const SEC_KEYS = { kind: 'kind', 種類: 'kind', title: 'title', 題名: 'title' };
function parseBlocks(blocks) {
  const head = {}, lang = {};
  let cur = null;
  const sec = l => (lang[l] ??= { meta: {}, title: '', body: [] });
  for (const b of blocks) {
    const t = blockText(b).trim();
    const hasImg = b.segs.some(s => s.img);
    if (!t && !hasImg) continue;
    if (AUTO_TR.test(t)) continue;
    const mk = t.match(/^(?:-{2,}\s*)?(ja|en|日本語|english|英語)$/i);
    if (mk && (t.startsWith('-') || !/^(ja|en)$/i.test(t))) { cur = sec(/^(ja|日本語)$/i.test(mk[1]) ? 'ja' : 'en'); continue; }
    const s = cur || sec('ja');
    const kvm = !hasImg && !s.body.length && t.match(/^([^:：]{1,12}?)\s*[:：]\s*(.*)$/);
    if (kvm) {
      const k0 = kvm[1].trim().toLowerCase();
      if (HEAD_KEYS[k0]) { head[HEAD_KEYS[k0]] = kvm[2].trim(); continue; }
      if (SEC_KEYS[k0]) { if (SEC_KEYS[k0] === 'title') s.title = kvm[2].trim(); else s.meta[SEC_KEYS[k0]] = kvm[2].trim(); continue; }
    }
    const isHead = b.style !== 'p' || /^#\s/.test(t);
    if (isHead && !hasImg && !s.title && !s.body.length) { s.title = t.replace(/^#\s*/, ''); continue; }
    if (isHead && !hasImg) { s.body.push([{ text: t.replace(/^#+\s*/, ''), b: true }]); continue; }
    s.body.push(b.segs);
  }
  return { head, lang };
}

/* ---------- images ---------- */

const images = {};
// HEIC photos and PDF boards are turned into a raster first (needs libheif-examples / poppler-utils on the runner).
async function toRaster(name, buf) {
  const ext = (name.match(/\.([^.]+)$/) || [])[1]?.toLowerCase();
  if (ext !== 'pdf' && ext !== 'heic' && ext !== 'heif') return buf;
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'raster-'));
  const inp = path.join(dir, 'in.' + ext);
  await fs.writeFile(inp, buf);
  try {
    if (ext === 'pdf') { execFileSync('pdftoppm', ['-f', '1', '-l', '1', '-r', '150', '-jpeg', '-singlefile', inp, path.join(dir, 'out')]); return await fs.readFile(path.join(dir, 'out.jpg')); }
    execFileSync('heif-convert', ['-q', '92', inp, path.join(dir, 'out.jpg')]); return await fs.readFile(path.join(dir, 'out.jpg'));
  } finally { await fs.rm(dir, { recursive: true, force: true }); }
}
// every image in a folder, including ones tucked into subfolders, in name order
function imagesIn(children, prefix = '') {
  const out = [];
  for (const e of [...children].sort(byName)) {
    if (e.kind === 'dir') out.push(...imagesIn(e.children, prefix + e.name + '/'));
    else if (IMAGE_EXT.test(e.name)) out.push(e);
    else if (!SKIP_EXT.test(e.name)) skipped.push(prefix + e.name);
  }
  return out;
}
const skipped = [];
async function addImage(f) {
  const buf = f.md5 ? null : await f.read();
  const key = f.md5 || crypto.createHash('md5').update(buf).digest('hex');
  if (images[key]) return key;
  const cached = path.join(CACHE, key + '.jpg');
  let out, meta;
  try { out = await fs.readFile(cached); }
  catch {
    const src = await toRaster(f.name, buf || await f.read());
    out = await sharp(src, { failOn: 'none' }).rotate().resize({ width: MAX, height: MAX, fit: 'inside', withoutEnlargement: true })
      .flatten({ background: '#ffffff' }).jpeg({ quality: QUALITY, mozjpeg: true }).toBuffer();
    await fs.writeFile(cached, out);
  }
  meta = await sharp(out).metadata();
  await fs.writeFile(path.join(DIST, 'img', key + '.jpg'), out);
  for (const [sz, px] of Object.entries(SIZES)) {
    const cachedSz = path.join(CACHE, `${key}.${sz}.jpg`);
    let small;
    try { small = await fs.readFile(cachedSz); }
    catch {
      small = await sharp(out).resize({ width: px, height: px, fit: 'inside', withoutEnlargement: true }).jpeg({ quality: QUALITY, mozjpeg: true }).toBuffer();
      await fs.writeFile(cachedSz, small);
    }
    await fs.writeFile(path.join(DIST, 'img', sz, key + '.jpg'), small);
  }
  images[key] = { url: `img/${key}.jpg`, ar: +(meta.width / meta.height).toFixed(4) };
  return key;
}

/* ---------- build ---------- */

async function main() {
  const fromDrive = DRIVE_FOLDER_ID && GOOGLE_API_KEY;
  console.log(fromDrive ? `Reading Google Drive folder ${DRIVE_FOLDER_ID}` : `Reading local folder ${CONTENT_DIR}`);
  const tree = fromDrive ? await driveTree(DRIVE_FOLDER_ID) : await localTree(CONTENT_DIR);

  await fs.rm(DIST, { recursive: true, force: true });
  await fs.mkdir(path.join(DIST, 'img'), { recursive: true });
  for (const sz of Object.keys(SIZES)) await fs.mkdir(path.join(DIST, 'img', sz), { recursive: true });
  await fs.mkdir(path.join(DIST, 'data'), { recursive: true });
  await fs.mkdir(CACHE, { recursive: true });
  await fs.cp(path.join(ROOT, 'site'), DIST, { recursive: true });

  const warn = [];

  // works: newest number first
  const worksDir = dir(tree, 'works');
  const workFolders = (worksDir ? worksDir.children.filter(e => e.kind === 'dir') : []).sort(byName).reverse();
  const works = [];
  for (const wf of workFolders) {
    const no = leadNo(wf.name) || String(works.length + 1).padStart(3, '0');
    const doc = parseDoc(await text(textFile(wf.children, 'info')));
    const before = skipped.length;
    const pics = imagesIn(wf.children);
    skipped.splice(before).forEach(n => warn.push(`works/${wf.name}/${n}: not a supported image`));
    if (!pics.length) { warn.push(`works/${wf.name}: no images, skipped`); continue; }
    const sheets = [];
    for (const [k, pf] of pics.entries()) {
      try {
        const key = await addImage(pf);
        sheets.push({ img: key, no: String(k + 1).padStart(2, '0'), t: esc(stripNo(pf.name)), s: '', ar: images[key].ar, _file: pf.name });
      } catch (e) { warn.push(`works/${wf.name}/${pf.name}: could not read (${e.message})`); }
    }
    if (!sheets.length) continue;
    const coverPref = doc.head.cover;
    let cover = coverPref ? sheets.findIndex(s => s._file.startsWith(coverPref) || s.no === coverPref.padStart(2, '0')) : 0;
    if (cover < 0) cover = 0;
    const tx = {};
    for (const l of ['ja', 'en']) {
      const sct = doc.lang[l];
      if (sct && (sct.title || sct.body.length)) tx[l] = { title: esc(sct.title), body: sct.body.map(esc) };
    }
    if (works.some(w => +w.no === +no)) warn.push(`works/${wf.name}: number ${no} is used by more than one work folder`);
    works.push({
      no, title: esc(doc.head.title || stripNo(wf.name)), period: esc(doc.head.year || ''), type: esc(doc.head.type || ''), tag: esc(doc.head.tag || ''),
      cover, sheets: sheets.map(({ _file, ...s }) => s), ...(Object.keys(tx).length ? { text: tx } : {}),
      _folder: wf.name, _id: wf.id, _title: doc.head.title || stripNo(wf.name)
    });
  }
  // a work named in a sheet or a text: "006", "006_Elevated Deck", "006 Elevated Deck" or just "Elevated Deck"
  const findWork = v => {
    v = String(v ?? '').normalize('NFKC').trim();
    if (!v) return -1;
    let i = works.findIndex(w => norm(w._folder) === norm(v));
    if (i >= 0) return i;
    const n = leadNo(v), rest = norm(v.replace(/^\d+[\s_.\-]*/, ''));
    const all = works.map((w, k) => k);
    const cands = n ? all.filter(k => +works[k].no === +n) : all;
    if (rest) {
      const t = cands.find(k => norm(works[k]._title) === rest || norm(stripNo(works[k]._folder)) === rest);
      if (t != null) return t;
      if (!n) return -1;
    }
    return cands.length ? cands[0] : -1;
  };
  const workIndex = findWork;
  // work names may contain commas ("006_Town, Kindergarten and Workplace"): split only before a number, or where a whole piece is a work
  const workList = v => {
    const parts = list(v);
    const out = [];
    for (const p of parts) {
      const prev = out[out.length - 1];
      if (prev != null && !/^\d/.test(p) && (findWork(p) < 0 || findWork(prev + ', ' + p) >= 0)) out[out.length - 1] = prev + ', ' + p;
      else out.push(p);
    }
    return out;
  };

  // references: a spreadsheet (Google Sheets or .xlsx) in the top folder or in references/, plus the old references.txt
  const refDir = dir(tree, 'references');
  const isSheet = e => e.kind === 'file' && /\.xlsx$/i.test(e.name) && !e.name.startsWith('~$');
  const sheetFiles = [...tree.filter(isSheet), ...(refDir ? refDir.children.filter(isSheet) : [])];
  const sheetFile = sheetFiles.find(e => /リファレンス|reference/i.test(e.name)) || sheetFiles[0];
  const seen = new Set(['works', 'references', 'texts', 'about.txt', 'about.docx', 'readme.txt', 'readme.docx']);
  if (sheetFile) seen.add(sheetFile.name.toLowerCase());
  tree.filter(e => !seen.has(e.name.toLowerCase())).forEach(e => warn.push(`${e.name}: outside works/references/texts, ignored`));
  if (worksDir) worksDir.children.filter(e => e.kind === 'file' && !SKIP_EXT.test(e.name)).forEach(e => warn.push(`works/${e.name}: put files inside a project folder`));

  const refs = [];
  const refMeta = [];   // per reference: { id, file, words: [] } (not published)
  const linkLabel = u => /spotify\.com/.test(u) ? 'Spotify' : /music\.apple\.com/.test(u) ? 'Apple Music' : /youtu\.?be/.test(u) ? 'YouTube' : /soundcloud\.com/.test(u) ? 'SoundCloud' : /bandcamp\.com/.test(u) ? 'Bandcamp' : 'Link';
  const pushRef = (key, o) => {
    const tags = list(o.tags).map(esc);
    const rel = [];
    for (const n of workList(o.works)) { const i = workIndex(n); if (i >= 0) { if (!rel.includes(i)) rel.push(i); } else warn.push(`references: "${o.title || o.file}" — work "${n}" not found`); }
    refs.push({
      img: key, no: 'REF-' + String(refs.length + 1).padStart(2, '0'), title: esc(o.title), kind: esc((o.kind || 'IMAGE').toUpperCase()),
      rel, note: o.note || '', tags: tags.length ? tags : ['Other'],
      links: (o.links || []).filter(l => /^https?:\/\//.test(l.url)).map(l => ({ label: esc(l.label), url: esc(l.url) })), ar: images[key].ar
    });
    refMeta.push({ id: o.id || '', file: o.file || '', words: list(o.words), text: !!o.fromText });
  };
  if (refDir || sheetFile) {
    // old format: references.txt with one [file name] block per image
    const blocks = {};
    const rt = refDir ? await text(file(refDir.children, 'references.txt')) : '';
    let cur = null;
    for (const line of rt.split('\n')) {
      const h = line.match(/^\s*\[(.+)\]\s*$/);
      if (h) { cur = blocks[norm(h[1])] = []; continue; }
      if (cur) cur.push(line);
    }
    // new format: one row per reference
    let rows = [];
    if (sheetFile) {
      try {
        const book = readXlsx(await sheetFile.read());
        const H = { img: /^(画像|ファイル|ファイル名|image|file|filename)$/, title: /^(題名|タイトル|名前|title|name)$/, kind: /^(種類|kind|type)$/, tags: /^(タグ|tags?)$/,
          works: /^(作品|関係する作品|works?)$/, words: /^(言葉|本文で結ぶ言葉|結ぶ言葉|キーワード|words?|keywords?)$/, note: /^(一言|メモ|説明|note|noteja)$/,
          noteEn: /^(一言英|一言英語|英語の一言|noteen|english)$/, links: /^(リンク|links?|urls?)$/, id: /^(id|ファイルid)$/ };
        const hn = s => String(s).normalize('NFKC').toLowerCase().replace(/[\s()()・_\-:]/g, '');
        for (const [sname, grid] of Object.entries(book)) {
          const hi = grid.findIndex(r => r && r.some(c => H.img.test(hn(c))));
          if (hi < 0) continue;
          const cols = {};
          grid[hi].forEach((c, k) => { for (const [f, re] of Object.entries(H)) if (cols[f] == null && re.test(hn(c))) cols[f] = k; });
          rows = grid.slice(hi + 1).map((r, k) => ({ r, line: hi + k + 2 })).filter(({ r }) => r && (r[cols.img] || r[cols.title]))
            .map(({ r, line }) => Object.fromEntries([['line', line], ...Object.entries(cols).map(([f, k]) => [f, String(r[k] ?? '').trim()])]));
          console.log(`Reference sheet "${sheetFile.name}" / ${sname}: ${rows.length} rows`);
          break;
        }
        if (!rows.length) warn.push(`${sheetFile.name}: no header row with a 画像 column found`);
      } catch (e) { warn.push(`${sheetFile.name}: could not read (${e.message})`); }
    }
    const before = skipped.length;
    const pics = refDir ? imagesIn(refDir.children) : [];
    skipped.splice(before).forEach(n => warn.push(`references/${n}: not a supported image`));
    const used = new Set();
    const findPic = (name, id) => {
      let k = id ? pics.findIndex(p => p.id === id) : -1;
      if (k < 0 && name) k = pics.findIndex(p => norm(p.name) === norm(name));
      if (k < 0 && name) k = pics.findIndex(p => norm(stripNo(p.name)) === norm(name) || norm(stripNo(p.name)) === norm(stripNo(name + '.x')));
      return k;
    };
    const noteOf = (ja, en) => {
      const paras = s => String(s || '').split(/\n\s*\n/).map(p => p.split('\n').map(x => x.trim()).filter(Boolean)
        .reduce((t, l) => t + (/[\x21-\x7E]$/.test(t) && /^[\x21-\x7E]/.test(l) ? ' ' : '') + l, '')).filter(Boolean).map(esc);
      const o = { ja: paras(ja), en: paras(en) };
      return o.ja.length || o.en.length ? o : '';
    };
    for (const row of rows) {
      const k = findPic(row.img, row.id || driveId(row.img));
      if (k < 0) { warn.push(`${sheetFile.name} row ${row.line}: image "${row.img}" not found in references/`); continue; }
      if (used.has(k)) { warn.push(`${sheetFile.name} row ${row.line}: "${row.img}" is listed twice`); continue; }
      used.add(k);
      const pf = pics[k];
      try {
        const key = await addImage(pf);
        const links = String(row.links || '').split(/[\s,、]+/).filter(u => /^https?:\/\//.test(u)).map(u => ({ label: linkLabel(u), url: u }));
        pushRef(key, { title: row.title || stripNo(pf.name), kind: row.kind, tags: row.tags, works: row.works, words: row.words, note: noteOf(row.note, row.noteEn), links, id: pf.id, file: pf.name });
      } catch (e) { warn.push(`references/${pf.name}: could not read (${e.message})`); }
    }
    // images without a row: described by references.txt if they have a block there, otherwise shown as they are
    for (const [k, pf] of pics.entries()) {
      if (used.has(k)) continue;
      try {
        const key = await addImage(pf);
        const blk = blocks[norm(pf.name)] || [];
        const o = kv(blk.join('\n'));
        const links = [];
        for (const [label, name] of [['spotify', 'Spotify'], ['apple music', 'Apple Music'], ['youtube', 'YouTube'], ['link', 'Link'], ['url', 'Link']]) {
          if (o[label]) links.push({ label: name, url: o[label] });
        }
        pushRef(key, { title: o.title || stripNo(pf.name), kind: o.kind, tags: o.tags, works: o.works, words: o.words || o['言葉'], note: refNote(blk), links, id: pf.id, file: pf.name });
      } catch (e) { warn.push(`references/${pf.name}: could not read (${e.message})`); }
    }
  }
  // a reference named in a text: "04", "REF-04", its title or its image file name
  const findRef = v => {
    v = String(v ?? '').normalize('NFKC').trim();
    if (!v) return -1;
    const m = v.match(/^(?:ref-?)?(\d{1,3})$/i);
    if (m) return refs.findIndex(r => +r.no.slice(4) === +m[1]);
    const n = norm(v);
    return refs.findIndex((r, k) => norm(unesc(r.title)) === n || norm(refMeta[k].file) === n || norm(stripNo(refMeta[k].file)) === n);
  };
  const refByLink = href => {
    if (!href) return -1;
    const id = driveId(href);
    if (id) { const k = refMeta.findIndex(m => m.id === id); if (k >= 0) return k; }
    let h = href;
    try { h = decodeURIComponent(href); } catch {}
    const m = h.match(/^(?:#|ref:)\s*(.+)$/i);
    if (m) return findRef(m[1]);
    const base = h.split(/[?#]/)[0].split('/').pop();
    if (base && IMAGE_EXT.test(base)) return findRef(base);
    return -1;
  };
  // pictures pasted into a Doc: the same picture as a reference (by a small fingerprint) or a new one
  const prints = {};
  const fingerprint = async key => prints[key] ??= await sharp(path.join(DIST, 'img', 's', key + '.jpg')).resize(9, 8, { fit: 'fill' }).grayscale().raw().toBuffer()
    .then(px => { let bits = ''; for (let y = 0; y < 8; y++) for (let x = 0; x < 8; x++) bits += px[y * 9 + x] < px[y * 9 + x + 1] ? '1' : '0'; return bits; });
  const near = (a, b) => [...a].reduce((d, c, k) => d + (c !== b[k]), 0) <= 8;

  // texts: newest number first. A .docx or Google Doc wins over a .txt of the same name.
  const textDir = dir(tree, 'texts');
  const texts = [];
  const textFiles = (textDir ? textDir.children.filter(e => e.kind === 'file' && /\.(txt|md|docx)$/i.test(e.name)) : []).sort(byName).reverse();
  const tfiles = textFiles.filter(f => {
    if (/\.docx$/i.test(f.name)) return true;
    const twin = textFiles.find(g => g !== f && /\.docx$/i.test(g.name) && norm(g.name) === norm(f.name));
    if (twin) warn.push(`texts/${f.name}: ignored, "${twin.name.replace(/\.docx$/, '')}" (document) is used instead`);
    return !twin;
  });
  for (const tf of tfiles) {
    const doc = /\.docx$/i.test(tf.name) ? parseBlocks(readDocx(await tf.read())) : parseDoc(await text(tf));
    if (!doc.lang.ja && !doc.lang.en) { const s = section(await text(tf)); doc.lang.ja = { ...s, body: s.body.map(b => [{ text: b }]) }; }
    for (const s of Object.values(doc.lang)) s.body = s.body.map(b => typeof b === 'string' ? [{ text: b }] : b);
    const entry = { no: 'T-' + (leadNo(tf.name) || String(texts.length + 1).padStart(2, '0')), date: esc(doc.head.date || ''), draft: /^(yes|true|1)$/i.test(doc.head.draft || ''),
      works: workList(doc.head.works).map(workIndex).filter(i => i >= 0), refs: [] };
    const order = [];   // references in the order they appear in the text
    const use = j => { if (!order.includes(j)) order.push(j); return j; };
    const tname = (doc.lang.ja?.title || doc.lang.en?.title || stripNo(tf.name));
    let figNo = 0;
    const renderPara = async segs => {
      let html = '';
      for (const s of segs) {
        if (s.img) {
          try {
            const buf = s.img.data;
            const key = await addImage({ name: 'pasted.' + s.img.ext, read: async () => buf });
            const fp = await fingerprint(key);
            let j = refs.findIndex(r => r.img === key);
            if (j < 0) for (const [k, r] of refs.entries()) if (!refMeta[k].text && near(fp, await fingerprint(r.img))) { j = k; break; }
            if (j < 0) { pushRef(key, { title: s.img.alt || `${tname} — ${++figNo}`, kind: 'FIGURE', tags: 'Text', works: doc.head.works, fromText: true }); j = refs.length - 1; }
            html += `<span class="rlink rpin" data-r="${use(j)}"></span>`;
          } catch (e) { warn.push(`texts/${tf.name}: a pasted picture could not be read (${e.message})`); }
          continue;
        }
        let t = esc(s.text).replace(/\n/g, '<br>');
        if (s.b) t = `<b>${t}</b>`;
        if (s.i) t = `<i>${t}</i>`;
        const href = s.href && unwrap(s.href);
        const j = refByLink(href);
        if (j >= 0) html += `<span class="rlink" data-r="${use(j)}">${t}</span>`;
        else if (s.href && /^https?:\/\//.test(s.href) && !driveId(s.href)) html += `<a href="${esc(s.href)}" target="_blank" rel="noopener">${t}</a>`;
        else {
          if (s.href) warn.push(`texts/${tf.name}: the link on "${s.text.trim()}" does not point to a reference`);
          // old style: [後景](04) or ［後景］（04）
          html += t.replace(/[\[［]([^\]］\n]+)[\]］][(（]\s*([^)）\n]{1,60})\s*[)）]/gi, (m, w, n) => {
            const k = findRef(unesc(n));
            return k < 0 ? m : `<span class="rlink" data-r="${use(k)}">${w}</span>`;
          });
        }
      }
      return html;
    };
    for (const l of ['ja', 'en']) {
      const s = doc.lang[l];
      if (!s) continue;
      const body = [];
      for (const b of s.body) { const h = await renderPara(b); if (h.trim()) body.push(h); }
      entry[l] = { kind: esc(s.meta.kind || ''), title: esc(s.title || stripNo(tf.name)), body };
    }
    // words listed for a reference in the sheet ("言葉") link themselves: the first place each one appears, in each language
    const linkWord = (l, k, w) => {
      for (let p = 0; p < entry[l].body.length; p++) {
        const parts = entry[l].body[p].split(/(<[^>]+>)/);
        let depth = 0;
        for (let q = 0; q < parts.length; q++) {
          const part = parts[q];
          if (part.startsWith('<')) { if (/^<(span|a)\b/.test(part)) depth++; else if (/^<\/(span|a)>/.test(part)) depth--; continue; }
          const at = depth ? -1 : part.indexOf(esc(w));
          if (at < 0) continue;
          parts[q] = part.slice(0, at) + `<span class="rlink" data-r="${use(k)}">${esc(w)}</span>` + part.slice(at + esc(w).length);
          entry[l].body[p] = parts.join('');
          return true;
        }
      }
      return false;
    };
    for (const l of ['ja', 'en']) {
      if (!entry[l]) continue;
      for (const [k, m] of refMeta.entries()) {
        if (entry[l].body.some(b => b.includes(`data-r="${k}"`))) continue;
        for (const w of m.words) if (linkWord(l, k, w)) break;
      }
    }
    for (const n of list(doc.head.refs)) { const j = findRef(n); if (j >= 0) use(j); else warn.push(`texts/${tf.name}: reference "${n}" not found`); }
    // margin pictures follow the order the references appear in the text; ones only named in "refs:" come last
    const seenOrder = [];
    for (const l of ['ja', 'en']) for (const b of entry[l]?.body || []) for (const m of b.matchAll(/data-r="(\d+)"/g)) if (!seenOrder.includes(+m[1])) seenOrder.push(+m[1]);
    entry.refs = [...seenOrder, ...order.filter(j => !seenOrder.includes(j))];
    texts.push(entry);
  }

  // about
  const ab = parseDoc(await text(textFile(tree, 'about')), true);
  const about = {};
  for (const l of ['ja', 'en']) {
    const s = ab.lang[l];
    if (!s) continue;
    about[l] = { role: list(s.meta.role).length ? (s.meta.role || '').split('/').map(x => esc(x.trim())) : [], lead: esc(s.body[0] || ''), body: s.body.slice(1).map(esc).join('</p><p>'),
      rows: s.rows.filter(r => r[1]).map(([k, v]) => [esc(k), /^https?:\/\//.test(v) ? `<a href="${esc(v)}" target="_blank" rel="noopener">${esc(v.replace(/^https?:\/\//, ''))}</a>` : esc(v)]) };
  }
  if (!about.ja && !about.en) about.ja = { role: [], lead: '', body: '', rows: [] };

  const name = ab.head.name || 'Aga Shunsei';
  const duck = {};
  for (const p of ['stand', 'read', 'carry', 'wave', 'sleep']) for (const t of ['day', 'night']) duck[`${p}_${t}`] = `assets/duck/${p}_${t}.png`;
  const imageUrls = Object.fromEntries(Object.entries(images).map(([k, v]) => [k, v.url]));

  if (!works.length) {
    // Never replace a live site with an empty one: if the folder looks empty or unreadable, stop and keep the last published version.
    if (fromDrive) throw new Error('No works found in the Drive folder (is works/ there, and is the folder shared as "anyone with the link"?). The site was not changed.');
    warn.push('no works found: add folders to works/');
  }
  const site = { site: { name: esc(name), built: new Date().toISOString() }, images: imageUrls, sizes: Object.keys(SIZES), duck, works: works.map(({ _folder, _id, _title, ...w }) => w), refs, texts, about };
  await fs.writeFile(path.join(DIST, 'data', 'site.json'), JSON.stringify(site));
  await fs.writeFile(path.join(DIST, '.nojekyll'), '');
  console.log(`Built ${works.length} works, ${refs.length} references, ${texts.length} texts, ${Object.keys(images).length} images.`);
  if (process.env.GITHUB_ACTIONS) console.log(`::notice title=Build::${fromDrive ? 'Google Drive' : 'content/ (Drive not connected)'}: ${works.length} works, ${refs.length} references, ${texts.length} texts, ${Object.keys(images).length} images${warn.length ? ' | ' + warn.slice(0, 40).join(' / ') : ''}`);
  if (process.env.GITHUB_ACTIONS) console.log(`::notice title=Works found::${works.map(w => w.no + ' ' + w.title + ' (' + w.sheets.length + ')').join(' / ')}`);
  if (warn.length) console.log('Notes:\n  ' + warn.join('\n  '));
}

main().catch(e => { console.error(e); process.exit(1); });
