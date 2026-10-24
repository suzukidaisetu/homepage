// Builds the site into dist/ from a content folder.
// Source: a public Google Drive folder (DRIVE_FOLDER_ID + GOOGLE_API_KEY), or a local folder (CONTENT_DIR, default ./content).
import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import sharp from 'sharp';

const ROOT = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
const DIST = path.join(ROOT, 'dist');
const CACHE = path.join(ROOT, '.cache', 'img');
const MAX = 1800;            // longest side of published images, px
const QUALITY = 82;
const IMAGE_EXT = /\.(jpe?g|png|webp|gif|tiff?|avif)$/i;

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
    const url = `${DRIVE}?q=${q}&fields=nextPageToken,files(id,name,mimeType,md5Checksum)&pageSize=1000&key=${GOOGLE_API_KEY}` + (pageToken ? `&pageToken=${pageToken}` : '');
    const j = await driveJSON(url);
    for (const f of j.files) {
      if (f.mimeType === 'application/vnd.google-apps.folder') {
        out.push({ name: f.name, kind: 'dir', children: await driveTree(f.id) });
      } else if (f.mimeType === 'application/vnd.google-apps.document') {
        // a Google Doc can stand in for a .txt file
        out.push({ name: f.name.replace(/(\.txt)?$/, '.txt'), kind: 'file', read: () => driveDownload(`${DRIVE}/${f.id}/export?mimeType=text/plain&key=${GOOGLE_API_KEY}`) });
      } else if (!f.mimeType.startsWith('application/vnd.google-apps.')) {
        out.push({ name: f.name, kind: 'file', md5: f.md5Checksum, read: () => driveDownload(`${DRIVE}/${f.id}?alt=media&key=${GOOGLE_API_KEY}`) });
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
const esc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const stripNo = s => s.replace(/\.[^.]+$/, '').replace(/^\d+[\s_-]*/, '').trim();
const leadNo = s => (s.match(/^(\d+)/) || [])[1] || '';
const text = async f => (f ? (await f.read()).toString('utf8').replace(/^﻿/, '').replace(/\r\n?/g, '\n') : '');

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
function section(block, withRows = false) {
  const lines = block.split('\n');
  const meta = {}, rows = [];
  let title = '', body = [], para = [];
  const flush = () => { if (para.length) { body.push(para.join(' ').trim()); para = []; } };
  for (const raw of lines) {
    const line = raw.trim();
    if (!line) { flush(); continue; }
    if (line.startsWith('# ') && !title) { flush(); title = line.slice(2).trim(); continue; }
    const m = line.match(/^([^:：]{1,24})\s*[:：]\s*(.*)$/);
    if (m && /^(kind|role)$/i.test(m[1].trim())) { meta[m[1].trim().toLowerCase()] = m[2].trim(); continue; }
    if (withRows && m && !/https?$/i.test(m[1]) && line.length < 140 && body.length) { flush(); rows.push([m[1].trim(), m[2].trim()]); continue; }
    para.push(line);
  }
  flush();
  return { meta, title, body, rows };
}
const list = s => (s || '').split(/[,、]/).map(x => x.trim()).filter(Boolean);

/* ---------- images ---------- */

const images = {};
async function addImage(f) {
  const buf = f.md5 ? null : await f.read();
  const key = f.md5 || crypto.createHash('md5').update(buf).digest('hex');
  if (images[key]) return key;
  const cached = path.join(CACHE, key + '.jpg');
  let out, meta;
  try { out = await fs.readFile(cached); }
  catch {
    const src = buf || await f.read();
    out = await sharp(src, { failOn: 'none' }).rotate().resize({ width: MAX, height: MAX, fit: 'inside', withoutEnlargement: true })
      .flatten({ background: '#ffffff' }).jpeg({ quality: QUALITY, mozjpeg: true }).toBuffer();
    await fs.writeFile(cached, out);
  }
  meta = await sharp(out).metadata();
  await fs.writeFile(path.join(DIST, 'img', key + '.jpg'), out);
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
    const doc = parseDoc(await text(file(wf.children, 'info.txt')));
    const pics = wf.children.filter(e => e.kind === 'file' && IMAGE_EXT.test(e.name)).sort(byName);
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
    works.push({
      no, title: esc(doc.head.title || stripNo(wf.name)), period: esc(doc.head.year || ''), type: esc(doc.head.type || ''), tag: esc(doc.head.tag || ''),
      cover, sheets: sheets.map(({ _file, ...s }) => s), ...(Object.keys(tx).length ? { text: tx } : {})
    });
  }
  const workIndex = no => works.findIndex(w => w.no === no || +w.no === +no);

  // references
  const refDir = dir(tree, 'references');
  const refs = [];
  if (refDir) {
    const blocks = {};
    const rt = await text(file(refDir.children, 'references.txt'));
    let cur = null;
    for (const line of rt.split('\n')) {
      const h = line.match(/^\s*\[(.+)\]\s*$/);
      if (h) { cur = blocks[h[1].trim()] = []; continue; }
      if (cur) cur.push(line);
    }
    const pics = refDir.children.filter(e => e.kind === 'file' && IMAGE_EXT.test(e.name)).sort(byName);
    for (const [k, pf] of pics.entries()) {
      try {
        const key = await addImage(pf);
        const o = kv((blocks[pf.name] || blocks[pf.name.replace(/\.[^.]+$/, '')] || []).join('\n'));
        const links = [];
        for (const [label, name] of [['spotify', 'Spotify'], ['apple music', 'Apple Music'], ['youtube', 'YouTube'], ['link', 'Link'], ['url', 'Link']]) {
          if (o[label]) links.push({ label: name, url: o[label] });
        }
        const rel = list(o.works).map(workIndex).filter(i => i >= 0);
        list(o.works).filter(n => workIndex(n) < 0).forEach(n => warn.push(`references/${pf.name}: work ${n} not found`));
        refs.push({
          img: key, no: 'REF-' + String(k + 1).padStart(2, '0'), title: esc(o.title || stripNo(pf.name)), kind: esc((o.kind || 'IMAGE').toUpperCase()),
          rel, note: '', tags: list(o.tags).map(esc).length ? list(o.tags).map(esc) : ['Other'], links: links.filter(l => /^https?:\/\//.test(l.url)).map(l => ({ label: l.label, url: esc(l.url) })), ar: images[key].ar
        });
      } catch (e) { warn.push(`references/${pf.name}: could not read (${e.message})`); }
    }
  }
  const refIndex = n => refs.findIndex(r => +r.no.slice(4) === +n);

  // texts: newest number first
  const textDir = dir(tree, 'texts');
  const texts = [];
  for (const tf of (textDir ? textDir.children.filter(e => e.kind === 'file' && /\.(txt|md)$/i.test(e.name)) : []).sort(byName).reverse()) {
    const doc = parseDoc(await text(tf));
    const entry = { no: 'T-' + (leadNo(tf.name) || String(texts.length + 1).padStart(2, '0')), date: esc(doc.head.date || ''), draft: /^(yes|true|1)$/i.test(doc.head.draft || ''),
      works: list(doc.head.works).map(workIndex).filter(i => i >= 0), refs: list(doc.head.refs).map(refIndex).filter(i => i >= 0) };
    for (const l of ['ja', 'en']) {
      const s = doc.lang[l];
      if (s) entry[l] = { kind: esc(s.meta.kind || ''), title: esc(s.title || stripNo(tf.name)), body: s.body.map(esc) };
    }
    if (!entry.ja && !entry.en) { const s = section((await text(tf))); entry.ja = { kind: '', title: esc(s.title || stripNo(tf.name)), body: s.body.map(esc) }; }
    texts.push(entry);
  }

  // about
  const ab = parseDoc(await text(file(tree, 'about.txt')), true);
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
  const site = { site: { name: esc(name), built: new Date().toISOString() }, images: imageUrls, duck, works, refs, texts, about };
  await fs.writeFile(path.join(DIST, 'data', 'site.json'), JSON.stringify(site));
  await fs.writeFile(path.join(DIST, '.nojekyll'), '');
  console.log(`Built ${works.length} works, ${refs.length} references, ${texts.length} texts, ${Object.keys(images).length} images.`);
  if (process.env.GITHUB_ACTIONS) console.log(`::notice title=Build::${fromDrive ? 'Google Drive' : 'content/ (Drive not connected)'}: ${works.length} works, ${refs.length} references, ${texts.length} texts, ${Object.keys(images).length} images${warn.length ? ' | ' + warn.slice(0, 5).join(' / ') : ''}`);
  if (warn.length) console.log('Notes:\n  ' + warn.join('\n  '));
}

main().catch(e => { console.error(e); process.exit(1); });
