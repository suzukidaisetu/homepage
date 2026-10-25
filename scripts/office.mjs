// Minimal readers for .xlsx (Google Sheets export) and .docx (Google Docs export), with no dependencies.
import zlib from 'node:zlib';

export function unzip(buf) {
  let e = buf.length - 22;
  while (e >= 0 && buf.readUInt32LE(e) !== 0x06054b50) e--;
  if (e < 0) throw new Error('not a zip file');
  const n = buf.readUInt16LE(e + 10);
  let p = buf.readUInt32LE(e + 16);
  const out = {};
  for (let i = 0; i < n; i++) {
    const method = buf.readUInt16LE(p + 10), csize = buf.readUInt32LE(p + 20);
    const nl = buf.readUInt16LE(p + 28), xl = buf.readUInt16LE(p + 30), cl = buf.readUInt16LE(p + 32), off = buf.readUInt32LE(p + 42);
    const name = buf.toString('utf8', p + 46, p + 46 + nl);
    p += 46 + nl + xl + cl;
    const start = off + 30 + buf.readUInt16LE(off + 26) + buf.readUInt16LE(off + 28);
    const data = buf.subarray(start, start + csize);
    out[name] = method === 0 ? data : zlib.inflateRawSync(data);
  }
  return out;
}

export const unxml = s => String(s).replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (m, c) =>
  c[0] === '#' ? String.fromCodePoint(c[1].toLowerCase() === 'x' ? parseInt(c.slice(2), 16) : +c.slice(1)) : { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[c.toLowerCase()]);
const attr = (tag, name) => { const m = tag.match(new RegExp('\\s' + name + '="([^"]*)"')); return m ? unxml(m[1]) : ''; };
function relsOf(z, p) {
  const x = z[p]?.toString('utf8') || '';
  const o = {};
  for (const m of x.matchAll(/<Relationship\b[^>]*>/g)) o[attr(m[0], 'Id')] = { target: attr(m[0], 'Target'), external: attr(m[0], 'TargetMode') === 'External' };
  return o;
}
const resolve = (base, t) => t.startsWith('/') ? t.slice(1) : (base + t).split('/').reduce((a, s) => (s === '..' ? a.pop() : s && s !== '.' && a.push(s), a), []).join('/');
const colNo = ref => { let n = 0; for (const ch of ref.replace(/\d+$/, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };

/* ---------- xlsx: { sheetName: [[cell, ...], ...] } ---------- */
export function readXlsx(buf) {
  const z = unzip(buf);
  const shared = [];
  for (const m of (z['xl/sharedStrings.xml']?.toString('utf8') || '').matchAll(/<si>([\s\S]*?)<\/si>/g)) {
    shared.push([...m[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, '').matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(t => unxml(t[1])).join(''));
  }
  const wb = z['xl/workbook.xml'].toString('utf8');
  const wr = relsOf(z, 'xl/_rels/workbook.xml.rels');
  const out = {};
  for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
    const rel = wr[attr(m[0], 'r:id')];
    if (!rel) continue;
    const path = resolve('xl/', rel.target);
    const sx = z[path]?.toString('utf8');
    if (!sx) continue;
    const srels = relsOf(z, path.replace(/([^/]+)$/, '_rels/$1.rels'));
    const links = {};
    for (const h of sx.matchAll(/<hyperlink\b[^>]*>/g)) { const r = srels[attr(h[0], 'r:id')]; if (r) links[attr(h[0], 'ref')] = r.target; }
    const rows = [];
    for (const r of sx.matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
      const ri = +attr(' ' + r[1], 'r') - 1;
      const row = [];
      for (const c of (r[2] || '').matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const ref = attr(' ' + c[1], 'r'), t = attr(' ' + c[1], 't'), inner = c[2] || '';
        const v = (inner.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
        let val = '';
        if (t === 's') val = shared[+v] ?? '';
        else if (t === 'inlineStr') val = [...inner.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(x => unxml(x[1])).join('');
        else if (t === 'b') val = v === '1' ? 'TRUE' : 'FALSE';
        else if (v != null) val = unxml(v);
        if (links[ref] && !/^https?:/.test(val)) val = links[ref];
        row[colNo(ref)] = val;
      }
      rows[ri >= 0 ? ri : rows.length] = Array.from(row, x => x ?? '');
    }
    out[attr(m[0], 'name')] = Array.from(rows, x => x || []);
  }
  return out;
}

/* ---------- docx: [{ style: 'title'|'heading'|'p', segs: [{text, href?, b?, i?} | {img: {data, ext, alt}}] }] ---------- */
export function readDocx(buf) {
  const z = unzip(buf);
  const doc = z['word/document.xml'].toString('utf8');
  const rels = relsOf(z, 'word/_rels/document.xml.rels');
  const styles = {};
  for (const m of (z['word/styles.xml']?.toString('utf8') || '').matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
    const nm = (m[2].match(/<w:name w:val="([^"]*)"/) || [])[1] || '';
    styles[attr(' ' + m[1], 'w:styleId')] = nm.toLowerCase();
  }
  const blocks = [];
  const body = doc.replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, '');
  for (const pm of body.matchAll(/<w:p\b[^>]*?(?:\/>|>([\s\S]*?)<\/w:p>)/g)) {
    const inner = pm[1] || '';
    const sid = (inner.match(/<w:pStyle w:val="([^"]*)"/) || [])[1] || '';
    const sname = styles[sid] || sid.toLowerCase();
    const style = /^title$/.test(sname) ? 'title' : /^(heading|見出し)\s*\d|^subtitle$/.test(sname.replace(/(\D)(\d)/, '$1 $2')) ? 'heading' : 'p';
    const segs = [];
    const hrefStack = [];
    let fld = null; // complex field being read: { instr, href }
    const fldStack = [];
    for (const t of inner.matchAll(/<w:hyperlink\b[^>]*>|<\/w:hyperlink>|<w:fldSimple\b[^>]*>|<\/w:fldSimple>|<w:r\b[^>]*?(?:\/>|>[\s\S]*?<\/w:r>)/g)) {
      const tok = t[0];
      if (tok.startsWith('<w:hyperlink')) {
        const r = rels[attr(tok, 'r:id')];
        const anchor = attr(tok, 'w:anchor');
        hrefStack.push(r ? r.target : anchor ? '#' + anchor : null);
        continue;
      }
      if (tok === '</w:hyperlink>' || tok === '</w:fldSimple>') { hrefStack.pop(); continue; }
      if (tok.startsWith('<w:fldSimple')) { hrefStack.push(fieldHref(attr(tok, 'w:instr'))); continue; }
      // a run
      const fc = (tok.match(/<w:fldChar\b[^>]*w:fldCharType="(\w+)"/) || [])[1];
      if (fc === 'begin') { fldStack.push(fld); fld = { instr: '', href: null, shown: false }; continue; }
      if (fc === 'separate') { if (fld) { fld.href = fieldHref(fld.instr); fld.shown = true; } continue; }
      if (fc === 'end') { fld = fldStack.pop() ?? null; continue; }
      if (fld && !fld.shown) { for (const it of tok.matchAll(/<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g)) fld.instr += unxml(it[1]); continue; }
      const href = (fld && fld.href) || [...hrefStack].reverse().find(Boolean) || [...fldStack].reverse().find(f => f && f.href)?.href || null;
      const rpr = (tok.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/) || [])[1] || '';
      const on = k => new RegExp(`<w:${k}(?: w:val="(?:1|true|on)")?\\s*/>`).test(rpr);
      for (const piece of tok.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:tab\/>|<w:br\b[^>]*\/>|<w:drawing>([\s\S]*?)<\/w:drawing>|<w:pict>([\s\S]*?)<\/w:pict>/g)) {
        const p = piece[0];
        if (p.startsWith('<w:t')) segs.push({ text: unxml(piece[1]), href, b: on('b'), i: on('i') });
        else if (p === '<w:tab/>') segs.push({ text: ' ', href });
        else if (p.startsWith('<w:br')) segs.push({ text: '\n', href });
        else {
          const x = piece[2] || piece[3] || '';
          const id = (x.match(/r:(?:embed|id)="([^"]+)"/) || [])[1];
          const r = id && rels[id];
          const file = r && !r.external ? z[resolve('word/', r.target)] : null;
          if (!file) continue;
          const docPr = (x.match(/<wp:docPr\b[^>]*>/) || [''])[0];
          segs.push({ img: { data: file, ext: (r.target.match(/\.(\w+)$/) || [])[1] || 'png', alt: (attr(docPr, 'descr') || attr(docPr, 'title')).trim() }, href });
        }
      }
    }
    blocks.push({ style, segs: mergeSegs(segs) });
  }
  return blocks;
}
function fieldHref(instr) {
  const m = String(instr || '').match(/HYPERLINK\s+(?:\\l\s+)?"([^"]+)"/i);
  return m ? m[1] : null;
}
function mergeSegs(segs) {
  const out = [];
  for (const s of segs) {
    const last = out[out.length - 1];
    if (last && !s.img && !last.img && last.href === s.href && !!last.b === !!s.b && !!last.i === !!s.i) last.text += s.text;
    else out.push({ ...s });
  }
  return out;
}
export const blockText = b => b.segs.map(s => s.text || '').join('');
// plain text with "# " for headings, paragraphs separated by blank lines (for about / info files written as Docs)
export const docxToPlain = buf => readDocx(buf).map(b => (b.style === 'p' ? '' : '# ') + blockText(b)).join('\n\n');
