
// PDF to Word (.docx), entirely in the browser.
// Each PDF page is drawn to a picture and placed on its own Word page, edge to
// edge, with the Word page set to the PDF page's size. Nothing reflows. The
// PDF's links become clear clickable pictures over the same spots, and the
// words of each link become real text there, since every app clicks text links.

const DPI = 200;                 // picture sharpness
const MAX_PX = 4000;             // longest side of one page picture
const MAX_PAGE_TW = 22 * 1440;   // Word's largest page, 22 inches, in twips
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

const esc = (s) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);

// Apple's preview (Mail on Mac and iPhone) lays the pages out with 5.4pt more
// below each page picture than Word does, but places floating things as if it
// didn't, so links on later pages drifted up. See pageParagraph.
const APPLE_GAP = 5.4 * 12700;   // EMU

// The page picture, in line or floating at the page's top left corner.
const pagePicture = (id, n, cx, cy, frame, place) =>
  `<w:r><w:rPr><w:sz w:val="2"/></w:rPr><w:drawing>${frame[0]}<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${place}` +
  `<wp:docPr id="${id}" name="Page ${n}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
  `<pic:nvPicPr><pic:cNvPr id="${id}" name="page${n}.jpeg"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rId${n}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
  `</pic:pic></a:graphicData></a:graphic>${frame[1]}</w:drawing></w:r>`;

// One page. The picture sits in line, which every app that opens .docx shows,
// made APPLE_GAP shorter so Apple's layout adds up to the page height and its
// links land where they belong. A full-size copy of the same picture floats on
// top at the page's corner, so the page still looks exact. The negative right
// indent keeps Apple's end-of-page mark on the picture's line, and the 1pt
// paragraph mark adds no height in Word. `extra` is the page's bookmark (when a
// link jumps here) and its links. Apple puts page n's floating things n - 1
// times its own height down, which is only right while every page so far has
// the same size. From the first page of another size on, a copy would cover
// the wrong page, so those pages (copyId 0) keep a full-size picture only.
const pageParagraph = (n, copyId, cx, cy, sectPr, extra = { mark: '', links: '' }) =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:right="-200"/>` +
  `<w:rPr><w:sz w:val="2"/></w:rPr>${sectPr}</w:pPr>${extra.mark}` +
  pagePicture(n, n, cx, copyId ? Math.max(cy / 2, cy - APPLE_GAP) : cy, ['<wp:inline distT="0" distB="0" distL="0" distR="0">', '</wp:inline>'], '') +
  (copyId ? pagePicture(copyId, n, cx, cy,
    [`<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${copyId}" behindDoc="0" locked="1" layoutInCell="1" allowOverlap="1">` +
      '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>', '</wp:anchor>'], '<wp:wrapNone/>') : '') +
  `${extra.links}</w:p>`;

// One link: a clear picture floating in front of the page picture, over the
// spot where the PDF had the link, carrying the link. Positions are in EMU from
// the page's top left corner.
const linkRun = (id, rid, x, y, cx, cy, label) =>
  `<w:r><w:rPr><w:sz w:val="2"/></w:rPr><w:drawing>` +
  `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${id}" behindDoc="0" locked="1" layoutInCell="1" allowOverlap="1">` +
  `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${x}</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="page"><wp:posOffset>${y}</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
  `<wp:docPr id="${id}" name="Link ${id}" descr="${esc(label)}"><a:hlinkClick r:id="${rid}"/></wp:docPr>` +
  `<wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
  `<pic:nvPicPr><pic:cNvPr id="${id}" name="link${id}.png"><a:hlinkClick r:id="${rid}"/></pic:cNvPr><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rIdClear"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
  `</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;

// One line of a link's words: a see-through text box at that spot holding a
// text link. Word 2010 and later read the first form, older readers the second.
const textBox = (id, x, y, cx, cy, para) =>
  `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>` +
  `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${id}" behindDoc="0" locked="1" layoutInCell="1" allowOverlap="1">` +
  `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${x}</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="page"><wp:posOffset>${y}</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
  `<wp:docPr id="${id}" name="Link text ${id}"/><wp:cNvGraphicFramePr/>` +
  `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>` +
  `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>` +
  `<wps:txbx><w:txbxContent>${para}</w:txbxContent></wps:txbx>` +
  `<wps:bodyPr rot="0" vert="horz" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit/></wps:bodyPr></wps:wsp>` +
  `</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict>` +
  `<v:shape id="LinkText${id}" type="#_x0000_t202" filled="f" stroked="f" style="position:absolute;margin-left:${x / 12700}pt;margin-top:${y / 12700}pt;` +
  `width:${cx / 12700}pt;height:${cy / 12700}pt;z-index:${id};mso-wrap-style:none;mso-position-horizontal-relative:page;mso-position-vertical-relative:page">` +
  `<v:textbox inset="0,0,0,0"><w:txbxContent>${para}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;

// The shape type the older form above refers to, written once.
const TEXT_BOX_TYPE = '<w:r><w:pict><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">' +
  '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype></w:pict></w:r>';

// Link words are written in a font every computer and phone has, picked to
// match the PDF's: ascent is how far the letters reach above the baseline.
const FALLBACK = {
  sans: { name: 'Arial', ascent: 0.905 },
  serif: { name: 'Times New Roman', ascent: 0.891 },
  mono: { name: 'Courier New', ascent: 0.833 },
  // fonts a web page may name that Word and Apple's apps have too
  georgia: { name: 'Georgia', ascent: 0.917 },
  verdana: { name: 'Verdana', ascent: 1.005 },
  tahoma: { name: 'Tahoma', ascent: 1 },
  trebuchet: { name: 'Trebuchet MS', ascent: 0.939 },
};

// Letters drawn inside a link are kept out of the page picture, to come back
// as real text in the same spot. A link whose letters can't all be written as
// text (no Unicode for them, see-through ink) gets them drawn back instead.
async function linkLetters(page, ctx, links, scale) {
  const boxes = [];
  links.forEach((link, i) => { for (const [x, y, w, h] of link.boxes) boxes.push([i, x * scale, y * scale, (x + w) * scale, (y + h) * scale]); });
  // What each drawn character stands for, from the page's drawing commands.
  const ops = await page.getOperatorList();
  const unicode = new Map();
  let font = '';
  ops.fnArray.forEach((fn, i) => {
    if (fn === pdfjsLib.OPS.setFont) font = ops.argsArray[i][0];
    else if (fn === pdfjsLib.OPS.showText || fn === pdfjsLib.OPS.showSpacedText)
      for (const g of ops.argsArray[i][0]) if (g && typeof g === 'object' && g.fontChar) unicode.set(`${font}|${g.fontChar}`, g.unicode);
  });
  const letters = [], unwritable = new Set();
  const fillText = ctx.fillText;
  ctx.fillText = function (text, x, y, ...rest) {
    const m = this.getTransform();
    const em = parseFloat(/([\d.]+)px/.exec(this.font)?.[1]);
    if (em > 0 && m.a > 0 && m.d > 0 && Math.abs(m.b) + Math.abs(m.c) < 1e-6 * (m.a + m.d)) {   // upright text only
      const w = this.measureText(text).width;
      const left = m.a * x + m.e, right = m.a * (x + w) + m.e, base = m.d * y + m.f, size = em * m.d;
      const cx = (left + right) / 2, cy = base - size * 0.3;
      const box = boxes.find(([, l, t, r, b]) => cx >= l && cx <= r && cy >= t && cy <= b);
      if (box) {
        const name = /px\s+"?([^",]+)/.exec(this.font)?.[1].trim() || '';
        // ligatures (fi, fl, ff ...) as their letters, so the words can be searched and copied
        const u = (unicode.get(`${name}|${text}`) ?? text).replace(/[ﬀ-ﬆ]/g, (c) => c.normalize('NFKC'));
        if (typeof this.fillStyle !== 'string' || this.globalAlpha < 0.99 || !/^[^\u0000-\u001f-�]+$/u.test(u)) unwritable.add(box[0]);
        letters.push({ link: box[0], text: u, left, right, base, size, name, fill: this.fillStyle,
          redraw: [m, this.font, this.fillStyle, this.globalAlpha, text, x, y] });
        return;
      }
    }
    return fillText.call(this, text, x, y, ...rest);
  };
  return () => {
    delete ctx.fillText;
    for (const l of letters) {
      if (!unwritable.has(l.link)) continue;
      const [m, font, fill, alpha, text, x, y] = l.redraw;
      ctx.save();
      ctx.setTransform(m);
      Object.assign(ctx, { font, fillStyle: fill, globalAlpha: alpha });
      ctx.fillText(text, x, y);
      ctx.restore();
    }
    return linkLines(page, letters.filter(l => !unwritable.has(l.link)), scale);
  };
}

const measure = document.createElement('canvas').getContext('2d');

// The left-out letters as lines of text, per link and baseline, in reading
// order, with sizes and places in points.
function linkLines(page, letters, scale) {
  const lines = [];
  letters.sort((a, b) => a.link - b.link || a.base - b.base);
  for (const l of letters) {
    const line = lines[lines.length - 1];
    if (line && line.link === l.link && Math.abs(line.base - l.base) < Math.max(line.size, l.size) * 0.35) line.letters.push(l);
    else lines.push({ link: l.link, base: l.base, size: l.size, letters: [l] });
  }
  return lines.map(({ link, letters: ls }) => {
    ls.sort((a, b) => a.left - b.left);
    const runs = [];
    let prev = null;
    for (const l of ls) {
      if (prev && l.text === prev.text && Math.abs(l.left - prev.left) < l.size * 0.1) continue;   // drawn twice, for weight
      const style = letterStyle(page, l, scale);
      let text = l.text;
      if (prev && l.left - prev.right > l.size * 0.2 && !/\s$/.test(prev.text) && !/^\s/.test(text)) text = ' ' + text;
      const last = runs[runs.length - 1];
      if (last && last.key === style.key) last.text += text;
      else runs.push({ ...style, text });
      prev = l;
    }
    const left = ls[0].left / scale, right = ls[ls.length - 1].right / scale;
    const top = Math.min(...ls.map(l => l.base / scale - FALLBACK[letterStyle(page, l, scale).kind].ascent * l.size / scale));
    const size = Math.max(...runs.map(r => r.size));
    // Spacing between letters so the line keeps the PDF's width in the fallback font.
    let natural = 0, count = -1;
    for (const r of runs) {
      measure.font = `${r.italic ? 'italic ' : ''}${r.bold ? 'bold ' : ''}${r.size}px "${FALLBACK[r.kind].name}"`;
      natural += measure.measureText(r.text).width;
      count += [...r.text].length;
    }
    const spacing = count > 0 ? Math.max(-0.2 * size, Math.min(0.3 * size, (right - left - natural) / count)) : 0;
    return { link, left, top, width: right - left, size, spacing, runs };
  });
}

// Font kind, weight, slant, color and size of one left-out letter.
function letterStyle(page, l, scale) {
  let font = {};
  try { font = page.commonObjs.get(l.name) || {}; } catch {}
  const name = font.name || '';
  const kind = font.isMonospace || /mono|courier|code/i.test(name) ? 'mono'
    : font.isSerifFont && !/sans/i.test(name) ? 'serif' : 'sans';
  const bold = !!(font.black || font.bold) || /bold|black|heavy/i.test(name);
  const italic = !!font.italic || /italic|oblique/i.test(name);
  let color = '000000';
  const hex = /^#([0-9a-f]{6})$/i.exec(l.fill);
  const rgb = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(l.fill);
  if (hex) color = hex[1].toUpperCase();
  else if (rgb) color = rgb.slice(1, 4).map(v => (+v).toString(16).padStart(2, '0')).join('').toUpperCase();
  const size = l.size / scale;
  return { kind, bold, italic, color, size, key: `${kind}|${bold}|${italic}|${color}|${Math.round(size * 2)}` };
}

// A fully clear 1x1 PNG, stretched over each link.
async function clearPng() {
  const canvas = document.createElement('canvas');
  canvas.width = canvas.height = 1;
  return (await new Promise(resolve => canvas.toBlob(resolve, 'image/png'))).arrayBuffer();
}

// Which page (1-based) a link inside the PDF jumps to, or 0.
async function destPage(pdf, dest) {
  try {
    const explicit = typeof dest === 'string' ? await pdf.getDestination(dest) : dest;
    const ref = Array.isArray(explicit) ? explicit[0] : null;
    const index = Number.isInteger(ref) ? ref : ref && typeof ref === 'object' ? await pdf.getPageIndex(ref) : -1;
    return index >= 0 && index < pdf.numPages ? index + 1 : 0;
  } catch {
    return 0;
  }
}

// The page's links: { url } for web and mail addresses (pdf.js only gives safe
// ones), { page } for jumps inside the PDF, each with its boxes in points from
// the page's top left corner. A link over two lines has one box per line.
async function pageLinks(pdf, page) {
  const viewport = page.getViewport({ scale: 1 });
  const links = [];
  for (const a of await page.getAnnotations({ intent: 'display' })) {
    if (a.subtype !== 'Link') continue;
    let target = null;
    if (a.url) target = { url: a.url };
    else if (a.dest) target = { page: await destPage(pdf, a.dest) };
    else if (a.action) {
      const n = page.pageNumber;
      target = { page: { FirstPage: 1, LastPage: pdf.numPages, NextPage: n + 1, PrevPage: n - 1 }[a.action] || 0 };
    }
    if (!target || target.page === 0 || target.page > pdf.numPages) continue;
    const quads = Array.isArray(a.quadPoints) && a.quadPoints.every(q => Array.isArray(q) && q.length === 4)
      ? a.quadPoints.map(q => [Math.min(...q.map(p => p.x)), Math.min(...q.map(p => p.y)), Math.max(...q.map(p => p.x)), Math.max(...q.map(p => p.y))])
      : [a.rect];
    const boxes = [];
    for (const r of quads) {
      const [x1, y1, x2, y2] = viewport.convertToViewportRectangle(r);
      const left = Math.max(0, Math.min(x1, x2)), right = Math.min(viewport.width, Math.max(x1, x2));
      const top = Math.max(0, Math.min(y1, y2)), bottom = Math.min(viewport.height, Math.max(y1, y2));
      if (right - left >= 0.5 && bottom - top >= 0.5) boxes.push([left, top, right - left, bottom - top]);
    }
    if (boxes.length) links.push({ ...target, boxes });
  }
  return links;
}

// Page size of one Word section, no margins unless asked.
const section = (w, h, m = 0) =>
  `<w:sectPr><w:pgSz w:w="${w}" w:h="${h}"${w > h ? ' w:orient="landscape"' : ''}/>` +
  `<w:pgMar w:top="${m}" w:right="${m}" w:bottom="${m}" w:left="${m}" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>`;

// The optional last page: every link written out as a plain text link, which
// every app can click, including email previews that ignore the link spots.
function linkPage(pageLinkList) {
  const font = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/>';
  const seen = new Set();
  let rows = '';
  pageLinkList.forEach((links, i) => {
    for (const link of links) {
      const key = `${i}|${link.url || link.page}`;
      if (seen.has(key)) continue;
      seen.add(key);
      const text = link.url ? link.url.replace(/^mailto:/i, '') : `Go to page ${link.page}`;
      const to = link.url ? `r:id="${link.rid}"` : `w:anchor="_Page${link.page}"`;
      rows += `<w:p><w:pPr><w:spacing w:before="0" w:after="80"/></w:pPr><w:r><w:rPr>${font}</w:rPr><w:t xml:space="preserve">Page ${i + 1}: </w:t></w:r>` +
        `<w:hyperlink ${to}><w:r><w:rPr>${font}<w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t>${esc(text)}</w:t></w:r></w:hyperlink></w:p>`;
    }
  });
  return rows && `<w:p><w:pPr><w:spacing w:before="0" w:after="240"/></w:pPr><w:r><w:rPr><w:b/>${font.replace('22', '28')}</w:rPr>` +
    `<w:t>Links in this document</w:t></w:r></w:p>${rows}`;
}

// The page as a picture, without its links' words, and those words as lines.
async function renderPage(page, links) {
  const size = page.getViewport({ scale: 1 });   // in points, rotation applied
  const scale = Math.min(DPI / 72, MAX_PX / Math.max(size.width, size.height));
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement('canvas');
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  const finish = links.length ? await linkLetters(page, ctx, links, scale) : () => [];
  await page.render({ canvasContext: ctx, viewport }).promise;
  const lines = finish();
  const jpeg = await new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', 0.92));
  canvas.width = canvas.height = 0;
  page.cleanup();
  return { jpeg: await jpeg.arrayBuffer(), width: size.width, height: size.height, lines };
}

// PDF bytes in, { docx bytes, number of links kept } out. onPage(n, total)
// reports progress; options.linkPage adds the last page listing the links.
async function pdfToDocx(data, onPage = () => {}, options = {}) {
  let pdf;
  try {
    pdf = await pdfjsLib.getDocument({ data, isEvalSupported: false }).promise;
  } catch (err) {
    throw new Error(err?.name === 'PasswordException' ? 'This PDF is password protected.' : 'Could not read this PDF.');
  }
  const pages = [];
  for (let n = 1; n <= pdf.numPages; n++) {
    onPage(n, pdf.numPages);
    const page = await pdf.getPage(n);
    const links = await pageLinks(pdf, page);
    pages.push({ ...(await renderPage(page, links)), links });
  }
  await pdf.destroy();
  return pagesToDocx(pages, options);
}

// Pages in, each { jpeg, width and height in points, links, lines }, and
// { docx bytes, number of links kept } out.
async function pagesToDocx(pages, options = {}) {
  const zip = new JSZip();
  const add = (name, content) => zip.file(name, content, { createFolders: false });
  const total = pages.length;
  const sizes = [];
  const pageLinkList = pages.map(p => p.links), pageLines = pages.map(p => p.lines);
  let rels = '';
  pages.forEach(({ jpeg, width, height }, i) => {
    add(`word/media/page${i + 1}.jpeg`, jpeg);
    rels += `<Relationship Id="rId${i + 1}" Type="${REL}image" Target="media/page${i + 1}.jpeg"/>`;
    // Same shape as the page; anything over 22 inches shrinks to fit, ratio kept.
    const k = Math.min(1, MAX_PAGE_TW / (Math.max(width, height) * 20));
    sizes.push([Math.round(width * 20 * k), Math.round(height * 20 * k), k]);
  });

  // Links: each gets a relationship (a web address, or a bookmark on the page
  // it jumps to) and one clear picture per box.
  const targets = new Set(pageLinkList.flat().filter(l => l.page).map(l => l.page));
  let linkCount = 0, pictureId = 2 * total, textBoxes = 0;   // ids up to 2 * total are the page pictures
  const extras = pageLinkList.map((links, i) => {
    const n = i + 1, k = sizes[i][2];
    const emu = (pt) => Math.round(pt * 12700 * k);
    const mark = targets.has(n) ? `<w:bookmarkStart w:id="${n}" w:name="_Page${n}"/><w:bookmarkEnd w:id="${n}"/>` : '';
    let runs = '';
    for (const link of links) {
      const rid = link.rid = `rIdLink${++linkCount}`;
      rels += link.url
        ? `<Relationship Id="${rid}" Type="${REL}hyperlink" Target="${esc(link.url)}" TargetMode="External"/>`
        : `<Relationship Id="${rid}" Type="${REL}hyperlink" Target="#_Page${link.page}"/>`;
      const label = link.url || `Page ${link.page}`;
      for (const [x, y, w, h] of link.boxes) runs += linkRun(++pictureId, rid, emu(x), emu(y), emu(w), emu(h), label);
    }
    // The links' words as text, on top of the clear pictures.
    for (const line of pageLines[i]) {
      const link = links[line.link];
      const open = link.url ? `<w:hyperlink r:id="${link.rid}">` : `<w:hyperlink w:anchor="_Page${link.page}">`;
      const spacing = Math.round(line.spacing * k * 20);
      const words = line.runs.map(r => {
        const font = FALLBACK[r.kind].name, sz = Math.max(2, Math.round(r.size * k * 2));
        return `<w:r><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/>` +
          `${r.bold ? '<w:b/><w:bCs/>' : ''}${r.italic ? '<w:i/><w:iCs/>' : ''}<w:color w:val="${r.color}"/>` +
          `${spacing ? `<w:spacing w:val="${spacing}"/>` : ''}<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>` +
          `${r.underline ? `<w:u w:val="single" w:color="${r.underline}"/>` : ''}</w:rPr>` +
          `<w:t xml:space="preserve">${esc(r.text)}</w:t></w:r>`;
      }).join('');
      const para = `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>${open}${words}</w:hyperlink></w:p>`;
      if (!textBoxes++) runs += TEXT_BOX_TYPE;
      runs += textBox(++pictureId, emu(line.left), emu(line.top), emu(line.width * 1.2 + line.size), emu(line.size * 1.35), para);
    }
    return { mark, links: runs };
  });
  if (linkCount) {
    add('word/media/link.png', await clearPng());
    rels += `<Relationship Id="rIdClear" Type="${REL}image" Target="media/link.png"/>`;
  }

  // Each page is its own Word section, which also starts it on a new page.
  // Apple's preview (Mail on Mac and iPhone) places floating things from the
  // top of their section, so in one long section every link landed on page 1.
  // The last section's settings go at the end of the body.
  const listed = options.linkPage ? linkPage(pageLinkList) : '';
  let sameSize = true;
  let body = '';
  sizes.forEach(([w, h], i) => {
    const last = i === sizes.length - 1 && !listed;
    sameSize = sameSize && w === sizes[0][0] && h === sizes[0][1];
    body += pageParagraph(i + 1, sameSize ? total + i + 1 : 0, w * 635, h * 635, last ? '' : section(w, h), extras[i]);
  });
  const [lw, lh] = sizes[sizes.length - 1];
  body += listed;
  const lastSection = listed ? section(lw, lh, Math.min(1440, Math.round(Math.min(lw, lh) / 8))) : section(lw, lh);

  add('[Content_Types].xml', XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="jpeg" ContentType="image/jpeg"/>' +
    (linkCount ? '<Default Extension="png" ContentType="image/png"/>' : '') +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '</Types>');
  add('_rels/.rels', XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="rId1" Type="${REL}officeDocument" Target="word/document.xml"/></Relationships>`);
  add('word/_rels/document.xml.rels', XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`);
  add('word/document.xml', XML + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
    ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
    ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"' +
    ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"' +
    ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
    ' xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">' +
    `<w:body>${body}${lastSection}</w:body></w:document>`);
  return { docx: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), links: linkCount };
}
