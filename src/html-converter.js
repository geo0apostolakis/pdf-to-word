
// HTML to Word: a slide deck saved as a web page (Claude Slides and the like),
// or any page. Each slide, a top-level <section>, is laid out in a hidden
// frame where nothing from the file runs, drawn to a picture through an SVG
// foreignObject, and its links are measured there. A page without slides is
// cut into A4 pages. The pages then go through the same Word writer as PDFs.

const DECK_PT = 960;               // a slide is 960pt wide, as in the Claude Slides PDF export (1920px = 960pt)
const CSS_PT = 0.75;               // CSS pixels to points, for pages that are not slides
const A4_PX = [794, 1123];         // an A4 page in CSS pixels
const BLANK_GIF = 'data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==';
const TYPES = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp', svg: 'image/svg+xml',
  avif: 'image/avif', woff2: 'font/woff2', woff: 'font/woff', ttf: 'font/ttf', otf: 'font/otf', css: 'text/css' };

// The page's HTML, and the files beside it when it comes as a .zip.
async function readHtml(bytes) {
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) return { html: new TextDecoder().decode(bytes), dir: '', files: new Map() };
  let zip;
  try { zip = await JSZip.loadAsync(bytes); } catch { throw new Error('Could not open this .zip file.'); }
  const names = Object.keys(zip.files).filter(n => !zip.files[n].dir && !/(^|\/)(__MACOSX|\.)/.test(n));
  const pages = names.filter(n => /\.html?$/i.test(n))
    .sort((a, b) => a.split('/').length - b.split('/').length || /(^|\/)index\.html?$/i.test(b) - /(^|\/)index\.html?$/i.test(a) || a.length - b.length);
  if (!pages.length) throw new Error('There is no .html file in this .zip.');
  return { html: await zip.file(pages[0]).async('string'), dir: pages[0].slice(0, pages[0].lastIndexOf('/') + 1), files: new Map(names.map(n => [n, zip.files[n]])) };
}

// Everything the page loads, as data: URLs, so the hidden frame and the
// pictures need no network. Files come from the .zip, or from the web for web
// addresses (a deck's Google Fonts, say). Counts what could not be had.
function resources(src) {
  const cache = new Map();
  const res = { missing: 0 };
  const blobOf = async (url, base) => {
    if (/^data:/i.test(url)) return (await fetch(url)).blob();
    const abs = /^https?:/i.test(url) ? url : /^https?:/i.test(base) ? new URL(url, base).href : null;
    if (abs) {
      const r = await fetch(abs, { credentials: 'omit', referrerPolicy: 'no-referrer' });
      if (!r.ok) throw new Error(r.status);
      return r.blob();
    }
    const path = decodeURIComponent(new URL(url, 'file:///' + base).pathname.slice(1));
    const file = src.files.get(path);
    if (!file) throw new Error('missing');
    return new Blob([await file.async('uint8array')], { type: TYPES[path.split('.').pop().toLowerCase()] || '' });
  };
  // a URL as a data: URL, or null
  res.data = (url, base = src.dir) => {
    url = url.trim();
    const key = base + '\n' + url;
    if (!cache.has(key)) cache.set(key, blobOf(url, base).then(b => new Promise((ok, fail) => {
      const r = new FileReader();
      r.onload = () => ok(r.result);
      r.onerror = fail;
      r.readAsDataURL(b);
    })).catch(() => { res.missing++; return null; }));
    return cache.get(key);
  };
  // CSS with its @imports pulled in and every url() made a data: URL
  res.css = async (css, base = src.dir, depth = 0) => {
    const imports = [];
    css = css.replace(/@import\s+(?:url\()?\s*['"]?([^'")\s;]+)['"]?\s*\)?[^;]*;/gi, (m, url) => { imports.push(url); return ''; });
    let head = '';
    for (const url of imports) head += depth < 4 ? await res.sheet(url, base, depth + 1) : '';
    const urls = [...new Set([...css.matchAll(/url\(\s*(['"]?)([^'")]+)\1\s*\)/gi)].map(m => m[2]).filter(u => !/^(data:|#)/i.test(u.trim())))];
    const done = new Map(await Promise.all(urls.map(async u => [u, await res.data(u, base)])));
    return head + css.replace(/url\(\s*(['"]?)([^'")]+)\1\s*\)/gi, (m, q, u) => /^(data:|#)/i.test(u.trim()) ? m : `url("${done.get(u) || 'data:,'}")`);
  };
  // a linked style sheet
  res.sheet = async (href, base = src.dir, depth = 0) => {
    const url = await res.data(href, base);
    if (!url) return '';
    const text = await (await fetch(url)).text();
    const abs = /^https?:/i.test(href) ? href : /^https?:/i.test(base) ? new URL(href, base).href : new URL(href, 'file:///' + base).pathname.slice(1);
    return res.css(text, abs.slice(0, abs.lastIndexOf('/') + 1), depth);
  };
  return res;
}

// Live embeds (a deck's <x-embed>, an <iframe srcdoc>) are small pages of
// their own. Each is swapped for a picture placeholder and its page kept.
function takeEmbeds(doc) {
  const embeds = [];
  for (const el of doc.querySelectorAll('x-embed, iframe')) {
    const img = doc.createElement('img');
    for (const a of ['style', 'class', 'id', 'width', 'height']) if (el.hasAttribute(a)) img.setAttribute(a, el.getAttribute(a));
    if (el.localName === 'iframe' && !el.hasAttribute('width') && !/width/.test(el.getAttribute('style') || '')) img.setAttribute('width', '300');
    if (el.localName === 'iframe' && !el.hasAttribute('height') && !/height/.test(el.getAttribute('style') || '')) img.setAttribute('height', '150');
    img.setAttribute('alt', '');
    img.setAttribute('src', BLANK_GIF);
    img.dataset.p2wEmbed = embeds.length;
    embeds.push(el.localName === 'x-embed' ? el.innerHTML : el.getAttribute('srcdoc') || '');
    el.replaceWith(img);
  }
  return embeds;
}

// Nothing from the file runs in the frame: no scripts, handlers or javascript: links.
function scrub(doc) {
  doc.querySelectorAll('script, noscript, base, meta[http-equiv], object, embed, frame, frameset, template, dialog').forEach(e => e.remove());
  for (const el of doc.querySelectorAll('*')) for (const { name, value } of [...el.attributes]) {
    if (/^on/i.test(name) || /^\s*javascript:/i.test(value) || name === 'srcset' || name === 'autofocus') el.removeAttribute(name);
  }
}

// The page's style sheets as one, and its pictures as data: URLs.
async function inlineAll(doc, res) {
  let css = '';
  for (const el of doc.querySelectorAll('link, style')) {
    if (el.localName === 'style') css += await res.css(el.textContent) + '\n';
    else if (/\bstylesheet\b/i.test(el.getAttribute('rel') || '') && el.getAttribute('href') && !/^\s*print\s*$/i.test(el.getAttribute('media') || ''))
      css += await res.sheet(el.getAttribute('href')) + '\n';
    el.remove();
  }
  await Promise.all([
    ...[...doc.querySelectorAll('[style*="url("]')].map(async el => el.setAttribute('style', await res.css(el.getAttribute('style')))),
    ...[...doc.querySelectorAll('img:not([data-p2w-embed]), video[poster], input[type=image]')].map(async el => {
      const attr = el.localName === 'video' ? 'poster' : 'src', url = el.getAttribute(attr);
      el.setAttribute(attr, url && !/^data:/i.test(url) ? (await res.data(url)) || BLANK_GIF : url || BLANK_GIF);
      el.removeAttribute('loading');
    }),
    ...[...doc.querySelectorAll('image')].map(async el => {
      const url = el.getAttribute('href') || el.getAttribute('xlink:href');
      if (url && !/^(data:|#)/i.test(url)) el.setAttribute('href', (await res.data(url)) || BLANK_GIF);
      el.removeAttribute('xlink:href');
    }),
  ]);
  doc.querySelectorAll('picture source, video source, audio, track').forEach(e => e.remove());
  return css;
}

const lockFrame = (sandbox, w, h) => {
  const frame = document.createElement('iframe');
  frame.setAttribute('sandbox', sandbox);
  frame.style.cssText = `position:fixed;left:-50000px;top:0;width:${w}px;height:${h}px;border:0;opacity:0;pointer-events:none`;
  return frame;
};
const LOCKED = (scripts) => `<meta http-equiv="Content-Security-Policy" content="default-src 'none'; ${scripts ? "script-src 'unsafe-inline'; " : ''}style-src 'unsafe-inline'; img-src data:; font-src data:">`;
// Still pictures: no animation or transition, whatever the page says.
const STILL = '*,*::before,*::after{animation:none!important;transition:none!important}';

// Runs one embed alone, scripts on but no fetches and no way to reach this
// page, and returns its page as it stands a moment after loading. It can still
// send its own frame to a new address, which tells a server no more than the
// file's pictures already do when they are loaded.
function runEmbed(html, w, h) {
  return new Promise(resolve => {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    doc.querySelectorAll('meta[http-equiv], base').forEach(e => e.remove());
    doc.head.insertAdjacentHTML('afterbegin', LOCKED(true));
    const token = Math.random().toString(36).slice(2);
    doc.body.insertAdjacentHTML('beforeend', `<script>addEventListener('load', () => setTimeout(() => parent.postMessage({ p2w: '${token}', html: document.documentElement.outerHTML }, '*'), 300));<\/script>`);
    const frame = lockFrame('allow-scripts', w, h);
    const done = (out) => { clearTimeout(timer); removeEventListener('message', onMessage); frame.remove(); resolve(out); };
    const onMessage = (e) => { if (e.source === frame.contentWindow && e.data && e.data.p2w === token) done(e.data.html); };
    const timer = setTimeout(() => done('<!doctype html>' + doc.documentElement.outerHTML), 4000);
    addEventListener('message', onMessage);
    frame.srcdoc = '<!doctype html>' + doc.documentElement.outerHTML;
    document.body.append(frame);
  });
}

// Draws XHTML markup laid out on a fw x fh area, cut to the view rectangle,
// at `scale` canvas pixels per CSS pixel, on `background` (none if null).
// Only the top `cut` CSS pixels of the view are drawn.
async function drawMarkup(markup, fw, fh, view, scale, background = '#fff', cut = view[3]) {
  const [vx, vy, vw, vh] = view;
  const w = Math.max(1, Math.round(vw * scale)), h = Math.max(1, Math.round(vh * scale));
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="${vx} ${vy} ${vw} ${vh}">` +
    `<foreignObject x="0" y="0" width="${fw}" height="${fh}">${markup}</foreignObject></svg>`;
  const img = new Image();
  img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
  await img.decode();
  const canvas = document.createElement('canvas');
  canvas.width = w;
  canvas.height = h;
  const ctx = canvas.getContext('2d');
  if (background) {
    // white under it, as a JPEG has no see-through
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, w, h);
    ctx.fillStyle = background;
    ctx.fillRect(0, 0, w, h);
  }
  ctx.save();
  ctx.beginPath();
  ctx.rect(0, 0, w, Math.round(cut * scale));
  ctx.clip();
  ctx.drawImage(img, 0, 0, w, h);
  ctx.restore();
  return canvas;
}

// An embed's page as a PNG data: URL of its box.
async function embedPicture(html, w, h) {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  doc.querySelectorAll('script, noscript, link, meta, base').forEach(e => e.remove());
  scrub(doc);
  const ser = new XMLSerializer();
  const styles = [...doc.querySelectorAll('style')].map(s => { s.remove(); return s.textContent; }).join('\n');
  const body = [...doc.body.childNodes].map(n => ser.serializeToString(n)).join('');
  const markup = `<div xmlns="http://www.w3.org/1999/xhtml" style="width:${w}px;height:${h}px;overflow:hidden;position:relative;${esc(doc.body.getAttribute('style') || '')}">` +
    `<style>${esc(styles + STILL)}</style>${body}</div>`;
  try {
    // see-through where the embed paints nothing, as the deck shows it
    return (await drawMarkup(markup, w, h, [0, 0, w, h], Math.min(3, 4000 / Math.max(w, h)), null)).toDataURL('image/png');
  } catch {
    return BLANK_GIF;
  }
}

// Computed style that a plain element of the same kind does not already have,
// as CSS text. Defaults come from an empty frame with no page styles.
function styleDiffer(defaultsDoc) {
  const cache = new Map();
  const skip = /^(animation|transition|will-change|cursor|pointer-events|user-select|-webkit-user|caret|scroll|overscroll|view-transition|anchor|position-try|inset-area|touch-action|content-visibility)/;
  const base = (el, pseudo) => {
    const key = el.namespaceURI + '|' + el.localName + '|' + (pseudo || '');
    if (!cache.has(key)) {
      const d = defaultsDoc.createElementNS(el.namespaceURI, el.localName);
      const holder = el.namespaceURI === 'http://www.w3.org/2000/svg' && el.localName !== 'svg'
        ? defaultsDoc.body.appendChild(defaultsDoc.createElementNS(el.namespaceURI, 'svg')) : defaultsDoc.body;
      holder.appendChild(d);
      const cs = defaultsDoc.defaultView.getComputedStyle(d, pseudo);
      const map = new Map();
      for (let i = 0; i < cs.length; i++) map.set(cs[i], cs.getPropertyValue(cs[i]));
      cache.set(key, map);
      (holder === defaultsDoc.body ? d : holder).remove();
    }
    return cache.get(key);
  };
  return (el, cs, pseudo) => {
    const def = base(el, pseudo);
    let out = '';
    for (let i = 0; i < cs.length; i++) {
      const name = cs[i];
      if (skip.test(name) || name.startsWith('--')) continue;
      const v = cs.getPropertyValue(name);
      if (v !== def.get(name)) out += `${name}:${v};`;
    }
    return out;
  };
}

// A copy of `root` that looks the same with no style sheets: every element
// carries its computed style, ::before and ::after become rules on generated
// classes, and the letters inside `hideIn` links are see-through (they come
// back as real text in Word).
function styledCopy(root, fdoc, differ, hideIn) {
  const win = fdoc.defaultView;
  let pseudoCss = '', n = 0;
  const families = new Set();
  const copy = (node, hide) => {
    if (node.nodeType === 3) return fdoc.createTextNode(node.data);
    if (node.nodeType !== 1) return null;
    const cs = win.getComputedStyle(node);
    if (cs.display === 'none') return null;
    const el = node.cloneNode(false);
    hide = hide || hideIn.has(node);
    let style = differ(node, cs, null);
    if (hide) style += '-webkit-text-fill-color:transparent;text-shadow:none;text-decoration-color:transparent;';
    if (node === root) style += 'position:relative;left:0;top:0;right:auto;bottom:auto;margin:0;transform:none;';
    el.setAttribute('style', style);
    families.add(cs.fontFamily);
    for (const pseudo of ['::before', '::after']) {
      const ps = win.getComputedStyle(node, pseudo);
      if (!ps.content || ps.content === 'none' || ps.content === 'normal') continue;
      const cls = `p2w${++n}`;
      el.setAttribute('class', (el.getAttribute('class') || '') + ' ' + cls);
      pseudoCss += `.${cls}${pseudo}{content:${ps.content};${differ(node, ps, pseudo)}${hide ? '-webkit-text-fill-color:transparent;' : ''}}`;
      families.add(ps.fontFamily);
    }
    if (node.localName === 'img' && node.currentSrc) el.setAttribute('src', node.currentSrc);
    for (const child of node.childNodes) { const c = copy(child, hide); if (c) el.appendChild(c); }
    return el;
  };
  const el = copy(root, false);
  return { el, pseudoCss, families };
}

// The @font-face rules of the fonts in use, from the frame's style sheets.
function fontFaces(fdoc, families) {
  const used = new Set([...families].flatMap(f => f.split(',').map(s => s.trim().replace(/^['"]|['"]$/g, '').toLowerCase())));
  let out = '';
  const walk = (rules) => {
    for (const rule of rules) {
      if (rule.type === 5) {   // CSSRule.FONT_FACE_RULE
        const family = rule.style.getPropertyValue('font-family').trim().replace(/^['"]|['"]$/g, '').toLowerCase();
        if (used.has(family)) out += rule.cssText + '\n';
      } else if (rule.cssRules) walk(rule.cssRules);
    }
  };
  for (const sheet of fdoc.styleSheets) { try { walk(sheet.cssRules); } catch {} }
  return out;
}

// The FALLBACK font for a CSS font-family list: the font the page really used
// when Word has it too, otherwise a common one of the same kind. `webFonts` are
// the families the page loaded itself, which Word does not have.
const SHARED = { georgia: 'georgia', verdana: 'verdana', tahoma: 'tahoma', 'trebuchet ms': 'trebuchet', arial: 'sans', helvetica: 'sans',
  'times new roman': 'serif', times: 'serif', 'courier new': 'mono', courier: 'mono' };
const GENERIC = new Set(['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui', 'ui-serif', 'ui-sans-serif', 'ui-monospace', 'ui-rounded', 'math', 'emoji', 'fangsong']);
const installed = new Map();
const isInstalled = (name) => {
  if (!installed.has(name)) {
    const probe = (font) => { measure.font = font; return measure.measureText('mmmmmmmmmmlli1WQ@').width; };
    installed.set(name, probe(`72px "${name}", monospace`) !== probe('72px monospace') || probe(`72px "${name}", serif`) !== probe('72px serif'));
  }
  return installed.get(name);
};
function familyKind(list, webFonts = new Set()) {
  const names = list.split(',').map(s => s.trim().replace(/^['"]|['"]$/g, '').toLowerCase());
  for (const f of names) {
    if (SHARED[f]) return SHARED[f];
    if (GENERIC.has(f) || f.startsWith('-') || webFonts.has(f) || isInstalled(f)) break;
  }
  if (names.some(f => f === 'monospace' || /mono|courier|code|consol|menlo/.test(f))) return 'mono';
  if (names.some(f => f === 'sans-serif' || /sans|arial|helvetica|grotesk|gothic/.test(f))) return 'sans';
  if (names.some(f => f === 'serif' || /serif|times|georgia|garamond|roman/.test(f))) return 'serif';
  return 'sans';
}

const cssHex = (c) => {
  const m = /rgba?\(([\d.]+),\s*([\d.]+),\s*([\d.]+)(?:,\s*([\d.]+))?/.exec(c);
  return m ? { hex: m.slice(1, 4).map(v => Math.round(+v).toString(16).padStart(2, '0')).join('').toUpperCase(), alpha: m[4] === undefined ? 1 : +m[4] } : { hex: '000000', alpha: 1 };
};
const applyCase = (t, cs) => cs.textTransform === 'uppercase' ? t.toUpperCase() : cs.textTransform === 'lowercase' ? t.toLowerCase() : t;

// Every link inside `root`: its target, its boxes (CSS px, frame coordinates,
// one per line), and its letters with their boxes, styles and underline color.
// A link whose words are see-through, filtered or turned gets no letters: its
// words stay in the picture, since Word could not draw them the same.
function measureLinks(root, fdoc, targetOf) {
  const win = fdoc.defaultView, range = fdoc.createRange();
  const links = [];
  for (const a of root.querySelectorAll('a[href]')) {
    if (win.getComputedStyle(a).visibility === 'hidden') continue;
    const target = targetOf(a);
    if (!target) continue;
    range.selectNodeContents(a);
    const rects = [...a.getClientRects(), ...range.getClientRects(), ...[...a.querySelectorAll('img, svg')].flatMap(e => [...e.getClientRects()])]
      .filter(r => r.width >= 1 && r.height >= 1);
    const boxes = [];
    for (const r of rects.sort((p, q) => p.top - q.top)) {
      const b = boxes.find(b => Math.min(b.bottom, r.bottom) - Math.max(b.top, r.top) > 0.5 * Math.min(b.bottom - b.top, r.height) && r.left <= b.right + 2 && r.right >= b.left - 2);
      if (b) Object.assign(b, { left: Math.min(b.left, r.left), right: Math.max(b.right, r.right), top: Math.min(b.top, r.top), bottom: Math.max(b.bottom, r.bottom) });
      else boxes.push({ left: r.left, right: r.right, top: r.top, bottom: r.bottom });
    }
    if (!boxes.length) continue;
    let letters = [], writable = true;
    const walker = fdoc.createTreeWalker(a, 4);   // NodeFilter.SHOW_TEXT
    for (let node; (node = walker.nextNode());) {
      const cs = win.getComputedStyle(node.parentElement);
      if (cs.visibility !== 'visible' || !node.data.trim()) continue;
      let underline = '';
      for (let e = node.parentElement; e && e !== root; e = e.parentElement) {
        const es = win.getComputedStyle(e);
        if (+es.opacity < 0.99 || es.filter !== 'none' || es.transform !== 'none' || es.mixBlendMode !== 'normal' || es.writingMode !== 'horizontal-tb') writable = false;
        if (!underline && es.textDecorationLine.includes('underline')) underline = cssHex(es.textDecorationColor);
        if (e === a) break;
      }
      if (cssHex(cs.color).alpha < 0.99 || underline && underline.alpha < 0.99) writable = false;
      for (let i = 0; i < node.data.length;) {
        const len = node.data.codePointAt(i) > 0xffff ? 2 : 1;
        range.setStart(node, i);
        range.setEnd(node, i + len);
        const r = range.getBoundingClientRect();
        if (r.width > 0 && r.height > 0) letters.push({ node, offset: i, text: applyCase(node.data.slice(i, i + len), cs), left: r.left, right: r.right, top: r.top, bottom: r.bottom, cs, underline: underline && underline.hex });
        i += len;
      }
    }
    if (!writable) letters = [];
    links.push({ a, target, boxes, letters });
  }
  return links;
}

// The letters of each link as lines; the baseline of each line comes from a
// zero-size marker put before its first letter (this changes the frame, so it
// runs after the picture is taken). Marker and letter are kept together, or a
// marker at a line break would stay at the end of the line before.
function letterLines(links, fdoc) {
  const lines = [];
  links.forEach((link, li) => {
    let line = null;
    for (const l of link.letters) {
      if (line && Math.min(line.bottom, l.bottom) - Math.max(line.top, l.top) > 0.5 * Math.min(line.bottom - line.top, l.bottom - l.top) && l.left >= line.right - 2) {
        line.letters.push(l);
        line.right = Math.max(line.right, l.right);
        line.top = Math.min(line.top, l.top);
        line.bottom = Math.max(line.bottom, l.bottom);
      } else lines.push(line = { link: li, letters: [l], left: l.left, right: l.right, top: l.top, bottom: l.bottom });
    }
  });
  // markers, last first within each text, so the earlier offsets stay right
  const byNode = new Map();
  const first = (line) => line.letters.find(l => l.text.trim()) || line.letters[0];
  for (const line of lines) {
    const s = first(line);
    if (!byNode.has(s.node)) byNode.set(s.node, []);
    byNode.get(s.node).push(s);
  }
  const marks = new Map();
  for (const [node, starts] of byNode) for (const s of starts.sort((p, q) => q.offset - p.offset)) {
    const rest = s.offset ? node.splitText(s.offset) : node;
    const len = rest.data.codePointAt(0) > 0xffff ? 2 : 1;
    if (rest.length > len) rest.splitText(len);
    const keep = fdoc.createElement('p2w-keep');
    keep.style.cssText = 'white-space:nowrap';
    const m = fdoc.createElement('p2w-mark');
    m.style.cssText = 'display:inline-block;width:0;height:0;margin:0;padding:0;border:0;vertical-align:baseline';
    rest.before(keep);
    keep.append(m, rest);
    marks.set(s, m);
  }
  for (const line of lines) line.base = marks.get(first(line)).getBoundingClientRect().bottom;
  return lines;
}

// Lines in the Word writer's shape: points from the page's top left corner.
function wordLines(lines, x0, y0, pt, webFonts) {
  return lines.map(line => {
    const runs = [];
    for (const l of line.letters) {
      const cs = l.cs, kind = familyKind(cs.fontFamily, webFonts), size = parseFloat(cs.fontSize) * pt;
      const bold = parseInt(cs.fontWeight, 10) >= 600, italic = cs.fontStyle !== 'normal', color = cssHex(cs.color).hex, underline = l.underline;
      const key = `${kind}|${bold}|${italic}|${color}|${Math.round(size * 2)}|${underline}`;
      const last = runs[runs.length - 1];
      if (last && last.key === key) last.text += l.text;
      else runs.push({ kind, bold, italic, color, size, key, text: l.text, underline });
    }
    runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\s+$/, '');
    runs[0].text = runs[0].text.replace(/^\s+/, '');
    const kept = runs.filter(r => r.text);
    if (!kept.length) return null;
    const ls = line.letters.filter(l => l.text.trim());
    const left = (ls[0].left - x0) * pt, right = (ls[ls.length - 1].right - x0) * pt, base = (line.base - y0) * pt;
    const size = Math.max(...kept.map(r => r.size));
    const top = Math.min(...kept.map(r => base - FALLBACK[r.kind].ascent * r.size));
    let natural = 0, count = -1;
    for (const r of kept) {
      measure.font = `${r.italic ? 'italic ' : ''}${r.bold ? 'bold ' : ''}${r.size}px "${FALLBACK[r.kind].name}"`;
      natural += measure.measureText(r.text).width;
      count += [...r.text].length;
    }
    const spacing = count > 0 ? Math.max(-0.2 * size, Math.min(0.3 * size, (right - left - natural) / count)) : 0;
    return { link: line.link, left, top, width: right - left, size, spacing, runs: kept };
  }).filter(Boolean);
}

// HTML bytes (or a .zip with the page and its files) in, { docx bytes, number
// of links kept } out, like pdfToDocx.
// How far the copy's top margins push it down inside its box. On the page they
// folded into the root's own margin, which the copy leaves out.
function marginShift(markup, bdoc) {
  // read as XML, as the picture is: the HTML parser would drop a copied <body> tag
  const outer = bdoc.importNode(new DOMParser().parseFromString(markup, 'application/xml').documentElement, true);
  bdoc.body.append(outer);
  const dy = outer.lastElementChild.getBoundingClientRect().top - outer.getBoundingClientRect().top;
  outer.remove();
  return dy;
}

// Where each A4 page of a long web page starts and how much of it to fill:
// up to the last line or picture that fits whole, unless that would leave more
// than a quarter of the page empty.
function pageCuts(fdoc, height, pageH) {
  const spans = [];
  const walker = fdoc.createTreeWalker(fdoc.body, NodeFilter.SHOW_TEXT);
  const range = fdoc.createRange();
  for (let t; (t = walker.nextNode());) {
    if (!t.data.trim()) continue;
    range.selectNodeContents(t);
    for (const r of range.getClientRects()) if (r.height > 0) spans.push([r.top, r.bottom]);
  }
  for (const el of fdoc.body.querySelectorAll('img, svg, video, canvas, input, button, select, textarea, math')) {
    const r = el.getBoundingClientRect();
    if (r.height > 0) spans.push([r.top, r.bottom]);
  }
  const cuts = [];
  for (let y = 0; y < height;) {
    let end = y + pageH;
    if (end < height) {
      const low = y + pageH * 0.75;
      const near = spans.filter(([t, b]) => b > low && t < end);
      const splits = (c) => near.some(([t, b]) => t < c - 0.5 && b > c + 0.5);
      end = [end, ...near.flatMap(([t, b]) => [Math.floor(t), Math.ceil(b)])]
        .filter(c => c > low && c <= end).sort((a, b) => b - a).find(c => !splits(c)) ?? end;
    }
    cuts.push([y, Math.min(end, height) - y]);
    y = end;
  }
  return cuts;
}

async function htmlToDocx(data, onPage = () => {}, options = {}) {
  const src = await readHtml(data);
  const doc = new DOMParser().parseFromString(src.html, 'text/html');
  const res = resources(src);
  const embeds = takeEmbeds(doc);
  scrub(doc);
  onPage(0, 0);
  const css = await inlineAll(doc, res);

  // embeds first, as pictures in their places
  await Promise.all([...doc.querySelectorAll('img[data-p2w-embed]')].map(async img => {
    const probe = document.createElement('div');
    probe.setAttribute('style', img.getAttribute('style') || '');
    const w = parseFloat(probe.style.width) || +img.getAttribute('width') || 300, h = parseFloat(probe.style.height) || +img.getAttribute('height') || 150;
    const html = await runEmbed(embeds[img.dataset.p2wEmbed], w, h);
    img.setAttribute('src', await embedPicture(html, w, h));
  }));

  // the whole page, inert, in a hidden frame
  doc.head.insertAdjacentHTML('afterbegin', `${LOCKED(false)}<meta charset="utf-8">`);
  const style = doc.createElement('style');
  style.textContent = `${css}\n${STILL}\nsection > aside { display: none !important; }`;
  doc.head.append(style);
  const frame = lockFrame('allow-same-origin', 1920, 1080);
  const blank = lockFrame('allow-same-origin', 100, 100);
  const loaded = (f) => new Promise(r => f.addEventListener('load', r, { once: true }));
  const ready = Promise.all([loaded(frame), loaded(blank)]);
  frame.srcdoc = '<!doctype html>' + doc.documentElement.outerHTML;
  // locked too: the margin probe puts the copy in it, with any address the file had
  blank.srcdoc = `<!doctype html><html><head>${LOCKED(false)}</head><body></body></html>`;
  document.body.append(frame, blank);
  try {
    await ready;
    const fdoc = frame.contentDocument, win = frame.contentWindow;
    const differ = styleDiffer(blank.contentDocument);

    // slides, shown one way or another
    const reveal = (el) => {
      for (let e = el; e && e !== fdoc.documentElement; e = e.parentElement) {
        const cs = win.getComputedStyle(e);
        if (cs.display === 'none') e.style.setProperty('display', e.style.display && e.style.display !== 'none' ? e.style.display : 'block', 'important');
        if (cs.visibility !== 'visible') e.style.setProperty('visibility', 'visible', 'important');
        if (+cs.opacity < 1) e.style.setProperty('opacity', '1', 'important');
        if (cs.transform !== 'none') e.style.setProperty('transform', 'none', 'important');
      }
    };
    const all = [...fdoc.body.querySelectorAll('section')].filter(s => !s.parentElement.closest('section'));
    const shown = all.filter(s => !s.hidden);
    let slides = shown.length * 2 >= all.length ? shown : all;
    slides.forEach(reveal);
    for (const s of slides) {
      const r = s.getBoundingClientRect();
      if (r.width < 50 || r.height < 50) { s.style.setProperty('width', '1920px', 'important'); s.style.setProperty('height', '1080px', 'important'); }
    }
    await fdoc.fonts.ready;
    await Promise.all([...fdoc.images].map(i => i.decode().catch(() => {})));
    const sizes = slides.map(s => s.getBoundingClientRect()).map(r => [Math.round(r.width), Math.round(r.height)]);
    const deck = slides.length > 0 && sizes.every(([w, h]) => w === sizes[0][0] && h === sizes[0][1] && w > h);

    // regions to draw: one per slide, or A4 cuts of the whole page
    let regions;
    if (deck) {
      const pt = DECK_PT / sizes[0][0];
      regions = slides.map(s => { const r = s.getBoundingClientRect(); return { root: s, x: r.left, y: r.top, w: r.width, h: r.height, cut: r.height, pt }; });
    } else {
      frame.style.width = A4_PX[0] + 'px';
      await new Promise(r => requestAnimationFrame(r));
      const b = fdoc.documentElement.getBoundingClientRect();
      const height = Math.max(fdoc.documentElement.scrollHeight, fdoc.body.scrollHeight, 1);
      regions = pageCuts(fdoc, height, A4_PX[1]).map(([y, cut]) => ({ root: fdoc.body, x: b.left, y, w: A4_PX[0], h: A4_PX[1], cut, pt: CSS_PT }));
      if (!regions.length) throw new Error('This HTML page is empty.');
    }
    const pageOf = (el) => {
      if (deck) { const i = slides.findIndex(s => s.contains(el)); return i + 1; }
      const t = el.getBoundingClientRect().top, i = regions.findIndex(g => t < g.y + g.cut);
      return i < 0 ? regions.length : i + 1;
    };
    const targetOf = (a) => {
      const href = (a.getAttribute('href') || '').trim();
      if (href.startsWith('#')) {
        let id = href.slice(1);
        try { id = decodeURIComponent(id); } catch {}
        const el = id && fdoc.getElementById(id);
        const page = el ? pageOf(el) : 0;
        return page > 0 ? { page } : null;
      }
      try {
        const u = new URL(href);
        return /^(https?|mailto|tel):$/.test(u.protocol) ? { url: u.href } : null;
      } catch { return null; }
    };

    const webFonts = new Set([...fdoc.fonts].filter(f => f.status === 'loaded').map(f => f.family.replace(/^['"]|['"]$/g, '').toLowerCase()));

    // as in a browser, the body's background fills the page when the root has none
    const background = (el) => { const c = win.getComputedStyle(el).backgroundColor; return cssHex(c).alpha > 0 ? c : ''; };
    const pageBg = deck ? '#fff' : background(fdoc.documentElement) || background(fdoc.body) || '#fff';

    // measure and copy every root once, before the markers change the frame
    const roots = [...new Set(regions.map(r => r.root))];
    const measured = new Map(roots.map(root => {
      const links = measureLinks(root, fdoc, targetOf);
      const copy = styledCopy(root, fdoc, differ, new Set(links.filter(l => l.letters.length).map(l => l.a)));
      const r = root.getBoundingClientRect();
      // a slide clips to its box; a whole page keeps what spills past the body, on the page's own background
      const box = deck ? { left: r.left, top: r.top, width: r.width, height: r.height } : { left: r.left, top: r.top,
        width: Math.max(r.width, fdoc.documentElement.scrollWidth - r.left), height: Math.max(r.height, fdoc.documentElement.scrollHeight - r.top) };
      const wrap = (css) => `<div xmlns="http://www.w3.org/1999/xhtml" style="position:relative;display:flow-root;width:${box.width}px;height:${box.height}px;` +
        `${deck ? 'overflow:hidden;' : ''}"><style>${esc(css)}</style>${new XMLSerializer().serializeToString(copy.el)}</div>`;
      const shift = marginShift(wrap(copy.pseudoCss), blank.contentDocument);
      if (Math.abs(shift) > 0.01) copy.el.setAttribute('style', copy.el.getAttribute('style') + `top:${-shift}px;`);
      return [root, { links, markup: wrap(fontFaces(fdoc, copy.families) + copy.pseudoCss), rect: box }];
    }));
    for (const m of measured.values()) m.lines = letterLines(m.links, fdoc);

    const pages = [];
    for (const [i, g] of regions.entries()) {
      onPage(i + 1, regions.length);
      const m = measured.get(g.root);
      const ox = g.x - m.rect.left, oy = g.y - m.rect.top;
      const scale = Math.min(DPI / 72 * g.pt, MAX_PX / Math.max(g.w, g.h));
      const canvas = await drawMarkup(m.markup, m.rect.width, m.rect.height, [ox, oy, g.w, g.h], scale, pageBg, g.cut);
      const jpeg = await new Promise(r => canvas.toBlob(r, 'image/jpeg', 0.92));
      canvas.width = canvas.height = 0;
      // this page's links, in points from its corner
      const inside = (top, bottom) => (top + bottom) / 2 >= g.y && (top + bottom) / 2 < g.y + g.cut;
      const links = [], index = new Map();
      m.links.forEach((link, li) => {
        const boxes = link.boxes.filter(b => inside(b.top, b.bottom)).map(b => {
          const left = Math.max(0, (b.left - g.x) * g.pt), top = Math.max(0, (b.top - g.y) * g.pt);
          const right = Math.min(g.w * g.pt, (b.right - g.x) * g.pt), bottom = Math.min(g.cut * g.pt, (b.bottom - g.y) * g.pt);
          return [left, top, right - left, bottom - top];
        }).filter(([, , w, h]) => w >= 0.5 && h >= 0.5);
        if (!boxes.length) return;
        index.set(li, links.length);
        links.push({ ...link.target, boxes });
      });
      const lines = wordLines(m.lines.filter(l => index.has(l.link) && inside(l.base - 1, l.base + 1)), g.x, g.y, g.pt, webFonts)
        .map(l => ({ ...l, link: index.get(l.link) }));
      pages.push({ jpeg: await jpeg.arrayBuffer(), width: g.w * g.pt, height: g.h * g.pt, links, lines });
    }
    const out = await pagesToDocx(pages, options);
    return { ...out, missing: res.missing };
  } finally {
    frame.remove();
    blank.remove();
  }
}
