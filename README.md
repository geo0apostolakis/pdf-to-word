# PDF to Word

Drop a PDF (or an HTML page, or a slide deck saved as HTML) and get a `.docx`
where every page looks exactly like the original. Each page becomes a picture
on its own Word page of the same size, so nothing reflows, and the links keep
working on top of it.

**Live:** https://pdf-to-word-pages.vercel.app · or download
[`dist/pdf-to-word.html`](dist/pdf-to-word.html) and open it. It is one file,
it works offline, and nothing you drop on it leaves your browser.

## Why it exists

Converters that "turn a PDF into an editable Word file" rebuild the layout and
get it wrong: fonts shift, columns break, a designed newsletter turns into a
mess. Sometimes you do not want editable. You want the same pages, inside a
Word file, with the links still clickable. That is what this does.

## What it does

- **Pages as pictures, at the original size.** Each PDF page is drawn at
  200 DPI by pdf.js and placed on a Word page with the same dimensions.
- **Links survive.** Every link box gets a clear picture over the same spot that
  carries the link: web, mail, or a jump to another page (as a bookmark).
- **Link words are real text.** The words of a link are cut out of the page
  picture and written as real text in a see-through text box at the same spot,
  with letter spacing tuned so the line keeps the PDF's width.
- **HTML and decks too.** A `.html` (or a `.zip` with its pictures and fonts) is
  laid out in a hidden frame and drawn the same way. A deck of same-size
  landscape slides becomes one Word page per slide; anything else becomes A4.
- **Optional links page** at the end, listing every link as plain text.

## The hard parts

- **Apple's preview disagrees with Word.** Mail on Mac and iPhone adds about
  5.4 pt under each page and places floating objects from the top of their
  section. So every page is its own section, the in-line picture is 5.4 pt
  shorter, and a full-size copy floats on top at the page corner (same image,
  no extra bytes). Word and Apple then line up exactly.
- **Mixed page sizes.** Apple offsets page n's floating objects by n - 1 times
  that page's height, which is only right while every page so far has the same
  size. From the first page of another size on, pages switch to a full-size
  in-line picture.
- **Untrusted HTML.** Nothing in a dropped file runs in the app. Scripts,
  handlers, `javascript:` links, `<base>`, meta refresh, objects and frames are
  removed, and both hidden frames get no scripts and a policy that allows only
  `data:` pictures and fonts. Embeds run alone in a sandbox with no network and
  come back as pictures.
- **Margins that disappear.** A copied page root loses its own margin, so top
  margins that collapsed into it are measured in a blank frame and undone.
  Without this, a plain page's picture sat about 22 px low.

## Build

The live page is a single self-contained HTML file. The source is split for
reading:

```
src/index.html         page and styles
src/converter.js       PDF to .docx: drawing, link boxes, link text, Word XML
src/html-converter.js  HTML and decks: sanitising, layout, slicing into pages
src/ui.js              drop zone and file list
vendor/                pdf.js 3.11.174 and JSZip 3.10.1, unmodified
```

```bash
python3 scripts/build.py   # writes dist/pdf-to-word.html
```

## Known limits

- PPTX is not supported: a browser cannot draw it faithfully. Export it as PDF
  or HTML first.
- Apple's preview ignores fonts embedded in a `.docx` and ignores letter
  spacing, so link text there shows in a fallback font, slightly wider or
  narrower than the gap left for it.
- Gmail's viewer, Google Docs, Outlook and Word on the web are untested.

## License

MIT for this project. Bundled: [pdf.js](https://github.com/mozilla/pdf.js)
(Apache 2.0) and [JSZip](https://github.com/Stuk/jszip) (MIT or GPLv3), with
their license headers intact in `vendor/`.
