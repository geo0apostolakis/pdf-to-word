
const list = document.getElementById('list');
const pick = document.getElementById('pick');

pick.addEventListener('change', () => { convertAll([...pick.files]); pick.value = ''; });

// The whole window is a drop target, so a slightly missed drop still works.
let depth = 0;
addEventListener('dragenter', e => { e.preventDefault(); depth++; document.body.classList.add('over'); });
addEventListener('dragleave', () => { if (--depth <= 0) { depth = 0; document.body.classList.remove('over'); } });
addEventListener('dragover', e => e.preventDefault());
addEventListener('drop', e => {
  e.preventDefault();
  depth = 0;
  document.body.classList.remove('over');
  convertAll([...e.dataTransfer.files]);
});

// pdf.js and its worker are the last scripts on the page, so a file picked
// while the page is still loading waits for them.
const scriptsLoaded = new Promise(r => document.readyState === 'loading' ? addEventListener('DOMContentLoaded', r) : r());

async function convertAll(files) {
  await scriptsLoaded;
  for (const file of files) {
    const li = document.createElement('li');
    li.innerHTML = '<span class="name"></span><span class="status">Opening…</span>';
    li.querySelector('.name').textContent = file.name;
    const status = li.querySelector('.status');
    list.prepend(li);
    try {
      if (/\.(pptx?|key)$/i.test(file.name)) throw new Error('PowerPoint files can\'t be drawn here to look the same. Save it as PDF (File > Export > PDF), or for a Claude Slides deck download it as HTML, then drop that here.');
      const html = /\.(html?|zip)$/i.test(file.name);
      if (!html && !/\.pdf$/i.test(file.name)) throw new Error('Not a PDF or HTML file.');
      const progress = (n, total) => { status.textContent = total ? `${html ? 'Slide' : 'Page'} ${n} of ${total}…` : 'Loading fonts and pictures…'; };
      const { docx, links, missing } = await (html ? htmlToDocx : pdfToDocx)(new Uint8Array(await file.arrayBuffer()), progress,
        { linkPage: document.getElementById('link-page').checked });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([docx], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
      a.download = file.name.replace(/\.[^.]+$/, '') + '.docx';
      a.textContent = 'download again';
      a.click();
      li.className = 'ok';
      status.textContent = (links ? `Done, ${links} link${links > 1 ? 's' : ''} kept` : 'Done') +
        (missing ? `, ${missing} picture${missing > 1 ? 's' : ''} or font${missing > 1 ? 's' : ''} not found` : '') + ' · ';
      status.append(a);
    } catch (err) {
      console.error(err);
      li.className = 'err';
      status.textContent = err.message;
    }
  }
}
