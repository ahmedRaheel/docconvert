(() => {
  'use strict';
  const boot = JSON.parse(document.getElementById('boot').textContent);
  const BASE = (boot.base || '').replace(/\/$/, '');
  const MAX = boot.maxMb * 1024 * 1024;
  const FORMATS = boot.formats;
  const CANON = boot.canon;
  const COLORS = { pdf: '--pdf', docx: '--docx', xlsx: '--xlsx', pptx: '--pptx', html: '--html' };
  const ORDER = ['pdf', 'docx', 'xlsx', 'pptx'];

  const $ = (s) => document.querySelector(s);
  const drop = $('#drop'), input = $('#fileInput'), list = $('#files'), go = $('#go');
  const summary = $('#summary'), live = $('#live'), zipLink = $('#zipLink');
  const state = { files: [], target: 'docx', seq: 0, running: false };

  const canon = (ext) => CANON[ext] || ext;
  const ext = (name) => {
    const m = /\.([a-z0-9]+)$/i.exec(name);
    const e = m ? m[1].toLowerCase() : '';
    return e === 'htm' ? 'html' : e;
  };
  const size = (b) => b < 1024 ? `${b} B` : b < 1048576 ? `${Math.round(b / 1024)} KB` : `${(b / 1048576).toFixed(1)} MB`;
  const plural = (n, w) => `${n} ${w}${n === 1 ? '' : 's'}`;
  const pending = () => state.files.filter((f) => f.status === 'ready' || f.status === 'failed');
  const done = () => state.files.filter((f) => f.status === 'done');
  const color = (e) => `var(${COLORS[canon(e)] || '--ink'})`;

  function el(tag, cls, text) {
    const n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text != null) n.textContent = text;
    return n;
  }

  /* ---------- adding files */
  function addFiles(fileList) {
    for (const f of fileList) {
      const e = ext(f.name);
      const item = { id: ++state.seq, file: f, name: f.name, size: f.size, ext: e, status: 'ready', progress: 0 };
      if (!FORMATS[e]) {
        item.status = 'bad';
        item.error = `${e ? '.' + e : 'This'} file can't be converted. Use PDF, Word, Excel, PowerPoint or HTML.`;
      } else if (f.size > MAX) {
        item.status = 'bad';
        item.error = `Larger than ${boot.maxMb} MB.`;
      }
      state.files.push(item);
    }
    refresh();
  }

  /* ---------- derived state */
  function allowed() {
    const p = pending();
    if (!p.length) return new Set(ORDER);
    let set = null;
    for (const f of p) {
      const t = new Set(FORMATS[f.ext]);
      set = set ? new Set([...set].filter((x) => t.has(x))) : t;
    }
    return set;
  }

  function chooseTarget() {
    const ok = allowed();
    if (ok.has(state.target)) return;
    const p = pending();
    const prefer = p.length && p.every((f) => canon(f.ext) === 'pdf') ? 'docx' : 'pdf';
    state.target = ok.has(prefer) ? prefer : ORDER.find((t) => ok.has(t)) || null;
  }

  /* ---------- rendering */
  const KEEPS = {
    pdf: {
      docx: 'Headings, bullets, bold and italic text, tables and pictures come through as editable Word content.',
      xlsx: 'Tables become real cells with numbers you can calculate with. Pages without tables are split into columns.',
      pptx: 'Each page becomes a slide. Text stays editable, or pick page pictures for an exact copy.',
    },
    pdf_to: {
      docx: 'Fonts, headings, tables, pictures and page size carry over.',
      xlsx: 'Every sheet becomes a page, with column widths and cell shading kept.',
      pptx: 'Each slide becomes a page with its text, pictures and shapes in place.',
      html: 'Headings, links, lists, tables and images are laid out into pages.',
    },
  };
  function keepsText(srcs, t) {
    if (!t) return 'These files can\u2019t all go to one format. Convert them in separate batches.';
    if (srcs.size !== 1) return 'Each file is converted with the settings on this page.';
    const s = [...srcs][0];
    if (s === 'pdf') return KEEPS.pdf[t];
    if (t === 'pdf') return KEEPS.pdf_to[s];
    return 'Content and structure carry over. Layout is approximate because the file passes through PDF.';
  }

  function renderHero() {
    const p = pending();
    const exts = [...new Set(p.map((f) => f.ext))];
    const srcs = new Set(p.map((f) => canon(f.ext)));
    const tokSrc = $('#tokSrc'), tokDst = $('#tokDst');
    let srcColor = '--pdf';
    if (!p.length) { tokSrc.textContent = '.pdf'; }
    else if (exts.length === 1) { tokSrc.textContent = '.' + exts[0]; srcColor = COLORS[canon(exts[0])]; }
    else { tokSrc.textContent = 'files'; srcColor = '--ink'; }
    tokDst.textContent = state.target ? '.' + state.target : 'a format';
    const body = document.body;
    body.style.setProperty('--src', `var(${srcColor})`);
    body.style.setProperty('--dst', state.target ? `var(${COLORS[state.target]})` : 'var(--ink)');
    $('#keeps').textContent = keepsText(p.length ? srcs : new Set(['pdf']), state.target || (p.length ? null : 'docx'));
  }

  function renderTiles() {
    const ok = allowed();
    for (const t of ORDER) {
      const r = $('#to-' + t);
      r.disabled = !ok.has(t);
      r.checked = state.target === t;
    }
  }

  function renderOptions() {
    const hasPdf = pending().some((f) => canon(f.ext) === 'pdf');
    $('#optPages').hidden = !hasPdf;
    $('#optPptx').hidden = !(hasPdf && state.target === 'pptx');
    $('#optXlsx').hidden = !(hasPdf && state.target === 'xlsx');
  }

  function statusText(f) {
    switch (f.status) {
      case 'ready': return size(f.size);
      case 'bad': case 'failed': return f.error;
      case 'uploading': return `Uploading ${Math.round(f.progress * 100)}%`;
      case 'converting': return `Converting to .${f.to}`;
      case 'done': return `Converted to .${f.to}, ${size(f.result.size)}`;
      default: return '';
    }
  }

  function renderFiles() {
    list.replaceChildren();
    for (const f of state.files) {
      const li = el('li', `file is-${f.status}`);
      li.dataset.id = f.id;
      if (f.to) li.style.setProperty('--to', `var(${COLORS[f.to]})`);
      li.style.setProperty('--p', f.progress);

      const badge = el('span', 'kind', f.ext || '?');
      badge.style.background = color(f.ext);
      const info = el('div', 'min-w-0');
      info.append(el('div', 'file-name', f.name), el('div', 'file-meta', statusText(f)));

      const act = el('div', 'file-act');
      if (f.status === 'done') {
        const a = el('a', 'dl', 'Download');
        a.href = `${BASE}/api/download/${f.result.id}`;
        a.setAttribute('aria-label', `Download ${f.result.name}`);
        act.append(a);
      }
      const rm = el('button', 'icon-btn');
      rm.type = 'button';
      const busy = f.status === 'uploading' || f.status === 'converting';
      rm.setAttribute('aria-label', `${busy ? 'Cancel' : 'Remove'} ${f.name}`);
      rm.innerHTML = '<svg width="14" height="14" viewBox="0 0 14 14" aria-hidden="true"><path d="M2 2l10 10M12 2L2 12" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>';
      rm.addEventListener('click', () => removeFile(f.id));
      act.append(rm);

      li.append(badge, info, act, el('span', 'bar'));
      list.append(li);
    }
  }

  function renderActions() {
    const p = pending();
    const bytes = p.reduce((n, f) => n + f.size, 0);
    go.disabled = state.running || !p.length || !state.target;
    go.textContent = state.running ? 'Converting\u2026' : p.length ? `Convert ${plural(p.length, 'file')}` : 'Convert';
    if (state.running) summary.textContent = 'Working on it. You can keep this tab open.';
    else if (p.length && !state.target) summary.textContent = 'These files can\u2019t all go to one format.';
    else if (p.length) summary.textContent = `${plural(p.length, 'file')} ready, ${size(bytes)}`;
    else if (done().length) summary.textContent = `${plural(done().length, 'file')} converted.`;
    else summary.textContent = 'Add a file to begin.';
    const d = done();
    zipLink.hidden = d.length < 2;
    if (d.length >= 2) {
      zipLink.href = `${BASE}/api/download-all?ids=${d.map((f) => f.result.id).join(',')}`;
      zipLink.textContent = `Download all ${d.length} as zip`;
    }
  }

  function refresh() {
    chooseTarget();
    renderFiles();
    renderTiles();
    renderHero();
    renderOptions();
    renderActions();
  }

  function removeFile(id) {
    const f = state.files.find((x) => x.id === id);
    if (!f) return;
    if (f.xhr) f.xhr.abort();
    state.files = state.files.filter((x) => x.id !== id);
    refresh();
  }

  /* ---------- converting */
  function collectOptions() {
    const o = {};
    const pages = $('#pages').value.trim();
    if (pages && !$('#optPages').hidden) o.pages = pages;
    if (!$('#optPptx').hidden) o.mode = document.querySelector('input[name=mode]:checked').value;
    if (!$('#optXlsx').hidden) o.sheet_mode = document.querySelector('input[name=sheet_mode]:checked').value;
    const eng = $('#engine').value;
    if (eng !== 'auto') o.engine = eng;
    return o;
  }

  function paintProgress(f) {
    const row = list.querySelector(`[data-id="${f.id}"]`);
    if (!row) return;
    row.style.setProperty('--p', f.progress);
    row.querySelector('.file-meta').textContent = statusText(f);
  }

  function send(f, to, options) {
    return new Promise((resolve) => {
      f.status = 'uploading'; f.progress = 0; f.to = to; f.error = '';
      renderFiles();
      const fd = new FormData();
      fd.append('file', f.file, f.name);
      fd.append('to', to);
      fd.append('options', JSON.stringify(options));
      const xhr = new XMLHttpRequest();
      f.xhr = xhr;
      xhr.open('POST', `${BASE}/api/convert`);
      xhr.upload.onprogress = (e) => { if (e.lengthComputable) { f.progress = e.loaded / e.total; paintProgress(f); } };
      xhr.upload.onload = () => { f.status = 'converting'; renderFiles(); };
      xhr.onload = () => {
        let data = {};
        try { data = JSON.parse(xhr.responseText); } catch (_) { /* non-JSON error page */ }
        f.xhr = null;
        if (xhr.status === 200 && data.id) {
          f.status = 'done'; f.result = data; f.progress = 1;
          live.textContent = `${f.name} converted to .${to}.`;
        } else {
          f.status = 'failed';
          f.error = data.error || `The server couldn\u2019t finish this (error ${xhr.status}). Try again.`;
        }
        renderFiles(); resolve();
      };
      xhr.onerror = () => {
        f.xhr = null; f.status = 'failed';
        f.error = 'Couldn\u2019t reach the server. Check your connection and try again.';
        renderFiles(); resolve();
      };
      xhr.onabort = () => resolve();
      xhr.send(fd);
    });
  }

  async function run() {
    if (state.running || !state.target) return;
    const queue = pending();
    if (!queue.length) return;
    const to = state.target, options = collectOptions();
    state.running = true; renderActions();
    const worker = async () => { while (queue.length) await send(queue.shift(), to, options); };
    await Promise.all([worker(), worker()]);
    state.running = false;
    refresh();
  }

  /* ---------- events */
  go.addEventListener('click', run);
  document.querySelectorAll('input[name=to]').forEach((r) => r.addEventListener('change', () => {
    state.target = r.value; refresh();
  }));
  drop.addEventListener('click', () => input.click());
  drop.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
  });
  input.addEventListener('change', () => { if (input.files.length) addFiles(input.files); input.value = ''; });

  const hasFiles = (e) => e.dataTransfer && [...e.dataTransfer.types].includes('Files');
  let depth = 0;
  window.addEventListener('dragenter', (e) => { if (hasFiles(e)) { depth++; drop.classList.add('is-over'); } });
  window.addEventListener('dragover', (e) => { if (hasFiles(e)) e.preventDefault(); });
  window.addEventListener('dragleave', () => { depth = Math.max(0, depth - 1); if (!depth) drop.classList.remove('is-over'); });
  window.addEventListener('drop', (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault(); depth = 0; drop.classList.remove('is-over');
    if (e.dataTransfer.files.length) addFiles(e.dataTransfer.files);
  });

  refresh();
})();
