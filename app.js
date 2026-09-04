  pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const statusSection = document.getElementById('statusSection');
  const statusText = document.getElementById('statusText');
  const statusPct = document.getElementById('statusPct');
  const barFill = document.getElementById('barFill');
  const fileList = document.getElementById('fileList');
  const sensitivitySection = document.getElementById('sensitivitySection');
  const sensSlider = document.getElementById('sensSlider');
  const sensValueLabel = document.getElementById('sensValueLabel');
  const resultSection = document.getElementById('resultSection');
  const dataBody = document.getElementById('dataBody');
  const rowCount = document.getElementById('rowCount');
  const copyBtn = document.getElementById('copyBtn');
  const downloadBtn = document.getElementById('downloadBtn');
  const clearBtn = document.getElementById('clearBtn');
  const copiedFlag = document.getElementById('copiedFlag');
  const warningArea = document.getElementById('warningArea');

  let worker = null;
  let processing = false;
  // Each item: { group: 'archivo.pdf · pág 1', nombre, rut, inkRatio, manual: null|'si'|'no' }
  let allRows = [];
  // Pages where the header (Nombre/RUT/Firma) could not be located automatically
  let unrecognizedPages = [];

  dropzone.addEventListener('click', () => fileInput.click());
  dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('drag'); });
  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('drag'));
  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('drag');
    if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
  });
  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) handleFiles(e.target.files);
  });

  clearBtn.addEventListener('click', () => {
    allRows = [];
    unrecognizedPages = [];
    renderTable();
  });

  copyBtn.addEventListener('click', async () => {
    const lines = ['Nombre\tRUT\tFirma'];
    allRows.forEach(r => {
      lines.push(`${r.nombre}\t${r.rut}\t${firmaLabel(r)}`);
    });
    const text = lines.join('\n');
    try {
      await navigator.clipboard.writeText(text);
    } catch (err) {
      const ta = document.createElement('textarea');
      ta.value = text;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      document.body.removeChild(ta);
    }
    copiedFlag.classList.add('show');
    setTimeout(() => copiedFlag.classList.remove('show'), 1600);
  });

  downloadBtn.addEventListener('click', () => {
    if (!allRows.length) return;
    const aoa = [['Nombre', 'RUT', 'Firma']];
    allRows.forEach(r => aoa.push([r.nombre || '', r.rut || '', firmaLabel(r)]));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 32 }, { wch: 14 }, { wch: 8 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Asistencia');
    const stamp = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `asistencia_${stamp}.xlsx`);
  });

  sensSlider.addEventListener('input', () => {
    const v = Number(sensSlider.value);
    sensValueLabel.textContent = v < 33 ? 'Baja' : (v < 66 ? 'Media' : 'Alta');
    renderTable();
  });

  function currentThreshold(){
    // slider 0 -> requires strong ink (0.035), slider 100 -> very sensitive (0.002)
    const v = Number(sensSlider.value);
    const min = 0.002, max = 0.035;
    return max - (v / 100) * (max - min);
  }

  function firmaLabel(row){
    if (row.manual) return row.manual === 'si' ? 'Sí' : 'No';
    return row.inkRatio >= currentThreshold() ? 'Sí' : 'No';
  }

  function setStatus(text, pct){
    statusText.textContent = text;
    statusPct.textContent = Math.round(pct) + '%';
    barFill.style.width = Math.min(100, Math.max(0, pct)) + '%';
  }

  function addFileRow(name){
    const li = document.createElement('li');
    li.dataset.name = name;
    li.innerHTML = `<span>${name}</span><span class="state">En cola</span>`;
    fileList.appendChild(li);
    return li;
  }
  function setRowState(row, state, className){
    row.querySelector('.state').textContent = state;
    row.className = className || '';
  }

  async function getWorker(){
    if (worker) return worker;
    worker = await Tesseract.createWorker('spa', 1, { logger: () => {} });
    return worker;
  }

  function normalize(s){
    return (s || '')
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/[.\s°]/g, '')
      .toUpperCase();
  }

  function flattenLines(blocks){
    const lines = [];
    (blocks || []).forEach(block => {
      (block.paragraphs || []).forEach(par => {
        (par.lines || []).forEach(line => lines.push(line));
      });
    });
    return lines;
  }

  // Chilean RUT, with or without dots: 12.345.678-9 / 12345678-9 / 1.234.567-K
  const RUT_REGEX = /\d{1,2}\.?\d{3}\.?\d{3}-[\dkK]/;

  // Upscales small/blurry images and boosts contrast before OCR — helps a lot with
  // real phone photos (uneven lighting, small resolution) at basically no cost for
  // already-sharp PDF pages.
  function preprocessCanvas(srcCanvas){
    const targetWidth = 2200;
    const scale = srcCanvas.width < targetWidth ? targetWidth / srcCanvas.width : 1;
    const w = Math.round(srcCanvas.width * scale);
    const h = Math.round(srcCanvas.height * scale);
    const canvas = document.createElement('canvas');
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext('2d');
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(srcCanvas, 0, 0, w, h);

    const imgData = ctx.getImageData(0, 0, w, h);
    const d = imgData.data;
    const total = w * h;
    const lum = new Float32Array(total);
    let min = 255, max = 0;
    for (let i = 0, p = 0; i < d.length; i += 4, p++){
      const l = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
      lum[p] = l;
      if (l < min) min = l;
      if (l > max) max = l;
    }
    const range = Math.max(1, max - min);
    for (let i = 0, p = 0; i < d.length; i += 4, p++){
      const v = Math.max(0, Math.min(255, Math.round((lum[p] - min) * 255 / range)));
      d[i] = d[i + 1] = d[i + 2] = v;
    }
    ctx.putImageData(imgData, 0, 0);
    return canvas;
  }

  function findRutWord(words){
    const sorted = words.slice().sort((a, b) => a.bbox.x0 - b.bbox.x0);
    for (const w of sorted){
      const m = (w.text || '').match(RUT_REGEX);
      if (m) return { word: w, match: m[0] };
    }
    return null;
  }

  function buildNombre(words, rutWord){
    const before = words
      .filter(w => w.bbox.x1 <= rutWord.bbox.x0)
      .sort((a, b) => a.bbox.x0 - b.bbox.x0);
    // drop a leading row-number token like "12" or "N°" before the actual name
    if (before.length > 1 && /^\d{1,3}\.?°?$/.test(before[0].text.trim())){
      before.shift();
    }
    return before.map(w => w.text.trim()).filter(Boolean).join(' ');
  }

  function inkRatioInRegion(canvas, x0, y0, x1, y1){
    x0 = Math.max(0, Math.floor(x0));
    y0 = Math.max(0, Math.floor(y0));
    x1 = Math.min(canvas.width, Math.ceil(x1));
    y1 = Math.min(canvas.height, Math.ceil(y1));
    const w = Math.max(1, x1 - x0), h = Math.max(1, y1 - y0);
    if (w <= 1 || h <= 1) return 0;
    const ctx = canvas.getContext('2d');
    let data;
    try {
      data = ctx.getImageData(x0, y0, w, h).data;
    } catch (e) {
      return 0;
    }
    let dark = 0;
    const total = w * h;
    for (let i = 0; i < data.length; i += 4){
      const lum = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
      if (lum < 170) dark++;
    }
    return dark / total;
  }

  async function extractRowsFromCanvas(rawCanvas, groupLabel){
    const canvas = preprocessCanvas(rawCanvas);
    const { data } = await worker.recognize(canvas, {}, { blocks: true });
    const lines = flattenLines(data.blocks);
    const rightMargin = canvas.width * 0.03;
    let foundAny = false;

    lines.forEach(line => {
      const words = (line.words || []);
      const found = findRutWord(words);
      if (!found) return; // no RUT on this line -> not a data row (header, footer, etc.)
      foundAny = true;
      const nombre = buildNombre(words, found.word);
      const rut = found.match;
      const padY = (line.bbox.y1 - line.bbox.y0) * 0.15;
      const firmaX0 = found.word.bbox.x1 + (line.bbox.y1 - line.bbox.y0) * 0.3;
      const inkRatio = inkRatioInRegion(
        canvas,
        firmaX0, line.bbox.y0 - padY,
        canvas.width - rightMargin, line.bbox.y1 + padY
      );
      allRows.push({ group: groupLabel, nombre, rut, inkRatio, manual: null });
    });

    if (!foundAny){
      unrecognizedPages.push({ label: groupLabel, text: (data.text || '').trim() });
    }
  }

  async function processPdf(file, row){
    const buffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
    for (let i = 1; i <= pdf.numPages; i++){
      setRowState(row, `Página ${i}/${pdf.numPages}…`);
      const page = await pdf.getPage(i);
      const viewport = page.getViewport({ scale: 2.5 });
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      const ctx = canvas.getContext('2d');
      await page.render({ canvasContext: ctx, viewport }).promise;
      await extractRowsFromCanvas(canvas, `${file.name} · pág ${i}`);
    }
  }

  function fileToCanvas(file){
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        canvas.width = img.naturalWidth;
        canvas.height = img.naturalHeight;
        canvas.getContext('2d').drawImage(img, 0, 0);
        URL.revokeObjectURL(img.src);
        resolve(canvas);
      };
      img.onerror = reject;
      img.src = URL.createObjectURL(file);
    });
  }

  async function processImage(file, row){
    setRowState(row, 'Reconociendo…');
    const canvas = await fileToCanvas(file);
    await extractRowsFromCanvas(canvas, file.name);
  }

  function renderTable(){
    dataBody.innerHTML = '';
    let lastGroup = null;
    const showGroups = new Set(allRows.map(r => r.group)).size > 1;

    allRows.forEach((row, idx) => {
      if (showGroups && row.group !== lastGroup){
        const gtr = document.createElement('tr');
        gtr.className = 'group-row';
        gtr.innerHTML = `<td colspan="4">${row.group}</td>`;
        dataBody.appendChild(gtr);
        lastGroup = row.group;
      }
      const tr = document.createElement('tr');
      const label = firmaLabel(row);
      tr.innerHTML = `
        <td>${row.nombre || ''}</td>
        <td class="rut">${row.rut || ''}</td>
        <td>
          <select class="firma-select ${label === 'Sí' ? 'si' : 'no'}" data-idx="${idx}">
            <option value="si" ${label === 'Sí' ? 'selected' : ''}>Sí</option>
            <option value="no" ${label === 'No' ? 'selected' : ''}>No</option>
          </select>
        </td>
        <td><button class="row-remove" data-idx="${idx}" title="Quitar fila">✕</button></td>`;
      dataBody.appendChild(tr);
    });

    dataBody.querySelectorAll('select.firma-select').forEach(sel => {
      sel.addEventListener('change', (e) => {
        const idx = Number(e.target.dataset.idx);
        allRows[idx].manual = e.target.value;
        e.target.className = 'firma-select ' + e.target.value;
      });
    });

    dataBody.querySelectorAll('.row-remove').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const idx = Number(e.target.dataset.idx);
        allRows.splice(idx, 1);
        renderTable();
      });
    });

    rowCount.textContent = allRows.length ? `· ${allRows.length} fila(s)` : '';
    resultSection.classList.toggle('visible', allRows.length > 0 || unrecognizedPages.length > 0);
    sensitivitySection.classList.toggle('visible', allRows.length > 0);

    warningArea.innerHTML = '';
    if (unrecognizedPages.length){
      const banner = document.createElement('div');
      banner.className = 'warning-banner';
      banner.textContent = `No se pudo detectar automáticamente el encabezado (Nombre / RUT / Firma) en ${unrecognizedPages.length} página(s). Puedes revisar el texto tal como lo leyó el OCR abajo y pasarlo a mano.`;
      warningArea.appendChild(banner);

      unrecognizedPages.forEach(p => {
        const det = document.createElement('details');
        det.className = 'raw-fallback';
        const sum = document.createElement('summary');
        sum.textContent = `Ver texto sin procesar — ${p.label}`;
        const ta = document.createElement('textarea');
        ta.value = p.text;
        ta.readOnly = false;
        det.appendChild(sum);
        det.appendChild(ta);
        warningArea.appendChild(det);
      });
    }
  }

  async function handleFiles(fileListRaw){
    if (processing) return;
    const files = Array.from(fileListRaw);
    processing = true;
    fileList.innerHTML = '';
    statusSection.classList.add('visible');
    setStatus('Preparando…', 2);
    await getWorker();

    const rows = files.map(f => addFileRow(f.name));

    for (let idx = 0; idx < files.length; idx++){
      const file = files[idx];
      const row = rows[idx];
      setRowState(row, 'Procesando…');
      setStatus(`Procesando ${file.name} (${idx + 1} de ${files.length})`, (idx / files.length) * 100);
      try {
        if (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')) {
          await processPdf(file, row);
        } else {
          await processImage(file, row);
        }
        setRowState(row, 'Listo', 'done');
      } catch (err) {
        console.error(err);
        setRowState(row, 'Error al leer', 'error');
      }
      setStatus(`Procesando (${idx + 1} de ${files.length})`, ((idx + 1) / files.length) * 100);
      renderTable();
    }

    setStatus('Listo', 100);
    processing = false;
    fileInput.value = '';
    renderTable();
  }
