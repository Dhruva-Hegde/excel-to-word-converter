/* ============================================================
   XL2DOC — app.js
   Excel → Word converter using SheetJS + docx.js
   ============================================================ */

(function () {
  'use strict';

  // Guard: ensure CDN libraries loaded before running
  if (typeof XLSX === 'undefined' || typeof docx === 'undefined') {
    document.addEventListener('DOMContentLoaded', () => {
      const t = document.getElementById('toast');
      if (t) {
        t.textContent = '⚠️ Libraries failed to load. Check your internet and refresh.';
        t.className = 'toast error show';
      }
    });
    console.error('[XL2DOC] Missing libs — XLSX:', typeof XLSX, '| docx:', typeof docx);
    return;
  }


  let workbook = null;        // SheetJS workbook
  let sheetNames = [];        // all sheet names from file
  let docxBlob = null;        // generated docx blob
  let currentPreviewSheet = 0;
  let previewSheets = [];     // processed sheet data for preview

  // ── DOM refs ───────────────────────────────────────────────
  const html          = document.documentElement;
  const themeToggle   = document.getElementById('themeToggle');
  const dropZone      = document.getElementById('dropZone');
  const fileInput     = document.getElementById('fileInput');
  const browseBtn     = document.getElementById('browseBtn');
  const fileInfo      = document.getElementById('fileInfo');
  const fileName      = document.getElementById('fileName');
  const fileSize      = document.getElementById('fileSize');
  const removeFile    = document.getElementById('removeFile');
  const sheetList     = document.getElementById('sheetList');
  const convertBtn    = document.getElementById('convertBtn');
  const stepUpload    = document.getElementById('stepUpload');
  const stepConfigure = document.getElementById('stepConfigure');
  const stepResult    = document.getElementById('stepResult');
  const previewContent    = document.getElementById('previewContent');
  const previewSheetCount = document.getElementById('previewSheetCount');
  const previewRowCount   = document.getElementById('previewRowCount');
  const sheetIndicator    = document.getElementById('sheetIndicator');
  const prevSheetBtn  = document.getElementById('prevSheet');
  const nextSheetBtn  = document.getElementById('nextSheet');
  const downloadBtn   = document.getElementById('downloadBtn');
  const outputFileName= document.getElementById('outputFileName');
  const outputStats   = document.getElementById('outputStats');
  const startOver     = document.getElementById('startOver');
  const progressOverlay = document.getElementById('progressOverlay');
  const progressText  = document.getElementById('progressText');
  const progressBar   = document.getElementById('progressBar');
  const toast         = document.getElementById('toast');

  // ── Theme ──────────────────────────────────────────────────
  const savedTheme = localStorage.getItem('xl2doc-theme') || 'dark';
  html.setAttribute('data-theme', savedTheme);

  themeToggle.addEventListener('click', () => {
    const next = html.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
    html.setAttribute('data-theme', next);
    localStorage.setItem('xl2doc-theme', next);
  });

  // ── File Upload ────────────────────────────────────────────
  browseBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    fileInput.click();
  });

  dropZone.addEventListener('click', () => fileInput.click());

  dropZone.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') fileInput.click();
  });

  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });

  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));

  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });

  fileInput.addEventListener('change', () => {
    if (fileInput.files[0]) handleFile(fileInput.files[0]);
  });

  removeFile.addEventListener('click', () => resetFile());

  function handleFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['xlsx', 'xls'].includes(ext)) {
      showToast('❌ Please upload a .xlsx or .xls file', 'error');
      return;
    }

    // Show file info
    fileName.textContent = file.name;
    fileSize.textContent = formatBytes(file.size);
    fileInfo.classList.remove('hidden');
    dropZone.style.display = 'none';

    // Parse with SheetJS
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        workbook = XLSX.read(ev.target.result, { type: 'array', cellStyles: true, cellDates: true });
        sheetNames = workbook.SheetNames;
        renderSheetList();
        activateStep(stepConfigure);
        convertBtn.disabled = false;
        showToast(`✅ File loaded — ${sheetNames.length} sheet(s) found`, 'success');
      } catch (err) {
        showToast('❌ Failed to parse Excel file: ' + err.message, 'error');
        resetFile();
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function resetFile() {
    workbook = null;
    sheetNames = [];
    docxBlob = null;
    previewSheets = [];
    fileInput.value = '';
    fileInfo.classList.add('hidden');
    dropZone.style.display = '';
    sheetList.innerHTML = '<p class="placeholder-text">Upload a file to see sheets</p>';
    convertBtn.disabled = true;
    deactivateStep(stepConfigure);
    deactivateStep(stepResult);
    activateStep(stepUpload);
  }

  // ── Sheet list ─────────────────────────────────────────────
  function renderSheetList() {
    sheetList.innerHTML = '';
    sheetNames.forEach((name, i) => {
      const item = document.createElement('label');
      item.className = 'sheet-item';
      item.innerHTML = `
        <input type="checkbox" class="sheet-check" data-index="${i}" checked />
        <span class="sheet-icon">📋</span>
        <span>${name}</span>
      `;
      sheetList.appendChild(item);
    });
  }

  function getSelectedSheets() {
    const checks = document.querySelectorAll('.sheet-check');
    const selected = [];
    checks.forEach((cb) => {
      if (cb.checked) selected.push(sheetNames[Number(cb.dataset.index)]);
    });
    return selected;
  }

  // ── Step management ────────────────────────────────────────
  function activateStep(step) { step.classList.add('active'); }
  function deactivateStep(step) { step.classList.remove('active'); }

  // ── Convert ────────────────────────────────────────────────
  convertBtn.addEventListener('click', async () => {
    const selected = getSelectedSheets();
    if (selected.length === 0) {
      showToast('⚠️ Select at least one sheet', 'error');
      return;
    }
    showProgress('Parsing sheets…', 10);
    await sleep(50);

    // Collect config
    const config = {
      docTitle:    document.getElementById('docTitle').value.trim()   || 'Excel Report',
      docAuthor:   document.getElementById('docAuthor').value.trim()  || 'XL2DOC',
      docSubject:  document.getElementById('docSubject').value.trim() || 'Generated from Excel',
      headerText:  document.getElementById('headerText').value.trim() || '',
      footerText:  document.getElementById('footerText').value.trim() || '',
      includeDate: document.getElementById('includeDate').checked,
      tableStyle:  document.querySelector('input[name="tableStyle"]:checked')?.value || 'blue',
      pageBreaks:  document.getElementById('pageBreaks').checked,
      orientation: document.querySelector('input[name="orientation"]:checked')?.value || 'portrait',
      fontSize:    parseInt(document.getElementById('fontSize').value, 10) || 11,
      boldHeaders: document.getElementById('boldHeaders').checked,
      includeSheetName: document.getElementById('includeSheetName').checked,
      stripEmpty:  document.getElementById('stripEmpty').checked,
      includeSummary: document.getElementById('includeSummary').checked,
      selectedSheets: selected,
    };

    try {
      // Process sheets into data
      showProgress('Reading data…', 25);
      await sleep(50);
      previewSheets = processSheets(config);

      showProgress('Building Word document…', 55);
      await sleep(50);
      docxBlob = await buildDocx(config, previewSheets);

      showProgress('Finalizing…', 90);
      await sleep(200);

      hideProgress();
      renderPreview(config);
      activateStep(stepResult);
      showToast('🎉 Document ready!', 'success');
    } catch (err) {
      hideProgress();
      showToast('❌ Conversion failed: ' + err.message, 'error');
      console.error(err);
    }
  });

  // ── Sheet Processing ───────────────────────────────────────
  function processSheets(config) {
    const results = [];
    config.selectedSheets.forEach((name) => {
      const ws = workbook.Sheets[name];
      let data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false, dateNF: 'dd/mm/yyyy' });

      if (config.stripEmpty) {
        // Remove rows where every cell is empty
        data = data.filter((row) => row.some((cell) => cell !== '' && cell !== null && cell !== undefined));
        // Trim trailing empty columns
        let maxCols = 0;
        data.forEach((row) => {
          let last = 0;
          row.forEach((cell, ci) => { if (cell !== '' && cell !== null && cell !== undefined) last = ci; });
          if (last + 1 > maxCols) maxCols = last + 1;
        });
        data = data.map((row) => row.slice(0, maxCols));
      }

      results.push({ name, data });
    });
    return results;
  }

  // ── DOCX Builder ───────────────────────────────────────────
  async function buildDocx(config, sheets) {
    const { Document, Paragraph, Table, TableRow, TableCell,
            TextRun, Header, Footer, PageBreak, AlignmentType,
            HeadingLevel, WidthType, BorderStyle,
            convertInchesToTwip } = docx;

    // Color palettes
    const palettes = {
      blue:   { header: '1E3A8A', headerText: 'FFFFFF', alt: 'DBEAFE', border: '93C5FD', title: '1E40AF' },
      green:  { header: '14532D', headerText: 'FFFFFF', alt: 'DCFCE7', border: '86EFAC', title: '166534' },
      purple: { header: '4C1D95', headerText: 'FFFFFF', alt: 'F3E8FF', border: 'D8B4FE', title: '6B21A8' },
      gray:   { header: '1F2937', headerText: 'FFFFFF', alt: 'F3F4F6', border: 'D1D5DB', title: '374151' },
    };
    const pal = palettes[config.tableStyle] || palettes.blue;
    const fontSizePt = config.fontSize * 2; // half-points

    const dateStr = new Date().toLocaleDateString('en-GB', {
      day: '2-digit', month: 'long', year: 'numeric',
    });

    // Header
    const headerContent = new Header({
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: config.headerText || config.docTitle,
              bold: true, size: 20, color: pal.title,
            }),
            ...(config.includeDate ? [
              new TextRun({ text: `    |    ${dateStr}`, size: 18, color: '6B7280' }),
            ] : []),
          ],
          alignment: AlignmentType.LEFT,
          border: { bottom: { color: pal.header, size: 6, style: BorderStyle.SINGLE } },
        }),
      ],
    });

    // Footer
    const footerContent = new Footer({
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: config.footerText
                ? config.footerText.replace('{n}', 'Page')
                : 'Generated by XL2DOC  ·  ' + new Date().toLocaleDateString('en-GB'),
              size: 18, color: '9CA3AF',
            }),
          ],
          alignment: AlignmentType.CENTER,
          border: { top: { color: pal.border, size: 4, style: BorderStyle.SINGLE } },
        }),
      ],
    });

    const children = [];

    // ── Summary Page ──────────────────────────────────────────
    if (config.includeSummary) {
      children.push(
        new Paragraph({
          children: [new TextRun({ text: config.docTitle, bold: true, size: 56, color: pal.title })],
          alignment: AlignmentType.CENTER,
          spacing: { before: 800, after: 200 },
        }),
        new Paragraph({
          children: [new TextRun({ text: config.docSubject || 'Excel Report', size: 26, color: '6B7280' })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [new TextRun({ text: `Author: ${config.docAuthor}`, size: 22, color: '9CA3AF' })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
        }),
        new Paragraph({
          children: [new TextRun({ text: `Date: ${dateStr}`, size: 22, color: '9CA3AF' })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 600 },
        }),
        // Sheets summary
        new Paragraph({
          children: [new TextRun({ text: 'Contents', bold: true, size: 28, color: pal.title })],
          alignment: AlignmentType.LEFT,
          spacing: { before: 400, after: 200 },
          border: { bottom: { color: pal.border, size: 4, style: BorderStyle.SINGLE } },
        }),
        ...sheets.map((s, i) => new Paragraph({
          children: [
            new TextRun({ text: `${i + 1}.  `, bold: true, size: 22, color: pal.title }),
            new TextRun({ text: s.name, size: 22, color: '374151' }),
            new TextRun({ text: `  —  ${s.data.length} rows`, size: 20, color: '9CA3AF' }),
          ],
          spacing: { after: 100 },
        })),
        // page break after summary
        new Paragraph({ children: [new PageBreak()] }),
      );
    }

    // ── Sheet Sections ─────────────────────────────────────────
    sheets.forEach((sheet, si) => {
      const { name, data } = sheet;
      if (data.length === 0) return;

      // Section heading
      if (config.includeSheetName) {
        children.push(
          new Paragraph({
            children: [new TextRun({ text: name, bold: true, size: 32, color: pal.title })],
            heading: HeadingLevel.HEADING_1,
            spacing: { before: si === 0 && !config.includeSummary ? 0 : 400, after: 240 },
          })
        );
      }

      // Build table
      const maxCols = Math.max(...data.map((r) => r.length));

      const rows = data.map((row, rowIdx) => {
        const isHeader = rowIdx === 0 && config.boldHeaders;
        const cells = [];

        for (let ci = 0; ci < maxCols; ci++) {
          const cellVal = row[ci] !== undefined && row[ci] !== null ? String(row[ci]) : '';
          cells.push(
            new TableCell({
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: cellVal,
                      bold: isHeader,
                      size: fontSizePt,
                      color: isHeader ? pal.headerText : '1F2937',
                    }),
                  ],
                  spacing: { before: 60, after: 60 },
                }),
              ],
              shading: isHeader
                ? { fill: pal.header }
                : (rowIdx % 2 === 0 ? { fill: 'FFFFFF' } : { fill: pal.alt }),
              borders: {
                top:    { style: BorderStyle.SINGLE, size: 1, color: pal.border },
                bottom: { style: BorderStyle.SINGLE, size: 1, color: pal.border },
                left:   { style: BorderStyle.SINGLE, size: 1, color: pal.border },
                right:  { style: BorderStyle.SINGLE, size: 1, color: pal.border },
              },
              margins: { top: 60, bottom: 60, left: 100, right: 100 },
            })
          );
        }

        return new TableRow({ children: cells, tableHeader: rowIdx === 0 && config.boldHeaders });
      });

      const table = new Table({
        rows,
        width: { size: 100, type: WidthType.PERCENTAGE },
      });
      children.push(table);

      // Page break between sheets
      if (config.pageBreaks && si < sheets.length - 1) {
        children.push(new Paragraph({ children: [new PageBreak()] }));
      } else if (si < sheets.length - 1) {
        children.push(new Paragraph({ spacing: { after: 400 } }));
      }
    });

    // ── Build Document ─────────────────────────────────────────
    const doc = new Document({
      creator: config.docAuthor,
      title:   config.docTitle,
      subject: config.docSubject,
      sections: [{
        properties: {
          page: {
            size: config.orientation === 'landscape'
              ? { width: convertInchesToTwip(11), height: convertInchesToTwip(8.5),
                  orientation: docx.PageOrientation ? docx.PageOrientation.LANDSCAPE : 'landscape' }
              : { width: convertInchesToTwip(8.5), height: convertInchesToTwip(11) },
            margin: {
              top:    convertInchesToTwip(1),
              bottom: convertInchesToTwip(1),
              left:   convertInchesToTwip(1),
              right:  convertInchesToTwip(1),
            },
          },
        },
        headers: { default: headerContent },
        footers: { default: footerContent },
        children,
      }],
    });

    const buffer = await docx.Packer.toBlob(doc);
    return buffer;
  }

  // ── Preview ────────────────────────────────────────────────
  function renderPreview(config) {
    if (previewSheets.length === 0) return;
    currentPreviewSheet = 0;

    const totalRows = previewSheets.reduce((sum, s) => sum + s.data.length, 0);
    previewSheetCount.textContent = `${previewSheets.length} sheet(s)`;
    previewRowCount.textContent   = `${totalRows} rows total`;

    renderPreviewSheet(config);
    updatePreviewNav();

    // Output file name
    const baseName = fileName.textContent.replace(/\.[^.]+$/, '');
    const outName = `${baseName}_converted.docx`;
    outputFileName.textContent = outName;
    outputStats.textContent = `${previewSheets.length} sheet(s) · ${totalRows} rows · ${formatBytes(docxBlob.size)}`;

    downloadBtn.onclick = () => {
      const a = document.createElement('a');
      a.href = URL.createObjectURL(docxBlob);
      a.download = outName;
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 5000);
      showToast('✅ Download started!', 'success');
    };
  }

  function renderPreviewSheet(config) {
    const sheet = previewSheets[currentPreviewSheet];
    if (!sheet) return;
    const { name, data } = sheet;
    const style = config?.tableStyle || document.querySelector('input[name="tableStyle"]:checked')?.value || 'blue';

    previewContent.innerHTML = '';

    const heading = document.createElement('h3');
    heading.textContent = `📋 ${name}`;
    previewContent.appendChild(heading);

    if (data.length === 0) {
      previewContent.innerHTML += '<p style="color:var(--text-3);font-size:0.85rem;">No data in this sheet.</p>';
      return;
    }

    const table = document.createElement('table');
    table.className = `preview-table-${style}`;

    const maxCols = data.length > 0 ? Math.max(...data.slice(0, 50).map((r) => r.length)) : 0;

    data.slice(0, 50).forEach((row, ri) => {
      const tr = document.createElement('tr');

      for (let ci = 0; ci < maxCols; ci++) {
        const cell = document.createElement(ri === 0 ? 'th' : 'td');
        cell.textContent = row[ci] !== undefined && row[ci] !== null ? String(row[ci]) : '';
        tr.appendChild(cell);
      }
      table.appendChild(tr);
    });

    if (data.length > 50) {
      const note = document.createElement('p');
      note.style.cssText = 'font-size:0.8rem;color:var(--text-3);margin-top:8px;font-style:italic;';
      note.textContent = `…and ${data.length - 50} more rows (all included in the .docx)`;
      previewContent.appendChild(table);
      previewContent.appendChild(note);
    } else {
      previewContent.appendChild(table);
    }
  }

  function updatePreviewNav() {
    sheetIndicator.textContent = `${currentPreviewSheet + 1} / ${previewSheets.length}`;
    prevSheetBtn.disabled = currentPreviewSheet === 0;
    nextSheetBtn.disabled = currentPreviewSheet === previewSheets.length - 1;
  }

  prevSheetBtn.addEventListener('click', () => {
    if (currentPreviewSheet > 0) {
      currentPreviewSheet--;
      renderPreviewSheet();
      updatePreviewNav();
    }
  });

  nextSheetBtn.addEventListener('click', () => {
    if (currentPreviewSheet < previewSheets.length - 1) {
      currentPreviewSheet++;
      renderPreviewSheet();
      updatePreviewNav();
    }
  });

  // ── Start Over ─────────────────────────────────────────────
  startOver.addEventListener('click', () => {
    resetFile();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  // ── Progress ───────────────────────────────────────────────
  function showProgress(text, pct) {
    progressText.textContent = text;
    progressBar.style.width  = pct + '%';
    progressOverlay.classList.remove('hidden');
  }

  function hideProgress() {
    progressBar.style.width = '100%';
    setTimeout(() => {
      progressOverlay.classList.add('hidden');
      progressBar.style.width = '0%';
    }, 300);
  }

  // ── Toast ──────────────────────────────────────────────────
  let toastTimer;
  function showToast(msg, type = '') {
    toast.textContent = msg;
    toast.className = `toast ${type} show`;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => toast.classList.remove('show'), 3200);
  }

  // ── Helpers ────────────────────────────────────────────────
  function formatBytes(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  }

  function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

})();
