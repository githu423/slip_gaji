(function () {
  'use strict';

  let rows = [];
  let cols = { tunjangan: [], potongan: [] };
  let uploadVersion = 0;
  let generating = false;

  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');
  const fileInfo = document.getElementById('fileInfo');
  const tableSection = document.getElementById('tableSection');
  const dataBadge = document.getElementById('dataBadge');
  const emptyContent = tableSection.innerHTML;
  const formatSelect = document.getElementById('formatSelect');

  dropZone.addEventListener('click', () => fileInput.click());
  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    if (!generating) dropZone.classList.add('drag');
  });
  dropZone.addEventListener('dragleave', (e) => {
    if (!dropZone.contains(e.relatedTarget)) dropZone.classList.remove('drag');
  });
  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag');
    if (!generating && e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) handleFile(e.target.files[0]);
    fileInput.value = '';
  });
  formatSelect.addEventListener('change', () => {
    document.querySelector('#formatNote span').textContent =
      currentFormat() === 'docx'
        ? 'Word: semua slip dalam satu file, halaman terpisah per pegawai.'
        : 'PDF: satu file per pegawai. Unduh semua akan dikemas dalam ZIP.';
    if (rows.length) renderTable(document.getElementById('searchBox')?.value || '');
  });
  document.querySelector('.help-link').addEventListener('click', () => {
    document.getElementById('panduan').open = true;
  });

  function showFileMessage(message, isError = false) {
    fileInfo.hidden = false;
    fileInfo.classList.toggle('is-error', isError);
    fileInfo.textContent = message;
  }

  async function handleFile(file) {
    const version = ++uploadVersion;
    // Clear old data so it cannot be mistaken for the newly selected file.
    rows = [];
    tableSection.innerHTML = emptyContent;
    dataBadge.textContent = 'Menunggu data';
    dataBadge.classList.remove('has-data');
    showFileMessage(`Membaca ${file.name}…`);
    try {
      if (!/\.(xlsx|xls)$/i.test(file.name))
        throw new Error('Pilih file Excel dengan format .xlsx atau .xls.');
      if (!window.XLSX)
        throw new Error(
          'Library Excel belum tersedia. Periksa koneksi internet lalu muat ulang halaman.'
        );
      const buffer = await file.arrayBuffer();
      if (version !== uploadVersion) return;
      const wb = XLSX.read(buffer, { type: 'array' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      if (!json.length) throw new Error('Lembar pertama kosong atau tidak terbaca.');
      const headers = Object.keys(json[0]);
      if (!['Nama', 'Gaji Pokok'].every((h) => headers.includes(h))) {
        throw new Error(
          'Kolom Nama dan Gaji Pokok wajib tersedia. Gunakan template Excel sebagai contoh.'
        );
      }
      const salaryHeaders = headers.filter(
        (h) => h === 'Gaji Pokok' || /^(tunjangan|potongan)/i.test(h)
      );
      json.forEach((row, index) => {
        if (!String(row.Nama).trim())
          throw new Error(`Nama pegawai pada data ke-${index + 1} belum diisi.`);
        if (
          String(row['Gaji Pokok']).trim() === '' ||
          salaryHeaders.some((h) => !Number.isFinite(Number(row[h])))
        ) {
          throw new Error(
            `Nominal gaji pada data ke-${index + 1} harus berupa angka tanpa Rp atau pemisah ribuan.`
          );
        }
      });
      rows = json;
      cols.tunjangan = headers.filter((h) => /^tunjangan/i.test(h));
      cols.potongan = headers.filter((h) => /^potongan/i.test(h));
      renderTable();
      dataBadge.textContent = `${rows.length} pegawai siap`;
      dataBadge.classList.add('has-data');
      showFileMessage(`${file.name} — ${rows.length} pegawai berhasil dibaca`);
    } catch (err) {
      if (version === uploadVersion) showFileMessage('Gagal membaca file: ' + err.message, true);
    }
  }

  function totals(row) {
    let pendapatan = Number(row['Gaji Pokok']) || 0;
    cols.tunjangan.forEach((c) => (pendapatan += Number(row[c]) || 0));
    let potongan = 0;
    cols.potongan.forEach((c) => (potongan += Number(row[c]) || 0));
    return { pendapatan, potongan, bersih: pendapatan - potongan };
  }

  function formatRp(n) {
    return 'Rp ' + (Number(n) || 0).toLocaleString('id-ID');
  }

  function escapeHtml(value) {
    return String(value ?? '').replace(
      /[&<>"']/g,
      (char) =>
        ({
          '&': '&amp;',
          '<': '&lt;',
          '>': '&gt;',
          '"': '&quot;',
          "'": '&#39;'
        })[char]
    );
  }

  function renderTable(filterText = '') {
    const query = filterText.toLocaleLowerCase('id-ID');
    const filtered = rows.filter((r) =>
      ['Nama', 'NIK', 'Jabatan', 'Departemen'].some((k) =>
        String(r[k] ?? '')
          .toLocaleLowerCase('id-ID')
          .includes(query)
      )
    );
    const previousSearch = document.getElementById('searchBox');
    const isSearching = document.activeElement === previousSearch;
    const selectionStart = previousSearch?.selectionStart;
    const selectionEnd = previousSearch?.selectionEnd;

    let html = `
      <div class="toolbar">
        <label class="sr-only" for="searchBox">Cari nama, NIK, jabatan, atau departemen</label>
        <input type="text" id="searchBox" placeholder="Cari nama / NIK / jabatan…" value="${escapeHtml(filterText)}">
        <span class="count" role="status">${filtered.length} dari ${rows.length} pegawai</span>
        <button class="primary-btn" id="downloadAllBtn" type="button"><svg class="icon" aria-hidden="true"><use href="#icon-download"/></svg>${allButtonLabel()}</button>
      </div>
      <div class="table-wrap" tabindex="0" role="region" aria-label="Data gaji pegawai"><table>
        <caption class="sr-only">Rincian gaji pegawai dari file Excel</caption>
        <thead><tr><th scope="col">Nama pegawai</th><th scope="col">NIK</th><th scope="col">Jabatan</th><th scope="col" class="num">Gaji kotor</th><th scope="col" class="num">Gaji bersih</th><th scope="col"><span class="sr-only">Unduh slip</span></th></tr></thead><tbody>`;
    filtered.forEach((r) => {
      const t = totals(r);
      html += `<tr>
        <td>${escapeHtml(r.Nama || '-')}</td><td>${escapeHtml(r.NIK || '-')}</td><td>${escapeHtml(r.Jabatan || '-')}</td>
        <td class="num">${formatRp(t.pendapatan)}</td><td class="num net-pay">${formatRp(t.bersih)}</td>
        <td class="actions"><button type="button" data-idx="${rows.indexOf(r)}" aria-label="Unduh slip ${escapeHtml(r.Nama)}">Unduh</button></td>
      </tr>`;
    });
    if (!filtered.length)
      html +=
        '<tr><td colspan="6" class="no-results">Tidak ada pegawai yang cocok. Coba kata kunci lain.</td></tr>';
    html +=
      '</tbody></table></div><p class="status" id="genStatus" role="status" aria-live="polite"></p>';
    tableSection.innerHTML = html;
    const search = document.getElementById('searchBox');
    search.addEventListener('input', (e) => renderTable(e.target.value));
    if (isSearching) {
      search.focus();
      search.setSelectionRange(selectionStart, selectionEnd);
    }
    document
      .getElementById('downloadAllBtn')
      .addEventListener('click', () =>
        runDownload(() => (currentFormat() === 'docx' ? downloadAllDocx() : downloadAllZip()))
      );
    tableSection.querySelectorAll('td.actions button').forEach((btn) => {
      btn.addEventListener('click', () =>
        runDownload(async () => {
          await downloadOne(rows[Number(btn.dataset.idx)]);
          document.getElementById('genStatus').textContent = 'Slip gaji berhasil dibuat.';
        })
      );
    });
  }

  async function runDownload(action) {
    if (generating) return;
    generating = true;
    const controls = document.querySelectorAll(
      '#generator button, #generator input, #generator select'
    );
    const status = document.getElementById('genStatus');
    controls.forEach((control) => {
      control.disabled = true;
    });
    status.classList.remove('is-error');
    status.textContent = 'Menyiapkan slip gaji…';
    try {
      await action();
    } catch (err) {
      status.classList.add('is-error');
      status.textContent =
        'Gagal membuat slip. Periksa koneksi internet untuk memuat library, lalu coba lagi.';
      console.error('Pembuatan slip gagal:', err);
    } finally {
      generating = false;
      controls.forEach((control) => {
        control.disabled = false;
      });
    }
  }

  function currentFormat() {
    return document.getElementById('formatSelect').value;
  }

  function allButtonLabel() {
    return currentFormat() === 'docx' ? 'Unduh semua (Word, 1 file)' : 'Unduh semua (ZIP)';
  }

  async function downloadOne(row) {
    if (currentFormat() === 'docx') {
      const blob = await docx.Packer.toBlob(buildSlipDocx(row));
      triggerDownload(blob, slipFilename(row, 'docx'));
    } else {
      const doc = buildSlipPdf(row);
      doc.save(slipFilename(row, 'pdf'));
    }
  }

  function triggerDownload(blob, filename) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }

  function slipFilename(row, ext) {
    const periode = (document.getElementById('periode').value || 'periode').replace(/\s+/g, '_');
    const nama = String(row.Nama || 'pegawai').replace(/[^a-z0-9]+/gi, '_');
    return `Slip_Gaji_${nama}_${periode}.${ext}`;
  }

  function buildSlipPdf(row) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: 'mm', format: 'a5' });
    const company = document.getElementById('companyName').value || 'Nama Perusahaan';
    const address = document.getElementById('companyAddress').value;
    const periode = document.getElementById('periode').value || '-';
    const pageW = 148;
    let y = 16;

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(13);
    doc.text(company, pageW / 2, y, { align: 'center' });
    y += 5.5;
    if (address) {
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      doc.text(address, pageW / 2, y, { align: 'center' });
      y += 6;
    }
    doc.setDrawColor(60);
    doc.line(12, y, pageW - 12, y);
    y += 7;

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(11);
    doc.text('SLIP GAJI', pageW / 2, y, { align: 'center' });
    y += 5;
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    doc.text(`Periode: ${periode}`, pageW / 2, y, { align: 'center' });
    y += 9;

    doc.setFontSize(9.5);
    [
      ['Nama', row.Nama],
      ['NIK', row.NIK],
      ['Jabatan', row.Jabatan],
      ['Departemen', row.Departemen]
    ]
      .filter(([, v]) => v !== undefined && v !== '')
      .forEach(([label, val]) => {
        doc.text(label, 12, y);
        doc.text(': ' + val, 34, y);
        y += 5.5;
      });
    y += 2;
    doc.line(12, y, pageW - 12, y);
    y += 6;

    const t = totals(row);

    doc.setFont('helvetica', 'bold');
    doc.text('Pendapatan', 12, y);
    y += 5.5;
    doc.setFont('helvetica', 'normal');
    if (row['Gaji Pokok'] !== undefined && row['Gaji Pokok'] !== '') {
      doc.text('Gaji Pokok', 16, y);
      doc.text(formatRp(row['Gaji Pokok']), pageW - 12, y, { align: 'right' });
      y += 5.5;
    }
    cols.tunjangan.forEach((c) => {
      if (row[c] === undefined || row[c] === '') return;
      doc.text(c, 16, y);
      doc.text(formatRp(row[c]), pageW - 12, y, { align: 'right' });
      y += 5.5;
    });
    y += 1;
    doc.line(12, y, pageW - 12, y);
    y += 5.5;
    doc.setFont('helvetica', 'bold');
    doc.text('Gaji Kotor', 16, y);
    doc.text(formatRp(t.pendapatan), pageW - 12, y, { align: 'right' });
    y += 8;

    if (cols.potongan.length) {
      doc.setFont('helvetica', 'bold');
      doc.text('Potongan', 12, y);
      y += 5.5;
      doc.setFont('helvetica', 'normal');
      cols.potongan.forEach((c) => {
        if (row[c] === undefined || row[c] === '') return;
        doc.text(c, 16, y);
        doc.text(formatRp(row[c]), pageW - 12, y, { align: 'right' });
        y += 5.5;
      });
      y += 1;
      doc.line(12, y, pageW - 12, y);
      y += 5.5;
      doc.setFont('helvetica', 'bold');
      doc.text('Total Potongan', 16, y);
      doc.text(formatRp(t.potongan), pageW - 12, y, { align: 'right' });
      y += 8;
    }

    doc.setFontSize(11);
    doc.text('GAJI BERSIH', 12, y);
    doc.text(formatRp(t.bersih), pageW - 12, y, { align: 'right' });

    doc.setFontSize(7.5);
    doc.setFont('helvetica', 'italic');
    doc.text('Dokumen digenerate otomatis.', pageW / 2, 200, { align: 'center' });

    return doc;
  }

  // ---------- Word (.docx) generation ----------

  function row2col(label, value, opts) {
    opts = opts || {};
    return new docx.TableRow({
      children: [
        new docx.TableCell({
          width: { size: 60, type: docx.WidthType.PERCENTAGE },
          borders: noBorders(),
          children: [
            new docx.Paragraph({ children: [new docx.TextRun({ text: label, bold: !!opts.bold })] })
          ]
        }),
        new docx.TableCell({
          width: { size: 40, type: docx.WidthType.PERCENTAGE },
          borders: noBorders(),
          children: [
            new docx.Paragraph({
              alignment: docx.AlignmentType.RIGHT,
              children: [new docx.TextRun({ text: value, bold: !!opts.bold })]
            })
          ]
        })
      ]
    });
  }

  function noBorders() {
    const none = { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    return { top: none, bottom: none, left: none, right: none };
  }

  function divider() {
    return new docx.Paragraph({
      border: { bottom: { color: '999999', space: 1, style: docx.BorderStyle.SINGLE, size: 4 } },
      spacing: { after: 160 }
    });
  }

  function buildSlipSection(row) {
    const company = document.getElementById('companyName').value || 'Nama Perusahaan';
    const address = document.getElementById('companyAddress').value;
    const periode = document.getElementById('periode').value || '-';
    const t = totals(row);
    const children = [];

    children.push(
      new docx.Paragraph({
        alignment: docx.AlignmentType.CENTER,
        children: [new docx.TextRun({ text: company, bold: true, size: 28 })]
      })
    );
    if (address) {
      children.push(
        new docx.Paragraph({
          alignment: docx.AlignmentType.CENTER,
          children: [new docx.TextRun({ text: address, size: 18, color: '555555' })]
        })
      );
    }
    children.push(divider());
    children.push(
      new docx.Paragraph({
        alignment: docx.AlignmentType.CENTER,
        spacing: { after: 40 },
        children: [new docx.TextRun({ text: 'SLIP GAJI', bold: true, size: 24 })]
      })
    );
    children.push(
      new docx.Paragraph({
        alignment: docx.AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [new docx.TextRun({ text: `Periode: ${periode}`, size: 20, color: '555555' })]
      })
    );

    const infoRows = [
      ['Nama', row.Nama],
      ['NIK', row.NIK],
      ['Jabatan', row.Jabatan],
      ['Departemen', row.Departemen]
    ]
      .filter(([, v]) => v !== undefined && v !== '')
      .map(([label, val]) => row2col(label, String(val)));
    children.push(
      new docx.Table({ width: { size: 100, type: docx.WidthType.PERCENTAGE }, rows: infoRows })
    );
    children.push(divider());

    const earningRows = [];
    if (row['Gaji Pokok'] !== undefined && row['Gaji Pokok'] !== '') {
      earningRows.push(row2col('Gaji Pokok', formatRp(row['Gaji Pokok'])));
    }
    cols.tunjangan.forEach((c) => {
      if (row[c] === undefined || row[c] === '') return;
      earningRows.push(row2col(c, formatRp(row[c])));
    });
    earningRows.push(row2col('Gaji Kotor', formatRp(t.pendapatan), { bold: true }));
    children.push(
      new docx.Paragraph({
        spacing: { before: 80, after: 40 },
        children: [new docx.TextRun({ text: 'Pendapatan', bold: true, size: 20 })]
      })
    );
    children.push(
      new docx.Table({ width: { size: 100, type: docx.WidthType.PERCENTAGE }, rows: earningRows })
    );

    if (cols.potongan.length) {
      const deductionRows = [];
      cols.potongan.forEach((c) => {
        if (row[c] === undefined || row[c] === '') return;
        deductionRows.push(row2col(c, formatRp(row[c])));
      });
      deductionRows.push(row2col('Total Potongan', formatRp(t.potongan), { bold: true }));
      children.push(
        new docx.Paragraph({
          spacing: { before: 200, after: 40 },
          children: [new docx.TextRun({ text: 'Potongan', bold: true, size: 20 })]
        })
      );
      children.push(
        new docx.Table({
          width: { size: 100, type: docx.WidthType.PERCENTAGE },
          rows: deductionRows
        })
      );
    }

    children.push(divider());
    children.push(
      new docx.Table({
        width: { size: 100, type: docx.WidthType.PERCENTAGE },
        rows: [row2col('GAJI BERSIH', formatRp(t.bersih), { bold: true })]
      })
    );

    children.push(
      new docx.Paragraph({
        alignment: docx.AlignmentType.CENTER,
        spacing: { before: 300 },
        children: [
          new docx.TextRun({
            text: 'Dokumen digenerate otomatis.',
            italics: true,
            size: 16,
            color: '777777'
          })
        ]
      })
    );

    return children;
  }

  function buildSlipDocx(row) {
    return new docx.Document({ sections: [{ children: buildSlipSection(row) }] });
  }

  async function downloadAllDocx() {
    if (!rows.length) return;
    const statusEl = document.getElementById('genStatus');
    const btn = document.getElementById('downloadAllBtn');
    btn.disabled = true;
    statusEl.textContent = `Menyusun ${rows.length} lembar slip...`;

    const allChildren = [];
    rows.forEach((row, i) => {
      const section = buildSlipSection(row);
      if (i > 0) section.unshift(new docx.Paragraph({ children: [], pageBreakBefore: true }));
      allChildren.push(...section);
    });

    const doc = new docx.Document({ sections: [{ children: allChildren }] });
    const blob = await docx.Packer.toBlob(doc);
    const periode = (document.getElementById('periode').value || 'periode').replace(/\s+/g, '_');
    triggerDownload(blob, `Slip_Gaji_${periode}.docx`);
    statusEl.textContent = `Selesai — ${rows.length} lembar dalam satu file Word, satu halaman per orang.`;
    btn.disabled = false;
  }

  async function downloadAllZip() {
    if (!rows.length) return;
    const statusEl = document.getElementById('genStatus');
    const btn = document.getElementById('downloadAllBtn');
    btn.disabled = true;
    const zip = new JSZip();
    for (let i = 0; i < rows.length; i++) {
      statusEl.textContent = `Membuat slip ${i + 1} dari ${rows.length}...`;
      const doc = buildSlipPdf(rows[i]);
      zip.file(`${i + 1}_${slipFilename(rows[i], 'pdf')}`, doc.output('blob'));
      await new Promise((r) => setTimeout(r, 0)); // keep UI responsive
    }
    statusEl.textContent = 'Menyusun file ZIP...';
    const blob = await zip.generateAsync({ type: 'blob' });
    const periode = (document.getElementById('periode').value || 'periode').replace(/\s+/g, '_');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `Slip_Gaji_${periode}.zip`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    statusEl.textContent = `Selesai — ${rows.length} slip diunduh dalam satu ZIP.`;
    btn.disabled = false;
  }

  document.getElementById('downloadTemplate').addEventListener('click', () => {
    if (!window.XLSX) {
      showFileMessage(
        'Library Excel belum tersedia. Periksa koneksi internet lalu muat ulang halaman.',
        true
      );
      return;
    }
    const sample = [
      {
        Nama: 'Siti Aminah',
        NIK: '3201010001',
        Jabatan: 'Staf Administrasi',
        Departemen: 'Umum',
        'Gaji Pokok': 3500000,
        'Tunjangan Jabatan': 500000,
        'Tunjangan Transport': 300000,
        'Potongan BPJS': 175000
      },
      {
        Nama: 'Budi Santoso',
        NIK: '3201010002',
        Jabatan: 'Staf Keuangan',
        Departemen: 'Keuangan',
        'Gaji Pokok': 4000000,
        'Tunjangan Jabatan': 700000,
        'Tunjangan Transport': 300000,
        'Potongan BPJS': 200000
      }
    ];
    const ws = XLSX.utils.json_to_sheet(sample);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data Gaji');
    XLSX.writeFile(wb, 'Contoh_Data_Gaji.xlsx');
  });
})();
