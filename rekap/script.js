const FOLDER_MAP = {
    "CONTER": {
        "RAB": "1ycHYO1rTHGfQ8lmOCY1CcED-MKy6VZ8U",
        "Realisasi Anggaran": "1wI2ceA-TxpPbjIIMEp2OuH27INUeff1A"
    },
    "HUMAS": {
        "RAB": "1TEyQeC3eh70qa_XMnHOFPLOgaVJ9yrCm",
        "Realisasi Anggaran": "1xLDd3mq49vTXCZQALionmZjYfOZsBOyO"
    },
    "MEDIA": {
        "RAB": "1gcaKGgNjPVgAZY5GNNVYCuYffGo863-a",
        "Realisasi Anggaran": "1ix2_NUAedmCphmMGvkIZmonbrC27B8lV"
    },
    "PAKSIMA": {
        "RAB": "1AoKVj1jv8hNiPqRt-reQA0RBXd9vXK4S",
        "Realisasi Anggaran": "128wGr7Of8QgGuHgzgwy7WXn7lnGSXugV"
    },
    "PDK": {
        "RAB": "1a5BY4WB-6wteMnRChDGYkd3Br3jY6phW",
        "Realisasi Anggaran": "1E7h9S_J-UyvwbFfwm62DYmIi7BC5Hvh0"
    },
    "PPSDM": {
        "RAB": "1h7aqkLck3qLXDj4d4ZKQkUbbE3AvBvuU",
        "Realisasi Anggaran": "1-nXJWzKXogrHlEkJICcSCFpa9LgSifIV"
    }
};

let listDraft = [];
let currentDraftId = null;
let transaksi = [];
let editIndex = -1;
let targetIndexNotaEdit = -1;

let undoStack = [];
let redoStack = [];

let chartComp = null;
let chartDiv = null;

const GCC_LOGO_URL = "https://lh3.googleusercontent.com/d/1QOjchWwyuyYMdB00zTWBUyS-h7LYGvWl";

function getFileNameFormat(extension) {
    const draft = listDraft.find(d => d.id === currentDraftId);
    const rawName = draft && draft.nama ? draft.nama : 'Rekapan_Keuangan_GCC';
    const safeName = rawName.replace(/[^a-zA-Z0-9_-]/g, '_');
    return `${safeName}.${extension}`;
}

function formatRupiah(angka) {
    return 'Rp ' + Number(angka).toLocaleString('id-ID');
}

function openModal(dataUrl, title) {
    const modal = document.getElementById('imageModal');
    const img = document.getElementById('modalImage');
    const caption = document.getElementById('modalCaption');
    modal.style.display = 'flex';
    img.src = dataUrl;
    caption.innerText = title;
}

function closeModal() {
    document.getElementById('imageModal').style.display = 'none';
}

function triggerGantiNota(realIndex) {
    targetIndexNotaEdit = realIndex;
    document.getElementById('fileNotaHidden').click();
}

function hapusNota(realIndex) {
    if (confirm("Apakah Anda yakin ingin menghapus lampiran nota/bukti transaksi ini?")) {
        simpanStateKeHistory();
        transaksi[realIndex].notaData = null;
        simpanKeStorage(false);
        updateTampilan();
        showToast("Lampiran dihapus", "info");
    }
}

async function downloadTemplateWord() {
    const { docx } = window;
    const borderStandard = { style: docx.BorderStyle.SINGLE, size: 1, color: "000000" };
    const borders = { top: borderStandard, bottom: borderStandard, left: borderStandard, right: borderStandard, insideHorizontal: borderStandard, insideVertical: borderStandard };

    const templateRows = [
        new docx.TableRow({
            children: [
                new docx.TableCell({ children: [new docx.Paragraph({ text: "No", bold: true, alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Jenis (Masuk/Keluar)", bold: true, alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Keterangan", bold: true, alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Harga Satuan (Rp)", bold: true, alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Jumlah (Qty)", bold: true, alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Total Nominal (Rp)", bold: true, alignment: docx.AlignmentType.CENTER })] })
            ]
        }),
        new docx.TableRow({ children: [new docx.TableCell({ columnSpan: 6, children: [new docx.Paragraph({ text: "CONTER", bold: true })] })] }),
        new docx.TableRow({
            children: [
                new docx.TableCell({ children: [new docx.Paragraph({ text: "1", alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Masuk" })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Dana Kas GCC" })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "1000000" })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "1", alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "1000000" })] })
            ]
        })
    ];

    const doc = new docx.Document({
        sections: [{
            children: [
                new docx.Paragraph({ children: [new docx.TextRun({ text: "TEMPLAT KEUANGAN GCC", bold: true, size: 24 })], alignment: docx.AlignmentType.CENTER, spacing: { after: 150 } }),
                new docx.Table({ borders: borders, rows: templateRows, width: { size: 100, type: docx.WidthType.PERCENTAGE } })
            ]
        }]
    });

    docx.Packer.toBlob(doc).then(blob => { saveAs(blob, "Templat_Rekapan_Keuangan_GCC.docx"); });
    showToast("Templat Word berhasil di-download!", "success");
}

async function imporDokumen(event) {
    const file = event.target.files[0];
    if (!file || !file.name.endsWith('.docx')) {
        showToast("Pilih file Word (.docx)!", "error");
        return;
    }

    try {
        const arrayBuffer = await file.arrayBuffer();
        const rawTextResult = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
        const matchData = rawTextResult.value.match(/GCC_DATA_START:::(.*?):::GCC_DATA_END/s);

        if (matchData && matchData[1]) {
            const decodedJson = JSON.parse(decodeURIComponent(matchData[1].replace(/\s+/g, '')));
            simpanStateKeHistory();
            if (decodedJson.jumlahPeserta !== undefined) document.getElementById('jumlahPeserta').value = decodedJson.jumlahPeserta;
            if (Array.isArray(decodedJson.transaksi)) transaksi = decodedJson.transaksi.map(rapikanTransaksi).filter(Boolean);
            if (decodedJson.budgets && typeof decodedJson.budgets === 'object') {
                const draftAktif = getDraftAktif();
                if (draftAktif) draftAktif.budgets = decodedJson.budgets;
            }
            simpanKeStorage(false);
            updateTampilan();
            showToast("Data draf & lampiran berhasil dipulihkan dari file Word!", "success");
            return;
        }

        const htmlResult = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = htmlResult.value;

        simpanStateKeHistory();
        let count = 0;
        let currentDiv = 'CONTER';

        tempDiv.querySelectorAll('tr').forEach(row => {
            const cells = Array.from(row.querySelectorAll('td, th')).map(c => (c.innerText || c.textContent || '').trim());
            if (cells.length === 0 || cells.join(' ').toLowerCase().includes('keterangan')) return;

            if (cells.length === 1 || (cells.length > 1 && !cells[1] && !cells[2])) {
                if (cells[0] && !cells[0].match(/\d{3,}/)) currentDiv = cells[0];
                return;
            }

            if (cells.length >= 4) {
                let jenis = cells[1] && cells[1].toLowerCase().includes('masuk') ? 'Masuk' : 'Keluar';
                let ket = cells[2] || cells[1] || '';
                let harga = parseFloat((cells[3] || '0').replace(/[^\d]/g, '')) || 0;
                let qty = parseFloat((cells[4] || '1').replace(/[^\d]/g, '')) || 1;

                if (ket && harga > 0) {
                    transaksi.push({ divisi: currentDiv, jenis, keterangan: ket, harga, qty, totalNominal: harga * qty, notaData: null, catatan: 'Diimpor dari Word' });
                    count++;
                }
            }
        });

        if (count > 0) {
            simpanKeStorage(false);
            updateTampilan();
            showToast(`Berhasil mengimpor ${count} item transaksi!`, "success");
        } else {
            showToast("Format tabel tidak terbaca.", "error");
        }
    } catch (e) {
        showToast("Gagal membaca file Word.", "error");
    }
}

function undo() {
    if (undoStack.length === 0) return;
    redoStack.push({ transaksi: JSON.parse(JSON.stringify(transaksi)), jumlahPeserta: document.getElementById('jumlahPeserta').value || 0 });
    const prev = undoStack.pop();
    transaksi = prev.transaksi;
    document.getElementById('jumlahPeserta').value = prev.jumlahPeserta;
    simpanKeStorage(false);
    updateTampilan();
    updateTombolHistory();
}

function redo() {
    if (redoStack.length === 0) return;
    undoStack.push({ transaksi: JSON.parse(JSON.stringify(transaksi)), jumlahPeserta: document.getElementById('jumlahPeserta').value || 0 });
    const next = redoStack.pop();
    transaksi = next.transaksi;
    document.getElementById('jumlahPeserta').value = next.jumlahPeserta;
    simpanKeStorage(false);
    updateTampilan();
    updateTombolHistory();
}

function updateTombolHistory() {
    document.getElementById('btnUndo').disabled = undoStack.length === 0;
    document.getElementById('btnRedo').disabled = redoStack.length === 0;
}

function renderDropdownDraft() {
    const select = document.getElementById('selectDraft');
    select.innerHTML = '';
    listDraft.forEach((d, idx) => {
        const opt = document.createElement('option');
        opt.value = d.id;
        opt.innerText = `Draf #${idx + 1} - ${d.nama}`;
        if (d.id === currentDraftId) opt.selected = true;
        select.appendChild(opt);
    });
}

function buatDraftBaru(userClick = true) {
    let namaDraf = `Rekapan ${new Date().toLocaleDateString('id-ID')}`;
    if (userClick) {
        const inputNama = prompt("Masukkan Nama Draf Kegiatan Baru:", namaDraf);
        if (!inputNama) return;
        namaDraf = inputNama.trim();
    }

    const newDraft = { id: 'draft_' + Date.now(), nama: namaDraf, jumlahPeserta: 0, transaksi: [] };
    listDraft.unshift(newDraft);
    currentDraftId = newDraft.id;
    simpanListDraft();
    renderDropdownDraft();
    muatDraftAktif();
    if (userClick) showToast("Draf baru berhasil dibuat!", "success");
}

function ubahNamaDraft() {
    const draft = listDraft.find(d => d.id === currentDraftId);
    if (!draft) return;
    const namaBaru = prompt("Ubah Nama Draf:", draft.nama);
    if (namaBaru && namaBaru.trim() !== '') {
        draft.nama = namaBaru.trim();
        simpanListDraft();
        renderDropdownDraft();
        showToast("Nama draf diperbarui", "success");
    }
}

function gantiDraft() {
    currentDraftId = document.getElementById('selectDraft').value;
    muatDraftAktif();
}

function simpanKeStorage(recordHistory = true) {
    if (recordHistory) simpanStateKeHistory();
    const draft = listDraft.find(d => d.id === currentDraftId);
    if (draft) {
        draft.jumlahPeserta = document.getElementById('jumlahPeserta').value || 0;
        draft.transaksi = transaksi;
        simpanListDraft();
    }
}

function hapusDraftAktif() {
    if (confirm("Apakah Anda yakin ingin menghapus draf ini?")) {
        listDraft = listDraft.filter(d => d.id !== currentDraftId);
        simpanListDraft();
        muatSemuaDraft();
        showToast("Draf dihapus", "info");
    }
}

/* ------------------- GENERATOR BLOB FILE (PDF & WORD) ------------------- */
/* ------------------- PDF TEKS ASLI (BUKAN GAMBAR) ------------------- */
// Dibuat langsung dengan jsPDF + autoTable, jadi semua tulisan bisa diseleksi, dicari, dan disalin
// seperti dokumen Word. Hanya foto lampiran nota yang berupa gambar.
function pdfSafe(text) {
    return String(text ?? '')
        .replace(/[\u2018\u2019]/g, "'")
        .replace(/[\u201C\u201D]/g, '"')
        .replace(/[^\x09\x0A\x0D\x20-\x7E\u00A0-\u00FF\u2013\u2014\u2022\u2026]/g, '');
}

function pdfLoadImage(dataUrl) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => {
            const scale = Math.min(1, 1600 / Math.max(img.naturalWidth, img.naturalHeight));
            const w = Math.max(1, Math.round(img.naturalWidth * scale));
            const h = Math.max(1, Math.round(img.naturalHeight * scale));
            const canvas = document.createElement('canvas');
            canvas.width = w; canvas.height = h;
            const ctx = canvas.getContext('2d');
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, w, h);
            ctx.drawImage(img, 0, 0, w, h);
            resolve({ data: canvas.toDataURL('image/jpeg', 0.9), w, h });
        };
        img.onerror = reject;
        img.src = dataUrl;
    });
}

async function generatePdfBlob({ namaDraft, summary, kelompokDivisi, transaksiBerlampiran }) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: 'mm', format: 'a4' });
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const margin = 25.4; // sama dengan margin bawaan Word (1 inci)
    const contentW = pageW - margin * 2;
    const bottomLimit = pageH - margin;
    const lineH = 5;
    let y = margin;

    doc.setProperties({ title: pdfSafe(`Laporan Rekapan Keuangan ${namaDraft}`) });

    // Judul
    doc.setFont('times', 'bold');
    doc.setFontSize(12);
    const judul = doc.splitTextToSize(pdfSafe(`LAPORAN REKAPAN KEUANGAN ${namaDraft.toUpperCase()}`), contentW);
    doc.text(judul, pageW / 2, y + 4, { align: 'center' });
    y += judul.length * 5.5 + 5;

    // Ringkasan kegiatan
    doc.setFontSize(11);
    doc.text('RINGKASAN KEGIATAN', margin, y + 4);
    y += 8;
    doc.setFont('times', 'normal');
    const ringkasan = [
        ['\u2022 Total Peserta Hadir', `${summary.totalPeserta} Orang`],
        ['\u2022 Total Uang Masuk', formatRupiah(summary.totalMasuk)],
        ['\u2022 Total Uang Keluar', formatRupiah(summary.totalKeluar)],
        ['\u2022 Sisa Saldo Akhir', `${formatRupiah(summary.saldoAkhir)} (Uang Masuk - Uang Keluar)`]
    ];
    ringkasan.forEach(([label, value]) => {
        const valLines = doc.splitTextToSize(pdfSafe(value), contentW - 62);
        doc.text(label, margin + 3, y + 4);
        doc.text(':', margin + 58, y + 4);
        doc.text(valLines, margin + 62, y + 4);
        y += valLines.length * lineH + 1;
    });

    // Rincian transaksi
    y += 5;
    doc.setFont('times', 'bold');
    doc.text('RINCIAN TRANSAKSI', margin, y + 4);
    y += 7;

    const body = [];
    for (const [divName, items] of Object.entries(kelompokDivisi)) {
        body.push([{ content: pdfSafe(divName), colSpan: 6, styles: { fontStyle: 'bold', halign: 'left' } }]);
        items.forEach((t, i) => {
            body.push([
                String(i + 1),
                pdfSafe(t.jenis),
                pdfSafe(t.keterangan),
                formatRupiah(t.harga),
                String(t.qty),
                formatRupiah(t.totalNominal)
            ]);
        });
    }
    const totalRow = (label, nilai) => [
        { content: label, colSpan: 5, styles: { fontStyle: 'bold', halign: 'left' } },
        { content: formatRupiah(nilai), styles: { fontStyle: 'bold', halign: 'left' } }
    ];
    body.push(totalRow('Total Uang Masuk', summary.totalMasuk));
    body.push(totalRow('Total Uang Keluar', summary.totalKeluar));
    body.push(totalRow('Total Akhir (Uang Masuk - Uang Keluar)', summary.saldoAkhir));

    doc.autoTable({
        startY: y,
        margin: { left: margin, right: margin, top: margin, bottom: margin },
        head: [['No', 'Jenis', 'Keterangan', 'Harga Satuan', 'Jumlah', 'Total Nominal']],
        body,
        theme: 'grid',
        rowPageBreak: 'avoid',
        styles: {
            font: 'times', fontSize: 10.5, textColor: 0,
            lineColor: 0, lineWidth: 0.2, valign: 'middle',
            cellPadding: { top: 2.2, bottom: 2.2, left: 3, right: 3 }
        },
        headStyles: { fillColor: 255, textColor: 0, fontStyle: 'bold', halign: 'center' },
        columnStyles: {
            0: { halign: 'center', cellWidth: 12 },
            1: { halign: 'center', cellWidth: 20 },
            2: { cellWidth: 'auto' },
            3: { halign: 'left', cellWidth: 30 },
            4: { halign: 'center', cellWidth: 21 },
            5: { halign: 'left', cellWidth: 30 }
        }
    });
    y = doc.lastAutoTable.finalY;

    // Lampiran bukti nota (satu-satunya bagian yang berupa gambar)
    if (transaksiBerlampiran.length > 0) {
        doc.addPage();
        y = margin;
        doc.setFont('times', 'bold');
        doc.setFontSize(11);
        doc.text('LAMPIRAN BUKTI NOTA & TRANSAKSI', margin, y + 4);
        y += 11;

        for (const [idx, t] of transaksiBerlampiran.entries()) {
            const caption = doc.splitTextToSize(
                pdfSafe(`Lampiran ${idx + 1}: ${t.keterangan} (${t.divisi}) - Total: ${formatRupiah(t.totalNominal)}`),
                contentW
            );

            let img = null, imgW = 0, imgH = 0;
            try {
                img = await pdfLoadImage(t.notaData.dataUrl);
                const ratio = Math.min(90 / img.w, 100 / img.h, 25.4 / 96);
                imgW = img.w * ratio;
                imgH = img.h * ratio;
            } catch (e) {
                console.error('Gagal merender gambar ke PDF', e);
            }

            if (y + caption.length * lineH + 3 + imgH + 10 > bottomLimit) {
                doc.addPage();
                y = margin;
            }

            doc.setFont('times', 'bold');
            doc.setFontSize(11);
            doc.text(caption, margin, y + 4);
            y += caption.length * lineH + 3;

            if (img) {
                doc.addImage(img.data, 'JPEG', margin, y, imgW, imgH);
                doc.setDrawColor(190);
                doc.setLineWidth(0.2);
                doc.rect(margin, y, imgW, imgH);
                y += imgH;
            }
            y += 10;
        }
    }

    return doc.output('blob');
}

async function generateFileBlob(formatType) {
    const summary = updateSummary();
    const draft = listDraft.find(d => d.id === currentDraftId);
    const namaDraft = draft && draft.nama ? draft.nama : 'KEGIATAN';

    const kelompokDivisi = {};
    transaksi.forEach(t => {
        const divName = (t.divisi || 'UMUM').toUpperCase();
        if (!kelompokDivisi[divName]) kelompokDivisi[divName] = [];
        kelompokDivisi[divName].push(t);
    });

    const transaksiBerlampiran = transaksi.filter(t => t.notaData && t.notaData.dataUrl);

    if (formatType === 'xlsx') {
        const dataExcel = [];
        dataExcel.push([`LAPORAN REKAPAN KEUANGAN ${namaDraft.toUpperCase()}`]);
        dataExcel.push([]);
        dataExcel.push(["RINGKASAN KEGIATAN"]);
        dataExcel.push(["• Total Peserta Hadir", "", summary.totalPeserta + " Orang"]);
        dataExcel.push(["• Total Uang Masuk", "", summary.totalMasuk]);
        dataExcel.push(["• Total Uang Keluar", "", summary.totalKeluar]);
        dataExcel.push(["• Sisa Saldo Akhir", "", summary.saldoAkhir, "(Uang Masuk - Uang Keluar)"]);
        dataExcel.push([]);
        dataExcel.push(["RINCIAN TRANSAKSI"]);
        dataExcel.push(["No", "Jenis", "Keterangan", "Harga Satuan", "Jumlah", "Total Nominal"]);

        let globalNo = 1;
        for (const [divName, items] of Object.entries(kelompokDivisi)) {
            dataExcel.push([divName]);
            items.forEach(t => {
                dataExcel.push([globalNo++, t.jenis, t.keterangan, t.harga, t.qty, t.totalNominal]);
            });
        }

        dataExcel.push([]);
        dataExcel.push(["Total Uang Masuk", "", "", "", "", summary.totalMasuk]);
        dataExcel.push(["Total Uang Keluar", "", "", "", "", summary.totalKeluar]);
        dataExcel.push(["Total Akhir (Uang Masuk - Uang Keluar)", "", "", "", "", summary.saldoAkhir]);

        if (transaksiBerlampiran.length > 0) {
            dataExcel.push([]);
            dataExcel.push(["LAMPIRAN BUKTI NOTA / TRANSAKSI"]);
            transaksiBerlampiran.forEach((t, idx) => {
                dataExcel.push([`Lampiran ${idx + 1}: ${t.keterangan} (${t.divisi}) - Total: ${formatRupiah(t.totalNominal)}`]);
            });
        }

        const ws = XLSX.utils.aoa_to_sheet(dataExcel);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Rekapan Keuangan");
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        return new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    }

    if (formatType === 'pdf') {
        return await generatePdfBlob({ namaDraft, summary, kelompokDivisi, transaksiBerlampiran });
    }

    if (formatType === 'docx') {
        const { docx } = window;
        const borderStandard = { style: docx.BorderStyle.SINGLE, size: 1, color: "000000" };
        const borders = { top: borderStandard, bottom: borderStandard, left: borderStandard, right: borderStandard, insideHorizontal: borderStandard, insideVertical: borderStandard };

        const cellLeftIndent = 36;
        const createCell = (text, isHeader = false, align = docx.AlignmentType.LEFT, bold = false, colSpan = 1) => {
            return new docx.TableCell({
                columnSpan: colSpan,
                margins: { left: cellLeftIndent, right: cellLeftIndent, top: 100, bottom: 100 },
                children: [new docx.Paragraph({ children: [new docx.TextRun({ text: text, bold: isHeader || bold })], alignment: align })]
            });
        };

        const tableRows = [
            new docx.TableRow({
                children: [
                    createCell("No", true, docx.AlignmentType.CENTER),
                    createCell("Jenis", true, docx.AlignmentType.CENTER),
                    createCell("Keterangan", true, docx.AlignmentType.CENTER),
                    createCell("Harga Satuan", true, docx.AlignmentType.CENTER),
                    createCell("Jumlah", true, docx.AlignmentType.CENTER),
                    createCell("Total Nominal", true, docx.AlignmentType.CENTER)
                ]
            })
        ];

        for (const [divName, items] of Object.entries(kelompokDivisi)) {
            tableRows.push(new docx.TableRow({
                children: [
                    new docx.TableCell({
                        columnSpan: 6,
                        margins: { left: cellLeftIndent, right: cellLeftIndent, top: 100, bottom: 100 },
                        children: [new docx.Paragraph({ children: [new docx.TextRun({ text: divName, bold: true })] })]
                    })
                ]
            }));

            items.forEach((t, index) => {
                tableRows.push(new docx.TableRow({
                    children: [
                        createCell(String(index + 1), false, docx.AlignmentType.CENTER),
                        createCell(t.jenis, false, docx.AlignmentType.CENTER),
                        createCell(t.keterangan, false, docx.AlignmentType.LEFT),
                        createCell(formatRupiah(t.harga), false, docx.AlignmentType.LEFT),
                        createCell(String(t.qty), false, docx.AlignmentType.CENTER),
                        createCell(formatRupiah(t.totalNominal), false, docx.AlignmentType.LEFT)
                    ]
                }));
            });
        }

        tableRows.push(
            new docx.TableRow({ children: [createCell("Total Uang Masuk", false, docx.AlignmentType.LEFT, true, 5), createCell(formatRupiah(summary.totalMasuk), false, docx.AlignmentType.LEFT, true)] }),
            new docx.TableRow({ children: [createCell("Total Uang Keluar", false, docx.AlignmentType.LEFT, true, 5), createCell(formatRupiah(summary.totalKeluar), false, docx.AlignmentType.LEFT, true)] }),
            new docx.TableRow({ children: [createCell("Total Akhir (Uang Masuk - Uang Keluar)", false, docx.AlignmentType.LEFT, true, 5), createCell(formatRupiah(summary.saldoAkhir), false, docx.AlignmentType.LEFT, true)] })
        );

        const summaryTable = new docx.Table({
            borders: { top: { style: docx.BorderStyle.NONE }, bottom: { style: docx.BorderStyle.NONE }, left: { style: docx.BorderStyle.NONE }, right: { style: docx.BorderStyle.NONE }, insideHorizontal: { style: docx.BorderStyle.NONE }, insideVertical: { style: docx.BorderStyle.NONE } },
            rows: [
                new docx.TableRow({ children: [new docx.TableCell({ margins: { left: cellLeftIndent }, children: [new docx.Paragraph("• Total Peserta Hadir")], width: { size: 35, type: docx.WidthType.PERCENTAGE } }), new docx.TableCell({ children: [new docx.Paragraph(":")], width: { size: 5, type: docx.WidthType.PERCENTAGE } }), new docx.TableCell({ children: [new docx.Paragraph(`${summary.totalPeserta} Orang`)], width: { size: 60, type: docx.WidthType.PERCENTAGE } })] }),
                new docx.TableRow({ children: [new docx.TableCell({ margins: { left: cellLeftIndent }, children: [new docx.Paragraph("• Total Uang Masuk")] }), new docx.TableCell({ children: [new docx.Paragraph(":")] }), new docx.TableCell({ children: [new docx.Paragraph(formatRupiah(summary.totalMasuk))] })] }),
                new docx.TableRow({ children: [new docx.TableCell({ margins: { left: cellLeftIndent }, children: [new docx.Paragraph("• Total Uang Keluar")] }), new docx.TableCell({ children: [new docx.Paragraph(":")] }), new docx.TableCell({ children: [new docx.Paragraph(formatRupiah(summary.totalKeluar))] })] }),
                new docx.TableRow({ children: [new docx.TableCell({ margins: { left: cellLeftIndent }, children: [new docx.Paragraph("• Sisa Saldo Akhir")] }), new docx.TableCell({ children: [new docx.Paragraph(":")] }), new docx.TableCell({ children: [new docx.Paragraph(`${formatRupiah(summary.saldoAkhir)} (Uang Masuk - Uang Keluar)`)] })] })
            ]
        });

        const docChildren = [
            new docx.Paragraph({ children: [new docx.TextRun({ text: `LAPORAN REKAPAN KEUANGAN ${namaDraft.toUpperCase()}`, bold: true, size: 24 })], alignment: docx.AlignmentType.CENTER, spacing: { after: 150 } }),
            new docx.Paragraph({ children: [new docx.TextRun({ text: "RINGKASAN KEGIATAN", bold: true, size: 20 })], spacing: { after: 100 } }),
            summaryTable,
            new docx.Paragraph({ children: [new docx.TextRun({ text: "RINCIAN TRANSAKSI", bold: true, size: 20 })], spacing: { before: 200, after: 100 } }),
            new docx.Table({ borders: borders, rows: tableRows, width: { size: 100, type: docx.WidthType.PERCENTAGE } })
        ];

        if (transaksiBerlampiran.length > 0) {
            docChildren.push(new docx.Paragraph({ children: [new docx.TextRun({ text: "LAMPIRAN BUKTI NOTA & TRANSAKSI", bold: true, size: 22 })], alignment: docx.AlignmentType.LEFT, spacing: { before: 300, after: 150 } }));
            
            for (const [idx, t] of transaksiBerlampiran.entries()) {
                docChildren.push(new docx.Paragraph({ children: [new docx.TextRun({ text: `Lampiran ${idx + 1}: ${t.keterangan} (${t.divisi}) - Total: ${formatRupiah(t.totalNominal)}`, bold: true })], alignment: docx.AlignmentType.LEFT, spacing: { after: 100 } }));
                try {
                    const base64Data = t.notaData.dataUrl.split(',')[1];
                    const imageBuffer = Uint8Array.from(atob(base64Data), c => c.charCodeAt(0));
                    docChildren.push(new docx.Paragraph({
                        children: [
                            new docx.ImageRun({
                                data: imageBuffer,
                                transformation: { width: 300, height: 200 }
                            })
                        ],
                        alignment: docx.AlignmentType.LEFT,
                        spacing: { after: 200 }
                    }));
                } catch (e) {
                    console.error("Gagal merender gambar ke Word", e);
                }
            }
        }

        const metadataString = encodeURIComponent(JSON.stringify({ jumlahPeserta: document.getElementById('jumlahPeserta').value || 0, transaksi: transaksi, budgets: (getDraftAktif() || {}).budgets || {} }));
        docChildren.push(new docx.Paragraph({ children: [new docx.TextRun({ text: `GCC_DATA_START:::${metadataString}:::GCC_DATA_END`, color: "FFFFFF", size: 2 })] }));

        const doc = new docx.Document({ sections: [{ children: docChildren }] });
        return await docx.Packer.toBlob(doc);
    }
}

async function exportToExcel() {
    const fileName = getFileNameFormat('xlsx');
    const blob = await generateFileBlob('xlsx');
    saveAs(blob, fileName);
    showToast(`File ${fileName} berhasil di-download!`, "success");
}

async function exportToPDF() {
    const fileName = getFileNameFormat('pdf');
    const blob = await generateFileBlob('pdf');
    saveAs(blob, fileName);
    showToast(`File ${fileName} berhasil di-download!`, "success");
}

async function exportToWord() {
    const fileName = getFileNameFormat('docx');
    const blob = await generateFileBlob('docx');
    saveAs(blob, fileName);
    showToast(`File ${fileName} berhasil di-download!`, "success");
}

/* ---------------------------------------------------------------------
   UPLOAD KE GOOGLE DRIVE
   Agar file terbaca saat dibuka dengan Google Docs / Sheets / Drive:
   - tipe MIME dikirim eksplisit per format (tidak bergantung pada blob)
   - isi file diperiksa dulu (Word/Excel = arsip ZIP "PK", PDF = "%PDF")
   - opsional: minta Drive membuat salinan Google Docs / Google Sheets
   --------------------------------------------------------------------- */
const MIME_DRIVE = {
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    pdf: 'application/pdf'
};
const GOOGLE_TIPE = { docx: 'document', xlsx: 'spreadsheet', pdf: null };
const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyt9AmgH3Ks5dLjPNDi2oK3gR5PD3M9MDKbtLGrKUaKMgCsmB3WfVS4Fc3MTjHJd1_H/exec";

function bacaBlobSebagai(blob, mode) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onerror = () => reject(reader.error || new Error('Gagal membaca file'));
        reader.onload = () => resolve(reader.result);
        if (mode === 'buffer') reader.readAsArrayBuffer(blob);
        else reader.readAsDataURL(blob);
    });
}

async function fileDriveValid(blob, formatType) {
    if (!blob || !blob.size) return false;
    const kepala = new Uint8Array(await bacaBlobSebagai(blob.slice(0, 5), 'buffer'));
    const teks = String.fromCharCode.apply(null, kepala);
    return formatType === 'pdf' ? teks.indexOf('%PDF') === 0 : teks.indexOf('PK') === 0;
}

async function uploadKeDrive(formatType) {
    const dept = document.getElementById('driveDeptSelect').value;
    const category = document.getElementById('driveCategorySelect').value;
    const fileName = getFileNameFormat(formatType);

    const targetFolderId = FOLDER_MAP[dept]?.[category];

    if (!targetFolderId) {
        showToast("ID Folder Google Drive belum terdaftar!", "error");
        return;
    }

    showToast(`⏳ Menyiapkan dan mengunggah ${fileName} ke Drive...`, "info");

    try {
        const fileBlob = await generateFileBlob(formatType);

        if (!(await fileDriveValid(fileBlob, formatType))) {
            showToast("File gagal dibuat dengan benar, unggahan dibatalkan.", "error");
            return;
        }

        const dataUrl = await bacaBlobSebagai(fileBlob, 'url');
        const base64Data = String(dataUrl).split(',')[1];

        const kolomKonversi = document.getElementById('driveKonversi');
        const googleType = GOOGLE_TIPE[formatType];
        const konversi = !!(kolomKonversi && kolomKonversi.checked && googleType);

        const payload = {
            folderId: targetFolderId,
            fileName: fileName,
            ext: formatType,
            mimeType: MIME_DRIVE[formatType] || fileBlob.type,
            size: fileBlob.size,
            convert: konversi,
            googleType: konversi ? googleType : null,
            fileData: base64Data
        };

        try {
            const response = await fetch(GOOGLE_SCRIPT_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify(payload)
            });

            const result = await response.json();
            if (result.status === "success" || response.ok) {
                let pesan = `Berhasil mengirim "${fileName}" ke Google Drive!`;
                if (konversi && result.converted) pesan += formatType === 'docx' ? ' Salinan Google Docs juga dibuat.' : ' Salinan Google Sheets juga dibuat.';
                showToast(pesan, "success");
            } else {
                showToast(`Gagal: ${result.message || 'Unknown error'}`, "error");
            }
        } catch (err) {
            console.error(err);
            showToast("Permintaan unggah terkirim ke Google Drive!", "success");
        }
    } catch (err) {
        console.error(err);
        showToast("Gagal memproses file untuk Google Drive.", "error");
    }
}

/* =====================================================================
   ★ BAGIAN BARU & DIPERBARUI
   Semua fungsi lama yang dipakai tombol di index.html tetap ada dengan
   nama yang sama. Bagian di bawah ini menambah fitur baru dan
   menggantikan beberapa fungsi lama dengan versi yang lebih kuat.
   ===================================================================== */

const REDUCED_MOTION = !!(window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches);
const PALET_CHART = ['#6366f1', '#10b981', '#f59e0b', '#ec4899', '#8b5cf6', '#06b6d4', '#f43f5e', '#84cc16'];

let chartSaldo = null;
let chartTop = null;
let lastHealthScore = null;
let calcEkspresi = '';
let calcNilaiAktif = 0;
let paletHasil = [];
let paletIndex = 0;
let partikelKonfeti = [];
let konfetiRaf = null;

/* ---------------------------------------------------------------------
   A. UTILITAS UMUM
   --------------------------------------------------------------------- */
function getDraftAktif() {
    return listDraft.find(d => d.id === currentDraftId);
}

function klon(obj) {
    return JSON.parse(JSON.stringify(obj));
}

function esc(teks) {
    return String(teks ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function potong(teks, maks) {
    const s = String(teks || '');
    return s.length > maks ? s.slice(0, maks - 1) + '…' : s;
}

function setTeks(id, teks) {
    const el = document.getElementById(id);
    if (el) el.textContent = teks;
}

function singkatRupiah(n) {
    const a = Math.abs(n);
    const tanda = n < 0 ? '-' : '';
    if (a >= 1e9) return tanda + (a / 1e9).toLocaleString('id-ID', { maximumFractionDigits: 1 }) + ' M';
    if (a >= 1e6) return tanda + (a / 1e6).toLocaleString('id-ID', { maximumFractionDigits: 1 }) + ' jt';
    if (a >= 1e3) return tanda + (a / 1e3).toLocaleString('id-ID', { maximumFractionDigits: 0 }) + ' rb';
    return tanda + a;
}

function formatWaktu(ms) {
    return new Date(ms).toLocaleString('id-ID', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
}

function bacaFileSebagaiDataUrl(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

/* Foto nota dikecilkan otomatis (maks 1280px, JPEG) supaya penyimpanan browser tidak cepat penuh */
async function kompresGambar(file, maks = 1280, kualitas = 0.8) {
    const asli = await bacaFileSebagaiDataUrl(file);
    if (!/^data:image\/(png|jpe?g|webp|bmp)/i.test(asli)) return { dataUrl: asli };
    try {
        const img = await new Promise((resolve, reject) => {
            const i = new Image();
            i.onload = () => resolve(i);
            i.onerror = reject;
            i.src = asli;
        });
        const skala = Math.min(1, maks / Math.max(img.naturalWidth, img.naturalHeight));
        const w = Math.max(1, Math.round(img.naturalWidth * skala));
        const h = Math.max(1, Math.round(img.naturalHeight * skala));
        const canvas = document.createElement('canvas');
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, w, h);
        ctx.drawImage(img, 0, 0, w, h);
        const hasil = canvas.toDataURL('image/jpeg', kualitas);
        return { dataUrl: hasil.length < asli.length ? hasil : asli };
    } catch (e) {
        return { dataUrl: asli };
    }
}

async function salinKeClipboard(teks) {
    try {
        await navigator.clipboard.writeText(teks);
        return true;
    } catch (e) {
        const ta = document.createElement('textarea');
        ta.value = teks;
        ta.className = 'hidden-input';
        ta.style.position = 'fixed';
        document.body.appendChild(ta);
        ta.select();
        let ok = false;
        try { ok = document.execCommand('copy'); } catch (err) { ok = false; }
        ta.remove();
        return ok;
    }
}

/* ---------------------------------------------------------------------
   B. TERBILANG (angka → tulisan rupiah)
   --------------------------------------------------------------------- */
function terbilang(angka) {
    const n = Math.floor(Math.abs(Number(angka)) || 0);
    if (n === 0) return 'nol';
    const satuan = ['', 'satu', 'dua', 'tiga', 'empat', 'lima', 'enam', 'tujuh', 'delapan', 'sembilan', 'sepuluh', 'sebelas'];

    function baca(x) {
        if (x < 12) return satuan[x];
        if (x < 20) return satuan[x - 10] + ' belas';
        if (x < 100) return satuan[Math.floor(x / 10)] + ' puluh' + (x % 10 ? ' ' + satuan[x % 10] : '');
        if (x < 200) return 'seratus' + (x % 100 ? ' ' + baca(x % 100) : '');
        if (x < 1000) return satuan[Math.floor(x / 100)] + ' ratus' + (x % 100 ? ' ' + baca(x % 100) : '');
        if (x < 2000) return 'seribu' + (x % 1000 ? ' ' + baca(x % 1000) : '');
        if (x < 1e6) return baca(Math.floor(x / 1000)) + ' ribu' + (x % 1000 ? ' ' + baca(x % 1000) : '');
        if (x < 1e9) return baca(Math.floor(x / 1e6)) + ' juta' + (x % 1e6 ? ' ' + baca(x % 1e6) : '');
        if (x < 1e12) return baca(Math.floor(x / 1e9)) + ' miliar' + (x % 1e9 ? ' ' + baca(x % 1e9) : '');
        return baca(Math.floor(x / 1e12)) + ' triliun' + (x % 1e12 ? ' ' + baca(x % 1e12) : '');
    }

    if (n >= 1e15) return 'angka terlalu besar';
    return baca(n);
}

function formatTerbilang(angka) {
    const teks = terbilang(angka) + ' rupiah';
    return teks.charAt(0).toUpperCase() + teks.slice(1);
}

function updateTerbilangTool() {
    const nilai = parseFloat(document.getElementById('terbilangInput').value);
    const out = document.getElementById('terbilangOutput');
    if (isNaN(nilai) || nilai <= 0) {
        out.textContent = 'Masukkan nominal untuk melihat terbilangnya.';
        out.classList.add('kosong');
        return;
    }
    out.classList.remove('kosong');
    out.textContent = formatTerbilang(nilai);
}

async function salinTerbilang() {
    const out = document.getElementById('terbilangOutput');
    if (out.classList.contains('kosong')) {
        showToast('Isi nominal terlebih dahulu', 'error');
        return;
    }
    const ok = await salinKeClipboard(out.textContent);
    showToast(ok ? 'Terbilang disalin!' : 'Gagal menyalin teks', ok ? 'success' : 'error');
}

/* ---------------------------------------------------------------------
   C. NOTIFIKASI (TOAST) — kini mendukung tipe warn & tombol aksi
   --------------------------------------------------------------------- */
function showToast(message, type = 'info', aksi = null) {
    const container = document.getElementById('toastContainer');
    if (!container) return;
    while (container.children.length >= 4) container.firstElementChild.remove();

    const toast = document.createElement('div');
    toast.className = `toast ${type}` + (aksi ? ' has-action' : '');

    const titik = document.createElement('span');
    titik.className = 'toast-dot';
    const isi = document.createElement('span');
    isi.className = 'toast-msg';
    isi.textContent = message;
    toast.append(titik, isi);

    if (aksi) {
        const tombol = document.createElement('button');
        tombol.type = 'button';
        tombol.className = 'toast-btn';
        tombol.textContent = aksi.label;
        tombol.onclick = () => { aksi.fn(); toast.remove(); };
        toast.appendChild(tombol);
    }

    container.appendChild(toast);
    setTimeout(() => toast.remove(), aksi ? 5200 : 3200);
}

/* ---------------------------------------------------------------------
   D. TEMA & NAVIGASI
   --------------------------------------------------------------------- */
function terapkanLabelTema() {
    const gelap = document.body.classList.contains('dark-mode');
    const tombol = document.getElementById('themeToggle');
    if (tombol) tombol.innerText = gelap ? 'Mode Terang' : 'Mode Gelap';
    const meta = document.querySelector('meta[name="theme-color"]');
    if (meta) meta.setAttribute('content', gelap ? '#0b0716' : '#f7f3ff');
}

function initTheme() {
    const tersimpan = localStorage.getItem('gcc_theme');
    const gelap = tersimpan ? tersimpan === 'dark' : !!(window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches);
    document.body.classList.toggle('dark-mode', gelap);
    terapkanLabelTema();
}

function toggleTheme() {
    document.body.classList.toggle('dark-mode');
    const gelap = document.body.classList.contains('dark-mode');
    localStorage.setItem('gcc_theme', gelap ? 'dark' : 'light');
    terapkanLabelTema();
    initCharts();
    updateTampilan();
}

function switchTab(tabId, btnElement) {
    document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.nav-tab, .mobile-menu-item[data-tab]').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.tab === tabId);
    });
    document.getElementById(tabId).classList.add('active');
    window.scrollTo({ top: 0, behavior: REDUCED_MOTION ? 'auto' : 'smooth' });

    if (tabId === 'tab-analisis') setTimeout(() => updateTampilan(), 50);
    if (tabId === 'tab-tools') renderBudget();
}

function toggleMobileMenu() {
    const dropdown = document.getElementById('mobileMenuDropdown');
    dropdown.style.display = dropdown.style.display === 'flex' ? 'none' : 'flex';
}

function switchTabMobile(tabId, btnElement) {
    switchTab(tabId, btnElement);
    document.getElementById('mobileMenuDropdown').style.display = 'none';
}

/* ---------------------------------------------------------------------
   Rincian Transaksi: di HP tampil di halaman sendiri, di PC tetap di
   halaman Pencatatan. Elemen yang sama dipindah (bukan digandakan) agar
   semua fungsi (filter, edit, hapus, undo) tetap bekerja.
   --------------------------------------------------------------------- */
const MQ_HP = window.matchMedia('(max-width: 768px)');

function posisikanRincian() {
    const bagian = document.getElementById('rincianSection');
    const tujuan = document.getElementById(MQ_HP.matches ? 'rincianMobileSlot' : 'rincianDesktopSlot');
    if (bagian && tujuan && bagian.parentNode !== tujuan) tujuan.appendChild(bagian);

    // Layar melebar ke ukuran PC saat halaman Rincian terbuka -> kembali ke Pencatatan
    const halamanRincian = document.getElementById('tab-rincian');
    if (!MQ_HP.matches && halamanRincian && halamanRincian.classList.contains('active')) {
        switchTab('tab-pencatatan');
    }
}

function bukaRincian() {
    if (MQ_HP.matches) {
        switchTab('tab-rincian');
    } else {
        switchTab('tab-pencatatan');
        setTimeout(() => {
            const bagian = document.getElementById('rincianSection');
            if (bagian) bagian.scrollIntoView({ behavior: REDUCED_MOTION ? 'auto' : 'smooth', block: 'start' });
        }, 120);
    }
}

function updateRincianLink() {
    setTeks('rincianLinkSub', transaksi.length ? `${transaksi.length} transaksi tercatat` : 'Belum ada transaksi');
}

/* ---------------------------------------------------------------------
   E. ANGKA BERANIMASI (count-up) & RINGKASAN
   --------------------------------------------------------------------- */
function animasiAngka(id, target, format) {
    const el = document.getElementById(id);
    if (!el) return;
    const dari = typeof el._nilai === 'number' ? el._nilai : 0;
    el._nilai = target;
    cancelAnimationFrame(el._raf);
    if (REDUCED_MOTION || dari === target) {
        el.textContent = format(target);
        return;
    }
    const mulai = performance.now();
    const durasi = 650;
    const langkah = (t) => {
        const p = Math.min((t - mulai) / durasi, 1);
        const e = 1 - Math.pow(1 - p, 3);
        el.textContent = p >= 1 ? format(target) : format(Math.round(dari + (target - dari) * e));
        if (p < 1) el._raf = requestAnimationFrame(langkah);
    };
    el._raf = requestAnimationFrame(langkah);
}

function updateSummary() {
    const totalPeserta = Number(document.getElementById('jumlahPeserta').value) || 0;
    let totalMasuk = 0, totalKeluar = 0;

    transaksi.forEach(t => {
        if (t.jenis === 'Masuk') totalMasuk += Number(t.totalNominal) || 0;
        if (t.jenis === 'Keluar') totalKeluar += Number(t.totalNominal) || 0;
    });

    const saldoAkhir = totalMasuk - totalKeluar;
    animasiAngka('dispPeserta', totalPeserta, v => `${v} Orang`);
    animasiAngka('dispMasuk', totalMasuk, formatRupiah);
    animasiAngka('dispKeluar', totalKeluar, formatRupiah);
    animasiAngka('dispSaldo', saldoAkhir, formatRupiah);

    const saldoEl = document.getElementById('dispSaldo');
    saldoEl.classList.toggle('saldo-pos', saldoAkhir > 0);
    saldoEl.classList.toggle('saldo-neg', saldoAkhir < 0);

    const percent = totalMasuk > 0 ? Math.min(Math.round((totalKeluar / totalMasuk) * 100), 100) : 0;
    const bar = document.getElementById('progressBar');
    bar.style.width = percent + '%';
    document.getElementById('progressText').innerText = percent + '%';
    bar.style.background = percent > 90 ? '#f43f5e' : 'linear-gradient(135deg, #6366f1 0%, #4f46e5 100%)';

    return { totalPeserta, totalMasuk, totalKeluar, saldoAkhir };
}

function hitungSubtotal(jenis) {
    const isMasuk = jenis === 'Masuk';
    const harga = parseFloat(document.getElementById(isMasuk ? 'hargaMasuk' : 'hargaKeluar').value) || 0;
    const qty = parseFloat(document.getElementById(isMasuk ? 'jumlahMasuk' : 'jumlahKeluar').value) || 0;
    const total = harga * qty;

    document.getElementById(isMasuk ? 'subtotalMasuk' : 'subtotalKeluar').innerText = `Total Item: ${formatRupiah(total)}`;
    const terbilangEl = document.getElementById(isMasuk ? 'terbilangMasuk' : 'terbilangKeluar');
    if (terbilangEl) terbilangEl.textContent = total >= 1 ? formatTerbilang(total) : '';
    return total;
}

function hitungSimulasiHTM() {
    const summary = updateSummary();
    const targetPeserta = parseFloat(document.getElementById('simPeserta').value) || 0;
    const sponsor = parseFloat(document.getElementById('simSponsor').value) || 0;
    const hasilEl = document.getElementById('simulasiResult');

    if (targetPeserta <= 0) {
        hasilEl.innerText = "Masukkan jumlah estimasi peserta yang valid.";
        return;
    }

    const sisaBeban = summary.totalKeluar - sponsor;
    if (sisaBeban <= 0) {
        hasilEl.innerText = "Dana Sponsorship/Kas sudah menutupi seluruh pengeluaran! HTM Peserta bisa Rp 0 (Gratis).";
    } else {
        const htmPerPeserta = Math.ceil((sisaBeban / targetPeserta) / 1000) * 1000;
        hasilEl.innerHTML = `
            <b>Hasil Simulasi:</b><br>
            • Sisa Beban Pengeluaran: <b>${formatRupiah(sisaBeban)}</b><br>
            • Estimasi HTM Minimal per Peserta: <b class="hasil-hijau">${formatRupiah(htmPerPeserta)}</b>
        `;
    }
}

/* ---------------------------------------------------------------------
   F. GRAFIK (Chart.js) — mengikuti tema terang/gelap
   --------------------------------------------------------------------- */
function warnaChart() {
    const gelap = document.body.classList.contains('dark-mode');
    return {
        gelap,
        teks: gelap ? '#cfc8ee' : '#4a4070',
        grid: gelap ? 'rgba(200, 175, 255, 0.12)' : 'rgba(110, 90, 200, 0.14)'
    };
}

function nilaiTitik(ctx) {
    if (typeof ctx.parsed === 'number') return ctx.parsed;
    return ctx.chart.options.indexAxis === 'y' ? ctx.parsed.x : ctx.parsed.y;
}

/* Plugin kecil: tulis "Belum ada data" di tengah grafik yang masih kosong */
const pluginKosong = {
    id: 'pluginKosong',
    afterDraw(chart) {
        const total = chart.data.datasets.reduce((jml, ds) => jml + ds.data.reduce((a, v) => a + Math.abs(Number(v) || 0), 0), 0);
        if (total > 0 || !chart.chartArea) return;
        const { ctx, chartArea } = chart;
        ctx.save();
        ctx.fillStyle = warnaChart().teks;
        ctx.globalAlpha = 0.7;
        ctx.font = "500 13px 'Inter', sans-serif";
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText('Belum ada data', (chartArea.left + chartArea.right) / 2, (chartArea.top + chartArea.bottom) / 2);
        ctx.restore();
    }
};

function hancurkanChart() {
    [chartComp, chartDiv, chartSaldo, chartTop].forEach(c => { if (c) c.destroy(); });
    chartComp = chartDiv = chartSaldo = chartTop = null;
}

function initCharts() {
    if (typeof Chart === 'undefined') {
        console.warn('Chart.js belum termuat, grafik dinonaktifkan.');
        return;
    }
    hancurkanChart();
    const w = warnaChart();
    Chart.defaults.color = w.teks;
    Chart.defaults.borderColor = w.grid;
    Chart.defaults.font.family = "'Inter', sans-serif";

    const tooltipRupiah = { callbacks: { label: ctx => ' ' + formatRupiah(nilaiTitik(ctx)) } };
    const skalaUang = {
        beginAtZero: true,
        ticks: { color: w.teks, callback: v => singkatRupiah(v) },
        grid: { color: w.grid },
        border: { display: false }
    };
    const skalaKategori = { ticks: { color: w.teks }, grid: { display: false }, border: { display: false } };

    chartComp = new Chart(document.getElementById('chartComparison').getContext('2d'), {
        type: 'bar',
        data: {
            labels: ['Total Masuk', 'Total Keluar'],
            datasets: [{ label: 'Nominal (Rp)', data: [0, 0], backgroundColor: ['#10b981', '#f43f5e'], borderRadius: 10, maxBarThickness: 90 }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { display: false }, tooltip: tooltipRupiah },
            scales: { y: skalaUang, x: skalaKategori }
        },
        plugins: [pluginKosong]
    });

    chartDiv = new Chart(document.getElementById('chartDivisi').getContext('2d'), {
        type: 'doughnut',
        data: {
            labels: [],
            datasets: [{ data: [], backgroundColor: PALET_CHART, borderWidth: 2, borderColor: w.gelap ? '#160e2e' : '#ffffff', hoverOffset: 8 }]
        },
        options: {
            responsive: true, maintainAspectRatio: false, cutout: '64%',
            plugins: {
                legend: { position: 'bottom', labels: { color: w.teks, usePointStyle: true, boxWidth: 8, padding: 14 } },
                tooltip: tooltipRupiah
            }
        },
        plugins: [pluginKosong]
    });

    const ctxSaldo = document.getElementById('chartSaldo').getContext('2d');
    const gradSaldo = ctxSaldo.createLinearGradient(0, 0, 0, 240);
    gradSaldo.addColorStop(0, 'rgba(155, 92, 246, 0.45)');
    gradSaldo.addColorStop(1, 'rgba(155, 92, 246, 0)');
    chartSaldo = new Chart(ctxSaldo, {
        type: 'line',
        data: {
            labels: [],
            datasets: [{ label: 'Saldo', data: [], borderColor: '#9b5cf6', backgroundColor: gradSaldo, fill: true, tension: 0.35, borderWidth: 3, pointRadius: 3, pointHoverRadius: 6, pointBackgroundColor: '#ec5fa8' }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        title: ctx => { const t = transaksi[ctx[0].dataIndex]; return t ? t.keterangan : ''; },
                        label: ctx => ' Saldo: ' + formatRupiah(nilaiTitik(ctx))
                    }
                }
            },
            scales: { y: { ...skalaUang, beginAtZero: false }, x: { ...skalaKategori, ticks: { color: w.teks, maxTicksLimit: 8 } } }
        },
        plugins: [pluginKosong]
    });

    chartTop = new Chart(document.getElementById('chartTop').getContext('2d'), {
        type: 'bar',
        data: { labels: [], datasets: [{ label: 'Nominal (Rp)', data: [], backgroundColor: PALET_CHART, borderRadius: 8, maxBarThickness: 28 }] },
        options: {
            indexAxis: 'y', responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        title: ctx => { const it = ctx[0].chart._items; return it && it[ctx[0].dataIndex] ? it[ctx[0].dataIndex].keterangan : ''; },
                        label: ctx => ' ' + formatRupiah(nilaiTitik(ctx))
                    }
                }
            },
            scales: { x: skalaUang, y: skalaKategori }
        },
        plugins: [pluginKosong]
    });
}

function updateCharts() {
    if (!chartComp || !chartDiv || !chartSaldo || !chartTop) return;

    let masuk = 0, keluar = 0;
    const perDivisi = {};
    transaksi.forEach(t => {
        const nilai = Number(t.totalNominal) || 0;
        if (t.jenis === 'Masuk') {
            masuk += nilai;
        } else if (t.jenis === 'Keluar') {
            keluar += nilai;
            const k = String(t.divisi || 'CONTER').trim().toUpperCase();
            perDivisi[k] = (perDivisi[k] || 0) + nilai;
        }
    });

    chartComp.data.datasets[0].data = [masuk, keluar];
    chartComp.update();

    const urutDivisi = Object.entries(perDivisi).sort((a, b) => b[1] - a[1]);
    chartDiv.data.labels = urutDivisi.map(d => d[0]);
    chartDiv.data.datasets[0].data = urutDivisi.map(d => d[1]);
    chartDiv.update();

    let saldo = 0;
    const saldoBerjalan = transaksi.map(t => {
        saldo += (t.jenis === 'Masuk' ? 1 : -1) * (Number(t.totalNominal) || 0);
        return saldo;
    });
    chartSaldo.data.labels = transaksi.map((_, i) => `#${i + 1}`);
    chartSaldo.data.datasets[0].data = saldoBerjalan;
    chartSaldo.update();

    const top = transaksi.filter(t => t.jenis === 'Keluar').sort((a, b) => b.totalNominal - a.totalNominal).slice(0, 5);
    chartTop._items = top;
    chartTop.data.labels = top.map(t => potong(t.keterangan, 22));
    chartTop.data.datasets[0].data = top.map(t => t.totalNominal);
    chartTop.update();
}

/* ---------------------------------------------------------------------
   G. PAGU ANGGARAN PER DIVISI
   --------------------------------------------------------------------- */
function pengeluaranPerDivisi() {
    const peta = {};
    transaksi.forEach(t => {
        if (t.jenis === 'Keluar') {
            const k = String(t.divisi || 'CONTER').trim().toUpperCase();
            peta[k] = (peta[k] || 0) + (Number(t.totalNominal) || 0);
        }
    });
    return peta;
}

function renderBudget() {
    const wadah = document.getElementById('budgetList');
    const saran = document.getElementById('budgetSuggest');
    if (!wadah || !saran) return;

    const draft = getDraftAktif();
    const budgets = (draft && draft.budgets) || {};
    const pakai = pengeluaranPerDivisi();
    const kunci = Object.keys(budgets);

    if (!kunci.length) {
        wadah.innerHTML = '<p class="empty-note">Belum ada pagu untuk draf ini. Tambahkan di atas, atau pilih salah satu divisi yang sudah punya pengeluaran.</p>';
    } else {
        wadah.innerHTML = kunci.map(k => {
            const pagu = Number(budgets[k]) || 0;
            const terpakai = pakai[k] || 0;
            const persen = pagu > 0 ? (terpakai / pagu) * 100 : 0;
            const status = persen > 100 ? 'over' : persen >= 80 ? 'warn' : 'ok';
            const sisa = pagu - terpakai;
            const catatan = status === 'over'
                ? `Melebihi pagu ${formatRupiah(-sisa)}`
                : `Sisa ${formatRupiah(sisa)} · ${Math.round(persen)}% terpakai`;
            return `
                <div class="budget-item ${status}">
                    <div class="budget-top">
                        <strong>${esc(k)}</strong>
                        <span class="budget-nums">${formatRupiah(terpakai)} / ${formatRupiah(pagu)}</span>
                        <button type="button" class="budget-del" data-hapus-budget="${esc(k)}" title="Hapus pagu" aria-label="Hapus pagu ${esc(k)}">&times;</button>
                    </div>
                    <div class="budget-bar"><span data-w="${Math.min(persen, 100).toFixed(1)}"></span></div>
                    <div class="budget-note">${catatan}</div>
                </div>`;
        }).join('');
        setTimeout(() => {
            wadah.querySelectorAll('.budget-bar span[data-w]').forEach(el => { el.style.width = el.dataset.w + '%'; });
        }, 30);
    }

    const tanpaPagu = Object.keys(pakai).filter(k => !(k in budgets));
    saran.innerHTML = tanpaPagu.length
        ? '<span class="muted">Divisi belum punya pagu:</span> ' + tanpaPagu.map(k => `<button type="button" class="chip" data-divisi="${esc(k)}">${esc(k)}</button>`).join('')
        : '';
}

function simpanBudget() {
    const draft = getDraftAktif();
    if (!draft) return;
    const nama = document.getElementById('budgetDivisi').value.trim().toUpperCase();
    const nominal = parseFloat(document.getElementById('budgetNominal').value);
    if (!nama || isNaN(nominal) || nominal <= 0) {
        showToast('Isi nama divisi dan pagu yang valid!', 'error');
        return;
    }
    draft.budgets = draft.budgets || {};
    draft.budgets[nama] = nominal;
    simpanListDraft();
    document.getElementById('budgetDivisi').value = '';
    document.getElementById('budgetNominal').value = '';
    updateTampilan();
    showToast(`Pagu ${nama} disimpan: ${formatRupiah(nominal)}`, 'success');
}

function hapusBudget(kunci) {
    const draft = getDraftAktif();
    if (!draft || !draft.budgets) return;
    delete draft.budgets[kunci];
    simpanListDraft();
    updateTampilan();
    showToast(`Pagu ${kunci} dihapus`, 'info');
}

function cekPagu(jenis, divisi) {
    if (jenis !== 'Keluar') return;
    const draft = getDraftAktif();
    const budgets = draft && draft.budgets;
    if (!budgets) return;
    const k = String(divisi).trim().toUpperCase();
    const pagu = Number(budgets[k]);
    if (!pagu) return;
    const terpakai = pengeluaranPerDivisi()[k] || 0;
    if (terpakai > pagu) {
        showToast(`Pagu ${k} terlampaui ${formatRupiah(terpakai - pagu)}!`, 'error');
    } else if (terpakai >= pagu * 0.8) {
        showToast(`Pagu ${k} sudah terpakai ${Math.round((terpakai / pagu) * 100)}%`, 'warn');
    }
}

/* ---------------------------------------------------------------------
   H. WAWASAN OTOMATIS & KESIAPAN LAPORAN
   --------------------------------------------------------------------- */
function updateInsights(summary) {
    const keluarList = transaksi.filter(t => t.jenis === 'Keluar');
    const masukList = transaksi.filter(t => t.jenis === 'Masuk');
    const peserta = Number(summary.totalPeserta) || 0;

    setTeks('statBiaya', peserta > 0 ? formatRupiah(Math.round(summary.totalKeluar / peserta)) : '—');
    setTeks('statBiayaSub', peserta > 0 ? `dari ${peserta} peserta` : 'Isi jumlah peserta dulu');

    const terbesar = keluarList.reduce((a, t) => (!a || t.totalNominal > a.totalNominal ? t : a), null);
    setTeks('statTerbesar', terbesar ? formatRupiah(terbesar.totalNominal) : '—');
    setTeks('statTerbesarSub', terbesar ? potong(terbesar.keterangan, 34) : 'Belum ada pengeluaran');

    setTeks('statJumlah', String(transaksi.length));
    setTeks('statJumlahSub', `${masukList.length} masuk · ${keluarList.length} keluar`);

    const bernota = keluarList.filter(t => t.notaData && t.notaData.dataUrl).length;
    setTeks('statNota', keluarList.length ? `${bernota}/${keluarList.length}` : '—');
    setTeks('statNotaSub', keluarList.length ? `${Math.round((bernota / keluarList.length) * 100)}% pengeluaran punya nota` : 'Belum ada pengeluaran');

    const perDivisi = Object.entries(pengeluaranPerDivisi()).sort((a, b) => b[1] - a[1]);
    if (perDivisi.length) {
        const [nama, nilai] = perDivisi[0];
        setTeks('statBoros', potong(nama, 18));
        setTeks('statBorosSub', `${formatRupiah(nilai)} · ${Math.round((nilai / summary.totalKeluar) * 100)}% dari total keluar`);
    } else {
        setTeks('statBoros', '—');
        setTeks('statBorosSub', 'Belum ada pengeluaran');
    }

    setTeks('statRata', keluarList.length ? formatRupiah(Math.round(summary.totalKeluar / keluarList.length)) : '—');
    setTeks('statRataSub', `per transaksi keluar (${keluarList.length})`);
}

function hitungKesiapan(summary) {
    const keluarList = transaksi.filter(t => t.jenis === 'Keluar');
    const masukList = transaksi.filter(t => t.jenis === 'Masuk');
    const tanpaNota = keluarList.filter(t => !(t.notaData && t.notaData.dataUrl)).length;
    const draft = getDraftAktif();
    const budgets = (draft && draft.budgets) || {};
    const pakai = pengeluaranPerDivisi();
    const adaBudget = Object.keys(budgets).length > 0;
    const melebihi = Object.keys(budgets).filter(k => (pakai[k] || 0) > budgets[k]);

    const cek = [
        {
            ok: summary.totalPeserta > 0,
            judul: 'Jumlah peserta sudah diisi',
            detail: summary.totalPeserta > 0 ? `${summary.totalPeserta} peserta` : 'Isi pada kolom jumlah peserta'
        },
        {
            ok: masukList.length > 0,
            judul: 'Ada catatan uang masuk',
            detail: masukList.length > 0 ? `${masukList.length} transaksi masuk` : 'Belum ada pemasukan yang dicatat'
        },
        {
            ok: keluarList.length > 0 && tanpaNota === 0,
            judul: 'Semua pengeluaran punya nota',
            detail: keluarList.length === 0 ? 'Belum ada pengeluaran' : tanpaNota === 0 ? `${keluarList.length} nota terlampir` : `${tanpaNota} transaksi belum ada nota`
        },
        {
            ok: transaksi.length > 0 && summary.saldoAkhir >= 0,
            judul: 'Saldo akhir tidak defisit',
            detail: transaksi.length === 0 ? 'Belum ada transaksi' : summary.saldoAkhir >= 0 ? `Saldo ${formatRupiah(summary.saldoAkhir)}` : `Defisit ${formatRupiah(Math.abs(summary.saldoAkhir))}`
        },
        adaBudget
            ? { ok: melebihi.length === 0, judul: 'Pengeluaran sesuai pagu divisi', detail: melebihi.length === 0 ? 'Semua divisi dalam batas pagu' : `Melebihi pagu: ${melebihi.join(', ')}` }
            : { ok: null, judul: 'Pagu per divisi (opsional)', detail: 'Belum diatur, bisa diatur di tab Tools' }
    ];

    const dihitung = cek.filter(c => c.ok !== null);
    const lulus = dihitung.filter(c => c.ok).length;
    const skor = dihitung.length ? Math.round((lulus / dihitung.length) * 100) : 0;
    return { cek, skor };
}

function updateHealth(summary) {
    const { cek, skor } = hitungKesiapan(summary);

    const arc = document.getElementById('healthArc');
    if (arc) {
        const keliling = 2 * Math.PI * 42;
        arc.style.strokeDasharray = keliling;
        arc.style.strokeDashoffset = keliling * (1 - skor / 100);
    }
    animasiAngka('healthScore', skor, v => `${v}%`);
    setTeks('healthVerdict', skor === 100 ? 'Siap dilaporkan' : skor >= 60 ? 'Hampir siap' : 'Perlu dilengkapi');

    const daftar = document.getElementById('healthList');
    if (daftar) {
        daftar.innerHTML = cek.map(c => `
            <li class="health-item ${c.ok === null ? 'skip' : c.ok ? 'ok' : 'bad'}">
                <span class="health-mark"></span>
                <div><b>${esc(c.judul)}</b><small>${esc(c.detail)}</small></div>
            </li>`).join('');
    }

    if (lastHealthScore !== null && lastHealthScore < 100 && skor === 100 && transaksi.length) {
        luncurkanKonfeti();
        showToast('Laporan siap! Semua pengecekan terpenuhi.', 'success');
    }
    lastHealthScore = skor;
}

/* ---------------------------------------------------------------------
   I. BANNER SAMBUTAN
   --------------------------------------------------------------------- */
function updateHero(summary) {
    const draft = getDraftAktif();
    const jam = new Date().getHours();
    const sapa = jam < 11 ? 'Selamat pagi' : jam < 15 ? 'Selamat siang' : jam < 18 ? 'Selamat sore' : 'Selamat malam';
    setTeks('heroGreeting', `${sapa}!`);
    setTeks('heroTitle', draft ? draft.nama : 'Rekapan Keuangan');

    const jmlMasuk = transaksi.filter(t => t.jenis === 'Masuk').length;
    const jmlKeluar = transaksi.length - jmlMasuk;
    setTeks('heroSub', transaksi.length
        ? `${transaksi.length} transaksi · ${jmlMasuk} uang masuk · ${jmlKeluar} uang keluar`
        : 'Belum ada transaksi, mulai catat lewat form di bawah');

    let kelas = 'kosong';
    let teks = 'Belum ada transaksi';
    if (transaksi.length) {
        if (summary.saldoAkhir > 0) { kelas = 'surplus'; teks = `Surplus ${formatRupiah(summary.saldoAkhir)}`; }
        else if (summary.saldoAkhir < 0) { kelas = 'defisit'; teks = `Defisit ${formatRupiah(Math.abs(summary.saldoAkhir))}`; }
        else { kelas = 'seimbang'; teks = 'Saldo seimbang'; }
    }
    const status = document.getElementById('heroStatus');
    if (status) status.className = `hero-status ${kelas}`;
    setTeks('heroStatusText', teks);
}

/* ---------------------------------------------------------------------
   J. RINGKASAN WHATSAPP, CSV, CADANGAN
   --------------------------------------------------------------------- */
function susunRingkasanWA(summary) {
    const draft = getDraftAktif();
    const nama = draft ? draft.nama : 'Kegiatan';
    const tanggal = new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
    const baris = [
        `*REKAPAN KEUANGAN ${nama.toUpperCase()}*`,
        '_GAMATIKA Coding Club_',
        '',
        `Peserta : ${summary.totalPeserta} orang`,
        `Uang masuk : ${formatRupiah(summary.totalMasuk)}`,
        `Uang keluar : ${formatRupiah(summary.totalKeluar)}`,
        `*Sisa saldo : ${formatRupiah(summary.saldoAkhir)}*`
    ];

    const perDivisi = Object.entries(pengeluaranPerDivisi()).sort((a, b) => b[1] - a[1]);
    if (perDivisi.length) {
        baris.push('', '*Pengeluaran per divisi*');
        perDivisi.forEach(([k, v]) => baris.push(`• ${k}: ${formatRupiah(v)}`));
    }
    baris.push('', `Diperbarui ${tanggal}`);
    return baris.join('\n');
}

function updateWAPreview(summary) {
    const ta = document.getElementById('waPreview');
    if (ta) ta.value = susunRingkasanWA(summary);
}

async function salinRingkasan() {
    const ta = document.getElementById('waPreview');
    const ok = await salinKeClipboard(ta ? ta.value : '');
    showToast(ok ? 'Ringkasan disalin, tinggal tempel di WhatsApp!' : 'Gagal menyalin ringkasan', ok ? 'success' : 'error');
}

function bukaWhatsApp() {
    const ta = document.getElementById('waPreview');
    if (!ta || !ta.value) return;
    window.open('https://wa.me/?text=' + encodeURIComponent(ta.value), '_blank', 'noopener');
}

function exportCSV() {
    if (!transaksi.length) {
        showToast('Belum ada transaksi untuk diekspor', 'error');
        return;
    }
    const sel = v => {
        let s = String(v ?? '');
        if (typeof v === 'string' && /^[=+\-@]/.test(s)) s = "'" + s;
        return /[";\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    };
    const baris = [['Divisi', 'Jenis', 'Keterangan', 'Harga Satuan', 'Jumlah', 'Total Nominal', 'Catatan', 'Ada Nota', 'Waktu Dicatat'].join(';')];
    transaksi.forEach(t => {
        baris.push([
            t.divisi, t.jenis, t.keterangan, t.harga, t.qty, t.totalNominal, t.catatan,
            t.notaData ? 'Ya' : 'Tidak', t.waktu ? new Date(t.waktu).toISOString() : ''
        ].map(sel).join(';'));
    });
    const blob = new Blob(['\uFEFF' + baris.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    const namaFile = getFileNameFormat('csv');
    saveAs(blob, namaFile);
    showToast(`File ${namaFile} berhasil di-download!`, 'success');
}

function backupSemuaDraf() {
    simpanKeStorage(false);
    if (!listDraft.length) {
        showToast('Belum ada draf untuk dicadangkan', 'error');
        return;
    }
    const isi = { aplikasi: 'keuangan-gcc', versi: 1, diekspor: new Date().toISOString(), draf: listDraft };
    const blob = new Blob([JSON.stringify(isi)], { type: 'application/json' });
    saveAs(blob, `Cadangan_Keuangan_GCC_${new Date().toISOString().slice(0, 10)}.json`);
    showToast('Cadangan semua draf berhasil diunduh!', 'success');
}

function rapikanTransaksi(t) {
    if (!t || (t.jenis !== 'Masuk' && t.jenis !== 'Keluar')) return null;
    const harga = Number(t.harga) || 0;
    const qty = Number(t.qty) || 1;
    const adaNota = t.notaData && typeof t.notaData.dataUrl === 'string' && /^data:image\//i.test(t.notaData.dataUrl);
    const hasil = {
        divisi: String(t.divisi || 'CONTER'),
        jenis: t.jenis,
        keterangan: String(t.keterangan || ''),
        harga,
        qty,
        totalNominal: Number(t.totalNominal) || harga * qty,
        notaData: adaNota ? { dataUrl: t.notaData.dataUrl } : null,
        catatan: String(t.catatan || '-')
    };
    if (Number(t.waktu)) hasil.waktu = Number(t.waktu);
    return hasil;
}

async function restoreDraf(event) {
    const file = event.target.files[0];
    event.target.value = '';
    if (!file) return;
    try {
        const data = JSON.parse(await file.text());
        if (!data || !Array.isArray(data.draf)) throw new Error('format');
        const stamp = Date.now();
        const valid = data.draf.filter(d => d && Array.isArray(d.transaksi)).map((d, i) => ({
            id: `draft_${stamp}_${i}`,
            nama: String(d.nama || 'Draf Pulihan'),
            jumlahPeserta: Number(d.jumlahPeserta) || 0,
            budgets: d.budgets && typeof d.budgets === 'object' ? d.budgets : {},
            transaksi: d.transaksi.map(rapikanTransaksi).filter(Boolean)
        }));
        if (!valid.length) throw new Error('kosong');
        if (!confirm(`Tambahkan ${valid.length} draf dari cadangan ini ke daftar draf Anda?\n(Draf yang sudah ada tidak akan ditimpa.)`)) return;

        simpanKeStorage(false);
        listDraft = valid.concat(listDraft);
        currentDraftId = valid[0].id;
        simpanListDraft();
        renderDropdownDraft();
        muatDraftAktif();
        showToast(`${valid.length} draf berhasil dipulihkan!`, 'success');
    } catch (e) {
        console.error(e);
        showToast('File cadangan tidak valid.', 'error');
    }
}

function updateStorageMeter() {
    const teks = document.getElementById('storageText');
    const isi = document.getElementById('storageFill');
    if (!teks || !isi) return;
    let karakter = 0;
    try {
        for (let i = 0; i < localStorage.length; i++) {
            const k = localStorage.key(i);
            karakter += k.length + (localStorage.getItem(k) || '').length;
        }
    } catch (e) { /* abaikan */ }
    const persen = Math.min((karakter / 5e6) * 100, 100);
    teks.textContent = `≈ ${(karakter / 1e6).toFixed(2)} MB dari ±5 MB`;
    isi.style.width = persen + '%';
    isi.classList.toggle('warn', persen >= 70 && persen < 90);
    isi.classList.toggle('over', persen >= 90);
}

/* ---------------------------------------------------------------------
   K. KALKULATOR CEPAT
   --------------------------------------------------------------------- */
function evaluasiEkspresi(str) {
    const tok = str.match(/\d+(?:\.\d+)?|[+\-*/%()]/g) || [];
    let p = 0;
    const intip = () => tok[p];

    function ekspresi() {
        let v = suku();
        while (intip() === '+' || intip() === '-') {
            const op = tok[p++];
            const r = suku();
            v = op === '+' ? v + r : v - r;
        }
        return v;
    }
    function suku() {
        let v = faktor();
        while (intip() === '*' || intip() === '/') {
            const op = tok[p++];
            const r = faktor();
            if (op === '/' && r === 0) throw new Error('bagi nol');
            v = op === '*' ? v * r : v / r;
        }
        return v;
    }
    function faktor() {
        const t = tok[p++];
        if (t === '-') return -faktor();
        if (t === '+') return faktor();
        if (t === '(') {
            const v = ekspresi();
            if (tok[p++] !== ')') throw new Error('kurung');
            return v;
        }
        if (t === undefined || isNaN(Number(t))) throw new Error('sintaks');
        let v = Number(t);
        while (intip() === '%') { p++; v = v / 100; }
        return v;
    }

    const hasil = ekspresi();
    if (p < tok.length) throw new Error('sintaks');
    return hasil;
}

function calcRender() {
    const tampil = calcEkspresi.replace(/\*/g, '×').replace(/\//g, '÷').replace(/-/g, '−');
    document.getElementById('calcExpr').textContent = tampil || '\u00A0';
    try {
        if (calcEkspresi && /[\d%]$/.test(calcEkspresi)) {
            const v = evaluasiEkspresi(calcEkspresi);
            if (isFinite(v)) calcNilaiAktif = Math.round(v * 1e8) / 1e8;
        } else if (!calcEkspresi) {
            calcNilaiAktif = 0;
        }
    } catch (e) { /* biarkan nilai terakhir */ }
    setTeks('calcResult', Number(calcNilaiAktif).toLocaleString('id-ID', { maximumFractionDigits: 6 }));
}

function calcTekan(tombol) {
    const operator = ['+', '-', '*', '/'];
    const akhir = calcEkspresi.slice(-1);

    if (tombol === 'C') {
        calcEkspresi = '';
        calcNilaiAktif = 0;
    } else if (tombol === 'DEL') {
        calcEkspresi = calcEkspresi.slice(0, -1);
    } else if (tombol === '=') {
        try {
            const v = evaluasiEkspresi(calcEkspresi || '0');
            if (!isFinite(v)) throw new Error('tak hingga');
            calcNilaiAktif = Math.round(v * 1e8) / 1e8;
            calcEkspresi = calcNilaiAktif === 0 ? '' : String(calcNilaiAktif);
        } catch (e) {
            showToast('Ekspresi tidak valid', 'error');
        }
    } else if (operator.includes(tombol)) {
        if (!calcEkspresi) {
            if (tombol === '-') calcEkspresi = '-';
        } else if (akhir === '.') {
            calcEkspresi = calcEkspresi.slice(0, -1) + tombol;
        } else if (operator.includes(akhir)) {
            calcEkspresi = calcEkspresi.slice(0, -1) + tombol;
        } else {
            calcEkspresi += tombol;
        }
    } else if (tombol === '%') {
        if (/[\d%]$/.test(calcEkspresi)) calcEkspresi += '%';
    } else if (tombol === '.') {
        const angkaTerakhir = calcEkspresi.split(/[+\-*/]/).pop();
        if (!angkaTerakhir.includes('.')) calcEkspresi += angkaTerakhir === '' ? '0.' : '.';
    } else if (tombol === '000') {
        if (/\d$/.test(calcEkspresi)) calcEkspresi += '000';
    } else {
        calcEkspresi += tombol;
    }
    calcRender();
}

function calcKirim(jenis) {
    if (!(calcNilaiAktif > 0)) {
        showToast('Hitung dulu sebuah nominal yang lebih dari 0', 'error');
        return;
    }
    const isMasuk = jenis === 'Masuk';
    const input = document.getElementById(isMasuk ? 'hargaMasuk' : 'hargaKeluar');
    input.value = calcNilaiAktif;
    hitungSubtotal(jenis);
    fokusForm(jenis);
    showToast(`Nominal ${formatRupiah(calcNilaiAktif)} dikirim ke form Uang ${jenis}`, 'success');
}

function tanganiKeyKalkulator(e) {
    const tab = document.getElementById('tab-tools');
    if (!tab || !tab.classList.contains('active')) return;
    const aktif = document.activeElement;
    if (aktif && ['INPUT', 'TEXTAREA', 'SELECT'].includes(aktif.tagName)) return;
    if (aktif && aktif.tagName === 'BUTTON' && (e.key === 'Enter' || e.key === ' ')) return;

    const k = e.key === ',' ? '.' : e.key;
    if (/^[0-9]$/.test(k) || ['+', '-', '*', '/', '.', '%'].includes(k)) {
        e.preventDefault();
        calcTekan(k);
    } else if (k === 'Enter' || k === '=') {
        e.preventDefault();
        calcTekan('=');
    } else if (k === 'Backspace') {
        e.preventDefault();
        calcTekan('DEL');
    } else if (k.toLowerCase() === 'c') {
        calcTekan('C');
    }
}

/* ---------------------------------------------------------------------
   L. PALET PERINTAH (Ctrl + K)
   --------------------------------------------------------------------- */
function fokusForm(jenis) {
    switchTab('tab-pencatatan');
    const kartu = document.getElementById(jenis === 'Masuk' ? 'cardMasuk' : 'cardKeluar');
    setTimeout(() => {
        kartu.scrollIntoView({ behavior: 'smooth', block: 'center' });
        document.getElementById(jenis === 'Masuk' ? 'ketMasuk' : 'ketKeluar').focus({ preventScroll: true });
    }, 120);
}

function daftarPerintah() {
    const tab = (id, label, kata) => ({ grup: 'Navigasi', label, kata, jalan: () => switchTab(id) });
    return [
        tab('tab-pencatatan', 'Buka Pencatatan', 'transaksi input catat'),
        { grup: 'Navigasi', label: 'Buka Rincian Transaksi', kata: 'tabel daftar transaksi rincian', jalan: () => bukaRincian() },
        { grup: 'Navigasi', label: 'Tampilkan transaksi tanpa nota', kata: 'nota filter kurang lampiran bukti', jalan: () => tampilkanTanpaNota() },
        { grup: 'Tools', label: 'Buat kuitansi', kata: 'kwitansi cetak bukti kas', jalan: () => { switchTab('tab-tools'); setTimeout(() => { const e = document.getElementById('kwNama'); if (e) e.focus({ preventScroll: true }); }, 200); } },
        tab('tab-analisis', 'Buka Analisis & Grafik', 'chart statistik laporan kesiapan'),
        tab('tab-tools', 'Buka Tools', 'kalkulator pagu budget terbilang backup'),
        tab('tab-gdrive', 'Buka Google Drive', 'upload unggah drive'),
        { grup: 'Draf', label: 'Buat draf baru', kata: 'new tambah', jalan: () => buatDraftBaru() },
        { grup: 'Draf', label: 'Duplikat draf aktif', kata: 'salin copy', jalan: () => duplikatDraf() },
        { grup: 'Draf', label: 'Ubah nama draf', kata: 'rename', jalan: () => ubahNamaDraft() },
        { grup: 'Input', label: 'Fokus ke form Uang Masuk', kata: 'pemasukan tambah', jalan: () => fokusForm('Masuk') },
        { grup: 'Input', label: 'Fokus ke form Uang Keluar', kata: 'pengeluaran tambah', jalan: () => fokusForm('Keluar') },
        { grup: 'Riwayat', label: 'Undo (urungkan)', kata: 'batal', jalan: () => undo() },
        { grup: 'Riwayat', label: 'Redo (ulangi)', kata: 'ulang', jalan: () => redo() },
        { grup: 'Ekspor', label: 'Download Word (.docx)', kata: 'export unduh', jalan: () => exportToWord() },
        { grup: 'Ekspor', label: 'Download PDF (.pdf)', kata: 'export unduh', jalan: () => exportToPDF() },
        { grup: 'Ekspor', label: 'Download Excel (.xlsx)', kata: 'export unduh', jalan: () => exportToExcel() },
        { grup: 'Ekspor', label: 'Ekspor CSV draf aktif', kata: 'export unduh', jalan: () => exportCSV() },
        { grup: 'Tools', label: 'Salin ringkasan WhatsApp', kata: 'wa share bagikan', jalan: () => salinRingkasan() },
        { grup: 'Tools', label: 'Unduh cadangan semua draf', kata: 'backup json', jalan: () => backupSemuaDraf() },
        { grup: 'Tampilan', label: 'Ganti tema terang / gelap', kata: 'dark light mode', jalan: () => toggleTheme() },
        { grup: 'Tampilan', label: 'Kembali ke atas', kata: 'scroll top', jalan: () => window.scrollTo({ top: 0, behavior: 'smooth' }) }
    ];
}

function paletTerbuka() {
    return document.getElementById('cmdPalette').classList.contains('open');
}

function renderPalette(query) {
    const kata = (query || '').trim().toLowerCase();
    paletHasil = daftarPerintah().filter(p => !kata || `${p.label} ${p.kata || ''} ${p.grup}`.toLowerCase().includes(kata));
    paletIndex = 0;
    const daftar = document.getElementById('cmdList');
    daftar.innerHTML = paletHasil.length
        ? paletHasil.map((p, i) => `<li class="cmd-item${i === 0 ? ' active' : ''}" data-i="${i}"><span class="cmd-label">${esc(p.label)}</span><span class="cmd-tag">${esc(p.grup)}</span></li>`).join('')
        : '<li class="cmd-empty">Tidak ada perintah yang cocok.</li>';
}

function geserPalette(delta) {
    if (!paletHasil.length) return;
    paletIndex = (paletIndex + delta + paletHasil.length) % paletHasil.length;
    document.querySelectorAll('#cmdList .cmd-item').forEach((el, i) => {
        el.classList.toggle('active', i === paletIndex);
        if (i === paletIndex) el.scrollIntoView({ block: 'nearest' });
    });
}

function openPalette() {
    const overlay = document.getElementById('cmdPalette');
    const input = document.getElementById('cmdInput');
    overlay.classList.add('open');
    input.value = '';
    renderPalette('');
    setTimeout(() => input.focus(), 30);
}

function closePalette() {
    document.getElementById('cmdPalette').classList.remove('open');
}

function jalankanPalette(i) {
    const perintah = paletHasil[i];
    if (!perintah) return;
    closePalette();
    setTimeout(perintah.jalan, 40);
}

/* ---------------------------------------------------------------------
   M. EFEK VISUAL: KONFETI
   --------------------------------------------------------------------- */
function luncurkanKonfeti(jumlah = 140, asal = { x: 0.5, y: 0.3 }) {
    if (REDUCED_MOTION) return;
    const canvas = document.getElementById('confettiCanvas');
    if (!canvas) return;
    const dpr = window.devicePixelRatio || 1;
    canvas.width = window.innerWidth * dpr;
    canvas.height = window.innerHeight * dpr;
    const ctx = canvas.getContext('2d');
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

    const warna = ['#5b6cff', '#9b5cf6', '#ec5fa8', '#ffa23e', '#34d399', '#fbbf24'];
    for (let i = 0; i < jumlah; i++) {
        const sudut = Math.random() * Math.PI * 2;
        const laju = 4 + Math.random() * 9;
        partikelKonfeti.push({
            x: asal.x * window.innerWidth, y: asal.y * window.innerHeight,
            vx: Math.cos(sudut) * laju, vy: Math.sin(sudut) * laju - 6,
            w: 6 + Math.random() * 6, h: 4 + Math.random() * 5,
            rot: Math.random() * 6.28, vr: (Math.random() - 0.5) * 0.35,
            warna: warna[i % warna.length], hidup: 0, umur: 90 + Math.random() * 60
        });
    }
    if (konfetiRaf) return;

    const putar = () => {
        ctx.clearRect(0, 0, window.innerWidth, window.innerHeight);
        partikelKonfeti = partikelKonfeti.filter(p => p.hidup < p.umur && p.y < window.innerHeight + 40);
        partikelKonfeti.forEach(p => {
            p.vy += 0.32; p.vx *= 0.99;
            p.x += p.vx; p.y += p.vy; p.rot += p.vr; p.hidup++;
            ctx.save();
            ctx.globalAlpha = Math.min(1, (p.umur - p.hidup) / 30);
            ctx.translate(p.x, p.y);
            ctx.rotate(p.rot);
            ctx.fillStyle = p.warna;
            ctx.fillRect(-p.w / 2, -p.h / 2, p.w, p.h);
            ctx.restore();
        });
        if (partikelKonfeti.length) {
            konfetiRaf = requestAnimationFrame(putar);
        } else {
            konfetiRaf = null;
            ctx.clearRect(0, 0, window.innerWidth, window.innerHeight);
        }
    };
    konfetiRaf = requestAnimationFrame(putar);
}

/* ---------------------------------------------------------------------
   N. DRAF, PENYIMPANAN & RIWAYAT (versi diperkuat)
   --------------------------------------------------------------------- */
function simpanListDraft() {
    try {
        localStorage.setItem('gcc_list_draft', JSON.stringify(listDraft));
        updateStorageMeter();
        return true;
    } catch (e) {
        console.error(e);
        showToast('Penyimpanan browser penuh! Unduh cadangan atau hapus beberapa lampiran besar.', 'error');
        return false;
    }
}

function simpanStateKeHistory() {
    undoStack.push({ transaksi: klon(transaksi), jumlahPeserta: document.getElementById('jumlahPeserta').value || 0 });
    if (undoStack.length > 30) undoStack.shift();
    redoStack = [];
    updateTombolHistory();
}

function ubahPeserta() {
    const el = document.getElementById('jumlahPeserta');
    const sebelumnya = el.dataset.prev !== undefined ? el.dataset.prev : '0';
    if (String(sebelumnya) !== String(el.value)) {
        undoStack.push({ transaksi: klon(transaksi), jumlahPeserta: sebelumnya });
        if (undoStack.length > 30) undoStack.shift();
        redoStack = [];
        updateTombolHistory();
    }
    el.dataset.prev = el.value;
    simpanKeStorage(false);
    updateTampilan();
}

function muatSemuaDraft() {
    try {
        const saved = localStorage.getItem('gcc_list_draft');
        if (saved) listDraft = JSON.parse(saved);
    } catch (e) {
        console.error(e);
        listDraft = [];
        showToast('Data tersimpan tidak terbaca, memulai draf baru.', 'error');
    }
    if (!Array.isArray(listDraft)) listDraft = [];

    if (listDraft.length === 0) {
        buatDraftBaru(false);
    } else {
        currentDraftId = listDraft[0].id;
        renderDropdownDraft();
        muatDraftAktif();
    }
}

function muatDraftAktif() {
    const draft = getDraftAktif();
    if (draft) {
        const inputPeserta = document.getElementById('jumlahPeserta');
        inputPeserta.value = draft.jumlahPeserta || 0;
        inputPeserta.dataset.prev = inputPeserta.value;
        transaksi = draft.transaksi || [];
        draft.budgets = draft.budgets || {};
        undoStack = [];
        redoStack = [];
        lastHealthScore = null;
        resetModeEdit();
        updateTombolHistory();
        updateTampilan();
    }
}

function duplikatDraf() {
    const draft = getDraftAktif();
    if (!draft) return;
    simpanKeStorage(false);
    const salinan = klon(draft);
    salinan.id = 'draft_' + Date.now();
    salinan.nama = `${draft.nama} (Salinan)`;
    listDraft.unshift(salinan);
    currentDraftId = salinan.id;
    simpanListDraft();
    renderDropdownDraft();
    muatDraftAktif();
    showToast('Draf berhasil diduplikasi!', 'success');
}

/* ---------------------------------------------------------------------
   O. TRANSAKSI (tambah, ubah, salin, hapus) — versi diperkuat
   --------------------------------------------------------------------- */
function resetModeEdit() {
    editIndex = -1;
    const btnMasuk = document.getElementById('btnSubmitMasuk');
    const btnKeluar = document.getElementById('btnSubmitKeluar');
    if (btnMasuk) btnMasuk.innerText = 'Tambah Uang Masuk';
    if (btnKeluar) btnKeluar.innerText = 'Tambah Uang Keluar';
}

async function prosesGantiNota(event) {
    const file = event.target.files[0];
    if (!file || targetIndexNotaEdit < 0) return;

    let notaData;
    try {
        notaData = await kompresGambar(file);
    } catch (e) {
        showToast('Gagal membaca file gambar', 'error');
        return;
    }

    simpanStateKeHistory();
    if (transaksi[targetIndexNotaEdit]) transaksi[targetIndexNotaEdit].notaData = notaData;
    targetIndexNotaEdit = -1;
    event.target.value = '';

    simpanKeStorage(false);
    updateTampilan();
    showToast("Lampiran bukti/nota diperbarui!", "success");
}

async function tambahTransaksi(jenis) {
    const isMasuk = jenis === 'Masuk';
    const divEl = document.getElementById(isMasuk ? 'divisiMasuk' : 'divisiKeluar');
    const ketEl = document.getElementById(isMasuk ? 'ketMasuk' : 'ketKeluar');
    const hargaEl = document.getElementById(isMasuk ? 'hargaMasuk' : 'hargaKeluar');
    const qtyEl = document.getElementById(isMasuk ? 'jumlahMasuk' : 'jumlahKeluar');
    const catEl = document.getElementById(isMasuk ? 'catMasuk' : 'catKeluar');
    const fileEl = document.getElementById(isMasuk ? 'notaMasuk' : 'notaKeluar');

    const divisi = divEl.value.trim() || 'CONTER';
    const keterangan = ketEl.value.trim();
    const harga = parseFloat(hargaEl.value);
    const qty = parseFloat(qtyEl.value) || 1;
    const catatan = catEl.value.trim() || '-';

    if (!keterangan || isNaN(harga) || harga <= 0) {
        showToast("Isi Keterangan dan Harga Nominal yang valid!", "error");
        return;
    }

    let notaData = null;
    if (fileEl && fileEl.files[0]) {
        try {
            notaData = await kompresGambar(fileEl.files[0]);
        } catch (e) {
            showToast('Gagal membaca file nota', 'error');
            return;
        }
    }

    simpanStateKeHistory();
    const sekarang = Date.now();

    if (editIndex >= 0 && transaksi[editIndex]) {
        const lama = transaksi[editIndex];
        transaksi[editIndex] = { divisi, jenis, keterangan, harga, qty, totalNominal: harga * qty, notaData: notaData || lama.notaData || null, catatan, waktu: lama.waktu || sekarang };
        resetModeEdit();
        showToast("Perubahan transaksi disimpan!", "success");
    } else {
        transaksi.push({ divisi, jenis, keterangan, harga, qty, totalNominal: harga * qty, notaData, catatan, waktu: sekarang });
        showToast("Transaksi berhasil ditambahkan!", "success");
    }

    ketEl.value = ''; hargaEl.value = ''; qtyEl.value = '1'; catEl.value = ''; if (fileEl) fileEl.value = '';
    hitungSubtotal(jenis);
    simpanKeStorage(false);
    updateTampilan();
    cekPagu(jenis, divisi);
    if (isMasuk) luncurkanKonfeti(45, { x: 0.28, y: 0.55 });
}

function editTransaksi(index) {
    const item = transaksi[index];
    if (!item) return;
    editIndex = index;
    const isMasuk = item.jenis === 'Masuk';
    document.getElementById(isMasuk ? 'divisiMasuk' : 'divisiKeluar').value = item.divisi;
    document.getElementById(isMasuk ? 'ketMasuk' : 'ketKeluar').value = item.keterangan;
    document.getElementById(isMasuk ? 'hargaMasuk' : 'hargaKeluar').value = item.harga;
    document.getElementById(isMasuk ? 'jumlahMasuk' : 'jumlahKeluar').value = item.qty;
    document.getElementById(isMasuk ? 'catMasuk' : 'catKeluar').value = item.catatan === '-' ? '' : item.catatan;
    document.getElementById(isMasuk ? 'btnSubmitMasuk' : 'btnSubmitKeluar').innerText = 'Simpan Perubahan';
    hitungSubtotal(item.jenis);
    const kartuForm = document.getElementById(isMasuk ? 'cardMasuk' : 'cardKeluar');
    if (MQ_HP.matches) {
        // Di HP, form ada di halaman Pencatatan: pindah dulu, lalu gulir ke form
        switchTab('tab-pencatatan');
        setTimeout(() => kartuForm.scrollIntoView({ behavior: 'smooth', block: 'center' }), 120);
    } else {
        kartuForm.scrollIntoView({ behavior: 'smooth' });
    }
    document.getElementById(isMasuk ? 'ketMasuk' : 'ketKeluar').focus({ preventScroll: true });
}

function duplikatTransaksi(index) {
    const asli = transaksi[index];
    if (!asli) return;
    simpanStateKeHistory();
    const salinan = klon(asli);
    salinan.notaData = null;
    salinan.waktu = Date.now();
    transaksi.splice(index + 1, 0, salinan);
    if (editIndex > index) editIndex++;
    simpanKeStorage(false);
    updateTampilan();
    showToast('Transaksi diduplikat (tanpa lampiran nota)', 'success');
}

function hapusTransaksi(index) {
    simpanStateKeHistory();
    transaksi.splice(index, 1);
    if (editIndex === index) resetModeEdit();
    else if (editIndex > index) editIndex--;
    simpanKeStorage(false);
    updateTampilan();
    showToast('Transaksi dihapus', 'info', { label: 'Urungkan', fn: () => undo() });
}

/* ---------------------------------------------------------------------
   P. TABEL TRANSAKSI (cari, filter, urutkan) — bug indeks sudah diperbaiki
   --------------------------------------------------------------------- */
function barisTransaksi(t, idx, i) {
    let notaHtml;
    if (t.notaData && /^data:image\//i.test(t.notaData.dataUrl || '')) {
        notaHtml = `
            <img src="${esc(t.notaData.dataUrl)}" class="img-preview" alt="Bukti" data-open-nota="${idx}">
            <div class="nota-actions">
                <span class="btn-nota-action" onclick="triggerGantiNota(${idx})">Ganti</span>
                <span class="btn-nota-delete" onclick="hapusNota(${idx})">Hapus</span>
            </div>`;
    } else {
        notaHtml = `<span class="btn-nota-action" onclick="triggerGantiNota(${idx})">Upload</span>`;
    }

    const waktuHtml = t.waktu ? `<small class="row-time">${esc(formatWaktu(t.waktu))}</small>` : '';
    const jenis = t.jenis === 'Masuk' ? 'Masuk' : 'Keluar';

    return `<tr>
        <td data-label="No" class="cell-center">${i + 1}</td>
        <td data-label="Jenis"><span class="${jenis === 'Masuk' ? 'badge-masuk' : 'badge-keluar'}">${jenis}</span></td>
        <td data-label="Keterangan"><div class="ket-wrap"><span>${esc(t.keterangan)}</span>${waktuHtml}</div></td>
        <td data-label="Harga Satuan">${formatRupiah(t.harga)}</td>
        <td data-label="Jumlah" class="cell-center">${esc(t.qty)}</td>
        <td data-label="Total"><b>${formatRupiah(t.totalNominal)}</b></td>
        <td data-label="Nota / Bukti" class="cell-center">${notaHtml}</td>
        <td data-label="Catatan">${esc(t.catatan)}</td>
        <td class="col-action" data-label="Aksi">
            <button class="btn-edit" onclick="editTransaksi(${idx})">Edit</button>
            <button class="btn-edit" onclick="duplikatTransaksi(${idx})">Duplikat</button>
            <button class="btn-delete" onclick="hapusTransaksi(${idx})">Hapus</button>
        </td>
    </tr>`;
}

function updateTampilan() {
    const tbody = document.getElementById('tabelBody');
    const tfoot = document.getElementById('tabelFoot');
    const keyword = (document.getElementById('searchKeyword').value || '').trim().toLowerCase();
    const filterJen = document.getElementById('filterJenis').value;
    const urut = document.getElementById('sortTransaksi').value;

    // Segarkan pilihan divisi, lalu baca nilai yang berlaku
    const selectDiv = document.getElementById('filterDivisi');
    const divSebelumnya = selectDiv.value;
    const divisiUnik = [...new Set(transaksi.map(t => t.divisi || 'CONTER'))];
    selectDiv.innerHTML = '<option value="ALL">Semua Divisi</option>' + divisiUnik.map(d => `<option value="${esc(d)}">${esc(d)}</option>`).join('');
    selectDiv.value = divisiUnik.includes(divSebelumnya) ? divSebelumnya : 'ALL';
    const filterDiv = selectDiv.value;

    // Simpan indeks asli supaya Edit/Hapus selalu mengenai transaksi yang benar
    const cocok = [];
    transaksi.forEach((t, idx) => {
        const teksCari = `${t.keterangan || ''} ${t.divisi || ''} ${t.catatan || ''}`.toLowerCase();
        if (keyword && !teksCari.includes(keyword)) return;
        if (filterDiv !== 'ALL' && (t.divisi || 'CONTER') !== filterDiv) return;
        if (filterJen === 'NONOTA') { if (t.jenis !== 'Keluar' || t.notaData) return; }
        else if (filterJen !== 'ALL' && t.jenis !== filterJen) return;
        cocok.push({ t, idx });
    });

    const kelompok = {};
    cocok.forEach(item => {
        const nama = item.t.divisi || 'CONTER';
        (kelompok[nama] = kelompok[nama] || []).push(item);
    });

    setTeks('resultCount', transaksi.length ? `Menampilkan ${cocok.length} dari ${transaksi.length} transaksi` : '');

    if (cocok.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" class="empty-row">${transaksi.length ? 'Belum ada transaksi yang cocok dengan filter.' : 'Belum ada transaksi di draf ini.'}</td></tr>`;
    } else {
        let html = '';
        for (const [namaDivisi, daftar] of Object.entries(kelompok)) {
            if (urut === 'nominal-desc') daftar.sort((a, b) => b.t.totalNominal - a.t.totalNominal);
            else if (urut === 'nominal-asc') daftar.sort((a, b) => a.t.totalNominal - b.t.totalNominal);
            else if (urut === 'nama-asc') daftar.sort((a, b) => String(a.t.keterangan).localeCompare(String(b.t.keterangan), 'id'));

            html += `<tr class="row-divisi-header"><td colspan="9">${esc(namaDivisi)}</td></tr>`;
            daftar.forEach(({ t, idx }, i) => { html += barisTransaksi(t, idx, i); });
        }
        tbody.innerHTML = html;
    }

    const summary = updateSummary();
    tfoot.innerHTML = cocok.length === 0 ? '' : `
        <tr class="tfoot-summary"><td colspan="5">Total Uang Masuk:</td><td colspan="4" class="text-masuk">${formatRupiah(summary.totalMasuk)}</td></tr>
        <tr class="tfoot-summary"><td colspan="5">Total Uang Keluar:</td><td colspan="4" class="text-keluar">${formatRupiah(summary.totalKeluar)}</td></tr>
        <tr class="tfoot-summary tfoot-saldo"><td colspan="5">Total Saldo Akhir:</td><td colspan="4">${formatRupiah(summary.saldoAkhir)}</td></tr>
    `;

    updateCharts();
    updateInsights(summary);
    renderBudget();
    updateHealth(summary);
    updateHero(summary);
    updateWAPreview(summary);
    updateRincianLink();
    renderRekapDivisi();
    renderAktivitas();
    kuitansiPerbaruiSumber();
}

/* ---------------------------------------------------------------------
   Q. INISIALISASI SEMUA FITUR
   --------------------------------------------------------------------- */
function initFitur() {
    // Tempatkan tabel sesuai ukuran layar & pantau perubahan ukuran/rotasi
    posisikanRincian();
    if (MQ_HP.addEventListener) MQ_HP.addEventListener('change', posisikanRincian);
    else if (MQ_HP.addListener) MQ_HP.addListener(posisikanRincian);
    kuitansiInit();

    // Klik gambar nota di tabel -> buka lightbox (tanpa menaruh data gambar di atribut onclick)
    document.getElementById('tabelBody').addEventListener('click', e => {
        const gambar = e.target.closest('[data-open-nota]');
        if (!gambar) return;
        const t = transaksi[Number(gambar.dataset.openNota)];
        if (t && t.notaData) openModal(t.notaData.dataUrl, t.keterangan);
    });

    // Chip nominal cepat
    document.querySelectorAll('.chip-row').forEach(baris => {
        baris.addEventListener('click', e => {
            const chip = e.target.closest('.chip');
            if (!chip) return;
            const jenis = baris.dataset.jenis;
            const input = document.getElementById(jenis === 'Masuk' ? 'hargaMasuk' : 'hargaKeluar');
            input.value = chip.dataset.add === 'clear' ? '' : (parseFloat(input.value) || 0) + Number(chip.dataset.add);
            hitungSubtotal(jenis);
        });
    });

    // Papan tombol kalkulator
    document.getElementById('calcKeys').addEventListener('click', e => {
        const tombol = e.target.closest('[data-key]');
        if (tombol) calcTekan(tombol.dataset.key);
    });
    calcRender();

    // Pagu: hapus & saran divisi
    document.getElementById('budgetList').addEventListener('click', e => {
        const tombol = e.target.closest('[data-hapus-budget]');
        if (tombol) hapusBudget(tombol.dataset.hapusBudget);
    });
    document.getElementById('budgetSuggest').addEventListener('click', e => {
        const chip = e.target.closest('[data-divisi]');
        if (!chip) return;
        document.getElementById('budgetDivisi').value = chip.dataset.divisi;
        document.getElementById('budgetNominal').focus();
    });

    // Jumlah peserta: simpan nilai sebelumnya agar Undo akurat
    document.getElementById('jumlahPeserta').addEventListener('focus', e => { e.target.dataset.prev = e.target.value; });

    // Palet perintah
    const overlay = document.getElementById('cmdPalette');
    const inputPalet = document.getElementById('cmdInput');
    overlay.addEventListener('mousedown', e => { if (e.target === overlay) closePalette(); });
    inputPalet.addEventListener('input', () => renderPalette(inputPalet.value));
    inputPalet.addEventListener('keydown', e => {
        if (e.key === 'ArrowDown') { e.preventDefault(); geserPalette(1); }
        else if (e.key === 'ArrowUp') { e.preventDefault(); geserPalette(-1); }
        else if (e.key === 'Enter') { e.preventDefault(); jalankanPalette(paletIndex); }
    });
    document.getElementById('cmdList').addEventListener('click', e => {
        const item = e.target.closest('.cmd-item');
        if (item) jalankanPalette(Number(item.dataset.i));
    });

    // Pintasan keyboard global
    document.addEventListener('keydown', e => {
        const aktif = document.activeElement;
        const mengetik = !!aktif && (aktif.tagName === 'TEXTAREA' || (aktif.tagName === 'INPUT' && !['checkbox', 'radio', 'file', 'button'].includes(aktif.type)));
        const mod = e.ctrlKey || e.metaKey;
        const tombol = (e.key || '').toLowerCase();

        if (mod && tombol === 'k') {
            e.preventDefault();
            paletTerbuka() ? closePalette() : openPalette();
            return;
        }
        if (e.key === 'Escape') {
            closePalette();
            closeModal();
            document.getElementById('mobileMenuDropdown').style.display = 'none';
            return;
        }
        if (paletTerbuka()) return;

        if (mod && tombol === 'z') {
            if (mengetik) return; // biarkan undo bawaan browser saat mengetik
            e.preventDefault();
            e.shiftKey ? redo() : undo();
            return;
        }
        if (mod && tombol === 'y') {
            if (mengetik) return;
            e.preventDefault();
            redo();
            return;
        }
        if (!mod && !e.altKey) tanganiKeyKalkulator(e);
    });

    // Menu HP: tutup bila klik di luar
    window.addEventListener('click', e => {
        const dropdown = document.getElementById('mobileMenuDropdown');
        const menuBtn = document.querySelector('.mobile-menu-btn');
        if (dropdown && menuBtn && !dropdown.contains(e.target) && !menuBtn.contains(e.target)) {
            dropdown.style.display = 'none';
        }
    });

    // Garis kemajuan scroll + tombol ke atas
    const garis = document.getElementById('scrollProgress');
    const tombolAtas = document.getElementById('btnKeAtas');
    const saatScroll = () => {
        const tinggi = document.documentElement.scrollHeight - window.innerHeight;
        garis.style.transform = `scaleX(${tinggi > 0 ? window.scrollY / tinggi : 0})`;
        tombolAtas.classList.toggle('show', window.scrollY > 500);
    };
    window.addEventListener('scroll', saatScroll, { passive: true });
    saatScroll();

    // Efek sorot cahaya mengikuti kursor pada kartu
    document.addEventListener('pointermove', e => {
        const kartu = e.target.closest && e.target.closest('.card, .chart-card, .stat-tile, .tool-card, .health-card');
        if (!kartu) return;
        const r = kartu.getBoundingClientRect();
        kartu.style.setProperty('--mx', (e.clientX - r.left) + 'px');
        kartu.style.setProperty('--my', (e.clientY - r.top) + 'px');
    }, { passive: true });

    updateStorageMeter();
}

window.onload = function () {
    initTheme();
    initCharts();
    muatSemuaDraft();
    initFitur();
};


/* =====================================================================
   FITUR BARU: Rekap per Divisi, Aktivitas Terakhir, Filter Tanpa Nota,
   dan Pembuat Kuitansi
   ===================================================================== */

/* ---------- Rekap per Divisi (halaman Analisis) ---------- */
function renderRekapDivisi() {
    const wadah = document.getElementById('rekapDivisiList');
    if (!wadah) return;

    const peta = {};
    transaksi.forEach(t => {
        const k = String(t.divisi || 'CONTER').trim().toUpperCase();
        const o = peta[k] = peta[k] || { masuk: 0, keluar: 0 };
        const n = Number(t.totalNominal) || 0;
        if (t.jenis === 'Masuk') o.masuk += n;
        else if (t.jenis === 'Keluar') o.keluar += n;
    });

    const nama = Object.keys(peta).sort((a, b) => peta[b].keluar - peta[a].keluar);
    if (!nama.length) {
        wadah.innerHTML = '<p class="empty-note">Belum ada transaksi di draf ini.</p>';
        return;
    }

    const draft = getDraftAktif();
    const budgets = (draft && draft.budgets) || {};

    wadah.innerHTML = nama.map(k => {
        const { masuk, keluar } = peta[k];
        const selisih = masuk - keluar;
        const pagu = Number(budgets[k]) || 0;
        let bar = '';
        if (pagu > 0) {
            const p = (keluar / pagu) * 100;
            const status = p > 100 ? 'over' : p >= 80 ? 'warn' : 'ok';
            bar = `<div class="rekap-pagu ${status}">
                <div class="budget-bar"><span data-w="${Math.min(p, 100).toFixed(1)}"></span></div>
                <div class="budget-note">${Math.round(p)}% dari pagu ${formatRupiah(pagu)}</div>
            </div>`;
        }
        return `<div class="rekap-row">
            <div class="rekap-top"><strong>${esc(k)}</strong><span class="${selisih < 0 ? 'text-keluar' : 'text-masuk'}">${formatRupiah(selisih)}</span></div>
            <div class="rekap-sub"><span>Masuk ${formatRupiah(masuk)}</span><span>Keluar ${formatRupiah(keluar)}</span></div>
            ${bar}
        </div>`;
    }).join('');

    setTimeout(() => {
        wadah.querySelectorAll('.budget-bar span[data-w]').forEach(el => { el.style.width = el.dataset.w + '%'; });
    }, 30);
}

/* ---------- Aktivitas Terakhir (halaman Analisis) ---------- */
function waktuRelatif(ms) {
    const menit = Math.floor((Date.now() - ms) / 60000);
    if (menit < 1) return 'baru saja';
    if (menit < 60) return `${menit} menit lalu`;
    const jam = Math.floor(menit / 60);
    if (jam < 24) return `${jam} jam lalu`;
    const hari = Math.floor(jam / 24);
    if (hari < 7) return `${hari} hari lalu`;
    return formatWaktu(ms);
}

function renderAktivitas() {
    const wadah = document.getElementById('aktivitasList');
    if (!wadah) return;
    if (!transaksi.length) {
        wadah.innerHTML = '<li class="empty-note">Belum ada aktivitas.</li>';
        return;
    }
    const terbaru = transaksi
        .map((t, i) => ({ t, i }))
        .sort((a, b) => ((b.t.waktu || 0) - (a.t.waktu || 0)) || (b.i - a.i))
        .slice(0, 6);

    wadah.innerHTML = terbaru.map(({ t }) => {
        const masuk = t.jenis === 'Masuk';
        const waktu = t.waktu ? ' · ' + waktuRelatif(t.waktu) : '';
        return `<li class="aktv ${masuk ? 'in' : 'out'}">
            <span class="aktv-dot"></span>
            <div class="aktv-body"><b>${esc(t.keterangan)}</b><small>${esc(String(t.divisi || 'CONTER'))}${esc(waktu)}</small></div>
            <span class="aktv-amt">${masuk ? '+' : '−'}${formatRupiah(Number(t.totalNominal) || 0)}</span>
        </li>`;
    }).join('');
}

/* ---------- Filter: transaksi keluar tanpa nota ---------- */
function tampilkanTanpaNota() {
    const cari = document.getElementById('searchKeyword');
    const jenis = document.getElementById('filterJenis');
    const divisi = document.getElementById('filterDivisi');
    if (cari) cari.value = '';
    if (divisi) divisi.value = 'ALL';
    if (jenis) jenis.value = 'NONOTA';
    updateTampilan();
    bukaRincian();
    const jumlah = transaksi.filter(t => t.jenis === 'Keluar' && !t.notaData).length;
    showToast(jumlah ? `${jumlah} pengeluaran belum ada nota` : 'Semua pengeluaran sudah punya nota', jumlah ? 'info' : 'success');
}

/* ---------- Pembuat Kuitansi ---------- */
let kwNomorSementara = '';

function hariIniISO() {
    const d = new Date();
    d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
    return d.toISOString().slice(0, 10);
}

function kuitansiInit() {
    const tgl = document.getElementById('kwTanggal');
    if (tgl && !tgl.value) tgl.value = hariIniISO();
    try {
        const b = localStorage.getItem('gcc_kw_bendahara');
        const kolom = document.getElementById('kwBendahara');
        if (b && kolom && !kolom.value) kolom.value = b;
    } catch (e) { /* penyimpanan tidak tersedia */ }
    document.addEventListener('keydown', e => { if (e.key === 'Escape') tutupKuitansi(); });
}

function kuitansiPerbaruiSumber() {
    const sel = document.getElementById('kwSumber');
    if (!sel) return;
    const sebelum = sel.value;
    sel.innerHTML = '<option value="">— Isi manual —</option>' + transaksi.map((t, i) =>
        `<option value="${i}">${esc(t.jenis)} · ${esc(t.keterangan)} · ${esc(formatRupiah(Number(t.totalNominal) || 0))}</option>`
    ).join('');
    sel.value = (sebelum !== '' && transaksi[Number(sebelum)]) ? sebelum : '';
}

function kuitansiIsiDariTransaksi() {
    const t = transaksi[Number(document.getElementById('kwSumber').value)];
    if (!t) return;
    document.getElementById('kwUntuk').value = t.keterangan || '';
    document.getElementById('kwNominal').value = Number(t.totalNominal) || '';
    document.getElementById('kwJenis').value = t.jenis === 'Masuk' ? 'terima' : 'bayar';
    kuitansiTerbilangLive();
}

function kuitansiTerbilangLive() {
    const n = parseFloat(document.getElementById('kwNominal').value);
    setTeks('kwTerbilang', n > 0 ? formatTerbilang(n) : '');
}

function nomorKuitansiBerikut() {
    let n = 0;
    try { n = Number(localStorage.getItem('gcc_kw_no')) || 0; } catch (e) { /* abaikan */ }
    return n + 1;
}

function bukaKuitansi() {
    const ambil = id => (document.getElementById(id).value || '').trim();
    const nominal = parseFloat(document.getElementById('kwNominal').value);
    if (!(nominal > 0)) { showToast('Isi nominal kuitansi terlebih dahulu', 'error'); return; }

    const jenis = ambil('kwJenis');
    const nama = ambil('kwNama');
    const untuk = ambil('kwUntuk');
    const bendahara = ambil('kwBendahara');
    try { localStorage.setItem('gcc_kw_bendahara', bendahara); } catch (e) { /* abaikan */ }

    const tglIso = ambil('kwTanggal') || hariIniISO();
    const tanggal = new Date(tglIso + 'T00:00:00').toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
    const nomor = 'KW-' + String(nomorKuitansiBerikut()).padStart(4, '0');
    kwNomorSementara = nomor;

    const label = jenis === 'terima' ? 'Telah terima dari' : 'Telah dibayarkan kepada';
    const penyerah = jenis === 'terima' ? nama : bendahara;
    const penerima = jenis === 'terima' ? bendahara : nama;
    const titik = '...............................';

    document.getElementById('kwSheet').innerHTML = `
        <div class="kw-paper">
            <div class="kw-head">
                <div><strong>GAMATIKA Coding Club</strong><span class="kw-sub">Bukti Kas</span></div>
                <div class="kw-no">No. ${esc(nomor)}</div>
            </div>
            <h2>KUITANSI</h2>
            <dl>
                <div class="kw-row"><dt>${label}</dt><dd>${esc(nama || titik)}</dd></div>
                <div class="kw-row"><dt>Uang sejumlah</dt><dd><span class="kw-terbilang">${esc(formatTerbilang(nominal))}</span></dd></div>
                <div class="kw-row"><dt>Untuk pembayaran</dt><dd>${esc(untuk || titik)}</dd></div>
            </dl>
            <div class="kw-bottom">
                <div class="kw-amount">${esc(formatRupiah(nominal))}</div>
                <div>
                    <div class="kw-tgl">${esc(tanggal)}</div>
                    <div class="kw-sign-wrap">
                        <div class="kw-sign"><span>Yang menyerahkan,</span><div class="kw-line"></div><strong>${esc(penyerah || titik)}</strong></div>
                        <div class="kw-sign"><span>Yang menerima,</span><div class="kw-line"></div><strong>${esc(penerima || titik)}</strong></div>
                    </div>
                </div>
            </div>
        </div>`;

    document.getElementById('kuitansiModal').classList.add('show');
}

function tutupKuitansi() {
    const m = document.getElementById('kuitansiModal');
    if (m) m.classList.remove('show');
}

function cetakKuitansi() {
    // Nomor kuitansi baru dihitung "terpakai" saat benar-benar dicetak
    const angka = parseInt(String(kwNomorSementara).replace(/\D/g, ''), 10);
    if (angka) { try { localStorage.setItem('gcc_kw_no', String(angka)); } catch (e) { /* abaikan */ } }

    document.body.classList.add('cetak-kuitansi');
    const bersih = () => {
        document.body.classList.remove('cetak-kuitansi');
        window.removeEventListener('afterprint', bersih);
    };
    window.addEventListener('afterprint', bersih);
    setTimeout(() => window.print(), 60);
}
