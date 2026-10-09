let listDraft = [];
let currentDraftId = null;
let transaksi = [];
let editIndex = -1;
let targetIndexNotaEdit = -1;

let undoStack = [];
let redoStack = [];

const GCC_LOGO_URL = "https://lh3.googleusercontent.com/d/1QOjchWwyuyYMdB00zTWBUyS-h7LYGvWl";

window.onload = function() {
    muatSemuaDraft();
    
    document.addEventListener('keydown', function(e) {
        const activeElement = document.activeElement;
        const isInputFocused = activeElement && (activeElement.tagName === 'INPUT' || activeElement.tagName === 'TEXTAREA');

        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
            if (e.shiftKey) {
                e.preventDefault();
                redo();
            } else if (!isInputFocused || isInputFocused) {
                e.preventDefault();
                undo();
            }
        } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
            e.preventDefault();
            redo();
        }
    });
};

function formatRupiah(angka) {
    return 'Rp ' + Number(angka).toLocaleString('id-ID');
}

function hitungSubtotal(jenis) {
    const isMasuk = jenis === 'Masuk';
    const harga = parseFloat(document.getElementById(isMasuk ? 'hargaMasuk' : 'hargaKeluar').value) || 0;
    const qty = parseFloat(document.getElementById(isMasuk ? 'jumlahMasuk' : 'jumlahKeluar').value) || 0;
    const total = harga * qty;

    document.getElementById(isMasuk ? 'subtotalMasuk' : 'subtotalKeluar').innerText = `Total Item: ${formatRupiah(total)}`;
    return total;
}

/* ------------------- FUNGSI EDIT & HAPUS LAMPIRAN NOTA ------------------- */
function triggerGantiNota(realIndex) {
    targetIndexNotaEdit = realIndex;
    document.getElementById('fileNotaHidden').click();
}

async function prosesGantiNota(event) {
    const file = event.target.files[0];
    if (!file || targetIndexNotaEdit < 0) return;

    const notaData = await new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve({ dataUrl: e.target.result });
        reader.readAsDataURL(file);
    });

    simpanStateKeHistory();
    transaksi[targetIndexNotaEdit].notaData = notaData;
    targetIndexNotaEdit = -1;
    event.target.value = '';

    simpanKeStorage(false);
    updateTampilan();
}

function hapusNota(realIndex) {
    if (confirm("Apakah Anda yakin ingin menghapus lampiran nota transaksi ini?")) {
        simpanStateKeHistory();
        transaksi[realIndex].notaData = null;
        simpanKeStorage(false);
        updateTampilan();
    }
}

/* ------------------- FUNGSI DOWNLOAD TEMPLAT WORD RESMI ------------------- */
async function downloadTemplateWord() {
    const { docx } = window;
    const borderStandard = { style: docx.BorderStyle.SINGLE, size: 1, color: "000000" };
    const borders = {
        top: borderStandard, bottom: borderStandard, left: borderStandard, right: borderStandard,
        insideHorizontal: borderStandard, insideVertical: borderStandard
    };

    const templateRows = [
        new docx.TableRow({
            children: [
                new docx.TableCell({ children: [new docx.Paragraph({ text: "No", bold: true, alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Jenis (Masuk/Keluar)", bold: true, alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Keterangan / Nama Barang", bold: true, alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Harga Satuan (Rp)", bold: true, alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Jumlah (Qty)", bold: true, alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Total Nominal (Rp)", bold: true, alignment: docx.AlignmentType.CENTER })] })
            ]
        }),
        new docx.TableRow({
            children: [
                new docx.TableCell({ columnSpan: 6, children: [new docx.Paragraph({ text: "Acara", bold: true })] })
            ]
        }),
        new docx.TableRow({
            children: [
                new docx.TableCell({ children: [new docx.Paragraph({ text: "1", alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Keluar" })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Honor Pemateri" })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "500000" })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "1", alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "500000" })] })
            ]
        }),
        new docx.TableRow({
            children: [
                new docx.TableCell({ columnSpan: 6, children: [new docx.Paragraph({ text: "Konsumsi", bold: true })] })
            ]
        }),
        new docx.TableRow({
            children: [
                new docx.TableCell({ children: [new docx.Paragraph({ text: "1", alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Keluar" })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Snack Peserta" })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "15000" })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "35", alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "525000" })] })
            ]
        })
    ];

    const doc = new docx.Document({
        sections: [{
            children: [
                new docx.Paragraph({
                    children: [new docx.TextRun({ text: "TEMPLAT RAB & REKAPAN KEUANGAN WORKSHOP GCC", bold: true, size: 24 })],
                    alignment: docx.AlignmentType.CENTER,
                    spacing: { after: 150 }
                }),
                new docx.Paragraph({
                    children: [new docx.TextRun({ text: "Petunjuk: Isi data transaksi di bawah ini. Anda bisa menambah baris baru di bawah tiap divisi sesuai kebutuhan.", italic: true, size: 18 })],
                    spacing: { after: 200 }
                }),
                new docx.Table({
                    borders: borders,
                    rows: templateRows,
                    width: { size: 100, type: docx.WidthType.PERCENTAGE }
                })
            ]
        }]
    });

    docx.Packer.toBlob(doc).then(blob => {
        saveAs(blob, "Templat_RAB_Rekapan_GCC.docx");
    });
}

/* ------------------- FUNGSI IMPOR KHUSUS WORD (.DOCX) ------------------- */
async function imporDokumen(event) {
    const file = event.target.files[0];
    const statusEl = document.getElementById('importStatus');
    if (!file) return;

    if (!file.name.endsWith('.docx')) {
        alert("⚠️ Upload Ditolak!\n\nUntuk mengedit kembali laporan, gunakan file Word (.docx). File PDF hanya untuk keperluan cetak/laporan.");
        event.target.value = '';
        return;
    }

    statusEl.innerText = "⏳ Membaca dokumen Word...";
    statusEl.style.color = "#0056b3";

    try {
        const arrayBuffer = await file.arrayBuffer();
        
        // 1. Cek Metadata JSON Tersembunyi (Akurat 100% jika file Word buatan web kita)
        const rawTextResult = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
        const fullRawText = rawTextResult.value;

        const matchData = fullRawText.match(/GCC_DATA_START:::(.*?):::GCC_DATA_END/s);
        if (matchData && matchData[1]) {
            try {
                const cleanedJsonStr = decodeURIComponent(matchData[1].replace(/\s+/g, ''));
                const decodedJson = JSON.parse(cleanedJsonStr);
                simpanStateKeHistory();

                if (decodedJson.jumlahPeserta !== undefined) {
                    document.getElementById('jumlahPeserta').value = decodedJson.jumlahPeserta;
                }
                if (Array.isArray(decodedJson.transaksi)) {
                    transaksi = decodedJson.transaksi;
                }

                simpanKeStorage(false);
                updateTampilan();

                statusEl.innerText = "✅ Berhasil memuat draf Word buatan web!";
                statusEl.style.color = "#28a745";
                alert("🎉 Draf Berhasil Dipulihkan!\n\nSeluruh data transaksi berhasil dimuat kembali secara utuh.");
                return;
            } catch (e) {
                console.warn("Lanjut membaca sebagai file Word luar...");
            }
        }

        // 2. Parser HTML-Table untuk File Word Templat / Luar
        const htmlResult = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = htmlResult.value;

        simpanStateKeHistory();
        let countImported = 0;
        let currentDivisi = 'Acara';

        const rows = tempDiv.querySelectorAll('tr');
        rows.forEach(row => {
            const cells = Array.from(row.querySelectorAll('td, th')).map(c => (c.innerText || c.textContent || '').trim());
            if (cells.length === 0) return;

            const rowText = cells.join(' ');
            if (rowText.toLowerCase().includes('jenis') || rowText.toLowerCase().includes('keterangan') || rowText.toLowerCase().includes('total')) return;

            // Jika baris berisi judul Divisi (1 sel)
            if (cells.length === 1 || (cells.length > 1 && !cells[1] && !cells[2])) {
                if (cells[0] && !cells[0].match(/\d{3,}/)) {
                    currentDivisi = cells[0];
                    return;
                }
            }

            // Jika baris transaksi (No, Jenis, Keterangan, Harga, Qty, Total)
            if (cells.length >= 4) {
                let jenis = cells[1] && cells[1].toLowerCase().includes('masuk') ? 'Masuk' : 'Keluar';
                let ket = cells[2] || cells[1] || '';
                let harga = parseFloat((cells[3] || '0').replace(/[^\d]/g, '')) || 0;
                let qty = parseFloat((cells[4] || '1').replace(/[^\d]/g, '')) || 1;

                if (ket && harga > 0) {
                    transaksi.push({
                        divisi: currentDivisi,
                        jenis: jenis,
                        keterangan: ket,
                        harga: harga,
                        qty: qty,
                        totalNominal: harga * qty,
                        notaData: null,
                        catatan: 'Diimpor dari Templat Word'
                    });
                    countImported++;
                }
            }
        });

        if (countImported === 0) {
            alert("⚠️ Format tabel tidak terbaca. Silakan gunakan tombol 'Download Templat Word' untuk membuat file RAB yang valid.");
        } else {
            alert(`🎉 Berhasil Mengimpor ${countImported} item transaksi!`);
            simpanKeStorage(false);
            updateTampilan();
        }

        statusEl.innerText = "✅ Impor data selesai!";
        statusEl.style.color = "#28a745";

    } catch (err) {
        statusEl.innerText = `❌ Error: ${err.message}`;
        statusEl.style.color = "#dc3545";
        alert(`Gagal Memproses File Word:\n${err.message}`);
    }
}

/* ------------------- FUNGSI HISTORY & DRAFT ------------------- */
function simpanStateKeHistory() {
    const stateSaatIni = {
        transaksi: JSON.parse(JSON.stringify(transaksi)),
        jumlahPeserta: document.getElementById('jumlahPeserta').value || 0
    };
    undoStack.push(stateSaatIni);
    redoStack = [];
    updateTombolHistory();
}

function undo() {
    if (undoStack.length === 0) return;

    const stateSaatIni = {
        transaksi: JSON.parse(JSON.stringify(transaksi)),
        jumlahPeserta: document.getElementById('jumlahPeserta').value || 0
    };
    redoStack.push(stateSaatIni);

    const previousState = undoStack.pop();
    transaksi = previousState.transaksi;
    document.getElementById('jumlahPeserta').value = previousState.jumlahPeserta;

    simpanKeStorage(false);
    updateTampilan();
    updateTombolHistory();
}

function redo() {
    if (redoStack.length === 0) return;

    const stateSaatIni = {
        transaksi: JSON.parse(JSON.stringify(transaksi)),
        jumlahPeserta: document.getElementById('jumlahPeserta').value || 0
    };
    undoStack.push(stateSaatIni);

    const nextState = redoStack.pop();
    transaksi = nextState.transaksi;
    document.getElementById('jumlahPeserta').value = nextState.jumlahPeserta;

    simpanKeStorage(false);
    updateTampilan();
    updateTombolHistory();
}

function updateTombolHistory() {
    document.getElementById('btnUndo').disabled = undoStack.length === 0;
    document.getElementById('btnRedo').disabled = redoStack.length === 0;
}

function resetHistory() {
    undoStack = [];
    redoStack = [];
    updateTombolHistory();
}

function muatSemuaDraft() {
    const saved = localStorage.getItem('gcc_list_draft');
    if (saved) {
        listDraft = JSON.parse(saved);
    }

    if (listDraft.length === 0) {
        buatDraftBaru(false);
    } else {
        currentDraftId = listDraft[0].id;
        renderDropdownDraft();
        muatDraftAktif();
    }
}

function renderDropdownDraft() {
    const select = document.getElementById('selectDraft');
    select.innerHTML = '';
    listDraft.forEach((d, idx) => {
        const opt = document.createElement('option');
        opt.value = d.id;
        opt.innerText = `Draf #${idx + 1} - ${d.nama} (${d.tanggal})`;
        if (d.id === currentDraftId) opt.selected = true;
        select.appendChild(opt);
    });
}

function buatDraftBaru(userClick = true) {
    if (listDraft.length >= 10 && userClick) {
        alert("Maksimal 10 riwayat draf! Silakan hapus salah satu draf lama.");
        return;
    }

    let namaDraf = `Rekapan ${new Date().toLocaleDateString('id-ID')}`;
    if (userClick) {
        const inputNama = prompt("Masukkan Nama Draf Baru:", namaDraf);
        if (inputNama === null) return; 
        if (inputNama.trim() !== '') namaDraf = inputNama.trim();
    }

    const now = new Date();
    const newDraft = {
        id: 'draft_' + Date.now(),
        nama: namaDraf,
        tanggal: now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }),
        jumlahPeserta: 0,
        transaksi: []
    };

    listDraft.unshift(newDraft);
    currentDraftId = newDraft.id;
    simpanListDraft();
    renderDropdownDraft();
    muatDraftAktif();
}

function ubahNamaDraft() {
    const draft = listDraft.find(d => d.id === currentDraftId);
    if (!draft) return;

    const namaBaru = prompt("Ubah Nama Draf:", draft.nama);
    if (namaBaru && namaBaru.trim() !== '') {
        draft.nama = namaBaru.trim();
        simpanListDraft();
        renderDropdownDraft();
    }
}

function gantiDraft() {
    const select = document.getElementById('selectDraft');
    currentDraftId = select.value;
    muatDraftAktif();
}

function muatDraftAktif() {
    const draft = listDraft.find(d => d.id === currentDraftId);
    if (draft) {
        document.getElementById('jumlahPeserta').value = draft.jumlahPeserta || 0;
        transaksi = draft.transaksi || [];
        resetHistory();
        updateTampilan();
    }
}

function simpanKeStorage(recordHistory = true) {
    if (recordHistory) {
        simpanStateKeHistory();
    }

    const draft = listDraft.find(d => d.id === currentDraftId);
    if (draft) {
        draft.jumlahPeserta = document.getElementById('jumlahPeserta').value || 0;
        draft.transaksi = transaksi;
        const now = new Date();
        draft.tanggal = now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
        simpanListDraft();
    }
}

function simpanListDraft() {
    localStorage.setItem('gcc_list_draft', JSON.stringify(listDraft));
}

function hapusDraftAktif() {
    if (confirm("Apakah Anda yakin ingin menghapus draf rekapan ini?")) {
        listDraft = listDraft.filter(d => d.id !== currentDraftId);
        simpanListDraft();
        muatSemuaDraft();
    }
}

async function tambahTransaksi(jenis) {
    const isMasuk = jenis === 'Masuk';
    const divisiInput = document.getElementById(isMasuk ? 'divisiMasuk' : 'divisiKeluar');
    const ketInput = document.getElementById(isMasuk ? 'ketMasuk' : 'ketKeluar');
    const hargaInput = document.getElementById(isMasuk ? 'hargaMasuk' : 'hargaKeluar');
    const qtyInput = document.getElementById(isMasuk ? 'jumlahMasuk' : 'jumlahKeluar');
    const catInput = document.getElementById(isMasuk ? 'catMasuk' : 'catKeluar');
    const fileInput = isMasuk ? null : document.getElementById('notaKeluar');

    const divisi = divisiInput.value.trim() || 'Umum';
    const keterangan = ketInput.value.trim();
    const harga = parseFloat(hargaInput.value);
    const qty = parseFloat(qtyInput.value) || 1;
    const catatan = catInput.value.trim() || '-';

    if (!keterangan) {
        alert("Keterangan tidak boleh kosong!");
        return;
    }
    if (isNaN(harga) || harga <= 0) {
        alert("Masukkan harga angka yang valid dan lebih dari 0!");
        return;
    }

    let notaData = null;
    if (!isMasuk && fileInput && fileInput.files[0]) {
        const file = fileInput.files[0];
        notaData = await new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve({ dataUrl: e.target.result });
            reader.readAsDataURL(file);
        });
    }

    simpanStateKeHistory();

    const totalNominal = harga * qty;

    if (editIndex >= 0) {
        transaksi[editIndex] = {
            divisi,
            jenis,
            keterangan,
            harga,
            qty,
            totalNominal,
            notaData: notaData || transaksi[editIndex].notaData,
            catatan
        };
        editIndex = -1;
        document.getElementById('btnSubmitMasuk').innerText = '+ Tambah Uang Masuk';
        document.getElementById('btnSubmitKeluar').innerText = '+ Tambah Uang Keluar';
    } else {
        transaksi.push({
            divisi,
            jenis,
            keterangan,
            harga,
            qty,
            totalNominal,
            notaData,
            catatan
        });
    }

    ketInput.value = '';
    hargaInput.value = '';
    qtyInput.value = '1';
    catInput.value = '';
    if (fileInput) fileInput.value = '';
    hitungSubtotal(jenis);

    simpanKeStorage(false);
    updateTampilan();
}

function editTransaksi(index) {
    const item = transaksi[index];
    editIndex = index;
    const isMasuk = item.jenis === 'Masuk';

    if (isMasuk) {
        const divEl = document.getElementById('divisiMasuk');
        divEl.value = item.divisi;
        document.getElementById('ketMasuk').value = item.keterangan;
        document.getElementById('hargaMasuk').value = item.harga;
        document.getElementById('jumlahMasuk').value = item.qty;
        document.getElementById('catMasuk').value = item.catatan === '-' ? '' : item.catatan;
        document.getElementById('btnSubmitMasuk').innerText = '💾 Simpan Perubahan (Masuk)';
        hitungSubtotal('Masuk');

        const card = document.getElementById('cardMasuk');
        card.scrollIntoView({ behavior: 'smooth', block: 'center' });
        setTimeout(() => divEl.focus(), 400);
    } else {
        const divEl = document.getElementById('divisiKeluar');
        divEl.value = item.divisi;
        document.getElementById('ketKeluar').value = item.keterangan;
        document.getElementById('hargaKeluar').value = item.harga;
        document.getElementById('jumlahKeluar').value = item.qty;
        document.getElementById('catKeluar').value = item.catatan === '-' ? '' : item.catatan;
        document.getElementById('btnSubmitKeluar').innerText = '💾 Simpan Perubahan (Keluar)';
        hitungSubtotal('Keluar');

        const card = document.getElementById('cardKeluar');
        card.scrollIntoView({ behavior: 'smooth', block: 'center' });
        setTimeout(() => divEl.focus(), 400);
    }
}

function hapusTransaksi(index) {
    simpanStateKeHistory();
    transaksi.splice(index, 1);
    simpanKeStorage(false);
    updateTampilan();
}

function updateSummary() {
    const totalPeserta = document.getElementById('jumlahPeserta').value || 0;
    let totalMasuk = 0;
    let totalKeluar = 0;

    transaksi.forEach(t => {
        if (t.jenis === 'Masuk') totalMasuk += t.totalNominal;
        if (t.jenis === 'Keluar') totalKeluar += t.totalNominal;
    });

    const saldoAkhir = totalMasuk - totalKeluar;

    document.getElementById('dispPeserta').innerText = `${totalPeserta} Orang`;
    document.getElementById('dispMasuk').innerText = formatRupiah(totalMasuk);
    document.getElementById('dispKeluar').innerText = formatRupiah(totalKeluar);
    document.getElementById('dispSaldo').innerText = formatRupiah(saldoAkhir);

    return { totalPeserta, totalMasuk, totalKeluar, saldoAkhir };
}

function updateTampilan() {
    const tbody = document.getElementById('tabelBody');
    const tfoot = document.getElementById('tabelFoot');
    tbody.innerHTML = '';
    tfoot.innerHTML = '';

    if (transaksi.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" style="text-align: center; color: #888;">Belum ada transaksi di draf ini.</td></tr>';
        updateSummary();
        return;
    }

    const kelompokDivisi = {};
    transaksi.forEach((t, realIdx) => {
        const divName = t.divisi || 'Umum';
        if (!kelompokDivisi[divName]) {
            kelompokDivisi[divName] = [];
        }
        kelompokDivisi[divName].push({ ...t, realIndex: realIdx });
    });

    let grandTotalMasuk = 0;
    let grandTotalKeluar = 0;

    for (const [divisiName, items] of Object.entries(kelompokDivisi)) {
        tbody.innerHTML += `
            <tr class="row-divisi-header">
                <td colspan="9">${divisiName}</td>
            </tr>
        `;

        items.forEach((t, i) => {
            if (t.jenis === 'Masuk') grandTotalMasuk += t.totalNominal;
            if (t.jenis === 'Keluar') grandTotalKeluar += t.totalNominal;

            let notaHtml = '-';
            if (t.notaData) {
                notaHtml = `
                    <img src="${t.notaData.dataUrl}" class="img-preview" alt="Nota">
                    <div style="text-align: center;">
                        <span class="btn-nota-action" onclick="triggerGantiNota(${t.realIndex})">✏️ Ganti</span>
                        <span class="btn-nota-delete" onclick="hapusNota(${t.realIndex})">🗑️ Hapus</span>
                    </div>
                `;
            } else {
                notaHtml = `
                    <div style="text-align: center;">
                        <span style="color: #888; font-size: 11px;">Tidak ada</span><br>
                        <span class="btn-nota-action" onclick="triggerGantiNota(${t.realIndex})">➕ Upload</span>
                    </div>
                `;
            }

            tbody.innerHTML += `<tr>
                <td style="text-align: center;">${i + 1}</td>
                <td><span class="${t.jenis === 'Masuk' ? 'badge-masuk' : 'badge-keluar'}">${t.jenis}</span></td>
                <td>${t.keterangan}</td>
                <td>${formatRupiah(t.harga)}</td>
                <td style="text-align: center;">${t.qty}</td>
                <td><b>${formatRupiah(t.totalNominal)}</b></td>
                <td style="text-align: center;">${notaHtml}</td>
                <td>${t.catatan}</td>
                <td class="col-action">
                    <button class="btn-edit" onclick="editTransaksi(${t.realIndex})">Edit</button>
                    <button class="btn-delete" onclick="hapusTransaksi(${t.realIndex})">Hapus</button>
                </td>
            </tr>`;
        });
    }

    const summary = updateSummary();
    tfoot.innerHTML = `
        <tr class="tfoot-summary">
            <td colspan="5" style="text-align: left;">Total Uang Masuk:</td>
            <td colspan="4" style="color: #28a745;">${formatRupiah(summary.totalMasuk)}</td>
        </tr>
        <tr class="tfoot-summary">
            <td colspan="5" style="text-align: left;">Total Uang Keluar:</td>
            <td colspan="4" style="color: #dc3545;">${formatRupiah(summary.totalKeluar)}</td>
        </tr>
        <tr class="tfoot-summary" style="background-color: #e2e8f0; font-size: 14px;">
            <td colspan="5" style="text-align: left;">Total Akhir (Uang Masuk - Uang Keluar):</td>
            <td colspan="4">${formatRupiah(summary.saldoAkhir)}</td>
        </tr>
    `;

    updateSummary();
}

function dataURLtoUint8Array(dataurl) {
    const arr = dataurl.split(',');
    const bstr = atob(arr[1]);
    let n = bstr.length;
    const u8arr = new Uint8Array(n);
    while (n--) {
        u8arr[n] = bstr.charCodeAt(n);
    }
    return u8arr;
}

/* ------------------- EKSPOR PDF (UNTUK DI-DOWNLOAD / CETAK) ------------------- */
function exportToPDF() {
    const draft = listDraft.find(d => d.id === currentDraftId);
    const fileName = draft ? `${draft.nama.replace(/\s+/g, '_')}.pdf` : 'Rekapan_Keuangan.pdf';
    
    const summary = updateSummary();
    const element = document.createElement('div');
    element.style.padding = '10px';
    element.style.fontFamily = "'Times New Roman', Times, serif";
    element.style.backgroundColor = '#ffffff';
    element.style.color = '#000000';
    element.style.fontSize = '11pt';

    const kelompokDivisi = {};
    transaksi.forEach(t => {
        const divName = t.divisi || 'Umum';
        if (!kelompokDivisi[divName]) kelompokDivisi[divName] = [];
        kelompokDivisi[divName].push(t);
    });

    const logoHtml = `<img src="${GCC_LOGO_URL}" style="height: 60px; width: auto; margin-bottom: 10px;" crossorigin="anonymous"><br>`;

    let htmlContent = `
        <style>
            .pdf-block { page-break-inside: avoid !important; margin-bottom: 15px; }
            .pdf-table { width: 100%; border-collapse: collapse; font-size: 10.5pt; table-layout: fixed; }
            .pdf-table th, .pdf-table td { border: 1px solid #000; padding: 5px 6px; text-align: left; vertical-align: middle; word-wrap: break-word; }
            .pdf-table th { text-align: center; font-weight: bold; background-color: #f2f2f2; }
            .pdf-tr { page-break-inside: avoid !important; page-break-after: auto !important; }
        </style>

        <div class="pdf-block" style="text-align: center;">
            ${logoHtml}
            <h2 style="font-size: 15pt; font-weight: bold; margin: 0 0 15px 0;">LAPORAN REKAPAN KEUANGAN WORKSHOP GCC</h2>
        </div>
        
        <div class="pdf-block">
            <h3 style="font-size: 12pt; font-weight: bold; margin: 0 0 8px 0;">RINGKASAN KEGIATAN</h3>
            <table style="width: 100%; border-collapse: collapse; border: none; font-size: 10.5pt;">
                <tr>
                    <td style="border: none; padding: 3px 0; width: 35%;">• Total Peserta Hadir</td>
                    <td style="border: none; padding: 3px 0; width: 65%;">: ${summary.totalPeserta} Orang</td>
                </tr>
                <tr>
                    <td style="border: none; padding: 3px 0;">• Total Uang Masuk</td>
                    <td style="border: none; padding: 3px 0;">: ${formatRupiah(summary.totalMasuk)}</td>
                </tr>
                <tr>
                    <td style="border: none; padding: 3px 0;">• Total Uang Keluar</td>
                    <td style="border: none; padding: 3px 0;">: ${formatRupiah(summary.totalKeluar)}</td>
                </tr>
                <tr>
                    <td style="border: none; padding: 3px 0; font-weight: bold;">• Sisa Saldo Akhir</td>
                    <td style="border: none; padding: 3px 0; font-weight: bold;">: ${formatRupiah(summary.saldoAkhir)} (Uang Masuk - Uang Keluar)</td>
                </tr>
            </table>
        </div>

        <div class="pdf-block" style="margin-bottom: 5px;">
            <h3 style="font-size: 12pt; font-weight: bold; margin: 0;">RINCIAN TRANSAKSI</h3>
        </div>
    `;

    let lampiranNotaHtml = '';
    let globalNotaIndex = 1;

    htmlContent += `<table class="pdf-table">
        <thead>
            <tr class="pdf-tr">
                <th style="width: 5%;">No</th>
                <th style="width: 12%;">Jenis</th>
                <th style="width: 38%;">Keterangan</th>
                <th style="width: 18%;">Harga Satuan</th>
                <th style="width: 9%;">Jumlah</th>
                <th style="width: 18%;">Total Nominal</th>
            </tr>
        </thead>
        <tbody>`;

    for (const [divisiName, items] of Object.entries(kelompokDivisi)) {
        htmlContent += `
            <tr class="pdf-tr">
                <td colspan="6" style="font-weight: bold; background-color: #f9f9f9; padding-left: 8px;">${divisiName}</td>
            </tr>
        `;

        items.forEach((t, idx) => {
            htmlContent += `
                <tr class="pdf-tr">
                    <td style="text-align: center;">${idx + 1}</td>
                    <td style="padding-left: 6px;">${t.jenis}</td>
                    <td style="padding-left: 6px;">${t.keterangan}</td>
                    <td style="padding-left: 6px;">${formatRupiah(t.harga)}</td>
                    <td style="text-align: center;">${t.qty}</td>
                    <td style="padding-left: 6px;">${formatRupiah(t.totalNominal)}</td>
                </tr>
            `;

            if (t.notaData) {
                lampiranNotaHtml += `
                    <div class="pdf-block">
                        <p style="font-size: 10.5pt; font-weight: bold; margin: 0 0 6px 0;">
                            • Nota Transaksi #${globalNotaIndex++} (${divisiName}): ${t.keterangan} (${formatRupiah(t.totalNominal)})
                        </p>
                        <img src="${t.notaData.dataUrl}" style="max-width: 250px; max-height: 250px; border: 1px solid #ccc; padding: 4px; border-radius: 4px;">
                    </div>
                `;
            }
        });
    }

    htmlContent += `
            <tr class="pdf-tr">
                <td colspan="5" style="font-weight: bold; padding-left: 6px;">Total Uang Masuk</td>
                <td style="font-weight: bold; padding-left: 6px;">${formatRupiah(summary.totalMasuk)}</td>
            </tr>
            <tr class="pdf-tr">
                <td colspan="5" style="font-weight: bold; padding-left: 6px;">Total Uang Keluar</td>
                <td style="font-weight: bold; padding-left: 6px;">${formatRupiah(summary.totalKeluar)}</td>
            </tr>
            <tr class="pdf-tr">
                <td colspan="5" style="font-weight: bold; padding-left: 6px;">Total Akhir (Uang Masuk - Uang Keluar)</td>
                <td style="font-weight: bold; padding-left: 6px;">${formatRupiah(summary.saldoAkhir)}</td>
            </tr>
        </tbody>
    </table>`;

    if (lampiranNotaHtml !== '') {
        htmlContent += `
            <div class="pdf-block" style="margin-top: 25px;">
                <h3 style="font-size: 12pt; font-weight: bold; margin: 0 0 12px 0;">LAMPIRAN NOTA PENGELUARAN</h3>
                ${lampiranNotaHtml}
            </div>
        `;
    }

    element.innerHTML = htmlContent;

    const opt = {
        margin:       [12, 12, 12, 12],
        filename:     fileName,
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2, useCORS: true, scrollY: 0 },
        jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' },
        pagebreak:    { mode: ['css', 'legacy'] }
    };

    html2pdf().set(opt).from(element).save();
}

/* ------------------- EKSPOR WORD (UNTUK UPLOAD / EDIT ULANG) ------------------- */
async function exportToWord() {
    const summary = updateSummary();
    const draft = listDraft.find(d => d.id === currentDraftId);
    const fileName = draft ? `${draft.nama.replace(/\s+/g, '_')}.docx` : 'Rekapan_Keuangan.docx';
    const { docx } = window;

    const borderKosong = { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" };
    const noBordersSummary = {
        top: borderKosong, bottom: borderKosong, left: borderKosong, right: borderKosong,
        insideHorizontal: borderKosong, insideVertical: borderKosong
    };

    const customIndentLeft = { left: 72 };

    const kelompokDivisi = {};
    transaksi.forEach(t => {
        const divName = t.divisi || 'Umum';
        if (!kelompokDivisi[divName]) kelompokDivisi[divName] = [];
        kelompokDivisi[divName].push(t);
    });

    const tableRows = [
        new docx.TableRow({
            children: [
                new docx.TableCell({ children: [new docx.Paragraph({ text: "No", bold: true, alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Jenis", bold: true, alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Keterangan", bold: true, alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Harga Satuan", bold: true, alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Jumlah", bold: true, alignment: docx.AlignmentType.CENTER })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: "Total Nominal", bold: true, alignment: docx.AlignmentType.CENTER })] })
            ]
        })
    ];

    const lampiranNotaChildren = [];
    let globalNotaIndex = 1;

    for (const [divisiName, items] of Object.entries(kelompokDivisi)) {
        tableRows.push(
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        columnSpan: 6,
                        children: [new docx.Paragraph({ text: divisiName, bold: true, indent: customIndentLeft })]
                    })
                ]
            })
        );

        items.forEach((t, index) => {
            tableRows.push(
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ children: [new docx.Paragraph({ text: String(index + 1), alignment: docx.AlignmentType.CENTER })] }),
                        new docx.TableCell({ children: [new docx.Paragraph({ text: t.jenis, indent: customIndentLeft })] }),
                        new docx.TableCell({ children: [new docx.Paragraph({ text: t.keterangan, indent: customIndentLeft })] }),
                        new docx.TableCell({ children: [new docx.Paragraph({ text: formatRupiah(t.harga), indent: customIndentLeft })] }),
                        new docx.TableCell({ children: [new docx.Paragraph({ text: String(t.qty), alignment: docx.AlignmentType.CENTER })] }),
                        new docx.TableCell({ children: [new docx.Paragraph({ text: formatRupiah(t.totalNominal), indent: customIndentLeft })] })
                    ]
                })
            );

            if (t.notaData) {
                try {
                    const imageBytes = dataURLtoUint8Array(t.notaData.dataUrl);
                    lampiranNotaChildren.push(
                        new docx.Paragraph({
                            children: [new docx.TextRun({ text: `• Nota Transaksi #${globalNotaIndex++} (${divisiName}): ${t.keterangan} (${formatRupiah(t.totalNominal)})`, bold: true, color: "000000" })],
                            spacing: { before: 100, after: 60, line: 276 }
                        }),
                        new docx.Paragraph({
                            children: [new docx.ImageRun({ data: imageBytes, transformation: { width: 250, height: 250 } })],
                            spacing: { after: 150, line: 276 }
                        })
                    );
                } catch (e) {}
            }
        });
    }

    tableRows.push(
        new docx.TableRow({
            children: [
                new docx.TableCell({ columnSpan: 5, children: [new docx.Paragraph({ text: "Total Uang Masuk", bold: true, alignment: docx.AlignmentType.LEFT, indent: customIndentLeft })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: formatRupiah(summary.totalMasuk), bold: true, indent: customIndentLeft })] })
            ]
        }),
        new docx.TableRow({
            children: [
                new docx.TableCell({ columnSpan: 5, children: [new docx.Paragraph({ text: "Total Uang Keluar", bold: true, alignment: docx.AlignmentType.LEFT, indent: customIndentLeft })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: formatRupiah(summary.totalKeluar), bold: true, indent: customIndentLeft })] })
            ]
        }),
        new docx.TableRow({
            children: [
                new docx.TableCell({ columnSpan: 5, children: [new docx.Paragraph({ text: "Total Akhir (Uang Masuk - Uang Keluar)", bold: true, alignment: docx.AlignmentType.LEFT, indent: customIndentLeft })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: formatRupiah(summary.saldoAkhir), bold: true, indent: customIndentLeft })] })
            ]
        })
    );

    const docChildren = [
        new docx.Paragraph({
            children: [new docx.TextRun({ text: "LAPORAN REKAPAN KEUANGAN WORKSHOP GCC", bold: true, color: "000000", size: 28 })],
            alignment: docx.AlignmentType.CENTER,
            spacing: { after: 200, line: 276 }
        }),
        
        new docx.Paragraph({
            children: [new docx.TextRun({ text: "RINGKASAN KEGIATAN", bold: true, color: "000000", size: 24 })],
            spacing: { before: 150, after: 100, line: 276 }
        }),

        new docx.Table({
            borders: noBordersSummary,
            rows: [
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ width: { size: 35, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ text: "• Total Peserta Hadir", color: "000000" })] }),
                        new docx.TableCell({ width: { size: 65, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ text: `: ${summary.totalPeserta} Orang`, color: "000000" })] })
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ width: { size: 35, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ text: "• Total Uang Masuk", color: "000000" })] }),
                        new docx.TableCell({ width: { size: 65, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ text: `: ${formatRupiah(summary.totalMasuk)}`, color: "000000" })] })
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ width: { size: 35, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ text: "• Total Uang Keluar", color: "000000" })] }),
                        new docx.TableCell({ width: { size: 65, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ text: `: ${formatRupiah(summary.totalKeluar)}`, color: "000000" })] })
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ width: { size: 35, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ text: "• Sisa Saldo Akhir", bold: true, color: "000000" })] }),
                        new docx.TableCell({ width: { size: 65, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ text: `: ${formatRupiah(summary.saldoAkhir)} (Uang Masuk - Uang Keluar)`, bold: true, color: "000000" })] })
                    ]
                })
            ]
        }),

        new docx.Paragraph({ text: "", spacing: { after: 150, line: 276 } }),

        new docx.Paragraph({
            children: [new docx.TextRun({ text: "RINCIAN TRANSAKSI", bold: true, color: "000000", size: 24 })],
            spacing: { before: 150, after: 100, line: 276 }
        }),

        new docx.Table({
            rows: tableRows,
            width: { size: 100, type: docx.WidthType.PERCENTAGE }
        })
    ];

    if (lampiranNotaChildren.length > 0) {
        docChildren.push(
            new docx.Paragraph({ text: "", spacing: { after: 200, line: 276 } }),
            new docx.Paragraph({
                children: [new docx.TextRun({ text: "LAMPIRAN NOTA PENGELUARAN", bold: true, color: "000000", size: 24 })],
                spacing: { before: 150, after: 100, line: 276 }
            }),
            ...lampiranNotaChildren
        );
    }

    const transaksiLight = transaksi.map(t => ({
        divisi: t.divisi,
        jenis: t.jenis,
        keterangan: t.keterangan,
        harga: t.harga,
        qty: t.qty,
        totalNominal: t.totalNominal,
        catatan: t.catatan,
        notaData: null
    }));

    const metadataObj = {
        jumlahPeserta: document.getElementById('jumlahPeserta').value || 0,
        transaksi: transaksiLight
    };
    const metadataString = encodeURIComponent(JSON.stringify(metadataObj));

    docChildren.push(
        new docx.Paragraph({
            children: [new docx.TextRun({ text: `GCC_DATA_START:::${metadataString}:::GCC_DATA_END`, color: "FFFFFF", size: 2 })]
        })
    );

    const doc = new docx.Document({
        styles: {
            default: {
                document: {
                    run: { font: "Times New Roman", size: 22, color: "000000" },
                    paragraph: { spacing: { line: 276, before: 60, after: 60 } }
                }
            }
        },
        sections: [{ properties: {}, children: docChildren }]
    });

    docx.Packer.toBlob(doc).then(blob => {
        saveAs(blob, fileName);
    });
}