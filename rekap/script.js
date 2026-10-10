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

window.onload = function() {
    initTheme();
    muatSemuaDraft();
    initCharts();
    
    document.addEventListener('keydown', function(e) {
        const activeElement = document.activeElement;
        const isInputFocused = activeElement && (activeElement.tagName === 'INPUT' || activeElement.tagName === 'TEXTAREA');

        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
            if (e.shiftKey) { e.preventDefault(); redo(); }
            else if (!isInputFocused || isInputFocused) { e.preventDefault(); undo(); }
        } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
            e.preventDefault();
            redo();
        }
    });
};

function switchTab(tabId, btnElement) {
    document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.nav-tab').forEach(btn => btn.classList.remove('active'));

    document.getElementById(tabId).classList.add('active');
    btnElement.classList.add('active');

    if (tabId === 'tab-analisis') {
        setTimeout(() => updateTampilan(), 50);
    }
}

function getFileNameFormat(extension) {
    const draft = listDraft.find(d => d.id === currentDraftId);
    const rawName = draft && draft.nama ? draft.nama : 'Rekapan_Keuangan_GCC';
    const safeName = rawName.replace(/[^a-zA-Z0-9_-]/g, '_');
    return `${safeName}.${extension}`;
}

function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = 'toast';
    toast.innerText = message;

    if (type === 'success') toast.style.borderLeftColor = '#10b981';
    if (type === 'error') toast.style.borderLeftColor = '#f43f5e';

    container.appendChild(toast);
    setTimeout(() => toast.remove(), 3200);
}

function initTheme() {
    const savedTheme = localStorage.getItem('gcc_theme');
    if (savedTheme === 'dark') {
        document.body.classList.add('dark-mode');
        document.getElementById('themeToggle').innerText = '☀️ Mode Terang';
    }
}

function toggleTheme() {
    document.body.classList.toggle('dark-mode');
    const isDark = document.body.classList.contains('dark-mode');
    localStorage.setItem('gcc_theme', isDark ? 'dark' : 'light');
    document.getElementById('themeToggle').innerText = isDark ? '☀️ Mode Terang' : '🌙 Mode Gelap';
}

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
    showToast("✅ Lampiran bukti/nota diperbarui!", "success");
}

function hapusNota(realIndex) {
    if (confirm("Apakah Anda yakin ingin menghapus lampiran nota/bukti transaksi ini?")) {
        simpanStateKeHistory();
        transaksi[realIndex].notaData = null;
        simpanKeStorage(false);
        updateTampilan();
        showToast("🗑️ Lampiran dihapus", "info");
    }
}

function hitungSimulasiHTM() {
    const summary = updateSummary();
    const targetPeserta = parseFloat(document.getElementById('simPeserta').value) || 0;
    const sponsor = parseFloat(document.getElementById('simSponsor').value) || 0;
    const totalPengeluaran = summary.totalKeluar;

    if (targetPeserta <= 0) {
        document.getElementById('simulasiResult').innerText = "Masukkan jumlah estimasi peserta yang valid.";
        return;
    }

    const sisaBeban = totalPengeluaran - sponsor;
    if (sisaBeban <= 0) {
        document.getElementById('simulasiResult').innerText = "🎉 Dana Sponsorship/Kas sudah menutupi seluruh pengeluaran! HTM Peserta bisa Rp 0 (Gratis).";
    } else {
        const htmPerPeserta = Math.ceil((sisaBeban / targetPeserta) / 1000) * 1000;
        document.getElementById('simulasiResult').innerHTML = `
            💡 <b>Hasil Simulasi:</b><br>
            • Sisa Beban Pengeluaran: <b>${formatRupiah(sisaBeban)}</b><br>
            • Estimasi HTM Minimal per Peserta: <b style="color:#10b981; font-size: 1.1rem;">${formatRupiah(htmPerPeserta)}</b>
        `;
    }
}

function initCharts() {
    const ctxComp = document.getElementById('chartComparison').getContext('2d');
    chartComp = new Chart(ctxComp, {
        type: 'bar',
        data: {
            labels: ['Total Masuk', 'Total Keluar'],
            datasets: [{
                label: 'Nominal (Rp)',
                data: [0, 0],
                backgroundColor: ['#10b981', '#f43f5e'],
                borderRadius: 8
            }]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });

    const ctxDiv = document.getElementById('chartDivisi').getContext('2d');
    chartDiv = new Chart(ctxDiv, {
        type: 'doughnut',
        data: {
            labels: [],
            datasets: [{
                data: [],
                backgroundColor: ['#6366f1', '#10b981', '#f59e0b', '#ec4899', '#8b5cf6', '#06b6d4', '#f43f5e']
            }]
        },
        options: { responsive: true, maintainAspectRatio: false }
    });
}

function updateCharts(totalMasuk, totalKeluar, kelompokDivisi) {
    if (!chartComp || !chartDiv) return;

    chartComp.data.datasets[0].data = [totalMasuk, totalKeluar];
    chartComp.update();

    const divLabels = [];
    const divTotals = [];

    for (const [divName, items] of Object.entries(kelompokDivisi)) {
        let sumKeluar = 0;
        items.forEach(t => { if (t.jenis === 'Keluar') sumKeluar += t.totalNominal; });
        if (sumKeluar > 0) {
            divLabels.push(divName);
            divTotals.push(sumKeluar);
        }
    }

    chartDiv.data.labels = divLabels;
    chartDiv.data.datasets[0].data = divTotals;
    chartDiv.update();
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
    showToast("📥 Templat Word berhasil di-download!", "success");
}

async function imporDokumen(event) {
    const file = event.target.files[0];
    if (!file || !file.name.endsWith('.docx')) {
        showToast("⚠️ Pilih file Word (.docx)!", "error");
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
            if (Array.isArray(decodedJson.transaksi)) transaksi = decodedJson.transaksi;
            simpanKeStorage(false);
            updateTampilan();
            showToast("🎉 Data draf dipulihkan dari file Word!", "success");
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
            showToast(`🎉 Berhasil mengimpor ${count} item transaksi!`, "success");
        } else {
            showToast("⚠️ Format tabel tidak terbaca.", "error");
        }
    } catch (e) {
        showToast("❌ Gagal membaca file Word.", "error");
    }
}

function simpanStateKeHistory() {
    undoStack.push({ transaksi: JSON.parse(JSON.stringify(transaksi)), jumlahPeserta: document.getElementById('jumlahPeserta').value || 0 });
    redoStack = [];
    updateTombolHistory();
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

function muatSemuaDraft() {
    const saved = localStorage.getItem('gcc_list_draft');
    if (saved) listDraft = JSON.parse(saved);

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
    if (userClick) showToast("➕ Draf baru berhasil dibuat!", "success");
}

function ubahNamaDraft() {
    const draft = listDraft.find(d => d.id === currentDraftId);
    if (!draft) return;
    const namaBaru = prompt("Ubah Nama Draf:", draft.nama);
    if (namaBaru && namaBaru.trim() !== '') {
        draft.nama = namaBaru.trim();
        simpanListDraft();
        renderDropdownDraft();
        showToast("✏️ Nama draf diperbarui", "success");
    }
}

function gantiDraft() {
    currentDraftId = document.getElementById('selectDraft').value;
    muatDraftAktif();
}

function muatDraftAktif() {
    const draft = listDraft.find(d => d.id === currentDraftId);
    if (draft) {
        document.getElementById('jumlahPeserta').value = draft.jumlahPeserta || 0;
        transaksi = draft.transaksi || [];
        undoStack = []; redoStack = [];
        updateTombolHistory();
        updateTampilan();
    }
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

function simpanListDraft() { localStorage.setItem('gcc_list_draft', JSON.stringify(listDraft)); }

function hapusDraftAktif() {
    if (confirm("Apakah Anda yakin ingin menghapus draf ini?")) {
        listDraft = listDraft.filter(d => d.id !== currentDraftId);
        simpanListDraft();
        muatSemuaDraft();
        showToast("🗑️ Draf dihapus", "info");
    }
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
        showToast("⚠️ Isi Keterangan dan Harga Nominal yang valid!", "error");
        return;
    }

    let notaData = null;
    if (fileEl && fileEl.files[0]) {
        notaData = await new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve({ dataUrl: e.target.result });
            reader.readAsDataURL(fileEl.files[0]);
        });
    }

    simpanStateKeHistory();

    if (editIndex >= 0) {
        transaksi[editIndex] = { divisi, jenis, keterangan, harga, qty, totalNominal: harga * qty, notaData: notaData || transaksi[editIndex].notaData, catatan };
        editIndex = -1;
        document.getElementById('btnSubmitMasuk').innerText = '+ Tambah Uang Masuk';
        document.getElementById('btnSubmitKeluar').innerText = '+ Tambah Uang Keluar';
        showToast("💾 Perubahan transaksi disimpan!", "success");
    } else {
        transaksi.push({ divisi, jenis, keterangan, harga, qty, totalNominal: harga * qty, notaData, catatan });
        showToast("✅ Transaksi berhasil ditambahkan!", "success");
    }

    ketEl.value = ''; hargaEl.value = ''; qtyEl.value = '1'; catEl.value = ''; if (fileEl) fileEl.value = '';
    hitungSubtotal(jenis);
    simpanKeStorage(false);
    updateTampilan();
}

function editTransaksi(index) {
    const item = transaksi[index];
    editIndex = index;
    const isMasuk = item.jenis === 'Masuk';
    document.getElementById(isMasuk ? 'divisiMasuk' : 'divisiKeluar').value = item.divisi;
    document.getElementById(isMasuk ? 'ketMasuk' : 'ketKeluar').value = item.keterangan;
    document.getElementById(isMasuk ? 'hargaMasuk' : 'hargaKeluar').value = item.harga;
    document.getElementById(isMasuk ? 'jumlahMasuk' : 'jumlahKeluar').value = item.qty;
    document.getElementById(isMasuk ? 'catMasuk' : 'catKeluar').value = item.catatan === '-' ? '' : item.catatan;
    document.getElementById(isMasuk ? 'btnSubmitMasuk' : 'btnSubmitKeluar').innerText = '💾 Simpan Perubahan';
    hitungSubtotal(item.jenis);
    document.getElementById(isMasuk ? 'cardMasuk' : 'cardKeluar').scrollIntoView({ behavior: 'smooth' });
}

function hapusTransaksi(index) {
    simpanStateKeHistory();
    transaksi.splice(index, 1);
    simpanKeStorage(false);
    updateTampilan();
    showToast("🗑️ Transaksi dihapus", "info");
}

function updateSummary() {
    const totalPeserta = document.getElementById('jumlahPeserta').value || 0;
    let totalMasuk = 0, totalKeluar = 0;

    transaksi.forEach(t => {
        if (t.jenis === 'Masuk') totalMasuk += t.totalNominal;
        if (t.jenis === 'Keluar') totalKeluar += t.totalNominal;
    });

    const saldoAkhir = totalMasuk - totalKeluar;
    document.getElementById('dispPeserta').innerText = `${totalPeserta} Orang`;
    document.getElementById('dispMasuk').innerText = formatRupiah(totalMasuk);
    document.getElementById('dispKeluar').innerText = formatRupiah(totalKeluar);
    document.getElementById('dispSaldo').innerText = formatRupiah(saldoAkhir);

    const percent = totalMasuk > 0 ? Math.min(Math.round((totalKeluar / totalMasuk) * 100), 100) : 0;
    const bar = document.getElementById('progressBar');
    bar.style.width = percent + '%';
    document.getElementById('progressText').innerText = percent + '%';
    bar.style.background = percent > 90 ? '#f43f5e' : 'linear-gradient(135deg, #6366f1 0%, #4f46e5 100%)';

    return { totalPeserta, totalMasuk, totalKeluar, saldoAkhir };
}

function updateTampilan() {
    const tbody = document.getElementById('tabelBody');
    const tfoot = document.getElementById('tabelFoot');
    tbody.innerHTML = ''; tfoot.innerHTML = '';

    const keyword = (document.getElementById('searchKeyword').value || '').toLowerCase();
    const filterDiv = document.getElementById('filterDivisi').value;
    const filterJen = document.getElementById('filterJenis').value;

    const filtered = transaksi.filter(t => {
        const matchKey = t.keterangan.toLowerCase().includes(keyword) || t.divisi.toLowerCase().includes(keyword);
        const matchDiv = filterDiv === 'ALL' || t.divisi === filterDiv;
        const matchJen = filterJen === 'ALL' || t.jenis === filterJen;
        return matchKey && matchDiv && matchJen;
    });

    const listDivisiUnik = [...new Set(transaksi.map(t => t.divisi))];
    const selectDiv = document.getElementById('filterDivisi');
    const valDivBefore = selectDiv.value;
    selectDiv.innerHTML = '<option value="ALL">Semua Divisi</option>';
    listDivisiUnik.forEach(d => {
        const opt = document.createElement('option'); opt.value = d; opt.innerText = d;
        if (d === valDivBefore) opt.selected = true;
        selectDiv.appendChild(opt);
    });

    const kelompokDivisi = {};
    filtered.forEach((t, realIdx) => {
        const divName = t.divisi || 'CONTER';
        if (!kelompokDivisi[divName]) kelompokDivisi[divName] = [];
        kelompokDivisi[divName].push({ ...t, realIndex: realIdx });
    });

    if (filtered.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" style="text-align: center; color: #888;">Belum ada transaksi yang cocok.</td></tr>';
        const summary = updateSummary();
        updateCharts(summary.totalMasuk, summary.totalKeluar, kelompokDivisi);
        return;
    }

    for (const [divisiName, items] of Object.entries(kelompokDivisi)) {
        tbody.innerHTML += `<tr class="row-divisi-header"><td colspan="9">${divisiName}</td></tr>`;

        items.forEach((t, i) => {
            let notaHtml = '-';
            if (t.notaData) {
                notaHtml = `
                    <img src="${t.notaData.dataUrl}" class="img-preview" alt="Bukti" onclick="openModal('${t.notaData.dataUrl}', '${t.keterangan}')">
                    <div style="text-align: center; margin-top:2px;">
                        <span class="btn-nota-action" onclick="triggerGantiNota(${t.realIndex})">✏️</span>
                        <span class="btn-nota-delete" onclick="hapusNota(${t.realIndex})">🗑️</span>
                    </div>`;
            } else {
                notaHtml = `<span class="btn-nota-action" onclick="triggerGantiNota(${t.realIndex})">➕ Upload</span>`;
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
        <tr class="tfoot-summary"><td colspan="5">Total Uang Masuk:</td><td colspan="4" style="color: #10b981;">${formatRupiah(summary.totalMasuk)}</td></tr>
        <tr class="tfoot-summary"><td colspan="5">Total Uang Keluar:</td><td colspan="4" style="color: #f43f5e;">${formatRupiah(summary.totalKeluar)}</td></tr>
        <tr class="tfoot-summary" style="background:var(--input-bg);"><td colspan="5">Total Saldo Akhir:</td><td colspan="4">${formatRupiah(summary.saldoAkhir)}</td></tr>
    `;

    updateCharts(summary.totalMasuk, summary.totalKeluar, kelompokDivisi);
}

/* ------------------- GENERATOR BLOB FILE (PDF & WORD) ------------------- */
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
        dataExcel.push([`${namaDraft.toUpperCase()}`]);
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
        const element = document.createElement('div');
        element.style.width = '100%';
        element.style.padding = '15px';
        element.style.fontFamily = "'Times New Roman', serif";
        element.style.color = '#000';
        element.style.fontSize = '11pt';
        element.style.background = '#ffffff';

        let html = `
            <div style="text-align: center; margin-bottom: 15px;">
                <img src="${GCC_LOGO_URL}" style="height: 45px; width: 45px; object-fit: contain; margin-bottom: 4px; display: inline-block;" crossorigin="anonymous"><br>
                <h3 style="margin: 0; text-transform: uppercase;">${namaDraft}</h3>
            </div>
            
            <div style="margin-bottom: 15px;">
                <h4 style="margin: 0 0 6px 0; text-transform: uppercase;">RINGKASAN KEGIATAN</h4>
                <table style="width: 100%; border: none; font-size: 11pt; border-collapse: collapse;">
                    <tr style="border: none;"><td style="width: 190px; border: none; padding: 2px 0 2px 5px; font-weight: bold;">• Total Peserta Hadir</td><td style="width: 15px; border: none; padding: 2px 0;">:</td><td style="border: none; padding: 2px 0;">${summary.totalPeserta} Orang</td></tr>
                    <tr style="border: none;"><td style="border: none; padding: 2px 0 2px 5px; font-weight: bold;">• Total Uang Masuk</td><td style="width: 15px; border: none; padding: 2px 0;">:</td><td style="border: none; padding: 2px 0;">${formatRupiah(summary.totalMasuk)}</td></tr>
                    <tr style="border: none;"><td style="border: none; padding: 2px 0 2px 5px; font-weight: bold;">• Total Uang Keluar</td><td style="width: 15px; border: none; padding: 2px 0;">:</td><td style="border: none; padding: 2px 0;">${formatRupiah(summary.totalKeluar)}</td></tr>
                    <tr style="border: none;"><td style="border: none; padding: 2px 0 2px 5px; font-weight: bold;">• Sisa Saldo Akhir</td><td style="width: 15px; border: none; padding: 2px 0;">:</td><td style="border: none; padding: 2px 0;">${formatRupiah(summary.saldoAkhir)} (Uang Masuk - Uang Keluar)</td></tr>
                </table>
            </div>

            <h4 style="margin: 0 0 6px 0; text-transform: uppercase;">RINCIAN TRANSAKSI</h4>
            <table border="1" style="width: 100%; border-collapse: collapse; font-size: 10pt;">
                <thead>
                    <tr style="background-color: #4b5563; color: #ffffff;">
                        <th style="padding: 6px 6px 6px 12px; text-align: center; width: 35px; color: #ffffff;">No</th>
                        <th style="padding: 6px 6px 6px 12px; text-align: center; width: 65px; color: #ffffff;">Jenis</th>
                        <th style="padding: 6px 6px 6px 12px; text-align: center; color: #ffffff;">Keterangan</th>
                        <th style="padding: 6px 6px 6px 12px; text-align: center; color: #ffffff;">Harga Satuan</th>
                        <th style="padding: 6px 6px 6px 12px; text-align: center; width: 50px; color: #ffffff;">Jumlah</th>
                        <th style="padding: 6px 6px 6px 12px; text-align: center; color: #ffffff;">Total Nominal</th>
                    </tr>
                </thead>
                <tbody>
        `;

        for (const [divName, items] of Object.entries(kelompokDivisi)) {
            html += `<tr><td colspan="6" style="background-color: #fafafa; font-weight: bold; padding: 6px 6px 6px 12px;"><b>${divName}</b></td></tr>`;
            items.forEach((t, i) => {
                html += `
                    <tr>
                        <td style="padding: 5px 5px 5px 12px; text-align: center;">${i + 1}</td>
                        <td style="padding: 5px 5px 5px 12px; text-align: center;">${t.jenis}</td>
                        <td style="padding: 5px 5px 5px 12px;">${t.keterangan}</td>
                        <td style="padding: 5px 5px 5px 12px; text-align: left;">${formatRupiah(t.harga)}</td>
                        <td style="padding: 5px 5px 5px 12px; text-align: center;">${t.qty}</td>
                        <td style="padding: 5px 5px 5px 12px; text-align: left;">${formatRupiah(t.totalNominal)}</td>
                    </tr>
                `;
            });
        }

        html += `
                </tbody>
            </table>
            
            <table border="1" style="width: 100%; border-collapse: collapse; margin-top: -1px; font-size: 10pt; font-weight: bold;">
                <tr>
                    <td style="padding: 6px 6px 6px 12px; text-align: left; font-weight: bold;">Total Uang Masuk</td>
                    <td style="padding: 6px 6px 6px 12px; text-align: left; width: 150px; font-weight: bold;">${formatRupiah(summary.totalMasuk)}</td>
                </tr>
                <tr>
                    <td style="padding: 6px 6px 6px 12px; text-align: left; font-weight: bold;">Total Uang Keluar</td>
                    <td style="padding: 6px 6px 6px 12px; text-align: left; font-weight: bold;">${formatRupiah(summary.totalKeluar)}</td>
                </tr>
                <tr>
                    <td style="padding: 6px 6px 6px 12px; text-align: left; font-weight: bold;">Total Akhir (Uang Masuk - Uang Keluar)</td>
                    <td style="padding: 6px 6px 6px 12px; text-align: left; font-weight: bold;">${formatRupiah(summary.saldoAkhir)}</td>
                </tr>
            </table>
        `;

        if (transaksiBerlampiran.length > 0) {
            html += `
                <div style="page-break-before: always; margin-top: 20px;">
                    <h4 style="text-align: left; text-transform: uppercase; margin-bottom: 15px;">LAMPIRAN BUKTI NOTA & TRANSAKSI</h4>
            `;
            transaksiBerlampiran.forEach((t, idx) => {
                html += `
                    <div style="margin-bottom: 25px;">
                        <p style="font-weight: bold; margin-bottom: 5px; text-align: left;">Lampiran ${idx + 1}: ${t.keterangan} (${t.divisi}) - Total: ${formatRupiah(t.totalNominal)}</p>
                        <div style="text-align: left; width: 100%;">
                            <img src="${t.notaData.dataUrl}" style="max-width: 350px; max-height: 350px; object-fit: contain; border: 1px solid #ccc; padding: 4px; display: inline-block;" crossorigin="anonymous">
                        </div>
                    </div>
                `;
            });
            html += `</div>`;
        }

        element.innerHTML = html;
        const wrapper = document.createElement('div');
        wrapper.style.position = 'fixed';
        wrapper.style.left = '-9999px';
        wrapper.style.top = '0';
        wrapper.style.width = '800px';
        wrapper.appendChild(element);
        document.body.appendChild(wrapper);

        const opt = { margin: 10, filename: getFileNameFormat('pdf'), html2canvas: { scale: 2, useCORS: true, scrollY: 0 }, jsPDF: { unit: 'mm', format: 'a4' } };
        const pdfBlob = await html2pdf().set(opt).from(element).output('blob');
        
        document.body.removeChild(wrapper);
        return pdfBlob;
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
            new docx.Paragraph({ children: [new docx.TextRun({ text: `${namaDraft.toUpperCase()}`, bold: true, size: 24 })], alignment: docx.AlignmentType.CENTER, spacing: { after: 150 } }),
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

        const transaksiLight = transaksi.map(t => ({ ...t, notaData: null }));
        const metadataString = encodeURIComponent(JSON.stringify({ jumlahPeserta: document.getElementById('jumlahPeserta').value || 0, transaksi: transaksiLight }));
        docChildren.push(new docx.Paragraph({ children: [new docx.TextRun({ text: `GCC_DATA_START:::${metadataString}:::GCC_DATA_END`, color: "FFFFFF", size: 2 })] }));

        const doc = new docx.Document({ sections: [{ children: docChildren }] });
        return await docx.Packer.toBlob(doc);
    }
}

async function exportToExcel() {
    const fileName = getFileNameFormat('xlsx');
    const blob = await generateFileBlob('xlsx');
    saveAs(blob, fileName);
    showToast(`📊 File ${fileName} berhasil di-download!`, "success");
}

async function exportToPDF() {
    const fileName = getFileNameFormat('pdf');
    const blob = await generateFileBlob('pdf');
    saveAs(blob, fileName);
    showToast(`📕 File ${fileName} berhasil di-download!`, "success");
}

async function exportToWord() {
    const fileName = getFileNameFormat('docx');
    const blob = await generateFileBlob('docx');
    saveAs(blob, fileName);
    showToast(`📄 File ${fileName} berhasil di-download!`, "success");
}

async function uploadKeDrive(formatType) {
    const dept = document.getElementById('driveDeptSelect').value;
    const category = document.getElementById('driveCategorySelect').value;
    const fileName = getFileNameFormat(formatType);

    const targetFolderId = FOLDER_MAP[dept]?.[category];

    if (!targetFolderId) {
        showToast("⚠️ ID Folder Google Drive belum terdaftar!", "error");
        return;
    }

    showToast(`⏳ Menyiapkan dan mengunggah ${fileName} ke Drive...`, "info");

    try {
        const fileBlob = await generateFileBlob(formatType);

        const reader = new FileReader();
        reader.readAsDataURL(fileBlob);
        reader.onloadend = async function () {
            const base64Data = reader.result.split(',')[1];

            const payload = {
                folderId: targetFolderId,
                fileName: fileName,
                mimeType: fileBlob.type,
                fileData: base64Data
            };

            const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyt9AmgH3Ks5dLjPNDi2oK3gR5PD3M9MDKbtLGrKUaKMgCsmB3WfVS4Fc3MTjHJd1_H/exec";

            try {
                const response = await fetch(GOOGLE_SCRIPT_URL, {
                    method: 'POST',
                    body: JSON.stringify(payload)
                });
                
                const result = await response.json();
                if (result.status === "success" || response.ok) {
                    showToast(`✅ Berhasil mengirim "${fileName}" ke Google Drive!`, "success");
                } else {
                    showToast(`❌ Gagal: ${result.message || 'Unknown error'}`, "error");
                }
            } catch (err) {
                console.error(err);
                showToast("✅ Permintaan unggah terkirim ke Google Drive!", "success");
            }
        };
    } catch (err) {
        console.error(err);
        showToast("❌ Gagal memproses file untuk Google Drive.", "error");
    }
}
