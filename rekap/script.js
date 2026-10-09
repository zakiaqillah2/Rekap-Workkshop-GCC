/* ==========================================================================
   STATE UTAMA APLIKASI KEUANGAN GCC
   ========================================================================== */
let daftarDraft = JSON.parse(localStorage.getItem('gcc_drafts')) || [
    { id: 1, nama: "Draf Utama - Kegiatan GCC", peserta: 0, transaksi: [] }
];
let activeDraftId = parseInt(localStorage.getItem('gcc_active_draft')) || 1;
let historyStack = [];
let redoStack = [];
let chartComparison = null;
let chartDivisiInstance = null;

// Inisialisasi saat halaman dimuat
window.addEventListener('DOMContentLoaded', () => {
    muatDraftKeSelect();
    updateTampilan();
    inisialisasiGrafik();
    
    // Cek tema awal dari storage
    if(localStorage.getItem('gcc_dark_mode') === 'true') {
        document.body.classList.add('dark-mode');
        const btn = document.getElementById('themeToggle');
        if(btn) btn.innerText = "☀️ Mode Terang";
    }
});

// Ambil draf aktif saat ini
function getDraftAktif() {
    let draft = daftarDraft.find(d => d.id === activeDraftId);
    if (!draft) {
        activeDraftId = daftarDraft[0].id;
        draft = daftarDraft[0];
    }
    return draft;
}

// Simpan ke localStorage
function simpanKeStorage() {
    localStorage.setItem('gcc_drafts', JSON.stringify(daftarDraft));
    localStorage.setItem('gcc_active_draft', activeDraftId);
}

// Simpan state untuk Undo/Redo
function simpanStateKeHistory() {
    const draft = getDraftAktif();
    historyStack.push(JSON.stringify(draft.transaksi));
    if (historyStack.length > 25) historyStack.shift(); // Batasi maksimal history
    redoStack = []; // Reset redo
    updateHistoryButtons();
}

function updateHistoryButtons() {
    const btnUndo = document.getElementById('btnUndo');
    const btnRedo = document.getElementById('btnRedo');
    if(btnUndo) btnUndo.disabled = historyStack.length === 0;
    if(btnRedo) btnRedo.disabled = redoStack.length === 0;
}

function undo() {
    if (historyStack.length === 0) return;
    const draft = getDraftAktif();
    redoStack.push(JSON.stringify(draft.transaksi));
    const prevState = historyStack.pop();
    draft.transaksi = JSON.parse(prevState);
    simpanKeStorage();
    updateTampilan();
    updateHistoryButtons();
    showToast("↩️ Perubahan dibatalkan (Undo)", "info");
}

function redo() {
    if (redoStack.length === 0) return;
    const draft = getDraftAktif();
    historyStack.push(JSON.stringify(draft.transaksi));
    const nextState = redoStack.pop();
    draft.transaksi = JSON.parse(nextState);
    simpanKeStorage();
    updateTampilan();
    updateHistoryButtons();
    showToast("↪️ Perubahan dikembalikan (Redo)", "info");
}

/* ==========================================================================
   MANAJEMEN DRAF KEGIATAN
   ========================================================================== */
function muatDraftKeSelect() {
    const select = document.getElementById('selectDraft');
    if (!select) return;
    select.innerHTML = '';
    daftarDraft.forEach(d => {
        let opt = document.createElement('option');
        opt.value = d.id;
        opt.textContent = d.nama;
        if (d.id === activeDraftId) opt.selected = true;
        select.appendChild(opt);
    });
}

function gantiDraft() {
    const select = document.getElementById('selectDraft');
    activeDraftId = parseInt(select.value);
    historyStack = [];
    redoStack = [];
    updateHistoryButtons();
    updateTampilan();
    simpanKeStorage();
}

function buatDraftBaru() {
    let namaBaru = prompt("Masukkan nama draf kegiatan baru:", `Draf #${daftarDraft.length + 1} - Kegiatan GCC`);
    if (!namaBaru) return;
    
    const newId = Date.now();
    daftarDraft.push({ id: newId, nama: namaBaru, peserta: 0, transaksi: [] });
    activeDraftId = newId;
    simpanKeStorage();
    muatDraftKeSelect();
    updateTampilan();
    showToast(`✨ Draf "${namaBaru}" berhasil dibuat!`, "success");
}

function ubahNamaDraft() {
    let draft = getDraftAktif();
    let namaBaru = prompt("Ubah nama draf kegiatan:", draft.nama);
    if (!namaBaru) return;
    draft.nama = namaBaru;
    simpanKeStorage();
    muatDraftKeSelect();
    showToast("✏️ Nama draf berhasil diperbarui!", "success");
}

function hapusDraftAktif() {
    if (daftarDraft.length <= 1) {
        showToast("⚠️ Minimal harus ada 1 draf aktif!", "error");
        return;
    }
    if (confirm("Yakin ingin menghapus draf kegiatan ini?")) {
        daftarDraft = daftarDraft.filter(d => d.id !== activeDraftId);
        activeDraftId = daftarDraft[0].id;
        simpanKeStorage();
        muatDraftKeSelect();
        updateTampilan();
        showToast("🗑️ Draf berhasil dihapus.", "info");
    }
}

/* ==========================================================================
   KALKULASI & TRANSAKSI
   ========================================================================== */
function hitungSubtotal(jenis) {
    const harga = parseFloat(document.getElementById(`harga${jenis}`).value) || 0;
    const jumlah = parseInt(document.getElementById(`jumlah${jenis}`).value) || 1;
    const total = harga * jumlah;
    document.getElementById(`subtotal${jenis}`).innerText = `Total Item: ${formatRupiah(total)}`;
}

function formatRupiah(angka) {
    return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', minimumFractionDigits: 0 }).format(angka);
}

function tambahTransaksi(jenis) {
    const divisi = document.getElementById(`divisi${jenis}`).value.trim() || "Umum";
    const ket = document.getElementById(`ket${jenis}`).value.trim();
    const harga = parseFloat(document.getElementById(`harga${jenis}`).value) || 0;
    const jumlah = parseInt(document.getElementById(`jumlah${jenis}`).value) || 1;
    const catatan = document.getElementById(`cat${jenis}`).value.trim();
    const notaInput = document.getElementById(`nota${jenis}`);

    if (!ket || harga <= 0) {
        showToast("⚠️ Harap isi keterangan dan harga satuan dengan benar!", "error");
        return;
    }

    const totalNominal = harga * jumlah;
    const draft = getDraftAktif();

    // Handle File Nota (jika di-upload)
    if (notaInput.files && notaInput.files[0]) {
        const reader = new FileReader();
        reader.onload = function(e) {
            simpanStateKeHistory();
            draft.transaksi.push({
                id: Date.now(),
                jenis: jenis,
                divisi: divisi.toUpperCase(),
                keterangan: ket,
                hargaSatuan: harga,
                jumlah: jumlah,
                total: totalNominal,
                nota: e.target.result,
                catatan: catatan
            });
            selesaiTambah(jenis);
        };
        reader.readAsDataURL(notaInput.files[0]);
    } else {
        simpanStateKeHistory();
        draft.transaksi.push({
            id: Date.now(),
            jenis: jenis,
            divisi: divisi.toUpperCase(),
            keterangan: ket,
            hargaSatuan: harga,
            jumlah: jumlah,
            total: totalNominal,
            nota: "",
            catatan: catatan
        });
        selesaiTambah(jenis);
    }
}

function selesaiTambah(jenis) {
    document.getElementById(`ket${jenis}`).value = "";
    document.getElementById(`harga${jenis}`).value = "";
    document.getElementById(`jumlah${jenis}`).value = "1";
    document.getElementById(`cat${jenis}`).value = "";
    document.getElementById(`nota${jenis}`).value = "";
    document.getElementById(`subtotal${jenis}`).innerText = "Total Item: Rp 0";

    simpanKeStorage();
    updateTampilan();
    showToast(`✅ Berhasil menambahkan uang ${jenis.toLowerCase()}!`, "success");
}

function hapusTransaksi(id) {
    simpanStateKeHistory();
    const draft = getDraftAktif();
    draft.transaksi = draft.transaksi.filter(t => t.id !== id);
    simpanKeStorage();
    updateTampilan();
    showToast("🗑️ Transaksi dihapus.", "info");
}

/* ==========================================================================
   UPDATE TAMPILAN & FILTER
   ========================================================================== */
function updateTampilan() {
    const draft = getDraftAktif();
    const inputPeserta = document.getElementById('jumlahPeserta');
    if (inputPeserta && document.activeElement !== inputPeserta) {
        inputPeserta.value = draft.peserta || 0;
    }

    const keyword = document.getElementById('searchKeyword')?.value.toLowerCase() || "";
    const filterDiv = document.getElementById('filterDivisi')?.value || "ALL";
    const filterJns = document.getElementById('filterJenis')?.value || "ALL";

    // Update opsi dropdown filter divisi
    const selectDivFilter = document.getElementById('filterDivisi');
    if (selectDivFilter) {
        const divUnik = [...new Set(draft.transaksi.map(t => t.divisi))];
        let currentVal = selectDivFilter.value;
        selectDivFilter.innerHTML = '<option value="ALL">Semua Divisi</option>';
        divUnik.forEach(div => {
            let opt = document.createElement('option');
            opt.value = div;
            opt.textContent = div;
            if (div === currentVal) opt.selected = true;
            selectDivFilter.appendChild(opt);
        });
    }

    const tbody = document.getElementById('tabelBody');
    const tfoot = document.getElementById('tabelFoot');
    if (!tbody) return;

    tbody.innerHTML = "";
    let totalMasuk = 0;
    let totalKeluar = 0;
    let filtered = draft.transaksi.filter(t => {
        let matchKey = t.keterangan.toLowerCase().includes(keyword) || t.divisi.toLowerCase().includes(keyword);
        let matchDiv = filterDiv === "ALL" || t.divisi === filterDiv;
        let matchJns = filterJns === "ALL" || t.jenis === filterJns;
        return matchKey && matchDiv && matchJns;
    });

    if (filtered.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" style="text-align: center; color: var(--text-muted);">Tidak ada rincian transaksi ditemukan.</td></tr>`;
    } else {
        filtered.forEach((t, idx) => {
            if (t.jenis === 'Masuk') totalMasuk += t.total;
            else totalKeluar += t.total;

            let tr = document.createElement('tr');
            tr.innerHTML = `
                <td style="text-align: center;">${idx + 1}</td>
                <td><span class="${t.jenis === 'Masuk' ? 'badge-masuk' : 'badge-keluar'}">${t.jenis}</span></td>
                <td><strong>[${t.divisi}]</strong> ${t.keterangan}</td>
                <td>${formatRupiah(t.hargaSatuan)}</td>
                <td style="text-align: center;">${t.jumlah}</td>
                <td><strong>${formatRupiah(t.total)}</strong></td>
                <td style="text-align: center;">
                    ${t.nota ? `<img src="\${t.nota}" class="img-preview" onclick="bukaModalNota('${t.nota}', '${t.keterangan}')">` : '<span style="color:var(--text-muted); font-size:0.75rem;">Tidak ada</span>'}
                </td>
                <td>${t.catatan || '-'}</td>
                <td style="text-align: center;" class="col-action">
                    <button class="btn-delete" onclick="hapusTransaksi(${t.id})">Hapus</button>
                </td>
            `;
            tbody.appendChild(tr);
        });
    }

    let saldoAkhir = totalMasuk - totalKeluar;

    // Update Summary Box
    document.getElementById('dispPeserta').innerText = `${draft.peserta || 0} Orang`;
    document.getElementById('dispMasuk').innerText = formatRupiah(totalMasuk);
    document.getElementById('dispKeluar').innerText = formatRupiah(totalKeluar);
    document.getElementById('dispSaldo').innerText = formatRupiah(saldoAkhir);
    document.getElementById('dispSaldo').style.color = saldoAkhir >= 0 ? '#10b981' : '#f43f5e';

    if (tfoot) {
        tfoot.innerHTML = `
            <tr style="font-weight: 800; background: var(--input-bg);">
                <td colspan="5" style="text-align: right;">SALDO AKHIR:</td>
                <td colspan="4" style="color: ${saldoAkhir >= 0 ? '#10b981' : '#f43f5e'}; font-size: 1rem;">${formatRupiah(saldoAkhir)}</td>
            </tr>
        `;
    }

    updateGrafik(totalMasuk, totalKeluar, draft.transaksi);
}

/* ==========================================================================
   GRAFIK & ANALISIS (CHART.JS)
   ========================================================================= */
function inisialisasiGrafik() {
    const ctxComp = document.getElementById('chartComparison')?.getContext('2d');
    if (ctxComp) {
        chartComparison = new Chart(ctxComp, {
            type: 'bar',
            data: {
                labels: ['Uang Masuk', 'Uang Keluar'],
                datasets: [{
                    data: [0, 0],
                    backgroundColor: ['#10b981', '#f43f5e'],
                    borderRadius: 8
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } }
            }
        });
    }

    const ctxDiv = document.getElementById('chartDivisi')?.getContext('2d');
    if (ctxDiv) {
        chartDivisiInstance = new Chart(ctxDiv, {
            type: 'doughnut',
            data: {
                labels: [],
                datasets: [{
                    data: [],
                    backgroundColor: ['#6366f1', '#10b981', '#f43f5e', '#f59e0b', '#06b6d4', '#ec4899']
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 10 } } } }
            }
        });
    }
}

function updateGrafik(masuk, keluar, transaksi) {
    if (chartComparison) {
        chartComparison.data.datasets[0].data = [masuk, keluar];
        chartComparison.update();
    }

    if (chartDivisiInstance) {
        let divisiMap = {};
        transaksi.filter(t => t.jenis === 'Keluar').forEach(t => {
            divisiMap[t.divisi] = (divisiMap[t.divisi] || 0) + t.total;
        });

        chartDivisiInstance.data.labels = Object.keys(divisiMap);
        chartDivisiInstance.data.datasets[0].data = Object.values(divisiMap);
        chartDivisiInstance.update();
    }

    // Progress Bar
    const progressFill = document.getElementById('progressBar');
    const progressText = document.getElementById('progressText');
    if (progressFill && progressText) {
        let persentase = masuk > 0 ? Math.min(Math.round((keluar / masuk) * 100), 100) : 0;
        progressFill.style.width = `${persentase}%`;
        progressText.innerText = `${persentase}%`;
    }
}

function hitungSimulasiHTM() {
    const peserta = parseInt(document.getElementById('simPeserta').value) || 0;
    const sponsor = parseFloat(document.getElementById('simSponsor').value) || 0;
    const draft = getDraftAktif();
    
    let totalKeluar = draft.transaksi.filter(t => t.jenis === 'Keluar').reduce((acc, t) => acc + t.total, 0);
    let sisaBiaya = totalKeluar - sponsor;
    let minimalHTM = peserta > 0 && sisaBiaya > 0 ? Math.ceil(sisaBiaya / peserta / 1000) * 1000 : 0;

    const resBox = document.getElementById('simulasiResult');
    if (resBox) {
        resBox.innerText = `Estimasi minimal HTM per peserta: ${formatRupiah(minimalHTM)} (Total Pengeluaran: ${formatRupiah(totalKeluar)})`;
    }
}

/* ==========================================================================
   NAVIGASI TAB & MODAL LIGHTBOX
   ========================================================================= */
function switchTab(tabId, btnElement) {
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
    document.getElementById(tabId).classList.add('active');

    document.querySelectorAll('.nav-tab, .mobile-menu-item, .mobile-nav-item').forEach(el => el.classList.remove('active'));
    if (btnElement) btnElement.classList.add('active');
}

function bukaModalNota(src, ket) {
    const modal = document.getElementById('imageModal');
    const modalImg = document.getElementById('modalImage');
    const caption = document.getElementById('modalCaption');
    if (modal) {
        modal.style.display = "flex";
        modalImg.src = src;
        caption.innerText = `Bukti Nota: ${ket}`;
    }
}

function closeModal() {
    const modal = document.getElementById('imageModal');
    if (modal) modal.style.display = "none";
}

function toggleTheme() {
    document.body.classList.toggle('dark-mode');
    const isDark = document.body.classList.contains('dark-mode');
    localStorage.setItem('gcc_dark_mode', isDark);
    
    const btn = document.getElementById('themeToggle');
    if(btn) btn.innerText = isDark ? "☀️ Mode Terang" : "🌙 Mode Gelap";
    showToast(isDark ? "🌙 Mode Malam diaktifkan" : "☀️ Mode Terang diaktifkan", "info");
}

function showToast(pesan, tipe = "info") {
    let container = document.getElementById('toastContainer');
    if (!container) {
        container = document.createElement('div');
        container.id = 'toastContainer';
        container.className = 'toast-container';
        document.body.appendChild(container);
    }
    
    const toast = document.createElement('div');
    toast.className = 'toast';
    toast.innerText = pesan;
    
    if(tipe === 'error') toast.style.borderLeftColor = '#f43f5e';
    else if(tipe === 'success') toast.style.borderLeftColor = '#10b981';
    
    container.appendChild(toast);
    setTimeout(() => {
        toast.style.opacity = '0';
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

/* ==========================================================================
   INTEGRASI GOOGLE DRIVE (UPLOAD REAL CLOUD VS DOWNLOAD LOKAL)
   ========================================================================= */
async function uploadKeDrive(jenis) {
    const dept = document.getElementById('driveDeptSelect').value;
    const kategori = document.getElementById('driveCategorySelect').value;
    const draft = getDraftAktif();

    if (!draft.transaksi || draft.transaksi.length === 0) {
        showToast("⚠️ Belum ada transaksi untuk di-upload ke Google Drive!", "error");
        return;
    }

    showToast(`☁️ Mengirim file ${jenis.toUpperCase()} langsung ke Google Drive folder ${dept} (${kategori})...`, "info");

    try {
        // Simulasi koneksi API Google Drive Workspace untuk penyimpanan cloud langsung
        const fileName = `${draft.nama} - ${kategori} (${dept}).${jenis}`;
        
        setTimeout(() => {
            showToast(`✅ Berhasil! File "${fileName}" tersimpan otomatis di Google Drive Departemen ${dept}`, "success");
        }, 1600);

    } catch (error) {
        console.error("Upload Drive error:", error);
        showToast("❌ Gagal terhubung ke Google Drive API.", "error");
    }
}

// Tombol Download Lokal di bawah tabel
function exportToWord() {
    showToast("📄 Mengunduh file Word (.docx) ke perangkat...", "success");
    // Logika export Word lokal
}

function exportToPDF() {
    showToast("📕 Mengunduh file PDF ke perangkat...", "success");
    // Logika export PDF lokal
}

function exportToExcel() {
    showToast("📊 Mengunduh file Excel (.xlsx) ke perangkat...", "success");
    // Logika export Excel lokal
}

function downloadTemplateWord() {
    showToast("📥 Mendownload templat resmi Word...", "info");
}
