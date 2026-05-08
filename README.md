<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Cetak Nama & Alamat Undangan Pernikahan</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    
    <!-- Firebase SDK -->
    <script src="https://www.gstatic.com/firebasejs/9.22.0/firebase-app-compat.js"></script>
    <script src="https://www.gstatic.com/firebasejs/9.22.0/firebase-database-compat.js"></script>

    <style>
        :root { --primary: #8E6C4F; --secondary: #D4A574; --bg: #F9F6F2; --white: #ffffff; --text: #2C2C2C; --border: #E0D5C7; }
        * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        body { background-color: var(--bg); color: var(--text); padding: 20px; }
        .header { text-align: center; margin-bottom: 30px; padding: 20px; background: var(--white); border-radius: 12px; box-shadow: 0 4px 15px rgba(142, 108, 79, 0.1); border-bottom: 4px solid var(--primary); }
        .header h1 { color: var(--primary); font-size: 24px; }
        .container { max-width: 1200px; margin: 0 auto; display: grid; grid-template-columns: 1fr 2fr; gap: 20px; }
        @media (max-width: 900px) { .container { grid-template-columns: 1fr; } }
        .card { background: var(--white); padding: 20px; border-radius: 12px; box-shadow: 0 4px 15px rgba(142, 108, 79, 0.05); border: 1px solid var(--border); }
        .card h2 { font-size: 18px; color: var(--primary); margin-bottom: 15px; border-bottom: 2px solid var(--border); padding-bottom: 10px; }
        .form-group { margin-bottom: 15px; }
        .form-group label { display: block; margin-bottom: 5px; font-weight: 600; font-size: 14px; }
        .form-group input, .form-group select { width: 100%; padding: 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 14px; }
        .form-group input:focus, .form-group select:focus { outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px rgba(142, 108, 79, 0.1); }
        .btn { display: inline-flex; align-items: center; justify-content: center; gap: 8px; padding: 10px 20px; border: none; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; transition: all 0.2s ease; text-decoration: none; }
        .btn-primary { background-color: var(--primary); color: var(--white); }
        .btn-primary:hover { background-color: #735640; }
        .btn-secondary { background-color: var(--secondary); color: var(--text); }
        .btn-secondary:hover { background-color: #c49464; }
        .btn-danger { background-color: #dc3545; color: var(--white); }
        .btn-danger:hover { background-color: #c82333; }
        .btn-success { background-color: #28a745; color: var(--white); }
        .btn-success:hover { background-color: #218838; }
        .action-buttons { display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 20px; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { padding: 12px 10px; text-align: left; border-bottom: 1px solid var(--border); font-size: 13px; }
        th { background-color: var(--bg); font-weight: 700; color: var(--primary); }
        tr:hover { background-color: #fdfcfa; }
        .badge { padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: bold; white-space: nowrap; }
        .badge-bapak { background-color: #e3f2fd; color: #1565c0; }
        .badge-ibu { background-color: #fce4ec; color: #c62828; }
        .badge-sdr { background-color: #e8f5e9; color: #2e7d32; }
        .btn-delete { background: none; border: none; color: #dc3545; cursor: pointer; font-size: 16px; }

        .print-area { display: none; background: white; width: 210mm; }
        .print-grid { display: grid; grid-template-columns: repeat(2, 302px); grid-template-rows: repeat(5, 189px); gap: 0px; justify-content: center; padding: 10mm 0; page-break-after: always; }
        .name-box { width: 302px; height: 189px; border: 1px solid #000; border-radius: 4px; padding: 20px 30px; display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; font-family: 'Times New Roman', serif; page-break-inside: avoid; overflow: hidden; }
        .name-box .kepada { font-size: 14px; font-weight: bold; margin-bottom: 8px; color: #000; }
        .name-box .panggilan { font-size: 22px; font-weight: bold; margin-bottom: 4px; color: #000; }
        .name-box .nama-undangan { font-size: 30px; font-weight: bold; line-height: 1.1; color: #000; text-transform: capitalize; margin-bottom: 12px; padding: 0 5px; }
        .name-box .alamat-undangan { font-size: 22px; font-weight: bold; color: #000; line-height: 1.3; width: 100%; word-wrap: break-word; text-align: center; }

        @media print {
            body * { visibility: hidden; }
            .print-area, .print-area * { visibility: visible; }
            .print-area { display: block !important; position: absolute; left: 0; top: 0; width: 210mm; padding: 0; }
            .name-box { border: 1px solid #000 !important; }
            @page { size: A4; margin: 0; }
        }

        .stats { display: flex; gap: 15px; margin-bottom: 20px; }
        .stat-box { flex: 1; background: var(--bg); padding: 15px; border-radius: 8px; text-align: center; }
        .stat-box h3 { font-size: 20px; color: var(--primary); margin-bottom: 5px; }
        .stat-box p { font-size: 12px; color: #666; }
        .empty-state { text-align: center; padding: 40px; color: #888; }
        .alert { padding: 10px 15px; border-radius: 6px; margin-bottom: 15px; font-size: 14px; display: none; }
        .alert-success { background-color: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .alert-danger { background-color: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .col-alamat { max-width: 200px; }
        
        /* CSS untuk Error Diagnostic */
        #debug-screen { display: none; background: #000; color: #0f0; padding: 20px; border-radius: 10px; margin-bottom: 20px; font-family: monospace; font-size: 14px; border: 2px solid red; max-height: 200px; overflow-y: auto; }
    </style>
</head>
<body>

    <!-- LAYAR DIAGNOSTIC ERROR -->
    <div id="debug-screen">
        <b>System Log:</b><br>
        <span id="log-text">Memulai sistem...</span>
    </div>

    <div class="header">
        <h1>👍 Aplikasi Cetak Label Undangan Pernikahan</h1>
        <p style="color: #666; font-size: 14px;">Ukuran Label: 8cm x 5cm (Landscape) | Maks 10 nama per lembar A4/HVS</p>
    </div>

    <div class="container">
        <div class="card">
            <h2>➕ Tambah Nama Tamu</h2>
            <div id="alertBox" class="alert alert-success"></div>
            <div class="form-group"><label>Panggilan</label><select id="inputPanggilan"><option value="Bapak">Bapak</option><option value="Ibu">Ibu</option><option value="Sdr/i">Sdr/i</option></select></div>
            <div class="form-group"><label>Nama Lengkap Tamu</label><input type="text" id="inputNama" placeholder="Contoh: Ahmad Sudirman"></div>
            <div class="form-group"><label>Alamat</label><input type="text" id="inputAlamat" placeholder="Contoh: Jl. Merdeka No. 10, Jakarta"></div>
            <button class="btn btn-primary" style="width: 100%; margin-bottom: 20px;" onclick="tambahNama()">Simpan ke Daftar</button>
            <hr style="border: 0; border-top: 1px solid var(--border); margin: 20px 0;">
            <h2>📁 Import / Export Excel</h2>
            <p style="font-size: 12px; color: #666; margin-bottom: 15px;">Format Excel: <br>Kolom A = Panggilan<br>Kolom B = Nama<br>Kolom C = Alamat</p>
            <div class="form-group"><input type="file" id="fileImport" accept=".xlsx, .xls" onchange="importExcel(event)"></div>
            <div class="action-buttons" style="flex-direction: column;">
                <button class="btn btn-secondary" style="width: 100%;" onclick="exportExcel()">⬇️ Export ke Excel (.xlsx)</button>
                <button class="btn btn-danger" style="width: 100%;" onclick="hapusSemua()">🗑️ Hapus Semua Data</button>
            </div>
        </div>

        <div class="card">
            <div class="stats">
                <div class="stat-box"><h3 id="totalNama">0</h3><p>Total Nama</p></div>
                <div class="stat-box"><h3 id="totalLembar">0</h3><p>Kebutuhan Lembar A4</p></div>
            </div>
            <div class="action-buttons">
                <button class="btn btn-success" onclick="prepareAndPrint()">🖨️ Print / Cetak Undangan</button>
            </div>
            <table>
                <thead><tr><th>No</th><th>Panggilan</th><th>Nama Tamu</th><th>Alamat</th><th>Aksi</th></tr></thead>
                <tbody id="tableBody"></tbody>
            </table>
            <div id="emptyState" class="empty-state">Belum ada data nama.<br>Silahkan isi form atau import dari Excel.</div>
        </div>
    </div>

    <div class="print-area" id="printArea"></div>

    <script>
        // FUNGSI TAMPILAN LOG ERROR (AGAR KITA TAU ERRORNYA APA)
        function addLog(msg, isError = false) {
            const debugScreen = document.getElementById('debug-screen');
            const logText = document.getElementById('log-text');
            debugScreen.style.display = 'block'; // Tampilkan layar hitam
            logText.innerHTML += `<br>-> ${msg}`;
            if(isError) debugScreen.style.border = "2px solid red";
        }

        // Tangkap error JavaScript yang tersembunyi
        window.onerror = function(msg, url, lineNo, columnNo, error) {
            addLog("FATAL ERROR DI BARIS " + lineNo + ": " + msg, true);
            return false;
        };
const firebaseConfig = {
  apiKey: "AIzaSyD5CACtHvSFw2ouyDp-ryK2D1F7yig3lJQ",
  authDomain: "cetak-undangan-4f659.firebaseapp.com",
  projectId: "cetak-undangan-4f659",
  databaseURL: "https://console.firebase.google.com/u/0/project/cetak-undangan-c9405",
  storageBucket: "cetak-undangan-4f659.firebasestorage.app",
  messagingSenderId: "599828746287",
  appId: "1:599828746287:web:bdad8e86dc3ba834fa76da"
        };

        let db;
        let dataNama = [];

        try {
            addLog("Mengecek Firebase Config...");
            if (firebaseConfig.apiKey.includes("XXXXXXX") || firebaseConfig.databaseURL.includes("project-anda")) {
                addLog("ERROR: Kode Firebase Belum Diganti! Tombol dimatikan.", true);
                document.querySelectorAll('.btn').forEach(btn => btn.disabled = true);
            } else {
                addLog("Menghubungkan ke Firebase...");
                firebase.initializeApp(firebaseConfig);
                db = firebase.database();
                addLog("Firebase Terhubung!");

                // LOAD DATA REAL-TIME
                db.ref('tamu').on('value', (snapshot) => {
                    addLog("Mengambil data dari cloud...");
                    const data = snapshot.val();
                    dataNama = [];
                    if (data) {
                        Object.keys(data).forEach(key => {
                            const item = data[key];
                            item.id = key; 
                            dataNama.push(item);
                        });
                    }
                    renderTable();
                }, (error) => {
                    addLog("FIREBASE REJECTED: " + error.message + " (Cek Firebase Rules Anda!)", true);
                });
            }
        } catch (e) {
            addLog("TRY-CATCH ERROR: " + e.message, true);
        }

        // FUNGSI TAMBAH NAMA
        function tambahNama() {
            addLog("Tombol Simpan Ditekan!");
            const panggilan = document.getElementById('inputPanggilan').value;
            const nama = document.getElementById('inputNama').value.trim();
            const alamat = document.getElementById('inputAlamat').value.trim();

            if (nama === '') {
                tampilkanAlert('Nama tidak boleh kosong!', 'danger');
                return;
            }

            if (!db) {
                addLog("Gagal simpan: Database tidak terhubung.", true);
                return;
            }

            db.ref('tamu').push({
                panggilan: panggilan,
                nama: nama,
                alamat: alamat
            }).then(() => {
                document.getElementById('inputNama').value = '';
                document.getElementById('inputAlamat').value = '';
                tampilkanAlert('Nama berhasil ditambahkan!', 'success');
                addLog("Data berhasil disimpan ke Cloud.");
            }).catch((error) => {
                addLog("Gagal simpan ke Firebase: " + error.message, true);
                tampilkanAlert('Gagal menyimpan! Cek layar log error.', 'danger');
            });
        }

        // FUNGSI HAPUS
        function hapusNama(id) {
            if(confirm('Hapus nama ini?')) {
                db.ref('tamu/' + id).remove();
            }
        }

        function hapusSemua() {
            if (confirm('Apakah Anda yakin ingin menghapus SEMUA data nama?')) {
                db.ref('tamu').remove();
            }
        }

        // FUNGSI RENDER TABEL
        function renderTable() {
            const tbody = document.getElementById('tableBody');
            const emptyState = document.getElementById('emptyState');
            tbody.innerHTML = '';

            if (dataNama.length === 0) {
                emptyState.style.display = 'block';
            } else {
                emptyState.style.display = 'none';
                dataNama.forEach((item, index) => {
                    const tr = document.createElement('tr');
                    let badgeClass = 'badge-sdr';
                    if (item.panggilan === 'Bapak') badgeClass = 'badge-bapak';
                    if (item.panggilan === 'Ibu') badgeClass = 'badge-ibu';
                    const tampilAlamat = item.alamat ? item.alamat : '-';

                    tr.innerHTML = `
                        <td>${index + 1}</td>
                        <td><span class="badge ${badgeClass}">${item.panggilan}</span></td>
                        <td style="text-transform: capitalize;">${item.nama}</td>
                        <td class="col-alamat">${tampilAlamat}</td>
                        <td><button class="btn-delete" onclick="hapusNama('${item.id}')" title="Hapus">✖</button></td>
                    `;
                    tbody.appendChild(tr);
                });
            }
            document.getElementById('totalNama').innerText = dataNama.length;
            document.getElementById('totalLembar').innerText = Math.ceil(dataNama.length / 10);
        }

        function tampilkanAlert(pesan, tipe) {
            const alertBox = document.getElementById('alertBox');
            alertBox.innerText = pesan;
            alertBox.className = `alert alert-${tipe}`;
            alertBox.style.display = 'block';
            setTimeout(() => { alertBox.style.display = 'none'; }, 5000);
        }

        // FUNGSI IMPORT EXCEL
        function importExcel(event) {
            const file = event.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                let jumlahBaru = 0;
                jsonData.forEach(row => {
                    if (row[1] && typeof row[1] === 'string' && row[1].toLowerCase() !== 'nama') {
                        let panggilan = row[0] ? row[0].toString() : 'Bapak';
                        if (panggilan.toLowerCase().includes('ibu')) panggilan = 'Ibu';
                        else if (panggilan.toLowerCase().includes('sdr')) panggilan = 'Sdr/i';
                        else panggilan = 'Bapak';
                        const nama = row[1].toString().trim();
                        const alamat = row[2] ? row[2].toString().trim() : '';
                        if (nama !== '') {
                            db.ref('tamu').push({ panggilan: panggilan, nama: nama, alamat: alamat });
                            jumlahBaru++;
                        }
                    }
                });
                tampilkanAlert(`Berhasil import ${jumlahBaru} nama baru!`, 'success');
                event.target.value = '';
            };
            reader.readAsArrayBuffer(file);
        }

        function exportExcel() {
            if (dataNama.length === 0) { tampilkanAlert('Tidak ada data untuk di-export!', 'danger'); return; }
            const dataToExport = dataNama.map((item, index) => ({ 'No': index + 1, 'Panggilan': item.panggilan, 'Nama': item.nama, 'Alamat': item.alamat }));
            const worksheet = XLSX.utils.json_to_sheet(dataToExport);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "Nama Undangan");
            worksheet['!cols'] = [{ wch: 5 }, { wch: 10 }, { wch: 30 }, { wch: 40 }];
            XLSX.writeFile(workbook, "Daftar_Nama_Undangan.xlsx");
            tampilkanAlert('File Excel berhasil di-download!', 'success');
        }

        // FUNGSI PRINT
        function prepareAndPrint() {
            addLog("Tombol Cetak Ditekan!");
            if (dataNama.length === 0) { 
                tampilkanAlert('Tidak ada nama untuk dicetak!', 'danger'); 
                return; 
            }
            const printArea = document.getElementById('printArea');
            printArea.innerHTML = '';
            const itemsPerPage = 10;
            for (let i = 0; i < dataNama.length; i += itemsPerPage) {
                const chunk = dataNama.slice(i, i + itemsPerPage);
                const printGrid = document.createElement('div');
                printGrid.className = 'print-grid';
                chunk.forEach(item => {
                    const box = document.createElement('div');
                    box.className = 'name-box';
                    const teksAlamat = item.alamat ? `di ${item.alamat}` : '';
                    box.innerHTML = `
                        <div class="kepada">Kepada Yth :</div>
                        <div class="panggilan">${item.panggilan}</div>
                        <div class="nama-undangan">${item.nama}</div>
                        <div class="alamat-undangan">${teksAlamat}</div>
                    `;
                    printGrid.appendChild(box);
                });
                printArea.appendChild(printGrid);
            }
            addLog("Membuka dialog print...");
            setTimeout(() => { window.print(); }, 500);
        }

        document.getElementById("inputAlamat").addEventListener("keypress", function(event) {
            if (event.key === "Enter") { event.preventDefault(); tambahNama(); }
        });
    </script>
</body>
</html>
