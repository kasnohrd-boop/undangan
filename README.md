<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Cetak Nama & Alamat Undangan Pernikahan</title>
    <!-- Library untuk Import/Export Excel -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        :root {
            --primary: #8E6C4F;
            --secondary: #D4A574;
            --bg: #F9F6F2;
            --white: #ffffff;
            --text: #2C2C2C;
            --border: #E0D5C7;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }

        body {
            background-color: var(--bg);
            color: var(--text);
            padding: 20px;
        }

        .header {
            text-align: center;
            margin-bottom: 30px;
            padding: 20px;
            background: var(--white);
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(142, 108, 79, 0.1);
            border-bottom: 4px solid var(--primary);
        }

        .header h1 {
            color: var(--primary);
            font-size: 24px;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            display: grid;
            grid-template-columns: 1fr 2fr;
            gap: 20px;
        }

        @media (max-width: 900px) {
            .container {
                grid-template-columns: 1fr;
            }
        }

        .card {
            background: var(--white);
            padding: 20px;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(142, 108, 79, 0.05);
            border: 1px solid var(--border);
        }

        .card h2 {
            font-size: 18px;
            color: var(--primary);
            margin-bottom: 15px;
            border-bottom: 2px solid var(--border);
            padding-bottom: 10px;
        }

        .form-group {
            margin-bottom: 15px;
        }

        .form-group label {
            display: block;
            margin-bottom: 5px;
            font-weight: 600;
            font-size: 14px;
        }

        .form-group input, .form-group select {
            width: 100%;
            padding: 10px;
            border: 1px solid var(--border);
            border-radius: 6px;
            font-size: 14px;
        }

        .form-group input:focus, .form-group select:focus {
            outline: none;
            border-color: var(--primary);
            box-shadow: 0 0 0 3px rgba(142, 108, 79, 0.1);
        }

        .btn {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            padding: 10px 20px;
            border: none;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s ease;
            text-decoration: none;
        }

        .btn-primary { background-color: var(--primary); color: var(--white); }
        .btn-primary:hover { background-color: #735640; }
        .btn-secondary { background-color: var(--secondary); color: var(--text); }
        .btn-secondary:hover { background-color: #c49464; }
        .btn-danger { background-color: #dc3545; color: var(--white); }
        .btn-danger:hover { background-color: #c82333; }
        .btn-success { background-color: #28a745; color: var(--white); }
        .btn-success:hover { background-color: #218838; }

        .action-buttons {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            margin-bottom: 20px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
        }

        th, td {
            padding: 12px 10px;
            text-align: left;
            border-bottom: 1px solid var(--border);
            font-size: 13px;
        }

        th {
            background-color: var(--bg);
            font-weight: 700;
            color: var(--primary);
        }

        tr:hover {
            background-color: #fdfcfa;
        }

        .badge {
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: bold;
            white-space: nowrap;
        }

        .badge-bapak { background-color: #e3f2fd; color: #1565c0; }
        .badge-ibu { background-color: #fce4ec; color: #c62828; }
        .badge-sdr { background-color: #e8f5e9; color: #2e7d32; }

        .btn-delete {
            background: none;
            border: none;
            color: #dc3545;
            cursor: pointer;
            font-size: 16px;
        }

        /* STYLE UNTUK PRINT AREA (5cm x 7cm) */
        .print-area {
            display: none;
            background: white;
            width: 210mm; 
            padding: 10mm;
        }

        .print-grid {
            display: grid;
            grid-template-columns: repeat(3, 189px); /* 3 kolom x 5cm */
            grid-template-rows: repeat(6, 264.6px); /* 6 baris x 7cm */
            gap: 0px;
            justify-content: center;
        }

        .name-box {
            width: 189px;
            height: 264.6px;
            
            /* GANTI: MENAMBAHKAN BINGKAI OVAL / ELIPS */
            border: 3px double var(--primary); /* Garis ganda elegan */
            border-radius: 50%; /* Membuat bentuk persegi menjadi oval/elips */
            
            /* Padding diperbesar agar teks tidak nabrak garis oval di pinggir */
            padding: 45px 25px 35px 25px; 
            
            display: flex;
            flex-direction: column;
            justify-content: center; /* Teks rata tengah vertikal */
            align-items: center;
            text-align: center;
            font-family: 'Times New Roman', serif;
            page-break-inside: avoid;
            overflow: hidden;
        }

        .name-box .kepada {
            font-size: 16px; /* Diperbesar dari 14px */
            font-weight: bold; /* Ditambahkan bold */
            margin-bottom: 12px;
            color: #000;
        }

        .name-box .panggilan {
            font-size: 20px; /* Diperbesar dari 14px */
            font-weight: bold;
            margin-bottom: 6px;
            color: #000;
        }

        .name-box .nama-undangan {
            font-size: 24px; /* Diperbesar dari 18px */
            font-weight: bold;
            line-height: 1.2;
            color: #000;
            text-transform: capitalize;
            margin-bottom: 15px;
            padding: 0 5px;
        }

        .name-box .alamat-undangan {
            font-size: 14px; /* Diperbesar dari 12px */
            color: #000;
            line-height: 1.4;
            width: 100%;
            word-wrap: break-word;
            text-align: center;
        }

        /* PRINT SPECIFIC STYLES */
        @media print {
            body * {
                visibility: hidden;
            }
            .print-area, .print-area * {
                visibility: visible;
            }
            .print-area {
                display: block !important;
                position: absolute;
                left: 0;
                top: 0;
                width: 210mm;
                padding: 5mm; 
            }
            .name-box {
                /* Pastikan garis oval tetap tercetak jelas berwarna hitam */
                border: 3px double #000 !important; 
                border-radius: 50% !important;
            }
            @page {
                size: A4;
                margin: 0;
            }
        }

        .stats {
            display: flex;
            gap: 15px;
            margin-bottom: 20px;
        }

        .stat-box {
            flex: 1;
            background: var(--bg);
            padding: 15px;
            border-radius: 8px;
            text-align: center;
        }

        .stat-box h3 {
            font-size: 20px;
            color: var(--primary);
            margin-bottom: 5px;
        }

        .stat-box p {
            font-size: 12px;
            color: #666;
        }

        .empty-state {
            text-align: center;
            padding: 40px;
            color: #888;
        }

        .alert {
            padding: 10px 15px;
            border-radius: 6px;
            margin-bottom: 15px;
            font-size: 14px;
            display: none;
        }
        .alert-success { background-color: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .alert-danger { background-color: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        
        .col-alamat {
            max-width: 200px;
        }
    </style>
</head>
<body>

    <div class="header">
        <h1>👍 Aplikasi Cetak Label Undangan Pernikahan</h1>
        <p style="color: #666; font-size: 14px;">Ukuran Label: 5cm x 7cm (Maks 18 nama per lembar A4/HVS) | Bingkai Oval</p>
    </div>

    <div class="container">
        <!-- PANEL KIRI: FORM INPUT -->
        <div class="card">
            <h2>➕ Tambah Nama Tamu</h2>
            
            <div id="alertBox" class="alert alert-success"></div>

            <div class="form-group">
                <label>Panggilan</label>
                <select id="inputPanggilan">
                    <option value="Bapak">Bapak</option>
                    <option value="Ibu">Ibu</option>
                    <option value="Sdr/i">Sdr/i</option>
                </select>
            </div>

            <div class="form-group">
                <label>Nama Lengkap Tamu</label>
                <input type="text" id="inputNama" placeholder="Contoh: Ahmad Sudirman">
            </div>

            <div class="form-group">
                <label>Alamat</label>
                <input type="text" id="inputAlamat" placeholder="Contoh: Jl. Merdeka No. 10, Jakarta">
            </div>

            <button class="btn btn-primary" style="width: 100%; margin-bottom: 20px;" onclick="tambahNama()">
                Simpan ke Daftar
            </button>

            <hr style="border: 0; border-top: 1px solid var(--border); margin: 20px 0;">

            <h2>📁 Import / Export Excel</h2>
            <p style="font-size: 12px; color: #666; margin-bottom: 15px;">
                Format Excel: <br>Kolom A = Panggilan<br>Kolom B = Nama<br>Kolom C = Alamat
            </p>

            <div class="form-group">
                <input type="file" id="fileImport" accept=".xlsx, .xls" onchange="importExcel(event)">
            </div>

            <div class="action-buttons" style="flex-direction: column;">
                <button class="btn btn-secondary" style="width: 100%;" onclick="exportExcel()">
                    ⬇️ Export ke Excel (.xlsx)
                </button>
                <button class="btn btn-danger" style="width: 100%;" onclick="hapusSemua()">
                    🗑️ Hapus Semua Data
                </button>
            </div>

            <hr style="border: 0; border-top: 1px solid var(--border); margin: 20px 0;">
            <div style="font-size: 12px; color: #888; background: #f9f9f9; padding: 10px; border-radius: 6px;">
                <strong>Catatan Akses Multi-User:</strong><br>
                Karena ini berbasis web lokal, untuk diisi banyak orang sekaligus seperti Google Form, unggah file HTML ini ke hosting gratis (Netlify/Vercel/GitHub) lalu sebar linknya.
            </div>
        </div>

        <!-- PANEL KANAN: TABEL DATA & PREVIEW -->
        <div class="card">
            <div class="stats">
                <div class="stat-box">
                    <h3 id="totalNama">0</h3>
                    <p>Total Nama</p>
                </div>
                <div class="stat-box">
                    <h3 id="totalLembar">0</h3>
                    <p>Kebutuhan Lembar A4</p>
                </div>
            </div>

            <div class="action-buttons">
                <button class="btn btn-success" onclick="prepareAndPrint()">
                    🖨️ Print / Cetak Undangan
                </button>
            </div>

            <table>
                <thead>
                    <tr>
                        <th>No</th>
                        <th>Panggilan</th>
                        <th>Nama Tamu</th>
                        <th>Alamat</th>
                        <th>Aksi</th>
                    </tr>
                </thead>
                <tbody id="tableBody">
                </tbody>
            </table>
            
            <div id="emptyState" class="empty-state">
                Belum ada data nama.<br>Silahkan isi form atau import dari Excel.
            </div>
        </div>
    </div>

    <!-- AREA TERSEMBUNYI UNTUK PRINTING -->
    <div class="print-area" id="printArea">
        <div class="print-grid" id="printGrid">
        </div>
    </div>

    <script>
        const STORAGE_KEY = 'data_undangan_pernikahan_v2';
        let dataNama = [];

        // LOAD DATA DARI LOCAL STORAGE
        window.onload = function() {
            const dataDariStorage = localStorage.getItem(STORAGE_KEY);
            if (dataDariStorage) {
                dataNama = JSON.parse(dataDariStorage);
            }
            renderTable();
        };

        // FUNGSI TAMBAH NAMA MANUAL
        function tambahNama() {
            const panggilan = document.getElementById('inputPanggilan').value;
            const nama = document.getElementById('inputNama').value.trim();
            const alamat = document.getElementById('inputAlamat').value.trim();

            if (nama === '') {
                tampilkanAlert('Nama tidak boleh kosong!', 'danger');
                return;
            }

            dataNama.push({ 
                panggilan: panggilan, 
                nama: nama, 
                alamat: alamat 
            });
            
            simpanKeStorage();
            renderTable();

            // Reset form
            document.getElementById('inputNama').value = '';
            document.getElementById('inputAlamat').value = '';
            tampilkanAlert('Nama berhasil ditambahkan!', 'success');
        }

        // FUNGSI HAPUS 1 NAMA
        function hapusNama(index) {
            dataNama.splice(index, 1);
            simpanKeStorage();
            renderTable();
        }

        // FUNGSI HAPUS SEMUA DATA
        function hapusSemua() {
            if (confirm('Apakah Anda yakin ingin menghapus SEMUA data nama?')) {
                dataNama = [];
                simpanKeStorage();
                renderTable();
            }
        }

        // FUNGSI SIMPAN KE LOCAL STORAGE
        function simpanKeStorage() {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(dataNama));
        }

        // FUNGSI RENDER TABEL HTML
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
                        <td><button class="btn-delete" onclick="hapusNama(${index})" title="Hapus">✖</button></td>
                    `;
                    tbody.appendChild(tr);
                });
            }

            // Update Stats
            document.getElementById('totalNama').innerText = dataNama.length;
            document.getElementById('totalLembar').innerText = Math.ceil(dataNama.length / 18);
        }

        // FUNGSI TAMPILKAN ALERT
        function tampilkanAlert(pesan, tipe) {
            const alertBox = document.getElementById('alertBox');
            alertBox.innerText = pesan;
            alertBox.className = `alert alert-${tipe}`;
            alertBox.style.display = 'block';
            
            setTimeout(() => {
                alertBox.style.display = 'none';
            }, 3000);
        }

        // FUNGSI IMPORT DARI EXCEL
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
                            dataNama.push({ panggilan: panggilan, nama: nama, alamat: alamat });
                            jumlahBaru++;
                        }
                    }
                });

                simpanKeStorage();
                renderTable();
                tampilkanAlert(`Berhasil import ${jumlahBaru} nama baru!`, 'success');
                event.target.value = '';
            };
            reader.readAsArrayBuffer(file);
        }

        // FUNGSI EXPORT KE EXCEL
        function exportExcel() {
            if (dataNama.length === 0) {
                tampilkanAlert('Tidak ada data untuk di-export!', 'danger');
                return;
            }

            const dataToExport = dataNama.map((item, index) => ({
                'No': index + 1,
                'Panggilan': item.panggilan,
                'Nama': item.nama,
                'Alamat': item.alamat
            }));

            const worksheet = XLSX.utils.json_to_sheet(dataToExport);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "Nama Undangan");

            worksheet['!cols'] = [
                { wch: 5 },  // No
                { wch: 10 }, // Panggilan
                { wch: 30 }, // Nama
                { wch: 40 }  // Alamat
            ];

            XLSX.writeFile(workbook, "Daftar_Nama_Undangan.xlsx");
            tampilkanAlert('File Excel berhasil di-download!', 'success');
        }

        // FUNGSI SIAPKAN DAN PRINT
        function prepareAndPrint() {
            if (dataNama.length === 0) {
                tampilkanAlert('Tidak ada nama untuk dicetak!', 'danger');
                return;
            }

            const printGrid = document.getElementById('printGrid');
            printGrid.innerHTML = '';

            dataNama.forEach(item => {
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

            setTimeout(() => {
                window.print();
            }, 500);
        }

        // ALLOW ENTER KEY TO SUBMIT FORM
        document.getElementById("inputAlamat").addEventListener("keypress", function(event) {
            if (event.key === "Enter") {
                event.preventDefault();
                tambahNama();
            }
        });
    </script>
</body>
</html>
