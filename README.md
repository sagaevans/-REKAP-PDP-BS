# PROMPT SISTEM & DOKUMENTASI REKAP PDP
> **UNTUK AI:** Jika pengguna meminta Anda untuk me-recreate, update, atau mendebug aplikasi ini, gunakan seluruh dokumen ini sebagai referensi utama Anda terkait "Context", "Business Logic", dan "Architecture".

## 1. Arsitektur & Tech Stack
Aplikasi ini adalah Single-Page Application (React JSX) yang berjalan sepenuhnya di sisi klien (browser).
- **Framework:** React 18+ (Hooks: `useState`, `useRef`, `useEffect`)
- **Styling:** Tailwind CSS (Utility classes)
- **Icons:** `lucide-react`
- **Data Processor (Read):** `xlsx` (SheetJS) - Dimuat via CDN (useEffect)
- **Data Formatter (Write):** `exceljs` - Dimuat via CDN (useEffect)
- **File Saver:** `FileSaver.js` - Dimuat via CDN (useEffect)

## 2. Struktur Input Data (Excel)
Aplikasi menerima 1 file `.xlsx` yang HARUS memiliki 2 sheet spesifik:
1. **`zrsd015`** (Data PGI & SJ)
   - Kolom penting: `Trans Date`, `Main SO Nbr.`, `Billing Nbr.`, `Material`, `Sales Description [Trs]`, `PGI. AUOM`, `Bill Qty`, `SPM/SJ LastDate`, `SJ`, `SPM`, `SPM AUOM`.
2. **`zrfi056`** (Data Invoice)
   - Kolom penting: `Tgl Nota`, `SO`, `Material`, `Quantity`.

## 3. Logika Bisnis Inti (CRITICAL)

### A. Pengelompokan Data (Grouping)
Semua baris direduksi (di-group) menggunakan *Map Key* unik: `${SO}_${Material}`.
Huruf pada key harus di-normalize (trim, lowercase).

### B. Tanggal dan Kategori Bulan
Input bulan dari UI adalah bulan laporan (1-12).
Setiap `Trans Date`, `SPM/SJ LastDate`, dan `Tgl Nota` harus dikonversi dari serial date Excel atau String DD/MM/YYYY ke objek date.
Berdasarkan bulan pada tanggal tersebut vs bulan laporan, data di-tag:
- `lalu` (Bulan Date < Bulan Laporan)
- `ini` (Bulan Date == Bulan Laporan)

### C. Kalkulasi Normal
- **`pgi_ini` / `pgi_lalu`**: Akumulasi dari `PGI. AUOM` (zrsd015) berdasarkan bulan `Trans Date`.
- **`sj_ini` / `sj_lalu`**: Akumulasi dari `SPM AUOM` (zrsd015) berdasarkan bulan `SPM/SJ LastDate`.
- **`inv_ini`**: Akumulasi dari `Quantity` (zrfi056) berdasarkan bulan `Tgl Nota`.

### D. KASUS KHUSUS (SO LEBIH MUDA DARI SJ / `hasSoMuda`)
Ini adalah Edge Case paling krusial!
**Deteksi:** Terjadi jika Tahun/Bulan `Trans Date` (SO) > Tahun/Bulan `SPM/SJ LastDate` (SJ). (Misal: SO dibentuk bulan Oktober, tapi mengambil stok/SJ dari bulan Agustus).
**Penanganan Logika (Jika hasSoMuda == true):**
1. **Tabel Kiri (Bulan Ini):**
   - Kolom E (SO Penuh) = `pgi_lalu` + `pgi_ini`
   - Kolom F (SJ) = `sj_ini`
   - Kolom G (Langsung Nota) = `inv_ini` - (jika ada sisa). *(Biasanya `inv_ini` lari ke sini secara utuh)*.
2. **Tabel Kanan (Bulan Lalu):**
   - **MENCEGAH DOUBLE COUNT:** Karena perhitungan sudah ditarik ke Tabel Kiri, maka Kolom S (Langsung Invoice) dan T (Belum Invoice) **HARUS DI-SET MENJADI 0**. Jangan lakukan perhitungan standar untuk S dan T jika record ini di-flag sebagai `hasSoMuda`.

### E. Error Handling / Audit Dokumen
1. **"Belum Tempel SJ"**: Jika `Billing Nbr.` ADA isinya, TETAPI `SJ` atau `SPM` KOSONG.
2. Flag `hasSoMuda` juga dicatat dalam Array `cekSoList` untuk ditampilkan di UI.

## 4. Ekspor Excel (ExcelJS Formatting)
- Output Excel bernama `REKAP_Bulan_Tahun.xlsx`.
- **Sheet `zrsd015`**: Ditambahkan 1 kolom paling kanan bernama "Status SJ" ("Benar" atau "Belum tempel SJ"). Sheet `zrfi056` dibiarkan as-is.
- **Sheet `REKAP` (Dibuat baru)**:
  - Baris 1: Kosong/dihapus agar UI mepet atas.
  - Baris 2: Grouping Header (Biru Muda `FF9DDCF6`, Border tipis, Bold, Tengah).
  - Baris 3 & 4: Sub Header (Di-merge/gabung vertikal tiap kolom, gaya sama seperti Baris 2).
  - Kolom 9 (Kolom I): Bertindak sebagai separator gelap. Warna background `FF404040`, TANPA TEXT, TANPA BORDER. Di-apply dari header sampai baris data.
  - Number Format: Excel harus mengenali angka desimal ribuan dengan format akuntansi: `_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-`.

> **END OF PROMPT**

Link Demo : https://share.gemini.google/UQXQASwaF2Mx
