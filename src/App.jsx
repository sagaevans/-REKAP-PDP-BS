import React, { useState, useRef, useEffect } from 'react';
import { Upload, ChevronDown, Download, AlertCircle, X, FileSpreadsheet, CheckCircle, Info } from 'lucide-react';

// --- HELPER FUNCTIONS ---
const parseDateString = (val) => {
  if (val === undefined || val === null || val === '') return null;
  if (typeof val === 'number') {
    const parsed = window.XLSX.SSF.parse_date_code(val);
    return { year: parsed.y, month: parsed.m, day: parsed.d };
  }
  if (typeof val === 'string') {
    const parts = val.split(/[-/]/);
    if (parts.length >= 3) {
      if (parts[0].length === 4) { 
        return { year: parseInt(parts[0], 10), month: parseInt(parts[1], 10), day: parseInt(parts[2], 10) };
      } else { 
        return { year: parseInt(parts[2], 10), month: parseInt(parts[1], 10), day: parseInt(parts[0], 10) };
      }
    }
  }
  return null;
};

const getMonthLabel = (dateObj, selectedMonthNum) => {
  if (!dateObj || !dateObj.month) return null;
  if (dateObj.month < selectedMonthNum) return 'lalu';
  if (dateObj.month === selectedMonthNum) return 'ini';
  return 'depan';
};

const normalizeKeys = (row) => {
  const normalized = {};
  for (let key in row) {
    const normKey = key.trim().toLowerCase().replace(/\s+/g, ' ');
    normalized[normKey] = row[key];
  }
  normalized._originalRow = row; 
  return normalized;
};

const formatNumber = (num) => {
  if (!num || num === 0) return '-';
  return new Intl.NumberFormat('id-ID').format(num);
};

// --- MAIN COMPONENT ---
export default function App() {
  const [file, setFile] = useState(null);
  const [month, setMonth] = useState('');
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [scriptsLoaded, setScriptsLoaded] = useState(false);
  
  const [results, setResults] = useState(null);
  const [activeTab, setActiveTab] = useState('kiri'); 
  
  const [modal, setModal] = useState({ isOpen: false, title: '', message: '', type: 'info' });
  const fileInputRef = useRef(null);

  useEffect(() => {
    const scripts = [
      'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
      'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js',
      'https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js'
    ];
    
    let loadedCount = 0;
    scripts.forEach(src => {
      if (document.querySelector(`script[src="${src}"]`)) {
        loadedCount++;
        if (loadedCount === scripts.length) setScriptsLoaded(true);
        return;
      }
      const script = document.createElement('script');
      script.src = src;
      script.onload = () => {
        loadedCount++;
        if (loadedCount === scripts.length) setScriptsLoaded(true);
      };
      document.head.appendChild(script);
    });
  }, []);

  const months = [
    'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
  ];

  const handleDragOver = (e) => { e.preventDefault(); setIsDragging(true); };
  const handleDragLeave = (e) => { e.preventDefault(); setIsDragging(false); };
  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFile(e.dataTransfer.files[0]);
    }
  };

  const handleFile = (uploadedFile) => {
    if (uploadedFile.name.endsWith('.xlsx')) {
      setFile(uploadedFile);
      setResults(null);
    } else {
      setModal({ isOpen: true, title: 'File Tidak Valid', message: 'Mohon upload file dengan ekstensi .xlsx', type: 'error' });
    }
  };

  const processData = async () => {
    if (!file || !month) return;
    setIsProcessing(true);

    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = window.XLSX.read(arrayBuffer, { type: 'array', cellDates: false });
      
      const sheetNamesLower = workbook.SheetNames.map(s => s.trim().toLowerCase());
      const sdIndex = sheetNamesLower.indexOf('zrsd015');
      const fiIndex = sheetNamesLower.indexOf('zrfi056');

      if (sdIndex === -1 || fiIndex === -1) {
        setModal({ isOpen: true, title: 'Sheet Hilang', message: 'File harus berisi sheet "zrsd015" dan "zrfi056".', type: 'error' });
        setIsProcessing(false);
        return;
      }

      const sdName = workbook.SheetNames[sdIndex];
      const fiName = workbook.SheetNames[fiIndex];

      const rawSd = window.XLSX.utils.sheet_to_json(workbook.Sheets[sdName], { defval: "", raw: true });
      const rawFi = window.XLSX.utils.sheet_to_json(workbook.Sheets[fiName], { defval: "", raw: true });

      const numBulan = parseInt(month, 10);
      const groupingMap = new Map();
      
      const cekSjList = [];
      const cekSoList = [];
      const statusSjMap = new Map();

      rawSd.forEach((rowRaw, i) => {
        const row = normalizeKeys(rowRaw);
        
        const materialStr = row['material']?.toString().trim();
        if (!materialStr) return; 

        const soStr = row['main so nbr.']?.toString().trim() || "";
        const key = `${soStr}_${materialStr}`;
        
        if (!groupingMap.has(key)) {
          groupingMap.set(key, { 
            so: soStr, 
            material: materialStr, 
            desc: row['sales description [trs]'] || '', 
            pgi_ini: 0, pgi_lalu: 0, sj_ini: 0, sj_lalu: 0, inv_ini: 0, 
            hasSoMuda: false, 
            labels: new Set() 
          });
        }
        
        const dataGroup = groupingMap.get(key);
        const transDateObj = parseDateString(row['trans date']);
        const sjDateObj = parseDateString(row['spm/sj lastdate']);
        
        const pgiQty = parseFloat(row['pgi. auom']) || 0;
        const sjQty = parseFloat(row['spm auom']) || 0;

        const transLabel = getMonthLabel(transDateObj, numBulan);
        const sjLabel = getMonthLabel(sjDateObj, numBulan);

        if (transLabel) dataGroup.labels.add(transLabel);

        if (transLabel === 'ini') dataGroup.pgi_ini += pgiQty;
        if (transLabel === 'lalu') dataGroup.pgi_lalu += pgiQty;
        if (sjLabel === 'ini') dataGroup.sj_ini += sjQty;
        if (sjLabel === 'lalu') dataGroup.sj_lalu += sjQty;

        const valSj = row['sj']?.toString().trim();
        const valSpm = row['spm']?.toString().trim();
        const valBill = row['billing nbr.']?.toString().trim();
        let statusSj = '';

        if (valSj && valSpm) {
          statusSj = 'Benar';
        } else if ((!valSj || !valSpm) && valBill) {
          statusSj = 'Belum tempel SJ';
          cekSjList.push({
            baris: i + 2,
            transDate: row['trans date'],
            so: soStr,
            bill: valBill,
            material: materialStr,
            desc: row['sales description [trs]'],
            pgiQty,
            billQty: parseFloat(row['bill qty']) || 0,
            sjDate: row['spm/sj lastdate'],
            sj: valSj,
            spm: valSpm,
            sjQty
          });
        }
        if (statusSj) statusSjMap.set(i + 2, statusSj);

        if (transDateObj && sjDateObj) {
          if (sjDateObj.year < transDateObj.year || (sjDateObj.year === transDateObj.year && sjDateObj.month < transDateObj.month)) {
            dataGroup.hasSoMuda = true;
            cekSoList.push({
              baris: i + 2,
              transDate: row['trans date'],
              so: soStr,
              bill: valBill,
              material: materialStr,
              desc: row['sales description [trs]'],
              pgiQty,
              billQty: parseFloat(row['bill qty']) || 0,
              sjDate: row['spm/sj lastdate'],
              sj: valSj,
              spm: valSpm,
              sjQty
            });
          }
        }
      });

      rawFi.forEach((rowRaw) => {
        const row = normalizeKeys(rowRaw);
        const materialStr = row['material']?.toString().trim();
        if (!materialStr) return;

        const soStr = row['so']?.toString().trim() || "";
        const key = `${soStr}_${materialStr}`;
        
        if (groupingMap.has(key)) {
          const tglObj = parseDateString(row['tgl nota']);
          const invLabel = getMonthLabel(tglObj, numBulan);
          if (invLabel === 'ini') {
            groupingMap.get(key).inv_ini += parseFloat(row['quantity']) || 0;
          }
        }
      });

      const tabelKiri = [];
      const tabelKanan = [];
      
      const allGroups = Array.from(groupingMap.values());
      allGroups.sort((a, b) => {
        const numA = parseInt(a.so, 10) || 0;
        const numB = parseInt(b.so, 10) || 0;
        return numA - numB;
      });

      allGroups.forEach(val => {
        if (val.labels.has('ini') || val.hasSoMuda) {
          let e = val.pgi_ini;
          let f = val.sj_ini;
          let g = val.inv_ini;
          let h = val.inv_ini > 0 ? (e - f) : f;

          if (val.hasSoMuda) {
             const r = val.inv_ini > 0 ? ((val.pgi_lalu + val.pgi_ini) - val.sj_ini) : 0;
             g = val.inv_ini - r;
             h = val.sj_ini - g;
          }

          tabelKiri.push({
            so: val.so, material: val.material, desc: val.desc, e, f, g, h
          });
        }

        if (val.labels.has('lalu') || val.hasSoMuda) {
          let n = val.hasSoMuda ? (val.pgi_lalu + val.pgi_ini) : val.pgi_lalu;
          let o = val.sj_lalu;
          let p = val.sj_ini;
          let q = val.inv_ini;
          let r = q > 0 ? (n - p) : 0;
          let s = q > 0 ? (n - o) : 0;
          let t = (p > 0 && q === 0) ? p : 0;

          if (val.hasSoMuda) {
            s = 0; 
            t = 0; 
          }

          tabelKanan.push({
             so: val.so, material: val.material, desc: val.desc, n, o, p, q, r, s, t
          });
        }
      });

      let alertMsg = "Lengkap, tidak ada kesalahan.";
      let alertType = 'success';
      if (cekSjList.length > 0 && cekSoList.length > 0) {
        alertMsg = `Ada SJ belum tempel dan ada ${cekSoList.length} SO yang tanggalnya lebih muda dari SJ-nya.`;
        alertType = 'warning';
      }
      else if (cekSjList.length > 0) {
        alertMsg = "Ada SJ belum tempel.";
        alertType = 'warning';
      }
      else if (cekSoList.length > 0) {
        alertMsg = `Ada ${cekSoList.length} SO yang tanggalnya lebih muda dari SJ-nya.`;
        alertType = 'warning';
      }

      setResults({
        tabelKiri,
        tabelKanan,
        cekSjList,
        cekSoList,
        statusSjMap, 
        originalBuffer: arrayBuffer,
        sdName,
        fiName
      });
      setActiveTab('kiri');
      
      setModal({ isOpen: true, title: 'Proses Selesai', message: alertMsg, type: alertType });

    } catch (err) {
      console.error(err);
      setModal({ isOpen: true, title: 'Terjadi Kesalahan', message: 'Gagal memproses file. Pastikan format sesuai.', type: 'error' });
    } finally {
      setIsProcessing(false);
    }
  };

  const handleDownloadExcel = async () => {
    if (!results || !file || !month) return;
    setIsProcessing(true);

    try {
      const monthName = months[parseInt(month, 10) - 1];
      const currentYear = new Date().getFullYear();

      const workbook = new window.ExcelJS.Workbook();
      await workbook.xlsx.load(results.originalBuffer);

      const sheetsToKeep = [results.sdName, results.fiName];
      workbook.worksheets.forEach(ws => {
        if (!sheetsToKeep.includes(ws.name)) {
          workbook.removeWorksheet(ws.id);
        }
      });

      const wsSd = workbook.getWorksheet(results.sdName);
      if (wsSd) {
        let maxCol = wsSd.columnCount;
        let headRow = wsSd.getRow(1);
        headRow.getCell(maxCol + 1).value = 'Status SJ';
        
        wsSd.eachRow((row, rowNumber) => {
          if (rowNumber > 1) {
            const status = results.statusSjMap.get(rowNumber);
            if (status) {
              row.getCell(maxCol + 1).value = status;
            }
          }
        });
      }

      const wsRekap = workbook.addWorksheet('REKAP');
      
      const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9DDCF6' } };
      const sepFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF404040' } };
      const borderThin = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
      };
      const numFormat = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';

      wsRekap.columns = [
        { key: 'A', width: 8 }, { key: 'B', width: 16.7 }, { key: 'C', width: 21 }, { key: 'D', width: 33 },
        { key: 'E', width: 19 }, { key: 'F', width: 13 }, { key: 'G', width: 24 }, { key: 'H', width: 18 },
        { key: 'I', width: 5 }, 
        { key: 'J', width: 14 }, { key: 'K', width: 16 }, { key: 'L', width: 21 }, { key: 'M', width: 33 },
        { key: 'N', width: 12 }, { key: 'O', width: 13 }, { key: 'P', width: 11 }, { key: 'Q', width: 10 },
        { key: 'R', width: 18 }, { key: 'S', width: 13 }, { key: 'T', width: 30 }
      ];

      const headerRow2 = [
        '', '', '', `BULAN INI - ${monthName.toUpperCase()}`, '', '', 'LANGSUNG NOTA', 'PDP MASUK',
        '', 
        '', 'BULAN LALU', 'PDP KELUAR', '', 'SO PENUH', 'SJ', 'SJ', 'INVOICE', 'PDP KELUAR', 'LANGSUNG NOTA', 'PDP MASUK'
      ];
      wsRekap.getRow(2).values = headerRow2;

      const headerRow3 = [
        'Periode', `SO ${monthName.toUpperCase()}`, 'Material', 'Material Description', 'SO Penuh (Bln ini)', 'SJ (Bln ini)', `Invoice -SO SJ Bulan ini (LANGSUNG NOTA ${monthName.toUpperCase()})`, 'SO SJ BULAN INI BELUM INVOICE (PDP Masuk)',
        '', 
        'SO Bulan sebelumnya , SJ bulan sebelumnya', 'SO BULAN LALU', 'Material', 'Material Description', 'bulan lalu', 'bulan lalu', 'bulan ini', 'bulan ini', `SO Bulan lalu SJ bulan lalu INVOICE bulan ini (PDP KELUAR ${monthName.toUpperCase()})`, `SO Bulan lalu SJ bulan ini, invoice bulan ini (LANGSUNG INVOICE ${monthName.toUpperCase()})`, 'SO Bulan lalu SJ bulan ini (BELUM INVOICE)'
      ];
      wsRekap.getRow(3).values = headerRow3;

      for (let i = 1; i <= 20; i++) {
        let cell2 = wsRekap.getCell(2, i);
        let cell3 = wsRekap.getCell(3, i);
        let cell4 = wsRekap.getCell(4, i);

        if (i === 9) { 
          cell2.fill = sepFill;
          cell3.fill = sepFill;
          cell4.fill = sepFill;
          continue;
        }

        cell2.fill = headerFill;
        cell2.border = borderThin;
        cell2.font = { bold: true };
        cell2.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

        cell3.fill = headerFill; cell3.border = borderThin; cell3.font = { bold: true }; cell3.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell4.fill = headerFill; cell4.border = borderThin;
        
        wsRekap.mergeCells(3, i, 4, i);
      }

      let rowIdx = 5;
      const maxRows = Math.max(results.tabelKiri.length, results.tabelKanan.length);
      
      let totE = 0, totF = 0, totG = 0, totH = 0;
      let totN = 0, totO = 0, totP = 0, totQ = 0, totR = 0, totS = 0, totT = 0;

      for (let i = 0; i < maxRows; i++) {
        const row = wsRekap.getRow(rowIdx);
        const kiri = results.tabelKiri[i];
        const kanan = results.tabelKanan[i];

        if (kiri) {
          row.getCell(1).value = monthName;
          row.getCell(2).value = kiri.so;
          row.getCell(3).value = kiri.material;
          row.getCell(4).value = kiri.desc;
          row.getCell(5).value = kiri.e; totE += kiri.e;
          row.getCell(6).value = kiri.f; totF += kiri.f;
          row.getCell(7).value = kiri.g; totG += kiri.g;
          row.getCell(8).value = kiri.h; totH += kiri.h;
        }

        row.getCell(9).fill = sepFill;

        if (kanan) {
          row.getCell(10).value = 'bulan lalu';
          row.getCell(11).value = kanan.so;
          row.getCell(12).value = kanan.material;
          row.getCell(13).value = kanan.desc;
          row.getCell(14).value = kanan.n; totN += kanan.n;
          row.getCell(15).value = kanan.o; totO += kanan.o;
          row.getCell(16).value = kanan.p; totP += kanan.p;
          row.getCell(17).value = kanan.q; totQ += kanan.q;
          row.getCell(18).value = kanan.r; totR += kanan.r;
          row.getCell(19).value = kanan.s; totS += kanan.s;
          row.getCell(20).value = kanan.t; totT += kanan.t;
        }

        [5,6,7,8, 14,15,16,17,18,19,20].forEach(colIdx => {
           row.getCell(colIdx).numFmt = numFormat;
        });

        rowIdx++;
      }

      if (results.tabelKiri.length > 0) {
        const trKiri = wsRekap.getRow(rowIdx);
        trKiri.getCell(4).value = 'TOTAL';
        trKiri.getCell(4).font = { bold: true };
        [5,6,7,8].forEach((c, idx) => {
          let cell = trKiri.getCell(c);
          cell.value = [totE, totF, totG, totH][idx];
          cell.font = { bold: true };
          cell.numFmt = numFormat;
        });
      }

      if (results.tabelKanan.length > 0) {
        const trKanan = wsRekap.getRow(rowIdx);
        trKanan.getCell(13).value = 'TOTAL';
        trKanan.getCell(13).font = { bold: true };
        [14,15,16,17,18,19,20].forEach((c, idx) => {
          let cell = trKanan.getCell(c);
          cell.value = [totN, totO, totP, totQ, totR, totS, totT][idx];
          cell.font = { bold: true };
          cell.numFmt = numFormat;
        });
      }

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      window.saveAs(blob, `REKAP_${monthName}_${currentYear}.xlsx`);

    } catch (err) {
      console.error(err);
      setModal({ isOpen: true, title: 'Gagal Download', message: 'Gagal membuat file Excel.', type: 'error' });
    } finally {
      setIsProcessing(false);
    }
  };

  const handleDownloadTemplate = async (e) => {
    e.preventDefault();
    const wb = new window.ExcelJS.Workbook();
    
    const wsSd = wb.addWorksheet('zrsd015');
    wsSd.addRow(['Trans Date', 'Contr Main', 'Main SO Nbr.', 'Billing Nbr.', 'Material', 'Sales Description [Trs]', 'PGI. AUOM', 'Bill Qty', 'SPM/SJ LastDate', 'SJ', 'SPM', 'SPM AUOM']);
    
    const wsFi = wb.addWorksheet('zrfi056');
    wsFi.addRow(['Nota', 'Tgl Nota', 'SO', 'Material', 'Material Description', 'Quantity', 'Sat']);

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    window.saveAs(blob, `Template_Rekap_PDP.xlsx`);
  };

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8 font-sans flex flex-col items-center">
      
      <div className="w-full max-w-5xl bg-white rounded-xl shadow-sm border border-slate-200 p-6 md:p-8 mb-6">
        <h1 className="text-2xl md:text-3xl font-bold text-slate-800 tracking-tight mb-2">
          Rekap Persediaan Dalam Perjalanan (PDP)
        </h1>
        <p className="text-slate-500 text-sm mb-8">
          Upload file Excel berisi sheet <span className="font-mono bg-slate-100 px-1 py-0.5 rounded">zrsd015</span> dan <span className="font-mono bg-slate-100 px-1 py-0.5 rounded">zrfi056</span>, pilih bulan, lalu klik Proses.
        </p>

        <div 
          className={`relative border-2 border-dashed rounded-xl p-8 flex flex-col items-center justify-center text-center cursor-pointer transition-all duration-200 mb-6 ${
            isDragging ? 'border-blue-500 bg-blue-50' : 'border-slate-300 hover:border-slate-400 bg-slate-50 hover:bg-slate-100/50'
          }`}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current.click()}
        >
          <input type="file" ref={fileInputRef} onChange={(e) => { if(e.target.files.length) handleFile(e.target.files[0]) }} accept=".xlsx" className="hidden" />
          
          <div className="bg-white p-3 rounded-full shadow-sm mb-4 border border-slate-100">
            {file ? <FileSpreadsheet className="w-8 h-8 text-green-600" /> : <Upload className="w-8 h-8 text-blue-600" />}
          </div>
          
          <h3 className="text-base font-semibold text-slate-800 mb-1">
            {file ? file.name : 'Klik atau seret file .xlsx ke sini'}
          </h3>
          <p className="text-sm text-slate-500">
            {file ? `Siap diproses. (Klik untuk ganti)` : 'Pastikan format data sesuai dengan template'}
          </p>
        </div>

        <div className="flex flex-col sm:flex-row gap-4 items-end">
          <div className="flex-1 w-full">
            <label className="block text-sm font-semibold text-slate-700 mb-2">Bulan laporan</label>
            <div className="relative">
              <select
                className="w-full appearance-none bg-white border border-slate-300 text-slate-700 text-sm rounded-lg px-4 py-2.5 pr-10 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 cursor-pointer"
                value={month}
                onChange={(e) => setMonth(e.target.value)}
              >
                <option value="" disabled>Pilih bulan</option>
                {months.map((m, i) => (
                  <option key={i} value={i + 1}>{m}</option>
                ))}
              </select>
              <ChevronDown className="w-4 h-4 text-slate-400 absolute right-3 top-1/2 -translate-y-1/2 pointer-events-none" />
            </div>
          </div>
          <button
            onClick={processData}
            disabled={!file || !month || isProcessing || !scriptsLoaded}
            className={`w-full sm:w-auto px-8 py-2.5 rounded-lg font-medium text-white transition-all shadow-sm flex items-center justify-center gap-2
              ${(!file || !month || isProcessing || !scriptsLoaded) ? 'bg-slate-300 cursor-not-allowed' : 'bg-slate-800 hover:bg-slate-700 active:scale-95'}
            `}
          >
            {isProcessing ? (
              <><span className="animate-spin rounded-full h-4 w-4 border-2 border-white border-t-transparent"></span> Memproses...</>
            ) : !scriptsLoaded ? 'Memuat Library...' : 'Proses Data'}
          </button>
        </div>
      </div>

      {results && (
        <div className="w-full max-w-5xl bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden mb-6 fade-in">
          
          <div className="bg-slate-800 px-6 py-4 flex flex-col sm:flex-row items-center justify-between gap-4">
            <h2 className="text-white font-semibold flex items-center gap-2">
              <CheckCircle className="w-5 h-5 text-green-400" /> Hasil Rekapitulasi Selesai
            </h2>
            <button
              onClick={handleDownloadExcel}
              disabled={isProcessing}
              className="bg-white text-slate-800 hover:bg-blue-50 px-4 py-2 rounded-lg text-sm font-bold flex items-center gap-2 transition-colors shadow-sm"
            >
              <Download className="w-4 h-4" /> Download Excel
            </button>
          </div>

          <div className="flex overflow-x-auto border-b border-slate-200 bg-slate-50/50 scrollbar-hide">
            {[
              { id: 'kiri', label: 'Bulan Ini' },
              { id: 'kanan', label: 'Bulan Lalu / PDP Keluar' },
              { id: 'sj', label: `Cek SJ (${results.cekSjList.length})` },
              { id: 'so', label: `Cek SO (${results.cekSoList.length})` },
            ].map(tab => (
              <button
                key={tab.id}
                onClick={() => setActiveTab(tab.id)}
                className={`px-6 py-3 text-sm font-medium whitespace-nowrap transition-colors border-b-2
                  ${activeTab === tab.id 
                    ? 'border-blue-600 text-blue-700 bg-white' 
                    : 'border-transparent text-slate-600 hover:text-slate-800 hover:bg-slate-100'}
                `}
              >
                {tab.label}
              </button>
            ))}
          </div>

          <div className="bg-white">
            {activeTab === 'kiri' && (
              results.tabelKiri.length === 0 ? <div className="p-4 text-slate-500">Tidak ada SO bulan ini untuk periode ini.</div> :
              <div className="overflow-x-auto max-h-[60vh]">
                <table className="w-full text-sm text-left whitespace-nowrap">
                  <thead className="text-xs text-slate-700 bg-blue-50 sticky top-0 border-b">
                    <tr>
                      <th className="px-4 py-3 border-r">SO</th>
                      <th className="px-4 py-3 border-r">Material</th>
                      <th className="px-4 py-3 border-r">Deskripsi</th>
                      <th className="px-4 py-3 text-right">SO Penuh (Bln ini)</th>
                      <th className="px-4 py-3 text-right">SJ (Bln ini)</th>
                      <th className="px-4 py-3 text-right">Langsung Nota</th>
                      <th className="px-4 py-3 text-right">PDP Masuk</th>
                    </tr>
                  </thead>
                  <tbody>
                    {results.tabelKiri.map((row, i) => (
                      <tr key={i} className="border-b hover:bg-slate-50">
                        <td className="px-4 py-2 border-r">{row.so}</td>
                        <td className="px-4 py-2 border-r">{row.material}</td>
                        <td className="px-4 py-2 border-r truncate max-w-[200px]" title={row.desc}>{row.desc}</td>
                        <td className="px-4 py-2 text-right">{formatNumber(row.e)}</td>
                        <td className="px-4 py-2 text-right">{formatNumber(row.f)}</td>
                        <td className="px-4 py-2 text-right">{formatNumber(row.g)}</td>
                        <td className="px-4 py-2 text-right">{formatNumber(row.h)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
            
            {activeTab === 'kanan' && (
              results.tabelKanan.length === 0 ? <div className="p-4 text-slate-500">Tidak ada SO bulan lalu untuk periode ini.</div> :
              <div className="overflow-x-auto max-h-[60vh]">
                <table className="w-full text-sm text-left whitespace-nowrap">
                  <thead className="text-xs text-slate-700 bg-amber-50 sticky top-0 border-b">
                    <tr>
                      <th className="px-4 py-3 border-r">SO Lalu</th>
                      <th className="px-4 py-3 border-r">Material</th>
                      <th className="px-4 py-3 border-r">Deskripsi</th>
                      <th className="px-4 py-3 text-right border-l-2 border-l-slate-300">SO (Lalu)</th>
                      <th className="px-4 py-3 text-right">SJ (Lalu)</th>
                      <th className="px-4 py-3 text-right border-l-2 border-l-slate-300">SJ (Ini)</th>
                      <th className="px-4 py-3 text-right">Inv (Ini)</th>
                      <th className="px-4 py-3 text-right border-l-2 border-l-slate-300 bg-amber-100">PDP Keluar</th>
                      <th className="px-4 py-3 text-right bg-amber-100">Langsung Invoice</th>
                      <th className="px-4 py-3 text-right bg-amber-100">Belum Invoice</th>
                    </tr>
                  </thead>
                  <tbody>
                    {results.tabelKanan.map((row, i) => (
                      <tr key={i} className="border-b hover:bg-slate-50">
                        <td className="px-4 py-2 border-r">{row.so}</td>
                        <td className="px-4 py-2 border-r">{row.material}</td>
                        <td className="px-4 py-2 border-r truncate max-w-[150px]">{row.desc}</td>
                        <td className="px-4 py-2 text-right border-l-2 border-l-slate-300">{formatNumber(row.n)}</td>
                        <td className="px-4 py-2 text-right">{formatNumber(row.o)}</td>
                        <td className="px-4 py-2 text-right border-l-2 border-l-slate-300">{formatNumber(row.p)}</td>
                        <td className="px-4 py-2 text-right">{formatNumber(row.q)}</td>
                        <td className="px-4 py-2 text-right border-l-2 border-l-slate-300 font-medium">{formatNumber(row.r)}</td>
                        <td className="px-4 py-2 text-right font-medium">{formatNumber(row.s)}</td>
                        <td className="px-4 py-2 text-right font-medium">{formatNumber(row.t)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
            
            {activeTab === 'sj' && (
              results.cekSjList.length === 0 ? <div className="p-4 text-slate-500">Tidak ada SJ yang belum ditempel.</div> :
              <div className="overflow-x-auto max-h-[60vh]">
                <table className="w-full text-sm text-left whitespace-nowrap">
                  <thead className="text-xs text-slate-700 bg-slate-100 sticky top-0 border-b">
                    <tr>
                      <th className="px-4 py-3">Baris</th>
                      <th className="px-4 py-3">Trans Date</th>
                      <th className="px-4 py-3">SO</th>
                      <th className="px-4 py-3">Billing</th>
                      <th className="px-4 py-3">Material</th>
                      <th className="px-4 py-3">Desc</th>
                      <th className="px-4 py-3 text-right">PGI Qty</th>
                      <th className="px-4 py-3 text-right">Bill Qty</th>
                      <th className="px-4 py-3">SJ Date</th>
                      <th className="px-4 py-3">SJ</th>
                      <th className="px-4 py-3">SPM</th>
                      <th className="px-4 py-3 text-right">SJ Qty</th>
                    </tr>
                  </thead>
                  <tbody>
                    {results.cekSjList.map((row, i) => (
                      <tr key={i} className="border-b hover:bg-slate-50">
                        <td className="px-4 py-2">{row.baris}</td>
                        <td className="px-4 py-2">{row.transDate}</td>
                        <td className="px-4 py-2">{row.so}</td>
                        <td className="px-4 py-2">{row.bill}</td>
                        <td className="px-4 py-2">{row.material}</td>
                        <td className="px-4 py-2 truncate max-w-[150px]">{row.desc}</td>
                        <td className="px-4 py-2 text-right">{row.pgiQty}</td>
                        <td className="px-4 py-2 text-right">{row.billQty}</td>
                        <td className="px-4 py-2">{row.sjDate}</td>
                        <td className="px-4 py-2">{row.sj}</td>
                        <td className="px-4 py-2">{row.spm}</td>
                        <td className="px-4 py-2 text-right">{row.sjQty}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
            
            {activeTab === 'so' && (
              results.cekSoList.length === 0 ? <div className="p-4 text-slate-500">Tidak ada SO yang tanggalnya lebih muda dari SJ-nya.</div> :
              <div className="overflow-x-auto max-h-[60vh]">
                <table className="w-full text-sm text-left whitespace-nowrap">
                  <thead className="text-xs text-slate-700 bg-slate-100 sticky top-0 border-b">
                    <tr>
                      <th className="px-4 py-3">Baris</th>
                      <th className="px-4 py-3">Trans Date</th>
                      <th className="px-4 py-3">SO</th>
                      <th className="px-4 py-3">Billing</th>
                      <th className="px-4 py-3">Material</th>
                      <th className="px-4 py-3">Desc</th>
                      <th className="px-4 py-3 text-right">PGI Qty</th>
                      <th className="px-4 py-3 text-right">Bill Qty</th>
                      <th className="px-4 py-3">SJ Date</th>
                      <th className="px-4 py-3">SJ</th>
                      <th className="px-4 py-3">SPM</th>
                      <th className="px-4 py-3 text-right">SJ Qty</th>
                    </tr>
                  </thead>
                  <tbody>
                    {results.cekSoList.map((row, i) => (
                      <tr key={i} className="border-b hover:bg-slate-50">
                        <td className="px-4 py-2">{row.baris}</td>
                        <td className="px-4 py-2">{row.transDate}</td>
                        <td className="px-4 py-2">{row.so}</td>
                        <td className="px-4 py-2">{row.bill}</td>
                        <td className="px-4 py-2">{row.material}</td>
                        <td className="px-4 py-2 truncate max-w-[150px]">{row.desc}</td>
                        <td className="px-4 py-2 text-right">{row.pgiQty}</td>
                        <td className="px-4 py-2 text-right">{row.billQty}</td>
                        <td className="px-4 py-2">{row.sjDate}</td>
                        <td className="px-4 py-2">{row.sj}</td>
                        <td className="px-4 py-2">{row.spm}</td>
                        <td className="px-4 py-2 text-right">{row.sjQty}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      )}

      <div className="w-full max-w-5xl mt-auto pt-8 pb-4 text-center border-t border-slate-200">
        <button onClick={handleDownloadTemplate} className="text-sm font-medium text-blue-600 hover:text-blue-700 hover:underline mb-2 transition-colors">
          Download Template Excel
        </button>
        <p className="text-xs text-slate-400 mt-1">
          — Contoh format file yang harus diupload (hanya header, tanpa data)
        </p>
        <p className="text-xs text-slate-400 mt-6">
          Made by Nas · Powered by Lovable (AI Replica)
        </p>
      </div>

      {modal.isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm p-4 animate-in fade-in duration-200">
          <div className="bg-white rounded-xl shadow-xl max-w-md w-full overflow-hidden scale-in">
            <div className={`p-4 border-b flex items-center gap-3 ${
              modal.type === 'error' ? 'bg-red-50 text-red-700 border-red-100' :
              modal.type === 'warning' ? 'bg-amber-50 text-amber-700 border-amber-100' :
              'bg-blue-50 text-blue-700 border-blue-100'
            }`}>
              {modal.type === 'error' ? <AlertCircle className="w-5 h-5" /> :
               modal.type === 'warning' ? <AlertCircle className="w-5 h-5" /> :
               <Info className="w-5 h-5" />}
              <h3 className="font-semibold">{modal.title}</h3>
              <button onClick={() => setModal({ ...modal, isOpen: false })} className="ml-auto text-slate-400 hover:text-slate-600">
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="p-6 text-slate-600 text-sm leading-relaxed">
              {modal.message}
            </div>
            <div className="p-4 border-t bg-slate-50 flex justify-end">
              <button 
                onClick={() => setModal({ ...modal, isOpen: false })}
                className="px-4 py-2 bg-slate-800 text-white rounded-lg text-sm font-medium hover:bg-slate-700 transition-colors"
              >
                Tutup
              </button>
            </div>
          </div>
        </div>
      )}

      <style dangerouslySetInnerHTML={{__html: `
        .scrollbar-hide::-webkit-scrollbar { display: none; }
        .scrollbar-hide { -ms-overflow-style: none; scrollbar-width: none; }
        @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
        @keyframes scaleIn { from { opacity: 0; transform: scale(0.95); } to { opacity: 1; transform: scale(1); } }
        .fade-in { animation: fadeIn 0.3s ease-out; }
        .scale-in { animation: scaleIn 0.2s ease-out; }
      `}} />
    </div>
  );
}
