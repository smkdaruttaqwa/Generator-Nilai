import React, { useState, useRef, DragEvent, useEffect } from 'react';
import * as XLSX from 'xlsx-js-style';
import { Upload, FileSpreadsheet, Download, RefreshCw, AlertCircle, X, CheckCircle2, LayoutTemplate, Database, AlertTriangle, Moon, Sun, ChevronRight, Layers } from 'lucide-react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

function normalizeNisn(val: string) {
  let clean = String(val || '').replace(/^'/, '').trim();
  if (clean.length > 0 && clean.length < 10 && /^\d+$/.test(clean)) {
    clean = clean.padStart(10, '0');
  }
  return clean;
}

interface MissingStudent {
  no: number;
  nisn: string;
  name: string;
  kelas: string;
}

interface SemesterResult {
  wb: XLSX.WorkBook;
  studentCount: number;
  missingStudents: MissingStudent[];
  totalSourceCount: number;
}

export default function App() {
  const [sourceFiles, setSourceFiles] = useState<File[]>([]);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<{ generated: Record<string, SemesterResult>, semesters: string[] } | null>(null);

  const [isDragSource, setIsDragSource] = useState(false);
  const [isDragTemplate, setIsDragTemplate] = useState(false);

  const sourceInputRef = useRef<HTMLInputElement>(null);
  const templateInputRef = useRef<HTMLInputElement>(null);

  const [isDarkMode, setIsDarkMode] = useState(() => {
    if (typeof window !== 'undefined') {
      return localStorage.getItem('theme') === 'dark' || 
        (!('theme' in localStorage) && window.matchMedia('(prefers-color-scheme: dark)').matches);
    }
    return false;
  });

  useEffect(() => {
    if (isDarkMode) {
      document.documentElement.classList.add('dark');
      localStorage.setItem('theme', 'dark');
    } else {
      document.documentElement.classList.remove('dark');
      localStorage.setItem('theme', 'light');
    }
  }, [isDarkMode]);

  const handleSourceChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setSourceFiles(prev => [...prev, ...Array.from(e.target.files!)]);
      setError(null);
      setResult(null);
    }
  };

  const handleTemplateChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setTemplateFile(e.target.files[0]);
      setError(null);
      setResult(null);
    }
  };

  const processData = async () => {
    if (sourceFiles.length === 0) {
      setError("Harap unggah Sumber Data (Master) terlebih dahulu.");
      return;
    }
    if (!templateFile) {
      setError("Harap unggah Template Output terlebih dahulu.");
      return;
    }
    setIsProcessing(true);
    setError(null);

    try {
      // --- LANGKAH 1: Ekstrak nilai dari file master ---
      let sourceGrades: Record<string, Record<string, Record<string, any>>> = {}; 
      let sourceInfo: Record<string, {name: string, kelas: string}> = {};

      for (const file of sourceFiles) {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        
        for (const sheetName of workbook.SheetNames) {
          const sheet = workbook.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];

          let dataStartIndex = -1;
          let semCol = -1;

          for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            if (!row) continue;
            for (let j = 0; j < row.length; j++) {
              const cell = String(row[j] || '').trim().toUpperCase();
              if (cell.startsWith('SM ') || cell === 'PSAJ' || cell === 'NS') {
                dataStartIndex = i;
                semCol = j;
                break;
              }
            }
            if (dataStartIndex !== -1) break;
          }

          if (dataStartIndex !== -1 && semCol !== -1) {
            let subjectRowIdx = Math.max(0, dataStartIndex - 1);
            let currentSubjects: string[] = [];
            if (rows[subjectRowIdx]) {
                for (let j = semCol + 1; j < rows[subjectRowIdx].length; j++) {
                    const subName = rows[subjectRowIdx][j];
                    if (subName && String(subName).trim() !== '') {
                        currentSubjects.push(String(subName).trim());
                    } else if (j < rows[subjectRowIdx].length) {
                        currentSubjects.push(`Mata Pelajaran ${j - semCol}`);
                    }
                }
            }

            let currentStudentNISN = '';

            for (let i = dataStartIndex; i < rows.length; i++) {
              const row = rows[i];
              if (!row || row.length === 0) continue;

              const semVal = String(row[semCol] || '').trim().toUpperCase();
              const isSemRow = semVal.startsWith('SM ') || semVal === 'PSAJ' || semVal === 'NS';
              
              // Ambil NISN dari template yang mengurutkan: URUT, NIS, NISN, NAMA
              const nisnCol = semCol - 4;
              const nisnValRaw = nisnCol >= 0 ? String(row[nisnCol] || '').trim() : '';
              
              // Hapus awalan petik satu jika ada (baik nama, maupun NISN) dan pad jadi 10 digit jika kurang
              const cleanNisn = normalizeNisn(nisnValRaw);
              
              const kelasCol = semCol - 1;
              const kelasValRaw = kelasCol >= 0 ? String(row[kelasCol] || '').trim() : '';
              const cleanKelas = kelasValRaw.replace(/^'/, '').trim();
              
              const namaCol = semCol - 3;
              const namaValRaw = namaCol >= 0 ? String(row[namaCol] || '').trim() : '';
              const cleanNama = namaValRaw.replace(/^'/, '').trim();

              if (cleanNisn !== '' && cleanNisn.toUpperCase() !== 'UNDEFINED' && cleanNisn.toUpperCase() !== 'NULL' && cleanNisn.toUpperCase() !== 'NISN') {
                currentStudentNISN = cleanNisn;
                if (!sourceInfo[cleanNisn] || sourceInfo[cleanNisn].name === '-') {
                  sourceInfo[cleanNisn] = { name: cleanNama || '-', kelas: cleanKelas || '-' };
                }
              }

              if (isSemRow && currentStudentNISN) {
                if (!sourceGrades[semVal]) sourceGrades[semVal] = {};
                if (!sourceGrades[semVal][currentStudentNISN]) sourceGrades[semVal][currentStudentNISN] = {};
                
                for (let k = 0; k < currentSubjects.length; k++) {
                  const gradeCol = semCol + 1 + k;
                  const gradeVal = row[gradeCol];
                  if (gradeVal !== undefined && gradeVal !== '') {
                      // Bersihkan value grade jika text
                      let cleanedVal = gradeVal;
                      if (typeof cleanedVal === 'string') {
                          cleanedVal = cleanedVal.replace(/^'/, '').trim();
                      }
                      sourceGrades[semVal][currentStudentNISN][currentSubjects[k]] = cleanedVal;
                  }
                }
              }
            }
          }
        }
      }

      if (Object.keys(sourceGrades).length === 0) {
        throw new Error("Tidak menemukan baris yang berawalan 'SM 1', 'SM 2', dst di file Master Sumber Data.");
      }

      // --- LANGKAH 2: Baca Template Output dan siapkan pemetaan kolom ---
      const templateData = await templateFile.arrayBuffer();
      // Gunakan cellStyles untuk mempertahankan format border dari XLSX-js-style
      const rawTemplateWb = XLSX.read(templateData, { type: 'array', cellDates: true, cellStyles: true });
      const rawTemplateSheet = rawTemplateWb.Sheets[rawTemplateWb.SheetNames[0]];
      const rawTemplateAoa = XLSX.utils.sheet_to_json(rawTemplateSheet, { header: 1 }) as any[][];

      let nisnColIdx = -1;
      let subjectCols: Record<string, number> = {};
      let headerRowIdx = -1;
      let semesterCellLocation = { r: -1, c: -1 };
      let maxTemplateColIdx = 0;

      for (let r = 0; r < Math.min(rawTemplateAoa.length, 30); r++) {
         const row = rawTemplateAoa[r];
         if (!row) continue;
         for (let c = 0; c < row.length; c++) {
           if (c > maxTemplateColIdx) maxTemplateColIdx = c;
             
           const cellText = String(row[c] || '').trim().toUpperCase();
           
           if (cellText === 'NISN') {
              nisnColIdx = c;
              headerRowIdx = r; // Asumsikan Header di mana NISN ada, itu patokannya
           }
           if (cellText === 'PA') subjectCols['Pendidikan Agama'] = c;
           if (cellText === 'KWN') subjectCols['Pendidikan Pancasila'] = c;
           if (cellText === 'IND') subjectCols['Bahasa Indonesia'] = c;
           if (cellText === 'MAT') subjectCols['Matematika'] = c;
           if (cellText === 'IPA') subjectCols['IPA'] = c;
           if (cellText === 'IPS') subjectCols['IPS'] = c;
           if (cellText === 'ING') subjectCols['Bahasa Inggris'] = c;
           
           if (cellText === 'SEMESTER:' || cellText === 'SEMESTER' || cellText === 'SMT') {
              semesterCellLocation = { r, c: c + 1 };
           }
         }
      }

      if (nisnColIdx === -1) {
         throw new Error("Tidak dapat menemukan header 'NISN' pada file Template Output.");
      }
      if (Object.keys(subjectCols).length === 0) {
         throw new Error("Tidak dapat menemukan header mata pelajaran (PA, KWN, IND, MAT, IPA, IPS, ING) pada Template Output.");
      }

      // Tentukan batas akhir kolom untuk pembuatan Border (misal kolom "Bhs Inggris")
      const endBorderCol = Math.max(nisnColIdx, ...Object.values(subjectCols));

      const generatedResults: Record<string, SemesterResult> = {};
      const semestersDesc = Object.keys(sourceGrades).sort();
      const dataStartRowIdx = headerRowIdx + 1;

      // --- LANGKAH 3: Cetak File Masing-Masing Semester ---
      for (const sem of semestersDesc) {
         // Kloning dengan membaca format array buffer lagi supaya fresh
         const templateWb = XLSX.read(templateData, { type: 'array', cellDates: true, cellStyles: true });
         const sheetName = templateWb.SheetNames[0];
         const sheet = templateWb.Sheets[sheetName];
         
         // Update info Semester di template atas jika ditemukan
         if (semesterCellLocation.r !== -1) {
            const semNum = sem.toUpperCase().replace('SM ', '').trim();
            const cellAddr = XLSX.utils.encode_cell({ r: semesterCellLocation.r, c: semesterCellLocation.c });
            if (sheet[cellAddr]) {
               sheet[cellAddr].v = semNum || sem;
               sheet[cellAddr].t = 's';
            } else {
               sheet[cellAddr] = { v: semNum || sem, t: 's' };
            }
         }

         let matchCount = 0;
         const processedNisns = new Set<string>();
         const rangeStr = sheet['!ref'];
         
         if (rangeStr) {
           const range = XLSX.utils.decode_range(rangeStr);
           for (let r = dataStartRowIdx; r <= range.e.r; r++) {
              const nisnCellAddr = XLSX.utils.encode_cell({ r, c: nisnColIdx });
              const nisnCell = sheet[nisnCellAddr];
              
              if (!nisnCell || !nisnCell.v) continue; // Skip jika tidak ada data di cell NISN
              
              const cleanNisn = normalizeNisn(String(nisnCell.v));
              if (cleanNisn.toUpperCase() === 'NISN' || cleanNisn === '') continue;
              
              const studentGrades = sourceGrades[sem][cleanNisn];
              
              if (studentGrades) {
                 matchCount++;
                 processedNisns.add(cleanNisn);
                 
                 for (const [sourceSubj, colIdx] of Object.entries(subjectCols)) {
                    let val: any = 0; 
                    let exactVal = studentGrades[sourceSubj];
                    if (exactVal !== undefined) {
                        val = exactVal;
                    } else {
                        const possibleKey = Object.keys(studentGrades).find(k => k.toUpperCase().includes(sourceSubj.toUpperCase()));
                        if (possibleKey) val = studentGrades[possibleKey];
                    }
                    
                    // Bersihkah quote satu (')
                    if (typeof val === 'string') {
                        val = val.replace(/^'/, '').trim();
                    }

                    const cellAddr = XLSX.utils.encode_cell({ r, c: colIdx });
                    const isNum = !isNaN(parseFloat(val)) && isFinite(val as any);
                    const setVal = isNum ? parseFloat(val) : val;

                    if (sheet[cellAddr]) {
                        sheet[cellAddr].v = setVal;
                        sheet[cellAddr].t = isNum ? 'n' : 's';
                    } else {
                        sheet[cellAddr] = { v: setVal, t: isNum ? 'n' : 's' };
                        if (colIdx > range.e.c) range.e.c = colIdx;
                    }
                 }
              }
           }

           const missingStudents: MissingStudent[] = [];
           let no = 1;
           const allSemNisns = Object.keys(sourceGrades[sem]);
           for (const nisn of allSemNisns) {
                if (!processedNisns.has(nisn)) {
                     const info = sourceInfo[nisn] || { name: '-', kelas: '-' };
                     missingStudents.push({ no: no++, nisn, name: info.name, kelas: info.kelas });
                }
           }

           // Tuliskan rekapitulasi data yang tidak masuk ke dalam excel di bagian bawah
           if (missingStudents.length > 0) {
               const startRowForMissing = range.e.r + 3;
               
               // Header rekapitulasi
               const headerNISNAddr = XLSX.utils.encode_cell({ r: startRowForMissing, c: nisnColIdx });
               sheet[headerNISNAddr] = { v: 'NISN (Tidak Masuk Template)', t: 's' };
               if (!sheet[headerNISNAddr].s) sheet[headerNISNAddr].s = {};
               sheet[headerNISNAddr].s.font = { bold: true };
               
               const headerNamaAddr = XLSX.utils.encode_cell({ r: startRowForMissing, c: nisnColIdx + 1 });
               sheet[headerNamaAddr] = { v: 'Nama Siswa', t: 's' };
               if (!sheet[headerNamaAddr].s) sheet[headerNamaAddr].s = {};
               sheet[headerNamaAddr].s.font = { bold: true };
               
               const headerKelasAddr = XLSX.utils.encode_cell({ r: startRowForMissing, c: nisnColIdx + 2 });
               sheet[headerKelasAddr] = { v: 'Kelas', t: 's' };
               if (!sheet[headerKelasAddr].s) sheet[headerKelasAddr].s = {};
               sheet[headerKelasAddr].s.font = { bold: true };

               for (let i = 0; i < missingStudents.length; i++) {
                   const r = startRowForMissing + 1 + i;
                   sheet[XLSX.utils.encode_cell({ r, c: nisnColIdx })] = { v: missingStudents[i].nisn, t: 's' };
                   sheet[XLSX.utils.encode_cell({ r, c: nisnColIdx + 1 })] = { v: missingStudents[i].name, t: 's' };
                   sheet[XLSX.utils.encode_cell({ r, c: nisnColIdx + 2 })] = { v: missingStudents[i].kelas, t: 's' };
               }
               
               const newMaxRow = startRowForMissing + missingStudents.length;
               if (newMaxRow > range.e.r) {
                   range.e.r = newMaxRow;
               }
           }

           sheet['!ref'] = XLSX.utils.encode_range(range);
           
           generatedResults[sem] = { 
               wb: templateWb, 
               studentCount: matchCount, 
               missingStudents,
               totalSourceCount: allSemNisns.length 
           };
         }
      }

      setResult({ generated: generatedResults, semesters: semestersDesc });

    } catch (err: any) {
      console.error(err);
      setError(err.message || "Terjadi kesalahan saat memproses data.");
    } finally {
      setIsProcessing(false);
    }
  };

  const exportTemplateMaster = () => {
      const wsData = [
          ["PENGOLAHAN NILAI SEKOLAH"],
          ["SMPN 2 SUKOREJO"],
          ["TAHUN AJARAN 2025/2026"],
          [],
          ["NO", "", "", "NAMA", "L/P", "KELAS", "NILAI", "MATA PELAJARAN", "", "", "", "", "", "", "", "", "", "JUMLAH", "RATA-RATA"],
          ["URUT", "NIS", "NISN", "", "", "", "", "Pendidikan Agama", "Pendidikan Pancasila", "Bahasa Indonesia", "Matematika", "IPA", "IPS", "Bahasa Inggris", "PJOK", "Informatika", "Seni Budaya", "Bahasa Daerah", "", ""],
          ["1", "3025", "0106867895", "ADE EKA SAFITRI", "P", "9A", "SM 1", 83, 81, 88, 81, 90, 88, 82, 90, 89, 82, 83, "", ""],
          ["", "", "", "", "", "", "SM 2", 86, 86, 82, 82, 85, 84, 80, 87, 86, 85, 86, "", ""],
      ];
      const ws = XLSX.utils.aoa_to_sheet(wsData);

      ws["!merges"] = [
          { s: {r: 0, c: 0}, e: {r: 0, c: 10} },
          { s: {r: 4, c: 0}, e: {r: 4, c: 2} },
          { s: {r: 4, c: 3}, e: {r: 5, c: 3} },
          { s: {r: 4, c: 7}, e: {r: 4, c: 17} },
      ];

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Master Data");
      XLSX.writeFile(wb, "Template_Sumber_Data.xlsx");
  };

  return (
    <div className="h-screen w-full flex flex-col bg-slate-50 dark:bg-slate-900 text-slate-900 dark:text-slate-100 font-sans transition-colors duration-300 overflow-hidden relative">
      {/* Background blobs for glassmorphism */}
      <div className="absolute top-0 -left-40 w-96 h-96 bg-blue-300 dark:bg-blue-900/40 rounded-full mix-blend-multiply dark:mix-blend-lighten filter blur-[100px] opacity-40 dark:opacity-20 animate-blob pointer-events-none"></div>
      <div className="absolute top-0 -right-40 w-96 h-96 bg-purple-300 dark:bg-purple-900/40 rounded-full mix-blend-multiply dark:mix-blend-lighten filter blur-[100px] opacity-40 dark:opacity-20 animate-blob animation-delay-2000 pointer-events-none"></div>
      <div className="absolute -bottom-40 left-20 w-96 h-96 bg-indigo-300 dark:bg-indigo-900/40 rounded-full mix-blend-multiply dark:mix-blend-lighten filter blur-[100px] opacity-40 dark:opacity-20 animate-blob animation-delay-4000 pointer-events-none"></div>
      
      {/* Top Navbar */}
      <nav className="h-16 bg-white/60 dark:bg-slate-900/60 backdrop-blur-xl border-b border-white/40 dark:border-slate-800/60 flex items-center justify-between px-6 md:px-8 z-20 shrink-0 shadow-sm transition-colors duration-300">
        <div className="flex items-center gap-4">
          <img 
            src="https://lh3.googleusercontent.com/a/ACg8ocKIyOSmUCibkuiYO0w4wo1Pl54QsoKUQBc3jfSxADJZEfuybRTZ=s288-c-no" 
            alt="Logo Aplikasi" 
            className="w-10 h-10 rounded-full shadow-md border-2 border-white dark:border-slate-700 object-cover" 
            referrerPolicy="no-referrer"
          />
          <div>
             <span className="font-bold text-slate-800 dark:text-slate-100 text-lg tracking-tight block leading-tight">Generator Nilai</span>
             <span className="text-xs font-medium text-blue-600 dark:text-blue-400">Tiap Semester v2.1</span>
          </div>
        </div>
        
        <div className="flex items-center gap-4 md:gap-6">
          <div className="hidden sm:flex items-center gap-2 text-sm text-slate-500 dark:text-slate-400 font-medium">
            <span className="w-2 h-2 rounded-full bg-green-500 animate-pulse"></span> Siap Diproses
          </div>
          <div className="hidden sm:block h-6 w-px bg-slate-200 dark:bg-slate-700"></div>

          {/* Theme Toggle */}
          <button 
             onClick={() => setIsDarkMode(!isDarkMode)}
             className="p-2 rounded-full hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors text-slate-600 dark:text-slate-300"
          >
             {isDarkMode ? <Sun className="w-5 h-5" /> : <Moon className="w-5 h-5" />}
          </button>

          <div className="hidden sm:flex items-center gap-3 pl-2 border-l border-slate-200 dark:border-slate-700">
            <div className="text-right">
              <div className="text-sm font-semibold text-slate-900 dark:text-slate-100">Pendidik</div>
              <div className="text-[10px] text-slate-500 dark:text-slate-400 uppercase tracking-widest">Administrator</div>
            </div>
            <div className="w-9 h-9 rounded-full bg-slate-100 dark:bg-slate-800 border-2 border-white dark:border-slate-700 flex items-center justify-center text-slate-600 dark:text-slate-300 font-bold text-sm shadow-sm">AD</div>
          </div>
        </div>
      </nav>

      <main className="flex-1 flex flex-col md:flex-row overflow-hidden p-4 md:p-6 lg:p-8 gap-6 max-w-[1500px] w-full mx-auto relative z-10 transition-colors duration-300">
        
        {/* Left Sidebar Control */}
        <div className="w-full md:w-96 flex flex-col gap-6 overflow-y-auto [&::-webkit-scrollbar]:hidden pb-10">
          
          {/* SECTION 1: Sumber Data */}
          <div className="relative rounded-2xl p-[1px] overflow-hidden group/border shadow-sm shrink-0 transition-all duration-300">
             <div className="absolute inset-[-100%] bg-[conic-gradient(from_0deg,transparent_0_340deg,rgba(59,130,246,0.6)_360deg)] dark:bg-[conic-gradient(from_0deg,transparent_0_340deg,rgba(59,130,246,0.8)_360deg)] animate-[spin_5s_linear_infinite] opacity-0 group-hover/border:opacity-100 transition-opacity duration-500" />
             <section className="relative bg-white/70 dark:bg-slate-800/60 backdrop-blur-md border border-white/50 dark:border-slate-700/50 rounded-2xl p-5 h-full transition-colors duration-300">
               <div className="flex justify-between items-center mb-5">
              <div className="flex items-center gap-2">
                 <div className="flex items-center justify-center w-6 h-6 rounded-full bg-slate-100 dark:bg-slate-700 text-xs font-bold text-slate-500 dark:text-slate-300">1</div>
                 <h2 className="text-sm font-semibold text-slate-800 dark:text-slate-100 tracking-wide">Sumber Data (Master)</h2>
              </div>
              <button 
                onClick={exportTemplateMaster}
                className="text-[10px] bg-blue-50 dark:bg-blue-500/10 text-blue-600 dark:text-blue-400 hover:bg-blue-100 dark:hover:bg-blue-500/20 px-2.5 py-1.5 rounded-md font-semibold transition-colors border border-blue-100 dark:border-blue-500/20"
              >
                Unduh Template
              </button>
            </div>
            
            <div
              className={cn(
                  "border-2 border-dashed rounded-xl p-8 flex flex-col items-center justify-center text-center transition-all cursor-pointer",
                  isDragSource 
                     ? "border-blue-500 bg-blue-50 dark:bg-blue-500/10 scale-[1.02]" 
                     : "border-slate-200 dark:border-slate-600 hover:border-blue-400 dark:hover:border-blue-400 bg-slate-50 dark:bg-slate-800/50 hover:bg-white dark:hover:bg-slate-800"
              )}
              onDragOver={(e) => { e.preventDefault(); setIsDragSource(true); }}
              onDragLeave={(e) => { e.preventDefault(); setIsDragSource(false); }}
              onDrop={(e) => {
                e.preventDefault(); setIsDragSource(false);
                if (e.dataTransfer.files) setSourceFiles(prev => [...prev, ...Array.from(e.dataTransfer.files)]);
              }}
              onClick={() => sourceInputRef.current?.click()}
            >
              <div className={cn("p-3 rounded-full mb-3 transition-colors", isDragSource ? "bg-blue-100 dark:bg-blue-500/20" : "bg-white dark:bg-slate-700 shadow-sm")}>
                 <Database className={cn("w-6 h-6", isDragSource ? "text-blue-600 dark:text-blue-400" : "text-slate-400 dark:text-slate-300")} />
              </div>
              <p className="text-sm font-medium text-slate-700 dark:text-slate-200">Klik atau Tarik File Nilai Master</p>
              <p className="text-xs text-slate-400 dark:text-slate-500 mt-1.5">Mendukung rekapitulasi nilai Semester 1-6</p>
              
              <input type="file" multiple accept=".xlsx, .xls, .csv" className="hidden" ref={sourceInputRef} onChange={handleSourceChange} />
            </div>

            {sourceFiles.length > 0 && (
              <div className="mt-4 space-y-2">
                {sourceFiles.map((file, index) => (
                  <div key={index} className="flex items-center justify-between text-xs p-2.5 bg-slate-50 dark:bg-slate-700/50 text-slate-700 dark:text-slate-200 rounded-lg border border-slate-200 dark:border-slate-600 transition-colors">
                    <div className="flex items-center gap-2.5 overflow-hidden">
                      <FileSpreadsheet className="w-4 h-4 text-blue-500 dark:text-blue-400 shrink-0" /> 
                      <span className="truncate max-w-[180px] font-medium">{file.name}</span>
                    </div>
                    <button onClick={(e) => { e.stopPropagation(); setSourceFiles(prev => prev.filter((_, i) => i !== index)); }} className="text-slate-400 dark:text-slate-500 hover:text-red-500 dark:hover:text-red-400 p-1 rounded-md hover:bg-slate-200 dark:hover:bg-slate-600 transition-colors">
                      <X className="w-3.5 h-3.5" />
                    </button>
                  </div>
                ))}
              </div>
            )}
             </section>
          </div>

          {/* SECTION 2: Template Output */}
          <div className="relative rounded-2xl p-[1px] overflow-hidden group/border shadow-sm shrink-0 transition-all duration-300">
             <div className="absolute inset-[-100%] bg-[conic-gradient(from_0deg,transparent_0_340deg,rgba(99,102,241,0.6)_360deg)] dark:bg-[conic-gradient(from_0deg,transparent_0_340deg,rgba(99,102,241,0.8)_360deg)] animate-[spin_5s_linear_infinite] opacity-0 group-hover/border:opacity-100 transition-opacity duration-500" />
             <section className="relative bg-white/70 dark:bg-slate-800/60 backdrop-blur-md border border-white/50 dark:border-slate-700/50 rounded-2xl p-5 h-full transition-colors duration-300">
               <div className="flex items-center gap-2 mb-5">
                 <div className="flex items-center justify-center w-6 h-6 rounded-full bg-slate-100 dark:bg-slate-700 text-xs font-bold text-slate-500 dark:text-slate-300">2</div>
                 <h2 className="text-sm font-semibold text-slate-800 dark:text-slate-100 tracking-wide">Template Cetak (Tujuan)</h2>
              </div>
            
            <div
              className={cn(
                  "border-2 border-dashed rounded-xl p-8 flex flex-col items-center justify-center text-center transition-all cursor-pointer",
                  isDragTemplate 
                     ? "border-indigo-500 bg-indigo-50 dark:bg-indigo-500/10 scale-[1.02]" 
                     : "border-slate-200 dark:border-slate-600 hover:border-indigo-400 dark:hover:border-indigo-400 bg-slate-50 dark:bg-slate-800/50 hover:bg-white dark:hover:bg-slate-800"
              )}
              onDragOver={(e) => { e.preventDefault(); setIsDragTemplate(true); }}
              onDragLeave={(e) => { e.preventDefault(); setIsDragTemplate(false); }}
              onDrop={(e) => {
                e.preventDefault(); setIsDragTemplate(false);
                if (e.dataTransfer.files && e.dataTransfer.files.length > 0) setTemplateFile(e.dataTransfer.files[0]);
              }}
              onClick={() => templateInputRef.current?.click()}
            >
              <div className={cn("p-3 rounded-full mb-3 transition-colors", isDragTemplate ? "bg-indigo-100 dark:bg-indigo-500/20" : "bg-white dark:bg-slate-700 shadow-sm")}>
                <LayoutTemplate className={cn("w-6 h-6", isDragTemplate ? "text-indigo-600 dark:text-indigo-400" : "text-slate-400 dark:text-slate-300")} />
              </div>
              <p className="text-sm font-medium text-slate-700 dark:text-slate-200">Upload Template Output Asli</p>
              <p className="text-xs text-slate-400 dark:text-slate-500 mt-1.5">Harus mengandung baris header (NISN, Nama Siswa, PA, dsb.)</p>
              
              <input type="file" accept=".xlsx, .xls" className="hidden" ref={templateInputRef} onChange={handleTemplateChange} />
            </div>

            {templateFile && (
              <div className="mt-4 flex items-center justify-between text-xs p-2.5 bg-indigo-50 dark:bg-indigo-500/10 text-indigo-700 dark:text-indigo-300 rounded-lg border border-indigo-100 dark:border-indigo-500/20 transition-colors">
                <div className="flex items-center gap-2.5 overflow-hidden">
                  <LayoutTemplate className="w-4 h-4 text-indigo-500 dark:text-indigo-400 shrink-0" /> 
                  <span className="truncate max-w-[180px] font-medium">{templateFile.name}</span>
                </div>
                <button onClick={() => setTemplateFile(null)} className="text-indigo-400 dark:text-indigo-500 hover:text-red-500 dark:hover:text-red-400 p-1 rounded-md hover:bg-indigo-100 dark:hover:bg-indigo-500/20 transition-colors">
                  <X className="w-3.5 h-3.5" />
                </button>
              </div>
            )}
             </section>
          </div>

          {/* SECTION 3: Pemrosesan */}
          {(sourceFiles.length > 0 && templateFile) && (
            <div className="relative rounded-2xl p-[1px] overflow-hidden group/border shadow-sm flex flex-col shrink-0 min-h-[fit-content] transition-all duration-300 animate-in fade-in slide-in-from-bottom-4">
              <div className="absolute inset-[-100%] bg-[conic-gradient(from_0deg,transparent_0_340deg,rgba(59,130,246,0.6)_360deg)] dark:bg-[conic-gradient(from_0deg,transparent_0_340deg,rgba(59,130,246,0.8)_360deg)] animate-[spin_3s_linear_infinite] opacity-100 transition-opacity duration-500" />
              <section className="relative bg-white/70 dark:bg-slate-800/60 backdrop-blur-md border border-white/50 dark:border-slate-700/50 rounded-2xl p-5 h-full flex flex-col transition-colors duration-300">
                <div className="flex items-center gap-2 mb-5">
                 <div className="flex items-center justify-center w-6 h-6 rounded-full bg-blue-100 dark:bg-blue-900 text-xs font-bold text-blue-600 dark:text-blue-400">3</div>
                 <h2 className="text-sm font-semibold text-slate-800 dark:text-slate-100 tracking-wide">Mulai Eksekusi</h2>
              </div>
              
              <div className="mt-auto">
                <button
                  onClick={processData}
                  disabled={isProcessing}
                  className="w-full bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white py-3.5 rounded-xl font-bold text-sm shadow-lg shadow-blue-500/25 transition-all disabled:opacity-70 disabled:cursor-not-allowed flex items-center justify-center gap-2"
                >
                  {isProcessing ? (
                    <>
                      <RefreshCw className="w-4 h-4 animate-spin" />
                      Mengeksekusi Formula...
                    </>
                  ) : (
                    'Proses Penyelarasan Otomatis'
                  )}
                </button>
                <div className="text-[10px] text-center text-slate-500 dark:text-slate-400 mt-4 px-2 py-2.5 bg-slate-50 dark:bg-slate-700/50 rounded-lg flex items-center justify-center gap-1.5 border border-slate-200 dark:border-slate-600/50">
                  <AlertTriangle className="w-3.5 h-3.5 text-amber-500 shrink-0" />
                   Akan mencocokkan nilai berdasarkan kecocokan NISN
                </div>
              </div>
            </section>
           </div>
          )}
        </div>

        {/* SECTION RESULT PREVIEW */}
        <div className="flex-1 flex flex-col bg-white/60 dark:bg-slate-800/40 backdrop-blur-xl border border-white/40 dark:border-slate-700/50 rounded-2xl shadow-sm overflow-hidden transition-colors duration-300 relative">
          
          {/* Header Preview */}
          <header className="h-[72px] border-b border-white/40 dark:border-slate-700/50 flex items-center justify-between px-6 lg:px-8 shrink-0 bg-white/30 dark:bg-slate-800/20">
            <div className="flex items-center gap-3">
              <h3 className="text-base font-bold text-slate-800 dark:text-slate-100">
                 Ruang Pratinjau
              </h3>
              {result && (
                <span className="flex items-center gap-1 text-xs font-semibold bg-emerald-100 dark:bg-emerald-500/10 text-emerald-700 dark:text-emerald-400 px-2.5 py-1 rounded-full border border-emerald-200 dark:border-emerald-500/20">
                  <CheckCircle2 className="w-3 h-3" /> Berhasil ({result.semesters.length} Semester)
                </span>
              )}
            </div>
          </header>

          <div className="flex-1 overflow-auto p-6 lg:p-8 relative">
            {/* Global Errors */}
            {error && (
              <div className="mb-6 p-4 bg-red-50 dark:bg-red-500/10 border border-red-200 dark:border-red-500/20 rounded-xl flex items-start gap-3 shadow-sm animate-in fade-in">
                <AlertCircle className="w-5 h-5 text-red-500 dark:text-red-400 mt-0.5 shrink-0" />
                <p className="text-red-700 dark:text-red-300 font-medium text-sm leading-relaxed">{error}</p>
              </div>
            )}

            {/* Results Grid */}
            {result ? (
              <div className="space-y-8 animate-in slide-in-from-bottom-8 duration-700 fade-in">
                {result.semesters.map((semester) => (
                  <div key={semester} className="flex flex-col border border-white/50 dark:border-slate-700/50 rounded-2xl bg-white/70 dark:bg-slate-800/60 backdrop-blur-md overflow-hidden shadow-sm">
                     <div className="p-6 md:p-8 flex flex-col md:flex-row justify-between items-start md:items-center border-b border-white/40 dark:border-slate-700/50 gap-4">
                        <div>
                          <span className="font-bold text-2xl text-slate-900 dark:text-white block mb-2">{semester}</span>
                          <div className="flex flex-wrap items-center gap-3">
                             <span className="text-sm font-semibold text-emerald-700 dark:text-emerald-400 bg-emerald-100 dark:bg-emerald-500/10 px-3 py-1.5 rounded-md inline-flex items-center gap-1">
                                <CheckCircle2 className="w-4 h-4" />
                                {result.generated[semester].studentCount} Siswa Masuk
                             </span>
                             {result.generated[semester].missingStudents.length > 0 && (
                               <span className="text-sm font-semibold text-rose-700 dark:text-rose-400 bg-rose-100 dark:bg-rose-500/10 px-3 py-1.5 rounded-md inline-flex items-center gap-1">
                                  <AlertCircle className="w-4 h-4" />
                                  {result.generated[semester].missingStudents.length} Siswa Tidak Masuk
                               </span>
                             )}
                             <span className="text-xs text-slate-500 dark:text-slate-400 font-medium">dari total {result.generated[semester].totalSourceCount} data master</span>
                          </div>
                        </div>
                        <button
                          onClick={() => XLSX.writeFile(result.generated[semester].wb, `${semester.replace(/\s+/g, '_')}_Result.xlsx`)}
                          className="flex items-center justify-center gap-2 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white font-bold py-3 px-6 rounded-xl shadow-lg shadow-blue-500/30 transition-all text-sm shrink-0 whitespace-nowrap"
                        >
                          <Download className="w-4 h-4" />
                          Unduh Output Excel
                        </button>
                     </div>

                     {/* Table of missed students */}
                     {result.generated[semester].missingStudents.length > 0 && (
                        <div className="p-6 md:p-8 bg-white/40 dark:bg-slate-800/40">
                           <h4 className="text-sm font-bold text-slate-800 dark:text-slate-200 mb-4 flex items-center gap-2">
                             <AlertTriangle className="w-4 h-4 text-rose-500" /> 
                             Daftar Siswa Tidak Masuk Template ({semester})
                           </h4>
                           <div className="overflow-x-auto border border-white/50 dark:border-slate-700/50 rounded-xl overflow-hidden bg-white/60 dark:bg-slate-900/40 backdrop-blur-md shadow-sm">
                             <table className="w-full text-left text-sm text-slate-700 dark:text-slate-300">
                                <thead className="text-xs uppercase bg-slate-100/50 dark:bg-slate-800/50 text-slate-800 dark:text-slate-400 border-b border-white/50 dark:border-slate-700/50">
                                   <tr>
                                      <th className="px-4 py-3 font-semibold text-center w-16">No</th>
                                      <th className="px-4 py-3 font-semibold w-32 border-x border-white/40 dark:border-slate-700/30">NISN</th>
                                      <th className="px-4 py-3 font-semibold">Nama Siswa</th>
                                      <th className="px-4 py-3 font-semibold w-24 text-center border-l border-white/40 dark:border-slate-700/30">Kelas</th>
                                   </tr>
                                </thead>
                                <tbody className="divide-y divide-white/50 dark:divide-slate-800/50">
                                   {result.generated[semester].missingStudents.map((ms) => (
                                       <tr key={ms.nisn} className="hover:bg-white/80 dark:hover:bg-slate-800/80 transition-colors">
                                          <td className="px-4 py-3 text-center text-slate-500 dark:text-slate-400">{ms.no}</td>
                                          <td className="px-4 py-3 font-mono text-xs border-x border-white/40 dark:border-slate-700/30">{ms.nisn}</td>
                                          <td className="px-4 py-3 font-medium text-slate-800 dark:text-slate-200">{ms.name}</td>
                                          <td className="px-4 py-3 text-center font-medium text-slate-700 dark:text-slate-300 border-l border-white/40 dark:border-slate-700/30">{ms.kelas}</td>
                                       </tr>
                                   ))}
                                </tbody>
                             </table>
                           </div>
                           <p className="text-xs text-slate-500 dark:text-slate-400 mt-3 ml-1">
                             * Siswa di atas ada di Data Master namun tidak ditemukan NISN-nya pada file Template. Informasi ini juga telah ditambahkan di bagian bawah file Excel yang diunduh.
                           </p>
                        </div>
                     )}
                  </div>
                ))}
              </div>
            ) : (
              <div className="w-full h-full min-h-[400px] flex flex-col items-center justify-center text-slate-400 dark:text-slate-500 animate-in fade-in duration-1000">
                 <div className="relative mb-6">
                    <div className="absolute inset-0 bg-blue-100 dark:bg-blue-900 blur-2xl rounded-full opacity-50"></div>
                    <Layers className="w-20 h-20 text-slate-200 dark:text-slate-700 relative z-10" />
                 </div>
                 <h4 className="text-lg font-semibold text-slate-700 dark:text-slate-300 mb-2">Area Tampil Kosong</h4>
                 <p className="text-sm text-slate-500 dark:text-slate-400 max-w-sm text-center leading-relaxed">
                   Pastikan Anda telah mengisi <strong className="text-slate-700 dark:text-slate-300">Sumber Data</strong> dan <strong className="text-slate-700 dark:text-slate-300">Template Cetak</strong> di panel kiri, kemudian klik Eksekusi.
                 </p>
              </div>
            )}
          </div>
        </div>
      </main>
    </div>
  );
}
