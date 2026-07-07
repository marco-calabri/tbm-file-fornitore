import React, { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { 
  Upload, 
  Folder, 
  FolderOpen, 
  FileSpreadsheet, 
  FileArchive, 
  Settings, 
  Play, 
  CheckCircle2, 
  XCircle, 
  Search, 
  Filter, 
  Download, 
  RefreshCw, 
  Info, 
  X, 
  ChevronRight, 
  Eye, 
  HelpCircle, 
  Check, 
  Layers, 
  Database, 
  FolderSync,
  AlertTriangle,
  Copy,
  Files
} from 'lucide-react';
import { ExcelRow, ProjectFile, ProcessedRow, ColumnMapping, AppStats } from './types';
import { 
  matchRowFiles, 
  computeStats, 
  generateUpdatedExcel, 
  generateZIP,
  autoFitColumns
} from './utils/fileMatcher';

export default function App() {
  // --- Custom Notification / Toast State ---
  const [toast, setToast] = useState<{ message: string; type: 'success' | 'error' | 'info' } | null>(null);

  const showNotification = (message: string, type: 'success' | 'error' | 'info' = 'error') => {
    setToast({ message, type });
    // Auto-dismiss after 5 seconds
    setTimeout(() => {
      setToast(prev => prev?.message === message ? null : prev);
    }, 5000);
  };

  // --- File Inputs State ---
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [sheets, setSheets] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  const [rawExcelData, setRawExcelData] = useState<any[]>([]);
  
  // Scanned files from folder picker
  const [projectFiles, setProjectFiles] = useState<ProjectFile[]>([]);
  const [projectFolderName, setProjectFolderName] = useState<string>('');

  // --- Column Mapping State ---
  const [columns, setColumns] = useState<string[]>([]);
  const [mapping, setMapping] = useState<ColumnMapping>({
    codice: '',
    configurazione: '',
    revisione: '',
    fornitore: ''
  });
  const [showMappingConfig, setShowMappingConfig] = useState(false);

  // --- Processing State ---
  const [isProcessing, setIsProcessing] = useState(false);
  const [processProgress, setProcessProgress] = useState(0);
  const [processStatusText, setProcessStatusText] = useState('');
  const [processedRows, setProcessedRows] = useState<ProcessedRow[]>([]);
  const [stats, setStats] = useState<AppStats | null>(null);

  // --- UI Filters State ---
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState<'ALL' | 'FOUND' | 'MISSING'>('ALL');
  const [selectedSupplierFilter, setSelectedSupplierFilter] = useState<string>('ALL');
  const [selectedExtensionFilter, setSelectedExtensionFilter] = useState<string>('ALL');
  const [ba1StpMissingOnly, setBa1StpMissingOnly] = useState(false);
  const [showBa1StpModal, setShowBa1StpModal] = useState(false);
  
  // --- Selected Row Modal ---
  const [selectedDetailRow, setSelectedDetailRow] = useState<ProcessedRow | null>(null);
  const [copiedRowId, setCopiedRowId] = useState<string | null>(null);



  // Refs
  const folderInputRef = useRef<HTMLInputElement>(null);
  const filesInputRef = useRef<HTMLInputElement>(null);
  const excelInputRef = useRef<HTMLInputElement>(null);

  // Auto-detect columns when sheet data or mapping columns change
  useEffect(() => {
    if (columns.length > 0) {
      const detectMapping = {
        codice: '',
        configurazione: '',
        revisione: '',
        fornitore: ''
      };

      // Simple heuristic mapping
      for (const col of columns) {
        const colLower = col.toLowerCase();
        if (!detectMapping.codice && (colLower === 'codice' || colLower === 'code' || colLower === 'art_code')) {
          detectMapping.codice = col;
        } else if (!detectMapping.configurazione && (colLower === 'configurazione' || colLower === 'config' || colLower === 'configuration')) {
          detectMapping.configurazione = col;
        } else if (!detectMapping.revisione && (colLower === 'revisione' || colLower === 'rev' || colLower === 'revision')) {
          detectMapping.revisione = col;
        } else if (!detectMapping.fornitore && (colLower === 'fornitore' || colLower === 'supplier' || colLower === 'vendor' || colLower === 'forn')) {
          detectMapping.fornitore = col;
        }
      }

      // Fallbacks if not detected
      setMapping({
        codice: detectMapping.codice || columns[0] || '',
        configurazione: detectMapping.configurazione || columns[1] || columns[0] || '',
        revisione: detectMapping.revisione || '', // allowed to be empty
        fornitore: detectMapping.fornitore || columns[4] || columns[3] || ''
      });
    }
  }, [columns]);

  // Handle Excel upload
  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;
    const file = files[0];
    processExcelFile(file);
  };

  const processExcelFile = (file: File) => {
    setExcelFile(file);
    
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const readWorkbook = XLSX.read(data, { type: 'array' });
        
        setWorkbook(readWorkbook);
        setSheets(readWorkbook.SheetNames);
        
        const firstSheet = readWorkbook.SheetNames[0];
        setSelectedSheet(firstSheet);
        loadSheetData(readWorkbook, firstSheet);
      } catch (err) {
        showNotification("Errore nel caricamento del file Excel: " + (err as Error).message, "error");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const loadSheetData = (wb: XLSX.WorkBook, sheetName: string) => {
    const sheet = wb.Sheets[sheetName];
    // Read raw rows as array of objects
    const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    setRawExcelData(jsonData);

    if (jsonData.length > 0) {
      // Find all column headers
      const headers = Object.keys(jsonData[0]);
      setColumns(headers);
    } else {
      setColumns([]);
    }
    // Clear previous processed data when new file loads
    setProcessedRows([]);
    setStats(null);
  };

  const handleSheetChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const sheetName = e.target.value;
    setSelectedSheet(sheetName);
    if (workbook) {
      loadSheetData(workbook, sheetName);
    }
  };

  // Handle Folder selection via WebkitDirectory
  const handleFolderUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;
    
    // Extract folder name from the first file relative path
    const firstPath = files[0].webkitRelativePath || '';
    const folderRootName = firstPath.split('/')[0] || 'Cartella Progetto';
    setProjectFolderName(folderRootName);

    setProcessStatusText("Lettura file all'interno della cartella...");
    const parsed = parseFolderFiles(files);
    setProjectFiles(parsed);
    showNotification(`Caricata con successo la cartella "${folderRootName}" (${parsed.length} file rilevati)`, 'success');
  };

  // Handle individual files selection (avoiding browser security popup)
  const handleFilesUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    setProjectFolderName("File Multipli Selezionati");
    setProcessStatusText("Lettura file selezionati...");
    const parsed = parseFolderFiles(files);
    setProjectFiles(parsed);
    showNotification(`Caricati con successo ${parsed.length} file disegni!`, 'success');
  };

  const parseFolderFiles = (fileList: FileList): ProjectFile[] => {
    const list: ProjectFile[] = [];
    for (let i = 0; i < fileList.length; i++) {
      const file = fileList[i];
      const path = file.webkitRelativePath || file.name;
      
      // Get filename and extension
      const lastSlash = path.lastIndexOf('/');
      const fileNameWithExt = lastSlash !== -1 ? path.substring(lastSlash + 1) : path;
      
      const lastDot = fileNameWithExt.lastIndexOf('.');
      let name = fileNameWithExt;
      let ext = '';
      if (lastDot !== -1) {
        name = fileNameWithExt.substring(0, lastDot);
        ext = fileNameWithExt.substring(lastDot).toLowerCase();
      }

      list.push({
        name,
        fullName: fileNameWithExt,
        relativePath: path,
        extension: ext,
        fileObject: file
      });
    }
    return list;
  };



  // Trigger main processing: Match files, copy mockly, generate stats
  const handleProcessSearch = async () => {
    if (rawExcelData.length === 0) {
      showNotification("Carica un file Excel prima di elaborare!", "error");
      return;
    }
    if (projectFiles.length === 0) {
      showNotification("Seleziona una cartella progetto contenente file prima di elaborare!", "error");
      return;
    }
    if (!mapping.codice) {
      showNotification("Seleziona almeno la colonna per il 'Codice'!", "error");
      return;
    }

    setIsProcessing(true);
    setProcessProgress(10);
    setProcessStatusText("Mappatura delle colonne ed estrazione dati...");

    setTimeout(() => {
      // 1. Map raw excel rows into strict ExcelRow structures
      const formattedExcelRows: ExcelRow[] = rawExcelData.map((row, idx) => {
        return {
          __rowNum__: idx + 1, // Excel index is 1-based (header is at row index 1, data starts from 2)
          codice: String(row[mapping.codice] || '').trim(),
          configurazione: mapping.configurazione ? String(row[mapping.configurazione] || '').trim() : '',
          revisione: mapping.revisione ? String(row[mapping.revisione] || '').trim() : '',
          fornitore: mapping.fornitore ? String(row[mapping.fornitore] || '').trim() : 'Senza_Fornitore',
          originalData: row
        };
      });

      setProcessProgress(40);
      setProcessStatusText("Ricerca ricorsiva dei file di progetto (.pdf, .dwg, .stp/.step)...");

      setTimeout(() => {
        // 2. Perform file matching
        const results = matchRowFiles(formattedExcelRows, projectFiles);
        setProcessedRows(results);

        // 3. Compute statistics
        const computedStats = computeStats(results);
        setStats(computedStats);

        setProcessProgress(100);
        setProcessStatusText("Elaborazione completata con successo!");
        
        setTimeout(() => {
          setIsProcessing(false);
        }, 800);
      }, 600);
    }, 400);
  };

  // Export updated Excel file
  const handleDownloadExcel = () => {
    if (!workbook || processedRows.length === 0) return;
    try {
      const excelBlob = generateUpdatedExcel(workbook, processedRows, selectedSheet);
      const url = URL.createObjectURL(excelBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = excelFile 
        ? `${excelFile.name.replace('.xlsx', '')}_elaborato.xlsx` 
        : 'distinta_elaborata.xlsx';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      showNotification("Errore nella generazione del file Excel: " + (err as Error).message, "error");
    }
  };

  // Export list of BA1 files without STP/STEP
  const handleExportBa1StpMissing = () => {
    const ba1StpMissingRows = processedRows.filter(row => {
      const isBa1 = row.excelRow.codice && row.excelRow.codice.toUpperCase().startsWith("BA1");
      const hasStp = row.matchResult.stpMatched || row.matchResult.stpFiles.length > 0;
      return isBa1 && !hasStp;
    });

    if (ba1StpMissingRows.length === 0) {
      showNotification("Nessun articolo BA1 con file .stp/.step mancante!", "info");
      return;
    }

    // Create worksheet data
    const data = ba1StpMissingRows.map(row => ({
      'Riga Excel': row.excelRow.__rowNum__ || '',
      'Codice Articolo': row.excelRow.codice || '',
      'Configurazione': row.excelRow.configurazione || '',
      'Revisione': row.excelRow.revisione || '',
      'Nome File Ricercato': row.matchResult.targetBaseName || '',
      'Fornitore': row.excelRow.fornitore || 'Senza Fornitore',
      'Nome File Atteso': `${row.matchResult.targetBaseName}.stp`
    }));

    try {
      const ws = XLSX.utils.json_to_sheet(data);
      autoFitColumns(ws);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'BA1 STP Mancanti');

      const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const sourceName = projectFolderName || 'Cartella Disegni';
      a.download = `${sourceName}_BA1 SENZA FILE STEP.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      showNotification("Errore nella generazione dell'esportazione: " + (err as Error).message, "error");
    }
  };

  // Export organized Supplier ZIP file
  const handleDownloadZIP = async () => {
    if (processedRows.length === 0) return;
    try {
      setIsProcessing(true);
      setProcessProgress(0);
      setProcessStatusText("Inizializzazione archivio ZIP...");

      const zipBlob = await generateZIP(processedRows, (percent, file) => {
        setProcessProgress(percent);
        setProcessStatusText(`Aggiunta file: ${file}`);
      });

      setProcessProgress(100);
      setProcessStatusText("Compressione ultimata!");

      const url = URL.createObjectURL(zipBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `File_Elaborati_Fornitori.zip`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      
      setTimeout(() => {
        setIsProcessing(false);
      }, 500);
    } catch (err) {
      setIsProcessing(false);
      showNotification("Errore nella generazione dello ZIP: " + (err as Error).message, "error");
    }
  };

  // Export organized ZIP file specifically for supplier 'FALEGNAMERIA'
  const handleDownloadZIPFalegnameria = async () => {
    if (processedRows.length === 0) return;
    try {
      const falegnameriaRows = processedRows.filter(r => {
        const fornitore = (r.excelRow.fornitore || '').trim().toUpperCase();
        return fornitore === 'FALEGNAMERIA';
      });

      if (falegnameriaRows.length === 0) {
        showNotification("Nessun articolo per il fornitore 'FALEGNAMERIA' trovato.", "info");
        return;
      }

      const hasMatchedFiles = falegnameriaRows.some(r => r.matchResult.status === 'Trovato e Copiato');
      if (!hasMatchedFiles) {
        showNotification("Nessun disegno trovato e copiato per il fornitore 'FALEGNAMERIA'.", "info");
        return;
      }

      setIsProcessing(true);
      setProcessProgress(0);
      setProcessStatusText("Inizializzazione archivio ZIP FALEGNAMERIA...");

      const zipBlob = await generateZIP(falegnameriaRows, (percent, file) => {
        setProcessProgress(percent);
        setProcessStatusText(`Aggiunta file: ${file}`);
      });

      setProcessProgress(100);
      setProcessStatusText("Compressione ultimata!");

      const url = URL.createObjectURL(zipBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `File_Disegni_FALEGNAMERIA.zip`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      
      setTimeout(() => {
        setIsProcessing(false);
      }, 500);
    } catch (err) {
      setIsProcessing(false);
      showNotification("Errore nella generazione dello ZIP FALEGNAMERIA: " + (err as Error).message, "error");
    }
  };

  // Export organized ZIP file for the currently selected supplier filter
  const handleDownloadZIPFiltrati = async () => {
    if (processedRows.length === 0) return;
    
    const supplierToFilter = selectedSupplierFilter;
    if (!supplierToFilter || supplierToFilter === 'ALL') {
      showNotification("Seleziona prima un fornitore specifico nel filtro a tendina dei Fornitori.", "info");
      return;
    }

    try {
      const filteredRowsList = processedRows.filter(r => {
        const fornitore = (r.excelRow.fornitore || '').trim().toUpperCase();
        return fornitore === supplierToFilter.trim().toUpperCase();
      });

      if (filteredRowsList.length === 0) {
        showNotification(`Nessun articolo trovato per il fornitore '${supplierToFilter}'.`, "info");
        return;
      }

      const hasMatchedFiles = filteredRowsList.some(r => r.matchResult.status === 'Trovato e Copiato');
      if (!hasMatchedFiles) {
        showNotification(`Nessun disegno trovato e copiato per il fornitore '${supplierToFilter}'.`, "info");
        return;
      }

      setIsProcessing(true);
      setProcessProgress(0);
      setProcessStatusText(`Inizializzazione archivio ZIP per ${supplierToFilter}...`);

      const zipBlob = await generateZIP(filteredRowsList, (percent, file) => {
        setProcessProgress(percent);
        setProcessStatusText(`Aggiunta file: ${file}`);
      });

      setProcessProgress(100);
      setProcessStatusText("Compressione ultimata!");

      const url = URL.createObjectURL(zipBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `File_Disegni_${supplierToFilter.replace(/\s+/g, '_')}.zip`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      
      setTimeout(() => {
        setIsProcessing(false);
      }, 500);
    } catch (err) {
      setIsProcessing(false);
      showNotification(`Errore nella generazione dello ZIP per ${supplierToFilter}: ` + (err as Error).message, "error");
    }
  };

  // Reset App
  const handleReset = () => {
    setExcelFile(null);
    setWorkbook(null);
    setSheets([]);
    setSelectedSheet('');
    setRawExcelData([]);
    setProjectFiles([]);
    setProjectFolderName('');
    setColumns([]);
    setMapping({
      codice: '',
      configurazione: '',
      revisione: '',
      fornitore: ''
    });
    setProcessedRows([]);
    setStats(null);
    setSearchTerm('');
    setStatusFilter('ALL');
    setSelectedSupplierFilter('ALL');
    setSelectedExtensionFilter('ALL');
    setBa1StpMissingOnly(false);
  };

  // --- Filtering Logic ---
  const suppliersList: string[] = Array.from(new Set(processedRows.map(r => r.excelRow.fornitore || 'Senza_Fornitore'))).filter(Boolean) as string[];
  
  const filteredRows = processedRows.filter(row => {
    // 1. Search term (matches codice, configurazione, fornitore)
    const matchSearch = searchTerm === '' || 
      row.excelRow.codice.toLowerCase().includes(searchTerm.toLowerCase()) ||
      row.excelRow.configurazione.toLowerCase().includes(searchTerm.toLowerCase()) ||
      row.excelRow.fornitore.toLowerCase().includes(searchTerm.toLowerCase()) ||
      row.matchResult.targetBaseName.toLowerCase().includes(searchTerm.toLowerCase());

    // 2. Status filter
    let matchStatus = true;
    if (statusFilter === 'FOUND') {
      matchStatus = row.matchResult.status === 'Trovato e Copiato';
    } else if (statusFilter === 'MISSING') {
      matchStatus = row.matchResult.status === 'Mancante';
    }

    // 3. Supplier filter
    const matchSupplier = selectedSupplierFilter === 'ALL' || 
      (row.excelRow.fornitore || 'Senza_Fornitore') === selectedSupplierFilter;

    // 4. Extension filter
    let matchExt = true;
    if (selectedExtensionFilter === 'PDF') {
      matchExt = row.matchResult.pdfMatched;
    } else if (selectedExtensionFilter === 'DWG') {
      matchExt = row.matchResult.dwgMatched;
    } else if (selectedExtensionFilter === 'STP') {
      matchExt = row.matchResult.stpMatched;
    }

    // 5. BA1 without STP filter
    let matchBa1StpMissing = true;
    if (ba1StpMissingOnly) {
      const isBa1 = row.excelRow.codice && row.excelRow.codice.toUpperCase().startsWith("BA1");
      const hasStp = row.matchResult.stpMatched || row.matchResult.stpFiles.length > 0;
      matchBa1StpMissing = isBa1 && !hasStp;
    }

    return matchSearch && matchStatus && matchSupplier && matchExt && matchBa1StpMissing;
  });

  const hasActiveFilters = searchTerm !== '' || statusFilter !== 'ALL' || selectedSupplierFilter !== 'ALL' || selectedExtensionFilter !== 'ALL' || ba1StpMissingOnly;

  return (
    <div className="flex flex-col h-screen overflow-hidden font-sans bg-[#f3f4f6] relative" id="main_app_layout">
      
      {/* Floating Notifications / Toast Banner */}
      {toast && (
        <div className="fixed top-4 right-4 z-[9999] max-w-sm w-full bg-white rounded-lg border shadow-2xl p-4 animate-in fade-in slide-in-from-top-4 duration-300 flex items-start gap-3" style={{ borderColor: toast.type === 'error' ? '#fecaca' : toast.type === 'success' ? '#bbf7d0' : '#bfdbfe' }}>
          <div className="flex-1">
            <h4 className="text-xs font-bold uppercase tracking-wider" style={{ color: toast.type === 'error' ? '#991b1b' : toast.type === 'success' ? '#166534' : '#1e40af' }}>
              {toast.type === 'error' ? 'Attenzione / Errore' : toast.type === 'success' ? 'Operazione Completata' : 'Informazione'}
            </h4>
            <p className="text-slate-600 text-[11px] mt-1 leading-relaxed font-sans font-medium">{toast.message}</p>
          </div>
          <button onClick={() => setToast(null)} className="text-slate-400 hover:text-slate-600 transition-colors p-1 cursor-pointer">
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      )}

      {/* Testata / Header (High Density Theme) */}
      <header className="h-14 bg-[#1e293b] text-white flex items-center justify-between px-6 shrink-0 shadow-md border-b border-slate-700" id="app_header">
        <div className="flex items-center gap-2.5">
          <FolderSync className="w-5 h-5 text-blue-400 animate-pulse-subtle" />
          <h1 className="text-sm font-bold tracking-tight uppercase">
            TBM System <span className="text-blue-400 font-normal">Gestione Preparazione Disegni Tecnici per Fornitori</span>
          </h1>
        </div>
        
        <div className="flex items-center gap-3">
          <div className="bg-slate-800 text-[10px] font-sans border border-slate-700 px-3 py-1 rounded text-slate-300 flex items-center gap-1.5">
            <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full animate-ping"></span>
            Stato: <span className="text-emerald-400 font-bold">Ready</span>
          </div>
          <div className="text-slate-400 uppercase tracking-widest text-[9px] font-semibold hidden sm:block">
            Rilascio: v2.4.0-Stabile
          </div>
        </div>
      </header>

      {/* Main Layout: Left Sidebar + Right Main Preview */}
      <main className="flex flex-1 overflow-hidden">
        
        {/* Left Sidebar: Controls & Uploads */}
        <aside className="w-80 bg-white border-r border-slate-200 p-4 flex flex-col gap-4 shrink-0 overflow-y-auto">
          
          {/* Section 1: Excel Upload */}
          <div className="space-y-2">
            <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">
              1. Carica Distinta Base (XLSX)
            </label>
            
            <div 
              className={`border-2 border-dashed rounded-lg p-3 text-center cursor-pointer transition-colors ${
                excelFile ? 'border-emerald-200 bg-emerald-50/40' : 'border-slate-200 hover:border-blue-400 bg-slate-50/50'
              }`} 
              onClick={() => excelInputRef.current?.click()}
            >
              <input 
                type="file" 
                ref={excelInputRef} 
                onChange={handleExcelUpload} 
                accept=".xlsx, .xls" 
                className="hidden" 
              />
              <FileSpreadsheet className={`w-6 h-6 mx-auto mb-1 ${excelFile ? 'text-emerald-500' : 'text-slate-400'}`} />
              {excelFile ? (
                <div>
                  <p className="text-xs font-semibold text-emerald-800 truncate px-2" title={excelFile.name}>
                    {excelFile.name}
                  </p>
                  <p className="text-[9px] text-emerald-600 font-medium">Excel pronto per l'analisi</p>
                </div>
              ) : (
                <div>
                  <p className="text-xs font-semibold text-slate-700">Trascina o clicca per caricare</p>
                  <p className="text-[9px] text-slate-400">Foglio excel della distinta (.xlsx)</p>
                </div>
              )}
            </div>

            {/* Foglio e Mappatura */}
            {sheets.length > 0 && (
              <div className="space-y-1.5 mt-2 bg-slate-50 p-2.5 rounded-lg border border-slate-150">
                <div className="flex items-center justify-between gap-2">
                  <span className="text-[9px] font-bold text-slate-500 uppercase">Foglio Excel:</span>
                  <select 
                    value={selectedSheet} 
                    onChange={handleSheetChange} 
                    className="bg-white border border-slate-200 text-[11px] rounded px-1.5 py-0.5 font-medium outline-hidden"
                  >
                    {sheets.map(sheet => (
                      <option key={sheet} value={sheet}>{sheet}</option>
                    ))}
                  </select>
                </div>

                <button 
                  onClick={() => setShowMappingConfig(!showMappingConfig)} 
                  className="w-full flex items-center justify-between text-[10px] font-bold text-slate-600 bg-white hover:bg-slate-100 px-2 py-1 rounded border border-slate-200 mt-2 cursor-pointer transition-colors"
                >
                  <span className="flex items-center gap-1">
                    <Settings className="w-3 h-3 text-slate-400" /> 
                    MAPPATURA COLONNE ({columns.length})
                  </span>
                  <ChevronRight className={`w-3 h-3 text-slate-400 transition-transform ${showMappingConfig ? 'rotate-90' : ''}`} />
                </button>

                {showMappingConfig && columns.length > 0 && (
                  <div className="space-y-2 mt-2 pt-2 border-t border-slate-200 text-[10px]">
                    <div>
                      <label className="block text-slate-500 font-bold uppercase tracking-wide mb-0.5">Codice Articolo <span className="text-rose-500">*</span></label>
                      <select 
                        value={mapping.codice} 
                        onChange={(e) => setMapping({...mapping, codice: e.target.value})} 
                        className="w-full bg-white border border-slate-200 rounded px-1.5 py-1 text-xs outline-hidden"
                      >
                        <option value="">-- Seleziona --</option>
                        {columns.map(col => <option key={col} value={col}>{col}</option>)}
                      </select>
                    </div>
                    
                    <div>
                      <label className="block text-slate-500 font-bold uppercase tracking-wide mb-0.5">Configurazione</label>
                      <select 
                        value={mapping.configurazione} 
                        onChange={(e) => setMapping({...mapping, configurazione: e.target.value})} 
                        className="w-full bg-white border border-slate-200 rounded px-1.5 py-1 text-xs outline-hidden"
                      >
                        <option value="">-- Nessuna (Solo Codice) --</option>
                        {columns.map(col => <option key={col} value={col}>{col}</option>)}
                      </select>
                    </div>

                    <div>
                      <label className="block text-slate-500 font-bold uppercase tracking-wide mb-0.5">Revisione</label>
                      <select 
                        value={mapping.revisione} 
                        onChange={(e) => setMapping({...mapping, revisione: e.target.value})} 
                        className="w-full bg-white border border-slate-200 rounded px-1.5 py-1 text-xs outline-hidden"
                      >
                        <option value="">-- Nessuna --</option>
                        {columns.map(col => <option key={col} value={col}>{col}</option>)}
                      </select>
                    </div>

                    <div>
                      <label className="block text-slate-500 font-bold uppercase tracking-wide mb-0.5">Fornitore (Copia)</label>
                      <select 
                        value={mapping.fornitore} 
                        onChange={(e) => setMapping({...mapping, fornitore: e.target.value})} 
                        className="w-full bg-white border border-slate-200 rounded px-1.5 py-1 text-xs outline-hidden"
                      >
                        <option value="">-- Nessuno --</option>
                        {columns.map(col => <option key={col} value={col}>{col}</option>)}
                      </select>
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>

          {/* Section 2: Folder Picker */}
          <div className="space-y-2.5 pt-2.5 border-t border-slate-100">
            <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">
              2. Seleziona Disegni / Cartella
            </label>
            
            {/* Split controls to give popup-free alternative */}
            <div className="grid grid-cols-2 gap-1.5 p-1 bg-slate-100 rounded-lg">
              <button
                type="button"
                onClick={() => filesInputRef.current?.click()}
                className="bg-white hover:bg-slate-50 text-slate-700 font-bold text-[9px] py-1.5 px-1 rounded border border-slate-200 shadow-xs flex flex-col items-center justify-center gap-1 cursor-pointer transition-all active:scale-95 text-center leading-tight"
                title="Consente di selezionare multipli file disegni. EVITA IL POPUP DI SICUREZZA."
              >
                <Files className="w-3.5 h-3.5 text-blue-500" />
                <span>Scegli File<br/><span className="text-[7px] text-emerald-600 font-medium font-mono">(Senza Popup)</span></span>
              </button>
              <button
                type="button"
                onClick={() => folderInputRef.current?.click()}
                className="bg-white hover:bg-slate-50 text-slate-700 font-bold text-[9px] py-1.5 px-1 rounded border border-slate-200 shadow-xs flex flex-col items-center justify-center gap-1 cursor-pointer transition-all active:scale-95 text-center leading-tight"
                title="Consente di selezionare un'intera cartella locale (Mostra popup browser)."
              >
                <FolderOpen className="w-3.5 h-3.5 text-indigo-500" />
                <span>Carica Cartella<br/><span className="text-[7px] text-slate-400 font-medium font-mono">(Popup Browser)</span></span>
              </button>
            </div>

            {/* Hidden native input elements */}
            <input 
              type="file" 
              ref={folderInputRef} 
              onChange={handleFolderUpload} 
              webkitdirectory="true" 
              directory="true" 
              multiple 
              className="hidden" 
            />
            <input 
              type="file" 
              ref={filesInputRef} 
              onChange={handleFilesUpload} 
              multiple 
              className="hidden" 
            />

            {/* Status container */}
            <div className={`border rounded-lg p-2.5 text-center transition-colors ${
              projectFiles.length > 0 ? 'border-indigo-200 bg-indigo-50/30' : 'border-slate-200 bg-slate-50/20'
            }`}>
              {projectFiles.length > 0 ? (
                <div className="flex items-center gap-2 text-left">
                  <div className="p-1.5 bg-indigo-100 rounded-lg text-indigo-600 shrink-0">
                    <FolderOpen className="w-4 h-4" />
                  </div>
                  <div className="min-w-0 flex-1">
                    <p className="text-[11px] font-bold text-slate-700 truncate" title={projectFolderName}>
                      {projectFolderName}
                    </p>
                    <p className="text-[9px] text-indigo-600 font-semibold">{projectFiles.length} file disegni pronti</p>
                  </div>
                </div>
              ) : (
                <div className="py-1 text-slate-400">
                  <p className="text-[11px] font-semibold text-slate-500">Nessun disegno caricato</p>
                  <p className="text-[9px] text-slate-400">Usa uno dei due metodi sopra per caricare i file dei disegni (.pdf, .dwg, .stp/.step)</p>
                </div>
              )}
            </div>
          </div>

          {/* Section 3: Azioni di Elaborazione */}
          <div className="space-y-2 mt-auto pt-4 border-t border-slate-100 shrink-0">
            <button 
              onClick={handleProcessSearch} 
              disabled={isProcessing} 
              className="w-full bg-blue-600 hover:bg-blue-700 disabled:bg-slate-300 text-white font-bold py-2.5 px-4 rounded-lg shadow-md flex items-center justify-center gap-2 transition-all transform active:scale-[0.98] cursor-pointer text-xs uppercase tracking-wider"
            >
              {isProcessing ? (
                <RefreshCw className="w-3.5 h-3.5 animate-spin" />
              ) : (
                <Play className="w-3.5 h-3.5" />
              )}
              ELABORA RICERCA
            </button>
            
            {(excelFile || projectFolderName || processedRows.length > 0) && (
              <button 
                onClick={handleReset} 
                className="w-full bg-slate-100 hover:bg-slate-200 text-slate-600 font-medium py-1.5 px-4 rounded-lg border border-slate-200 flex items-center justify-center gap-1.5 transition-all cursor-pointer text-[11px]"
              >
                <RefreshCw className="w-3 h-3 text-slate-400" />
                Svuota Sessione
              </button>
            )}
          </div>
        </aside>

        {/* Right Preview & Results Section */}
        <section className="flex-1 flex flex-col p-5 overflow-hidden gap-4">
          
          {/* Information & Sandbox warning bar */}
          <div className="bg-slate-100 border-l-4 border-slate-500 p-3 rounded-r-lg text-[11px] text-slate-600 leading-normal shrink-0">
            <div className="flex items-start gap-2">
              <Info className="w-4 h-4 text-slate-500 shrink-0 mt-0.5" />
              <div>
                <p className="font-bold text-slate-800">Nota sulle Limitazioni del Browser Sandbox</p>
                <p>I browser non possono scrivere file direttamente sul disco locale. Questo applicativo organizza i disegni tecnici trovati (PDF, DWG, STEP) smistandoli in cartelle corrispondenti ai Fornitori, per poi generare un file ZIP finale pronto da estrarre nella directory desiderata.</p>
              </div>
            </div>
          </div>

          {processedRows.length === 0 ? (
            /* Empty State Container */
            <div className="bg-white rounded-xl shadow-sm border border-slate-200 flex-1 flex flex-col items-center justify-center p-8 text-center" id="empty_state_panel">
              <div className="w-16 h-16 bg-blue-50 text-blue-500 rounded-full flex items-center justify-center mb-4 border border-blue-100 animate-pulse-subtle">
                <FolderSync className="w-8 h-8" />
              </div>
              <h3 className="text-sm font-bold text-slate-700 uppercase tracking-widest">Nessun file elaborato</h3>
              <p className="text-xs text-slate-400 max-w-md mt-1.5 leading-relaxed">
                Carica il file Excel della distinta ed indica la cartella dei disegni dal pannello di sinistra per elaborare la ricerca dei disegni e smistarli per fornitore.
              </p>
            </div>
          ) : (
            /* Real Data & Table Dashboard */
            <>
              {/* Stats Counters Grid (High Density Layout) */}
              {stats && (
                <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 shrink-0" id="stats_panel">
                  
                  {/* Total Rows */}
                  <div className="bg-white p-3 rounded-lg border border-slate-200 shadow-xs flex items-center justify-between">
                    <div>
                      <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">Articoli Distinta</p>
                      <h3 className="text-lg font-bold text-slate-800 mt-0.5">{stats.totalRows}</h3>
                    </div>
                    <div className="p-1.5 bg-blue-50 text-blue-600 rounded">
                      <FileSpreadsheet className="w-4 h-4" />
                    </div>
                  </div>

                  {/* BA1 with missing STP Rows */}
                  <div 
                    onClick={() => setShowBa1StpModal(true)}
                    className={`p-3 rounded-lg border shadow-xs flex items-center justify-between cursor-pointer select-none transition-all ${
                      ba1StpMissingOnly 
                        ? 'bg-amber-50 border-amber-500 ring-2 ring-amber-500/20' 
                        : 'bg-white border-slate-200 hover:border-amber-400 hover:bg-amber-50/10'
                    }`}
                    title="Clicca per aprire la segnalazione BA1 senza STP/STEP"
                  >
                    <div>
                      <div className="flex items-center gap-1.5">
                        <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">BA1 senza STP/STEP</p>
                        {ba1StpMissingOnly && (
                          <span className="w-1.5 h-1.5 bg-amber-500 rounded-full animate-ping"></span>
                        )}
                      </div>
                      <h3 className="text-lg font-bold text-amber-700 mt-0.5">
                        {stats.ba1StpMissingCount} <span className="text-[10px] font-normal text-slate-400">articoli</span>
                      </h3>
                    </div>
                    <div className={`p-1.5 rounded transition-colors ${
                      ba1StpMissingOnly ? 'bg-amber-500 text-white' : 'bg-amber-50 text-amber-600'
                    }`}>
                      <AlertTriangle className="w-4 h-4" />
                    </div>
                  </div>

                  {/* Missing Rows */}
                  <div className="bg-white p-3 rounded-lg border border-slate-200 shadow-xs flex items-center justify-between">
                    <div>
                      <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">Articoli Mancanti</p>
                      <h3 className="text-lg font-bold text-rose-700 mt-0.5">
                        {stats.missingRows} <span className="text-[10px] font-normal text-slate-400">({Math.round((stats.missingRows / stats.totalRows) * 100)}%)</span>
                      </h3>
                    </div>
                    <div className="p-1.5 bg-rose-50 text-rose-600 rounded">
                      <XCircle className="w-4 h-4" />
                    </div>
                  </div>

                  {/* Total Copied Files details */}
                  <div className="bg-white p-3 rounded-lg border border-slate-200 shadow-xs flex items-center justify-between">
                    <div>
                      <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">Disegni Trovati</p>
                      <h3 className="text-lg font-bold text-indigo-700 mt-0.5">{stats.totalFilesFound}</h3>
                      <p className="text-[8px] text-slate-400 font-sans mt-0.5">
                        PDF:{stats.pdfFound} | DWG:{stats.dwgFound} | STP:{stats.stpFound}
                      </p>
                    </div>
                    <div className="p-1.5 bg-indigo-50 text-indigo-600 rounded">
                      <Database className="w-4 h-4" />
                    </div>
                  </div>

                </div>
              )}

              {/* Table frame card */}
              <div className="bg-white rounded-xl shadow-sm border border-slate-200 flex flex-col flex-1 overflow-hidden" id="results_table_container">
                
                {/* Header Tabella & Filtri & Azioni */}
                <div className="p-3 border-b border-slate-150 flex flex-wrap items-center justify-between gap-3 bg-slate-50/50 shrink-0">
                  
                  {/* Filters block on the left */}
                  <div className="flex flex-wrap items-center gap-2">
                    
                    {/* Text Search */}
                    <div className="relative">
                      <input 
                        type="text" 
                        placeholder="Filtra per codice..." 
                        value={searchTerm} 
                        onChange={(e) => setSearchTerm(e.target.value)} 
                        className="bg-white border border-slate-200 rounded text-[11px] py-1 pl-6 pr-2 w-36 outline-hidden font-medium text-slate-700 focus:border-blue-400 transition-colors" 
                      />
                      <Search className="w-3 h-3 text-slate-400 absolute left-2 top-2" />
                    </div>
                    
                    {/* Status filter tabs */}
                    <div className="flex border border-slate-200 rounded overflow-hidden text-[10px] bg-white">
                      <button 
                        onClick={() => setStatusFilter('ALL')} 
                        className={`px-2 py-1 font-bold transition-all cursor-pointer ${
                          statusFilter === 'ALL' ? 'bg-[#1e293b] text-white' : 'text-slate-500 hover:text-slate-800'
                        }`}
                      >
                        TUTTI ({processedRows.length})
                      </button>
                      <button 
                        onClick={() => setStatusFilter('FOUND')} 
                        className={`px-2 py-1 font-bold border-l border-slate-200 transition-all cursor-pointer ${
                          statusFilter === 'FOUND' ? 'bg-[#1e293b] text-emerald-400' : 'text-slate-500 hover:text-emerald-600'
                        }`}
                      >
                        TROVATI ({stats?.matchedRows})
                      </button>
                      <button 
                        onClick={() => setStatusFilter('MISSING')} 
                        className={`px-2 py-1 font-bold border-l border-slate-200 transition-all cursor-pointer ${
                          statusFilter === 'MISSING' ? 'bg-[#1e293b] text-rose-400' : 'text-slate-500 hover:text-rose-600'
                        }`}
                      >
                        MANCANTI ({stats?.missingRows})
                      </button>
                    </div>

                    {/* Suppliers dropdown */}
                    <select 
                      value={selectedSupplierFilter} 
                      onChange={(e) => setSelectedSupplierFilter(e.target.value)} 
                      className="bg-white border border-slate-200 text-[10px] rounded px-1.5 py-1 outline-hidden font-bold text-slate-600 cursor-pointer"
                    >
                      <option value="ALL">FORNITORE: TUTTI ({suppliersList.length})</option>
                      {suppliersList.map(s => (
                        <option key={s} value={s}>{s.toUpperCase()}</option>
                      ))}
                    </select>

                    {/* Extension dropdown */}
                    <select 
                      value={selectedExtensionFilter} 
                      onChange={(e) => setSelectedExtensionFilter(e.target.value)} 
                      className="bg-white border border-slate-200 text-[10px] rounded px-1.5 py-1 outline-hidden font-bold text-slate-600 cursor-pointer"
                    >
                      <option value="ALL">ESTENSIONI: QUALSIASI</option>
                      <option value="PDF">CON FILE PDF</option>
                      <option value="DWG">CON FILE DWG</option>
                      <option value="STP">CON FILE STP/STEP</option>
                    </select>

                    {/* Clear Filters Button */}
                    {hasActiveFilters && (
                      <button
                        onClick={() => {
                          setSearchTerm('');
                          setStatusFilter('ALL');
                          setSelectedSupplierFilter('ALL');
                          setSelectedExtensionFilter('ALL');
                          setBa1StpMissingOnly(false);
                        }}
                        className="bg-rose-50 hover:bg-rose-100 text-rose-600 border border-rose-200 text-[10px] rounded px-2.5 py-1.5 font-bold flex items-center gap-1 cursor-pointer transition-all active:scale-[0.97]"
                        title="Resetta tutti i filtri di ricerca applicati"
                      >
                        <X className="w-3.5 h-3.5" /> Annulla Filtri
                      </button>
                    )}

                  </div>

                  {/* Actions buttons on the right */}
                  <div className="flex items-center gap-1.5">
                    <button 
                      onClick={handleDownloadExcel} 
                      className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-[10px] py-1.5 px-3 rounded flex items-center gap-1 shadow-xs transition-all active:scale-[0.97] cursor-pointer uppercase tracking-wider"
                      title="Scarica Excel aggiornato con l'esito dei match"
                    >
                      <Download className="w-3.5 h-3.5" /> XLSX RISULTATI
                    </button>
                    <button 
                      onClick={handleDownloadZIP} 
                      className="bg-blue-600 hover:bg-blue-700 text-white font-bold text-[10px] py-1.5 px-3 rounded flex items-center gap-1 shadow-xs transition-all active:scale-[0.97] cursor-pointer uppercase tracking-wider"
                      title="Esporta archivio ZIP strutturato per Fornitore"
                    >
                      <FileArchive className="w-3.5 h-3.5" /> ESPORTA ZIP
                    </button>
                    <button 
                      onClick={handleDownloadZIPFalegnameria} 
                      className="bg-amber-600 hover:bg-amber-700 text-white font-bold text-[10px] py-1.5 px-3 rounded flex items-center gap-1 shadow-xs transition-all active:scale-[0.97] cursor-pointer uppercase tracking-wider"
                      title="Esporta archivio ZIP contenente solo i file del fornitore FALEGNAMERIA"
                    >
                      <FileArchive className="w-3.5 h-3.5" /> ZIP FALEGNAMERIA
                    </button>
                    <button 
                      onClick={handleDownloadZIPFiltrati} 
                      disabled={selectedSupplierFilter === 'ALL'}
                      className="bg-indigo-600 hover:bg-indigo-700 disabled:opacity-50 disabled:cursor-not-allowed text-white font-bold text-[10px] py-1.5 px-3 rounded flex items-center gap-1 shadow-xs transition-all active:scale-[0.97] cursor-pointer uppercase tracking-wider"
                      title={selectedSupplierFilter === 'ALL' ? "Filtra per un Fornitore per scaricare lo ZIP filtrato" : `Esporta archivio ZIP per il fornitore selezionato: ${selectedSupplierFilter}`}
                    >
                      <FileArchive className="w-3.5 h-3.5" /> ZIP FILTRATI
                    </button>
                  </div>

                </div>

                {/* Table Area */}
                <div className="flex-1 overflow-auto">
                  <table className="w-full text-left border-collapse">
                    <thead className="bg-slate-50 sticky top-0 z-10 border-b border-slate-200 text-[9px] text-slate-500 uppercase tracking-widest font-bold">
                      <tr>
                        <th className="p-2 w-10 text-center border-r border-slate-100">Rig.</th>
                        <th className="p-2 border-r border-slate-100">Codice Articolo</th>
                        <th className="p-2 border-r border-slate-100">Configurazione</th>
                        <th className="p-2 w-10 text-center border-r border-slate-100">Rev.</th>
                        <th className="p-2 border-r border-slate-100">Nome File Ricercato</th>
                        <th className="p-2 border-r border-slate-100">Stato Matching</th>
                        <th className="p-2 text-center border-r border-slate-100">Tipi Trovati</th>
                        <th className="p-2">Destinazione Fornitore</th>
                        <th className="p-2 w-10 text-center">Dett.</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 text-[11px] font-sans text-slate-600">
                      {filteredRows.length === 0 ? (
                        <tr>
                          <td colSpan={9} className="py-10 text-center text-slate-400 font-medium">
                            Nessuna riga della distinta base corrisponde ai filtri impostati.
                          </td>
                        </tr>
                      ) : (
                        filteredRows.map((row) => {
                          const res = row.matchResult;
                          const excel = row.excelRow;
                          return (
                            <tr 
                              key={row.id} 
                              className="hover:bg-slate-50 transition-colors"
                            >
                              <td className="p-2 text-center text-slate-400 border-r border-slate-100 font-bold">
                                {excel.__rowNum__}
                              </td>
                              <td className="p-2 font-bold text-slate-900 border-r border-slate-100">
                                {excel.codice}
                              </td>
                              <td className="p-2 text-slate-600 border-r border-slate-100">
                                {excel.configurazione || <span className="text-slate-300">—</span>}
                              </td>
                              <td className="p-2 text-center border-r border-slate-100">
                                {excel.revisione ? (
                                  <span className="bg-blue-50 text-blue-700 font-semibold px-1 rounded text-[9px] border border-blue-200">
                                    {excel.revisione}
                                  </span>
                                ) : (
                                  <span className="text-slate-300">—</span>
                                )}
                              </td>
                              <td className="p-2 border-r border-slate-100">
                                <div className="flex items-center justify-between gap-1.5">
                                  <span className="bg-slate-100 border border-slate-200 px-1 py-0.5 rounded text-[10px] font-semibold text-slate-700">
                                    {res.targetBaseName}
                                  </span>
                                  <button
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      navigator.clipboard.writeText(res.targetBaseName);
                                      setCopiedRowId(row.id);
                                      setTimeout(() => setCopiedRowId(null), 1500);
                                    }}
                                    className="p-1 hover:bg-slate-200 text-slate-400 hover:text-slate-600 rounded transition-all active:scale-90 cursor-pointer flex items-center justify-center shrink-0"
                                    title="Copia codice completo del file"
                                  >
                                    {copiedRowId === row.id ? (
                                      <Check className="w-3.5 h-3.5 text-emerald-600" />
                                    ) : (
                                      <Copy className="w-3.5 h-3.5" />
                                    )}
                                  </button>
                                </div>
                              </td>
                              <td className="p-2 border-r border-slate-100">
                                {res.status === 'Trovato e Copiato' ? (
                                  <span className="bg-emerald-100 text-emerald-800 px-1.5 py-0.5 rounded text-[9px] font-bold uppercase">
                                    TROVATO
                                  </span>
                                ) : (
                                  <span className="bg-rose-100 text-rose-800 px-1.5 py-0.5 rounded text-[9px] font-bold uppercase">
                                    MANCANTE
                                  </span>
                                )}
                              </td>
                              <td className="p-2 border-r border-slate-100">
                                <div className="flex items-center justify-center gap-1">
                                  {/* PDF */}
                                  <span 
                                    className={`px-1 rounded text-[8px] font-bold uppercase ${
                                      res.pdfMatched 
                                        ? 'bg-emerald-500 text-white' 
                                        : 'bg-slate-100 text-slate-300'
                                    }`}
                                    title={res.pdfMatched ? `PDF Trovati (${res.pdfFiles.length}): ${res.pdfFiles.map(f => f.fullName).join(', ')}` : 'PDF Mancante'}
                                  >
                                    PDF
                                  </span>
                                  {/* DWG */}
                                  <span 
                                    className={`px-1 rounded text-[8px] font-bold uppercase ${
                                      res.dwgMatched 
                                        ? 'bg-sky-500 text-white' 
                                        : 'bg-slate-100 text-slate-300'
                                    }`}
                                    title={res.dwgMatched ? `DWG Trovati (${res.dwgFiles.length}): ${res.dwgFiles.map(f => f.fullName).join(', ')}` : 'DWG Mancante'}
                                  >
                                    DWG
                                  </span>
                                  {/* STP */}
                                  <span 
                                    className={`px-1 rounded text-[8px] font-bold uppercase ${
                                      res.stpMatched 
                                        ? 'bg-amber-500 text-white' 
                                        : 'bg-slate-100 text-slate-300'
                                    }`}
                                    title={res.stpMatched ? `STP Trovati (${res.stpFiles.length}): ${res.stpFiles.map(f => f.fullName).join(', ')}` : 'STP Mancante'}
                                  >
                                    STP
                                  </span>
                                </div>
                              </td>
                              <td className="p-2 text-slate-700 truncate max-w-[140px] font-semibold">
                                {excel.fornitore || <span className="text-slate-400 italic">Senza Fornitore</span>}
                              </td>
                              <td className="p-2 text-center">
                                <button
                                  onClick={() => setSelectedDetailRow(row)}
                                  className="p-1 text-slate-400 hover:text-blue-600 hover:bg-slate-100 rounded transition-colors cursor-pointer"
                                  title="Visualizza Analisi Dettagliata"
                                >
                                  <Eye className="w-3.5 h-3.5 mx-auto" />
                                </button>
                              </td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>

                {/* Table Footer Stats & Status Bar */}
                <div className="p-3 bg-slate-50 border-t border-slate-200 flex flex-col gap-2 shrink-0 text-[10px] font-bold">
                  {isProcessing ? (
                    <>
                      <div className="flex justify-between items-center text-slate-500 uppercase tracking-widest">
                        <span>Generazione file ed elaborazione in corso...</span>
                        <span>{processProgress}%</span>
                      </div>
                      <div className="w-full h-2.5 bg-slate-200 rounded-full overflow-hidden shadow-inner border border-slate-300">
                        <div 
                          className="h-full bg-gradient-to-r from-blue-500 to-blue-400 shadow-lg relative transition-all duration-300" 
                          style={{ width: `${processProgress}%` }}
                        ></div>
                      </div>
                      <div className="flex justify-between text-[9px] text-slate-400 italic font-normal tracking-wide">
                        <span>{processStatusText}</span>
                      </div>
                    </>
                  ) : (
                    <div className="flex justify-between items-center text-slate-400 font-normal">
                      <span>Visualizzati <b>{filteredRows.length}</b> di <b>{processedRows.length}</b> articoli</span>
                      <span className="font-sans text-[9px] uppercase tracking-wider bg-white border border-slate-200 px-2 py-0.5 rounded text-slate-500 font-bold">
                        Cartella Disegni: {projectFolderName} ({projectFiles.length} file)
                      </span>
                    </div>
                  )}
                </div>

              </div>
            </>
          )}

        </section>
      </main>

      {/* Detail Modal (High Density Style) */}
      {selectedDetailRow && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50" id="detail_modal_container">
          <div className="bg-white rounded-xl border border-slate-200 max-w-5xl w-full shadow-2xl overflow-hidden flex flex-col max-h-[90vh] font-sans">
            
            {/* Modal Header */}
            <div className="bg-slate-50 px-4 py-3 border-b border-slate-200 flex items-center justify-between shrink-0">
              <div className="flex items-center gap-2">
                <Info className="w-4 h-4 text-blue-600" />
                <h3 className="font-bold text-slate-900 text-xs uppercase tracking-wider">
                  Analisi Riga Excel {selectedDetailRow.excelRow.__rowNum__}
                </h3>
              </div>
              <button 
                onClick={() => setSelectedDetailRow(null)}
                className="text-slate-400 hover:text-slate-600 transition-colors p-1 cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Modal Body (Widescreen 3 Columns on desktop) */}
            <div className="p-4 overflow-y-auto flex-1 text-xs">
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4 items-start">
                
                {/* Column 1: Informazioni principali & Regole */}
                <div className="space-y-4">
                  <div>
                    <h4 className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-2">Dati Principali</h4>
                    {/* Values grid */}
                    <div className="grid grid-cols-2 gap-2">
                      <div className="bg-slate-50 p-2 rounded border border-slate-100 font-sans">
                        <p className="text-[9px] font-bold text-slate-400 uppercase tracking-wide">Codice</p>
                        <p className="text-slate-800 font-bold mt-0.5">{selectedDetailRow.excelRow.codice}</p>
                      </div>
                      <div className="bg-slate-50 p-2 rounded border border-slate-100 font-sans">
                        <p className="text-[9px] font-bold text-slate-400 uppercase tracking-wide">Configurazione</p>
                        <p className="text-slate-700 font-semibold mt-0.5">{selectedDetailRow.excelRow.configurazione || '—'}</p>
                      </div>
                      <div className="bg-slate-50 p-2 rounded border border-slate-100 font-sans">
                        <p className="text-[9px] font-bold text-slate-400 uppercase tracking-wide">Revisione</p>
                        <p className="text-slate-700 font-semibold mt-0.5">{selectedDetailRow.excelRow.revisione || '—'}</p>
                      </div>
                      <div className="bg-slate-50 p-2 rounded border border-slate-100 font-sans">
                        <p className="text-[9px] font-bold text-slate-400 uppercase tracking-wide">Fornitore</p>
                        <p className="text-slate-800 font-bold mt-0.5 text-blue-700 truncate" title={selectedDetailRow.excelRow.fornitore || 'Senza Fornitore'}>
                          {selectedDetailRow.excelRow.fornitore || 'Senza Fornitore'}
                        </p>
                      </div>
                    </div>
                  </div>

                  {/* Matching Rules details */}
                  <div className="bg-blue-50/50 rounded-lg p-3 border border-blue-100 space-y-1">
                    <h4 className="text-[10px] font-bold text-blue-900 uppercase tracking-wide">Regola Nome Applicata</h4>
                    <div className="text-[11px] text-slate-700 leading-relaxed font-sans">
                      <div className="flex justify-between mb-1">
                        <span className="text-slate-400 font-bold">Target Base Name:</span>
                        <span className="text-blue-900 font-bold">{selectedDetailRow.matchResult.targetBaseName}</span>
                      </div>
                      <div className="bg-white/80 p-2 rounded border border-blue-100 text-[10px] font-sans text-slate-600">
                        {selectedDetailRow.excelRow.codice === selectedDetailRow.excelRow.configurazione || !selectedDetailRow.excelRow.configurazione ? (
                          selectedDetailRow.excelRow.revisione ? (
                            <span><b>Regola 3:</b> Codice e Configurazione uguali con Revisione. Nome ricercato: <code>CodiceRevisione</code>.</span>
                          ) : (
                            <span><b>Regola 1:</b> Codice e Configurazione uguali senza Revisione. Nome ricercato: <code>Codice</code>.</span>
                          )
                        ) : (
                          selectedDetailRow.excelRow.revisione ? (
                            <span><b>Regola 4:</b> Codice e Configurazione diversi con Revisione. Nome ricercato: <code>Codice_ConfigurazioneRevisione</code>.</span>
                          ) : (
                            <span><b>Regola 2:</b> Codice e Configurazione diversi senza Revisione. Nome ricercato: <code>Codice_Configurazione</code>.</span>
                          )
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Destination directory info */}
                  {selectedDetailRow.matchResult.status === 'Trovato e Copiato' && (
                    <div className="bg-emerald-50 border border-emerald-150 rounded p-2.5 text-[11px] text-emerald-800">
                      <span className="font-bold block">Percorso di Smistamento ZIP:</span>
                      <span className="font-sans text-[10px] bg-white border border-emerald-100 px-1 py-0.5 rounded text-emerald-950 inline-block mt-1">
                        {`${selectedDetailRow.excelRow.fornitore || 'Senza_Fornitore'}/`}
                      </span>
                    </div>
                  )}
                </div>

                {/* Column 2: Physical Files matched list */}
                <div className="space-y-3">
                  <h4 className="text-[10px] font-bold text-slate-700 uppercase tracking-wider">Esito File Trovati per Estensione</h4>
                  
                  {/* PDF */}
                  <div className="flex flex-col p-2.5 rounded-lg border text-[11px] font-sans bg-slate-50 border-slate-200">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-1.5">
                        <span className="font-bold text-[8px] bg-red-100 text-red-800 px-1.5 py-0.5 rounded">PDF</span>
                        <span className="font-bold text-slate-700 uppercase text-[9px] tracking-wide">File PDF Associati</span>
                      </div>
                      {selectedDetailRow.matchResult.pdfMatched ? (
                        <span className="text-emerald-700 font-bold flex items-center gap-0.5 text-[10px]">
                          <Check className="w-3.5 h-3.5" /> TROVATO ({selectedDetailRow.matchResult.pdfFiles.length})
                        </span>
                      ) : (
                        <span className="text-slate-400 italic">NON TROVATO</span>
                      )}
                    </div>
                    {selectedDetailRow.matchResult.pdfMatched ? (
                      <div className="mt-1.5 pl-4 border-l-2 border-emerald-500 space-y-1.5 max-h-[80px] overflow-y-auto">
                        {selectedDetailRow.matchResult.pdfFiles.map((file, fIdx) => (
                          <div key={fIdx} className="text-[10px] leading-relaxed text-slate-600 break-all">
                            <span className="font-bold text-slate-900">{file.fullName}</span>
                            <span className="text-slate-400 block text-[9px]">Percorso: {file.relativePath}</span>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div className="mt-1 pl-4 border-l-2 border-slate-200 text-[10px] text-slate-400 italic">
                        Atteso: {selectedDetailRow.matchResult.targetBaseName}.pdf
                      </div>
                    )}
                  </div>

                  {/* DWG */}
                  <div className="flex flex-col p-2.5 rounded-lg border text-[11px] font-sans bg-slate-50 border-slate-200">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-1.5">
                        <span className="font-bold text-[8px] bg-sky-100 text-sky-800 px-1.5 py-0.5 rounded">DWG</span>
                        <span className="font-bold text-slate-700 uppercase text-[9px] tracking-wide">File DWG Associati</span>
                      </div>
                      {selectedDetailRow.matchResult.dwgMatched ? (
                        <span className="text-emerald-700 font-bold flex items-center gap-0.5 text-[10px]">
                          <Check className="w-3.5 h-3.5" /> TROVATO ({selectedDetailRow.matchResult.dwgFiles.length})
                        </span>
                      ) : (
                        <span className="text-slate-400 italic">NON TROVATO</span>
                      )}
                    </div>
                    {selectedDetailRow.matchResult.dwgMatched ? (
                      <div className="mt-1.5 pl-4 border-l-2 border-sky-500 space-y-1.5 max-h-[80px] overflow-y-auto">
                        {selectedDetailRow.matchResult.dwgFiles.map((file, fIdx) => (
                          <div key={fIdx} className="text-[10px] leading-relaxed text-slate-600 break-all">
                            <span className="font-bold text-slate-900">{file.fullName}</span>
                            <span className="text-slate-400 block text-[9px]">Percorso: {file.relativePath}</span>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div className="mt-1 pl-4 border-l-2 border-slate-200 text-[10px] text-slate-400 italic">
                        Atteso: {selectedDetailRow.matchResult.targetBaseName}.dwg
                      </div>
                    )}
                  </div>

                  {/* STP */}
                  <div className="flex flex-col p-2.5 rounded-lg border text-[11px] font-sans bg-slate-50 border-slate-200">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-1.5">
                        <span className="font-bold text-[8px] bg-amber-100 text-amber-800 px-1.5 py-0.5 rounded">STP/STEP</span>
                        <span className="font-bold text-slate-700 uppercase text-[9px] tracking-wide">File STP Associati</span>
                      </div>
                      {selectedDetailRow.matchResult.stpMatched ? (
                        <span className="text-emerald-700 font-bold flex items-center gap-0.5 text-[10px]">
                          <Check className="w-3.5 h-3.5" /> TROVATO ({selectedDetailRow.matchResult.stpFiles.length})
                        </span>
                      ) : (
                        <span className="text-slate-400 italic">NON TROVATO</span>
                      )}
                    </div>
                    {selectedDetailRow.matchResult.stpMatched ? (
                      <div className="mt-1.5 pl-4 border-l-2 border-amber-500 space-y-1.5 max-h-[80px] overflow-y-auto">
                        {selectedDetailRow.matchResult.stpFiles.map((file, fIdx) => (
                          <div key={fIdx} className="text-[10px] leading-relaxed text-slate-600 break-all">
                            <span className="font-bold text-slate-900">{file.fullName}</span>
                            <span className="text-slate-400 block text-[9px]">Percorso: {file.relativePath}</span>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div className="mt-1 pl-4 border-l-2 border-slate-200 text-[10px] text-slate-400 italic">
                        Atteso: {selectedDetailRow.matchResult.targetBaseName}.stp / .step
                      </div>
                    )}
                  </div>
                </div>

                {/* Column 3: Full Original Excel Row Data */}
                <div className="space-y-2 flex flex-col">
                  <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Chiavi Riga Excel Originale</h4>
                  <div className="bg-slate-50 rounded p-2.5 border border-slate-200 max-h-[480px] overflow-y-auto font-sans text-[9px] text-slate-700 divide-y divide-slate-150">
                    {Object.entries(selectedDetailRow.excelRow.originalData).map(([key, value]) => (
                      <div key={key} className="flex justify-between py-1">
                        <span className="text-slate-400 font-bold truncate pr-3">{key}:</span>
                        <span className="text-slate-900 break-all font-semibold">{String(value)}</span>
                      </div>
                    ))}
                  </div>
                </div>

              </div>
            </div>

            {/* Modal Footer */}
            <div className="bg-slate-50 px-4 py-2 border-t border-slate-200 flex justify-end shrink-0">
              <button
                onClick={() => setSelectedDetailRow(null)}
                className="bg-[#1e293b] hover:bg-[#0f172a] text-white font-bold text-xs px-4 py-1.5 rounded cursor-pointer transition-colors"
              >
                Chiudi
              </button>
            </div>

          </div>
        </div>
      )}

      {/* BA1 Without STP/STEP Modal */}
      {showBa1StpModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50" id="ba1_stp_modal_container">
          <div className="bg-white rounded-xl border border-slate-200 max-w-2xl w-full shadow-2xl overflow-hidden flex flex-col max-h-[85vh] font-sans">
            
            {/* Modal Header */}
            <div className="bg-amber-500 text-white px-4 py-3 border-b border-amber-600 flex items-center justify-between shrink-0">
              <div className="flex items-center gap-2">
                <AlertTriangle className="w-5 h-5 text-white animate-pulse" />
                <h3 className="font-bold text-white text-xs uppercase tracking-wider">
                  Articoli BA1 Senza Disegno STP/STEP
                </h3>
              </div>
              <button 
                onClick={() => setShowBa1StpModal(false)}
                className="text-white hover:text-amber-100 transition-colors p-1 cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Modal Body */}
            <div className="p-4 space-y-4 overflow-y-auto flex-1 flex flex-col min-h-0 text-xs">
              
              {/* Mandatory Action Message */}
              <div className="bg-amber-50 border border-amber-200 rounded-lg p-3.5 flex items-start gap-3">
                <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />
                <div>
                  <h4 className="font-bold text-amber-900 text-[11px] uppercase tracking-wide">Direttiva di Produzione</h4>
                  <p className="text-amber-800 text-xs font-bold mt-1 leading-relaxed">
                    Comunicare a Ufficio Tecnico la mancanza di questi file .step
                  </p>
                </div>
              </div>

              {/* Grid or List of items */}
              <div className="space-y-1.5 flex-1 flex flex-col min-h-0">
                <div className="flex items-center justify-between">
                  <h4 className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">
                    Elenco degli Articoli Interessati ({
                      processedRows.filter(row => {
                        const isBa1 = row.excelRow.codice && row.excelRow.codice.toUpperCase().startsWith("BA1");
                        const hasStp = row.matchResult.stpMatched || row.matchResult.stpFiles.length > 0;
                        return isBa1 && !hasStp;
                      }).length
                    })
                  </h4>
                </div>
                <div className="border border-slate-200 rounded-lg overflow-hidden flex-1 overflow-y-auto min-h-[150px] max-h-[300px] bg-white">
                  <table className="w-full text-left border-collapse">
                    <thead className="bg-slate-50 sticky top-0 z-10 border-b border-slate-200 text-[9px] text-slate-500 uppercase tracking-widest font-bold">
                      <tr>
                        <th className="p-2 w-10 text-center border-r border-slate-100">Rig.</th>
                        <th className="p-2 border-r border-slate-100">Nome File Ricercato</th>
                        <th className="p-2 border-r border-slate-100">Fornitore</th>
                        <th className="p-2">Nome File .stp Atteso</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 text-[11px] font-sans text-slate-600">
                      {processedRows
                        .filter(row => {
                          const isBa1 = row.excelRow.codice && row.excelRow.codice.toUpperCase().startsWith("BA1");
                          const hasStp = row.matchResult.stpMatched || row.matchResult.stpFiles.length > 0;
                          return isBa1 && !hasStp;
                        })
                        .map((row, idx) => (
                          <tr key={row.id || idx} className="hover:bg-slate-50 transition-colors bg-white">
                            <td className="p-2 text-center text-slate-400 border-r border-slate-100 font-bold">
                              {row.excelRow.__rowNum__}
                            </td>
                            <td className="p-2 text-slate-600 border-r border-slate-100 font-sans">
                              {row.matchResult.targetBaseName}
                            </td>
                            <td className="p-2 text-slate-600 border-r border-slate-100 truncate max-w-[120px]" title={row.excelRow.fornitore || 'Senza Fornitore'}>
                              {row.excelRow.fornitore || 'Senza Fornitore'}
                            </td>
                            <td className="p-2 text-amber-800 font-semibold">
                              {row.matchResult.targetBaseName}.stp
                            </td>
                          </tr>
                        ))
                      }
                      {processedRows.filter(row => {
                        const isBa1 = row.excelRow.codice && row.excelRow.codice.toUpperCase().startsWith("BA1");
                        const hasStp = row.matchResult.stpMatched || row.matchResult.stpFiles.length > 0;
                        return isBa1 && !hasStp;
                      }).length === 0 && (
                        <tr>
                          <td colSpan={4} className="p-4 text-center text-slate-400 italic">
                            Nessun articolo BA1 con disegno STP/STEP mancante riscontrato.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>

            </div>

            {/* Modal Footer */}
            <div className="bg-slate-50 px-4 py-3 border-t border-slate-200 flex flex-wrap justify-between items-center gap-2 shrink-0">
              <button
                onClick={handleExportBa1StpMissing}
                disabled={processedRows.filter(row => {
                  const isBa1 = row.excelRow.codice && row.excelRow.codice.toUpperCase().startsWith("BA1");
                  const hasStp = row.matchResult.stpMatched || row.matchResult.stpFiles.length > 0;
                  return isBa1 && !hasStp;
                }).length === 0}
                className="bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-200 disabled:text-slate-400 text-white font-bold text-xs px-4 py-2 rounded-lg cursor-pointer transition-all flex items-center gap-1.5 shadow-sm active:scale-95"
              >
                <Download className="w-4 h-4" /> Esporta Lista Excel
              </button>

              <div className="flex gap-2">
                <button
                  onClick={() => {
                    setBa1StpMissingOnly(true);
                    setShowBa1StpModal(false);
                  }}
                  disabled={processedRows.filter(row => {
                    const isBa1 = row.excelRow.codice && row.excelRow.codice.toUpperCase().startsWith("BA1");
                    const hasStp = row.matchResult.stpMatched || row.matchResult.stpFiles.length > 0;
                    return isBa1 && !hasStp;
                  }).length === 0}
                  className="bg-blue-600 hover:bg-blue-700 disabled:bg-slate-200 disabled:text-slate-400 text-white font-bold text-xs px-4 py-2 rounded-lg cursor-pointer transition-all shadow-sm active:scale-95"
                >
                  Applica Filtro Tabella
                </button>
                <button
                  onClick={() => setShowBa1StpModal(false)}
                  className="bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs px-4 py-2 rounded-lg cursor-pointer transition-colors"
                >
                  Annulla
                </button>
              </div>
            </div>

          </div>
        </div>
      )}

      {/* Premium System Footer (High Density Theme) */}
      <footer className="h-8 bg-slate-100 border-t border-slate-200 flex items-center px-4 justify-between text-[10px] text-slate-500 uppercase tracking-widest shrink-0 font-medium">
        <div>Logic Engine: Case_Based_V4</div>
        <div className="flex gap-6">
          <span>File Locali: {projectFiles.length}</span>
          <span>Sorgente: {projectFolderName || 'Nessuna'}</span>
          <span>Utente: Ufficio Tecnico</span>
        </div>
      </footer>

    </div>
  );
}
