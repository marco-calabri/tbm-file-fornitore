import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { ExcelRow, ProjectFile, FileMatchResult, ProcessedRow, ColumnMapping, AppStats } from '../types';

/**
 * Computes the target file base name (without extension) according to the user's 4 rules:
 * 1. Codice == Configurazione, Revisione is empty -> Codice
 * 2. Codice != Configurazione, Revisione is empty -> Codice_Configurazione
 * 3. Codice == Configurazione, Revisione has value -> CodiceRevisione
 * 4. Codice != Configurazione, Revisione has value -> Codice_ConfigurazioneRevisione
 */
export function getTargetFileName(codice: string, configurazione: string, revisione: string): string {
  const c = (codice || '').trim();
  const cf = (configurazione || '').trim();
  const r = (revisione || '').trim();

  // If configuration is empty, we treat it as equivalent to Code to avoid "Code_Empty"
  const isSameConfig = cf === '' || c.toLowerCase() === cf.toLowerCase();

  let baseName = '';
  if (isSameConfig) {
    baseName = c;
  } else {
    baseName = `${c}_${cf}`;
  }

  if (r !== '') {
    baseName = `${baseName}${r}`;
  }

  return baseName;
}

/**
 * Parses scanned browser files into ProjectFile entities.
 * Supports recursive scanning from input[webkitdirectory]
 */
export function parseUploadedFiles(files: FileList | File[]): ProjectFile[] {
  const list: ProjectFile[] = [];
  const items = Array.from(files);

  for (const file of items) {
    // Relative path or fallback to name
    const path = file.webkitRelativePath || file.name;
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
}

/**
 * Performs matching between Excel rows and Scanned Project Files
 */
export function matchRowFiles(
  excelRows: ExcelRow[],
  projectFiles: ProjectFile[]
): ProcessedRow[] {
  // Precalculate target names and lowercase keys for all excel rows
  const rowTargets = excelRows.map((row, idx) => {
    const targetBaseName = getTargetFileName(row.codice, row.configurazione, row.revisione);
    return {
      index: idx,
      row,
      targetBaseName,
      key: targetBaseName.toLowerCase()
    };
  });

  // Map to hold matched project files for each excel row index
  const rowCandidatesMap = new Map<number, ProjectFile[]>();
  for (let i = 0; i < excelRows.length; i++) {
    rowCandidatesMap.set(i, []);
  }

  // Assign each project file to the best matching row(s) to avoid false matches
  for (const f of projectFiles) {
    const fileNameLower = f.name.toLowerCase();
    
    // Find all rows that prefix-match this file (exact match or followed by '_')
    const matches = rowTargets.filter(rt => {
      return fileNameLower === rt.key || fileNameLower.startsWith(rt.key + '_');
    });

    if (matches.length > 0) {
      // Find the maximum target name length among matching rows to handle specific configs/revisions
      const maxLength = Math.max(...matches.map(m => m.key.length));
      
      // Keep only rows with the longest match (this prevents e.g. a row without config matching a file with config)
      const bestMatches = matches.filter(m => m.key.length === maxLength);
      
      for (const bm of bestMatches) {
        rowCandidatesMap.get(bm.index)!.push(f);
      }
    }
  }

  // Now construct the final ProcessedRow objects
  return rowTargets.map((rt) => {
    const candidates = rowCandidatesMap.get(rt.index) || [];

    // Separate candidate files by extension
    const pdfCandidates = candidates.filter(f => f.extension === '.pdf');
    const dwgCandidates = candidates.filter(f => f.extension === '.dwg');
    const stpCandidates = candidates.filter(f => f.extension === '.stp' || f.extension === '.step');

    // Helper to select the best candidate (preferring exact match over description suffix match)
    const selectBest = (files: ProjectFile[], target: string) => {
      if (files.length === 0) return undefined;
      const exact = files.find(f => f.name.toLowerCase() === target.toLowerCase());
      return exact || files[0];
    };

    const pdfFile = selectBest(pdfCandidates, rt.targetBaseName);
    const dwgFile = selectBest(dwgCandidates, rt.targetBaseName);
    const stpFile = selectBest(stpCandidates, rt.targetBaseName);

    const pdfMatched = pdfCandidates.length > 0;
    const dwgMatched = dwgCandidates.length > 0;
    const stpMatched = stpCandidates.length > 0;
    const isAnyMatched = pdfMatched || dwgMatched || stpMatched;

    const matchResult: FileMatchResult = {
      pdfMatched,
      dwgMatched,
      stpMatched,
      pdfFile,
      dwgFile,
      stpFile,
      pdfFiles: pdfCandidates,
      dwgFiles: dwgCandidates,
      stpFiles: stpCandidates,
      targetBaseName: rt.targetBaseName,
      status: isAnyMatched ? 'Trovato e Copiato' : 'Mancante',
      copiedTo: isAnyMatched ? (rt.row.fornitore ? rt.row.fornitore.trim() : 'Senza_Fornitore') : undefined
    };

    return {
      id: `row-${rt.index}-${rt.row.codice}`,
      excelRow: rt.row,
      matchResult
    };
  });
}

/**
 * Computes statistics from processed rows
 */
export function computeStats(processedRows: ProcessedRow[]): AppStats {
  let foundRows = 0;
  let missingRows = 0;
  let totalFilesFound = 0;
  let pdfFound = 0;
  let dwgFound = 0;
  let stpFound = 0;
  let ba1StpMissingCount = 0;

  for (const row of processedRows) {
    const res = row.matchResult;
    if (res.status === 'Trovato e Copiato') {
      foundRows++;
    } else {
      missingRows++;
    }

    if (res.pdfMatched) {
      pdfFound += res.pdfFiles.length;
      totalFilesFound += res.pdfFiles.length;
    }
    if (res.dwgMatched) {
      dwgFound += res.dwgFiles.length;
      totalFilesFound += res.dwgFiles.length;
    }
    if (res.stpMatched) {
      stpFound += res.stpFiles.length;
      totalFilesFound += res.stpFiles.length;
    }

    // Check if starts with BA1 and doesn't have STP/STEP reference
    const isBa1 = row.excelRow.codice && row.excelRow.codice.toUpperCase().startsWith("BA1");
    const hasStp = res.stpMatched || res.stpFiles.length > 0;
    if (isBa1 && !hasStp) {
      ba1StpMissingCount++;
    }
  }

  return {
    totalRows: processedRows.length,
    foundRows,
    matchedRows: foundRows,
    missingRows,
    totalFilesFound,
    pdfFound,
    dwgFound,
    stpFound,
    ba1StpMissingCount
  };
}

/**
 * Automatically adjusts column widths in a worksheet based on cell content length
 */
export function autoFitColumns(ws: XLSX.WorkSheet) {
  if (!ws || !ws['!ref']) return;
  try {
    const range = XLSX.utils.decode_range(ws['!ref']);
    const cols: { wch: number }[] = [];
    for (let C = range.s.c; C <= range.e.c; ++C) {
      let maxWidth = 10; // minimum character width
      for (let R = range.s.r; R <= range.e.r; ++R) {
        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = ws[cellAddress];
        if (cell) {
          const value = cell.w !== undefined ? String(cell.w) : (cell.v !== undefined && cell.v !== null ? String(cell.v) : '');
          if (value.length > maxWidth) {
            maxWidth = value.length;
          }
        }
      }
      cols.push({ wch: maxWidth + 3 }); // add padding
    }
    ws['!cols'] = cols;
  } catch (e) {
    console.error("Errore nell'adattamento delle colonne:", e);
  }
}

/**
 * Updates the loaded excel workbook with a new status column at the first available position
 * and returns the binary array of the updated workbook.
 */
export function generateUpdatedExcel(
  originalWorkbook: XLSX.WorkBook,
  processedRows: ProcessedRow[],
  activeSheetName: string
): Blob {
  // We make a copy of the workbook to avoid mutating original state
  const wb = XLSX.utils.book_new();
  
  for (const sheetName of originalWorkbook.SheetNames) {
    const originalSheet = originalWorkbook.Sheets[sheetName];
    
    // If it is the active sheet where rows were processed, we update it
    if (sheetName === activeSheetName) {
      // Convert sheet to json of arrays (to keep header and structure easily)
      const rows: any[][] = XLSX.utils.sheet_to_json(originalSheet, { header: 1, defval: '' });
      
      if (rows.length > 0) {
        const headers = rows[0] as string[];
        
        // Find the first free column. User suggests Column O, let's find the actual end,
        // or if headers has fewer than 15 columns, we can pad it up to Column O (index 14)
        let statusColIdx = headers.length;
        if (statusColIdx < 14) {
          // If less than column O, let's pad empty headers up to column O
          while (headers.length < 14) {
            headers.push('');
          }
          headers.push('Stato Ricerca File'); // This will be at index 14 (Column O)
          statusColIdx = 14;
        } else {
          // If equal or larger, append at the end
          headers.push('Stato Ricerca File');
        }

        // Now map our processed results back to the row indices.
        // We need to account for headers at rows[0]. So excel row data starts at rows[1].
        // Let's match by original excel row index (`__rowNum__`).
        for (const prow of processedRows) {
          const excelRowIdx = prow.excelRow.__rowNum__; // 0-based index of row in data (with header at 0)
          
          if (excelRowIdx < rows.length) {
            // Pad the row array if it's shorter than the status column index
            while (rows[excelRowIdx].length < statusColIdx) {
              rows[excelRowIdx].push('');
            }
            rows[excelRowIdx][statusColIdx] = prow.matchResult.status;
          }
        }
      }

      // Create new sheet from the updated array representation
      const updatedSheet = XLSX.utils.aoa_to_sheet(rows);
      autoFitColumns(updatedSheet);
      XLSX.utils.book_append_sheet(wb, updatedSheet, sheetName);
    } else {
      // Copy other sheets unchanged but auto-fit them too
      const sheetCopy = { ...originalSheet };
      autoFitColumns(sheetCopy);
      XLSX.utils.book_append_sheet(wb, sheetCopy, sheetName);
    }
  }

  // Write Excel file as array buffer
  const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  return new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

/**
 * Generates a ZIP file containing the files sorted by supplier folder.
 * Shows progression through a callback.
 */
export async function generateZIP(
  processedRows: ProcessedRow[],
  onProgress?: (percent: number, currentFile: string) => void
): Promise<Blob> {
  const zip = new JSZip();
  const foundRows = processedRows.filter(r => r.matchResult.status === 'Trovato e Copiato');
  const totalFilesToZip = foundRows.reduce((acc, r) => {
    const res = r.matchResult;
    let count = 0;
    if (res.pdfMatched) count += res.pdfFiles.length;
    if (res.dwgMatched) count += res.dwgFiles.length;
    if (res.stpMatched) count += res.stpFiles.length;
    return acc + count;
  }, 0);

  if (totalFilesToZip === 0) {
    throw new Error("Nessun file trovato da copiare nello ZIP.");
  }

  let zippedCount = 0;

  for (const prow of foundRows) {
    const fornitoreName = (prow.excelRow.fornitore || 'Senza_Fornitore').trim().replace(/[\/\\?%*:|"<>\.]/g, '_');
    const res = prow.matchResult;

    const allFiles = [
      ...res.pdfFiles,
      ...res.dwgFiles,
      ...res.stpFiles
    ];

    for (const projFile of allFiles) {
      // Update progress
      zippedCount++;
      if (onProgress) {
        const percent = Math.round((zippedCount / totalFilesToZip) * 100);
        onProgress(percent, `${fornitoreName}/${projFile.fullName}`);
      }

      // Add to ZIP under "SupplierName/FileName.ext"
      const content = await projFile.fileObject.arrayBuffer();
      zip.folder(fornitoreName)!.file(projFile.fullName, content);
    }
  }

  return await zip.generateAsync({ type: 'blob' }, (metadata) => {
    // Optional: metadata.percent can be used for zip compression progress
  });
}


