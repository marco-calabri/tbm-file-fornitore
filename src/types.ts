export interface ExcelRow {
  __rowNum__: number;
  codice: string;
  configurazione: string;
  revisione: string;
  fornitore: string;
  originalData: Record<string, any>;
}

export interface ProjectFile {
  name: string;        // e.g. "BA005689" (without extension)
  fullName: string;    // e.g. "BA005689.pdf"
  relativePath: string; // e.g. "ProgettoA/PDF/BA005689.pdf"
  extension: string;   // e.g. ".pdf"
  fileObject: File;    // The actual File object from webkitdirectory picker
}

export interface FileMatchResult {
  pdfMatched: boolean;
  dwgMatched: boolean;
  stpMatched: boolean; // stp or step
  pdfFile?: ProjectFile;
  dwgFile?: ProjectFile;
  stpFile?: ProjectFile;
  pdfFiles: ProjectFile[];
  dwgFiles: ProjectFile[];
  stpFiles: ProjectFile[];
  targetBaseName: string;
  status: 'Trovato e Copiato' | 'Mancante';
  copiedTo?: string; // Target path/folder inside ZIP
}

export interface ProcessedRow {
  id: string;
  excelRow: ExcelRow;
  matchResult: FileMatchResult;
}

export interface ColumnMapping {
  codice: string;
  configurazione: string;
  revisione: string;
  fornitore: string;
}

export interface AppStats {
  totalRows: number;
  foundRows: number;
  matchedRows?: number; // fallback/alias
  missingRows: number;
  totalFilesFound: number;
  pdfFound: number;
  dwgFound: number;
  stpFound: number;
  ba1StpMissingCount: number;
}
