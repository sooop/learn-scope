import type ExcelJS from 'exceljs';

declare global {
  interface Window {
    ExcelJS: typeof ExcelJS;
  }
}

export {};
