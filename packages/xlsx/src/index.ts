import { Workbook } from 'exceljs';

export interface InterpolateXlsxOptions {
  template: Buffer;
  data: Record<string, any>;
}

export async function interpolateXlsx(options: InterpolateXlsxOptions): Promise<Buffer> {
  const { template, data } = options;
  const workbook = new Workbook();
  await workbook.xlsx.load(template as any);

  for (const worksheet of workbook.worksheets) {
    const rowsToExpand: { rowNumber: number; arrayKey: string }[] = [];

    worksheet.eachRow((row, rowNumber) => {
      let arrayKey: string | null = null;

      row.eachCell((cell) => {
        if (typeof cell.value === 'string') {
          // Detect if there are array markers
          const arrayMatch = cell.value.match(/\[\[\s*([^\].]+)(?:\.[^\]]+)?\s*\]\]/);
          if (arrayMatch) {
            const key = arrayMatch[1];
            if (arrayKey && arrayKey !== key) {
              throw new Error(`Mixed array keys in row ${rowNumber}: ${arrayKey} vs ${key}`);
            }
            arrayKey = key;
          }
        }
      });

      if (arrayKey) {
        rowsToExpand.push({ rowNumber, arrayKey });
      }
    });

    // Process rows that must be expanded (from bottom to top to avoid index shifts)
    rowsToExpand.sort((a, b) => b.rowNumber - a.rowNumber);

    for (const { rowNumber, arrayKey } of rowsToExpand) {
      const array = data[arrayKey];
      if (array === undefined) {
        continue; // Leave markers untouched
      }
      if (!Array.isArray(array)) {
        throw new Error(`[[${arrayKey}.*]] requires '${arrayKey}' to be an array. Received: ${typeof array}`);
      }

      const originalRow = worksheet.getRow(rowNumber);

      // Remove the original row
      worksheet.spliceRows(rowNumber, 1);

      // Insert new rows
      for (let i = 0; i < array.length; i++) {
        const item = array[i];
        const newRowNumber = rowNumber + i;
        const newRow = worksheet.insertRow(newRowNumber, []);

        // Copy values and styles from the original row
        originalRow.eachCell((originalCell, colNumber) => {
          let value = originalCell.value;
          const newCell = newRow.getCell(colNumber);

          // Adjust formulas to point to the new row when they reference the template row
          if (value && typeof value === 'object' && 'formula' in value) {
            const originalFormula = (value as any).formula as string;
            const adjustedFormula = adjustFormulaForRow(originalFormula, rowNumber, newRowNumber);
            newCell.value = {
              ...(value as any),
              formula: adjustedFormula
            };
            // Styles & validation will be copied below; skip marker interpolation for formulas
          } else if (typeof value === 'string') {
            // Array interpolation: [[array.key]]
            value = value.replace(/\[\[\s*([^\].]+)\.([^\]]+)\s*\]\]/g, (_, arrKey, propPath) => {
              if (arrKey !== arrayKey) return `[[${arrKey}.${propPath}]]`; // dejar intacto
              const { found, value: resolved } = resolvePath(item, propPath);
              if (!found) return `[[${arrKey}.${propPath}]]`;
              return resolved == null ? '' : String(resolved);
            });

            // Root-level interpolation: {{key}}
            value = value.replace(/\{\{\s*([^\}]+)\s*\}\}/g, (_, path) => {
              const { found, value: resolved } = resolvePath(data, path);
              if (!found) return `{{${path}}}`;
              return resolved == null ? '' : String(resolved);
            });
          }

          if (typeof value !== 'undefined' && typeof value !== 'object') {
            newCell.value = value;
          }
          // Preserve basic styles; merges will be handled separately in a later step
          if (originalCell.style) {
            newCell.style = { ...originalCell.style };
          }
          if (originalCell.dataValidation) {
            newCell.dataValidation = { ...originalCell.dataValidation };
          }
          if (originalCell.protection) {
            newCell.protection = { ...originalCell.protection };
          }
        });
      }
    }

    // Second pass: interpolate root-level {{ }} markers in all cells
    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        if (typeof cell.value !== 'string') return;

        let value = cell.value;

        value = value.replace(/\{\{\s*([^\}]+)\s*\}\}/g, (_, path) => {
          const { found, value: resolved } = resolvePath(data, path);
          if (!found) return `{{${path}}}`;
          return resolved == null ? '' : String(resolved);
        });

        cell.value = value;
      });
    });
  }

  const result = await workbook.xlsx.writeBuffer();
  return result as any as Buffer;
}

// Reutilizar la función de core
function resolvePath(obj: any, path: string): { found: boolean; value: any } {
  const keys = path.split('.');
  let current = obj;

  for (const key of keys) {
    if (current == null || typeof current !== 'object') {
      return { found: false, value: undefined };
    }
    if (!(key in current)) {
      return { found: false, value: undefined };
    }
    current = current[key];
  }

  return { found: true, value: current };
}

// Adjust row-relative references in formulas when cloning rows.
// Example: fromRow=2, toRow=3, formula 'B2*C2' becomes 'B3*C3'.
function adjustFormulaForRow(formula: string, fromRow: number, toRow: number): string {
  if (fromRow === toRow) return formula;

  return formula.replace(/(\$?[A-Z]+)(\d+)/g, (match, col, rowStr) => {
    const row = Number(rowStr);
    if (Number.isNaN(row) || row !== fromRow) return match;
    return `${col}${toRow}`;
  });
}