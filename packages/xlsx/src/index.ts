import { Workbook } from 'exceljs';
import type { Buffer } from 'node:buffer';

export interface InterpolateXlsxOptions {
  template: Buffer;
   Record<string, any>;
}

export async function interpolateXlsx(options: InterpolateXlsxOptions): Promise<Buffer> {
  const { template,  } = options;
  const workbook = new Workbook();
  await workbook.xlsx.load(template);

  for (const worksheet of workbook.worksheets) {
    const rowsToExpand: { rowNumber: number; arrayKey: string }[] = [];

    worksheet.eachRow((row, rowNumber) => {
      let arrayKey: string | null = null;

      row.eachCell((cell) => {
        if (typeof cell.value === 'string') {
          // Detectar si hay marcadores de array
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

    // Procesar filas que deben expandirse
    for (const { rowNumber, arrayKey } of rowsToExpand) {
      const array = data[arrayKey];
      if (array === undefined) {
        continue; // Dejar marcadores intactos
      }
      if (!Array.isArray(array)) {
        throw new Error(`[[${arrayKey}.*]] requires '${arrayKey}' to be an array. Received: ${typeof array}`);
      }

      const originalRow = worksheet.getRow(rowNumber);
      const originalValues = originalRow.values as any[];
      const originalCellStyles = originalRow.getCell('A')._numberFormat; // Simplificado

      // Eliminar la fila original
      worksheet.spliceRows(rowNumber, 1);

      // Insertar nuevas filas
      for (let i = 0; i < array.length; i++) {
        const item = array[i];
        const newRow = worksheet.insertRow(rowNumber + i);

        // Copiar valores y reemplazar marcadores
        originalRow.eachCell((originalCell, colNumber) => {
          let value = originalCell.value;

          if (typeof value === 'string') {
            // Interpolación de array: [[array.key]]
            value = value.replace(/\[\[\s*([^\].]+)\.([^\]]+)\s*\]\]/g, (_, arrKey, propPath) => {
              if (arrKey !== arrayKey) return `[[${arrKey}.${propPath}]]`; // dejar intacto
              const { found, value: resolved } = resolvePath(item, propPath);
              return found && resolved != null ? String(resolved) : `[[${arrKey}.${propPath}]]`;
            });

            // Interpolación simple: {{key}}
            value = value.replace(/\{\{\s*([^\}]+)\s*\}\}/g, (_, path) => {
              const { found, value: resolved } = resolvePath(data, path);
              return found && resolved != null ? String(resolved) : `{{${path}}}`;
            });
          }

          newRow.getCell(colNumber).value = value;
          // Aquí puedes copiar estilos, fórmulas, etc.
        });
      }
    }
  }

  return await workbook.xlsx.writeBuffer();
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