import ExcelJS from 'exceljs';
import { resolvePath } from '@interpolator/core';

const { Workbook } = ExcelJS;

export interface InterpolateXlsxOptions {
  template: Buffer;
  data: Record<string, any>;
}

/**
 * Resolves a path from the data object or from the special Excel context markers.
 */
function resolveWithContext(
  path: string,
  data: any,
  ctx: { now: Date; sheet?: string; row?: number; col?: number },
): { found: boolean; value: any } {
  const trimmedPath = path.trim();
  if (trimmedPath === '$now') return { found: true, value: ctx.now };
  if (trimmedPath === '$sheet') return { found: true, value: ctx.sheet };
  if (trimmedPath === '$row') return { found: true, value: ctx.row };
  if (trimmedPath === '$col') return { found: true, value: ctx.col };

  return resolvePath(data, trimmedPath);
}

export async function interpolateXlsx(options: InterpolateXlsxOptions): Promise<Buffer> {
  const { template, data } = options;
  const workbook = new Workbook();
  await workbook.xlsx.load(template as any);

  const now = new Date();

  for (const worksheet of workbook.worksheets) {
    // Interpolate worksheet name
    worksheet.name = worksheet.name.replace(/\{\{\s*([^\}]+)\s*\}\}/g, (_, path) => {
      const { found, value: resolved } = resolveWithContext(path, data, {
        now,
        sheet: worksheet.name,
      });
      if (!found) return `{{${path}}}`;
      return resolved == null ? '' : String(resolved);
    });

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

    // Capture current worksheet merge ranges once
    const mergeRanges = getWorksheetMergeRanges(worksheet as any);

    for (const { rowNumber, arrayKey } of rowsToExpand) {
      const array = data[arrayKey];
      if (array === undefined) {
        continue; // Leave markers untouched
      }
      if (!Array.isArray(array)) {
        const sheetName = worksheet.name;
        throw new Error(
          `[[${arrayKey}.*]] requires "${arrayKey}" to be an array in worksheet "${sheetName}", row ${rowNumber}. Received: ${
            array === null ? 'null' : typeof array
          }`,
        );
      }

      const originalRow = worksheet.getRow(rowNumber);

      // Collect merges that involve the template row before deleting it
      const templateRowMerges = mergeRanges.filter((range) =>
        mergeRangeIncludesRow(range, rowNumber),
      );

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
            value = {
              ...(value as any),
              formula: adjustedFormula
            };
            // Styles & validation will be copied below; skip marker interpolation for formulas
          } else if (typeof value === 'string') {
            // Check if it's a single [[ ]] marker to preserve type
            const singleArMatch = value.match(/^\[\[\s*([^\].\s]+)(?:\.([^\]\s]+))?\s*\]\]$/);
            if (singleArMatch) {
              const [, arrKey, propPath] = singleArMatch;
              if (arrKey === arrayKey) {
                if (!propPath) {
                  value = item === undefined ? value : item;
                } else if (propPath === '$index' || propPath === '$index0') {
                  value = i;
                } else if (propPath === '$index1' || propPath === '$number') {
                  value = i + 1;
                } else if (propPath === '$first') {
                  value = i === 0;
                } else if (propPath === '$last') {
                  value = i === array.length - 1;
                } else if (propPath === '$length') {
                  value = array.length;
                } else if (propPath === '$even') {
                  value = (i + 1) % 2 === 0;
                } else if (propPath === '$odd') {
                  value = (i + 1) % 2 !== 0;
                } else if (propPath === '$row') {
                  value = newRowNumber;
                } else if (propPath === '$col') {
                  value = colNumber;
                } else {
                  const { found, value: resolved } = resolvePath(item, propPath);
                  value = found ? resolved : value;
                }
              }
            }

            // If it was not a single array marker, or it's still a string, process all markers
            if (typeof value === 'string') {
              // Check if it's a single {{ }} marker to preserve type
              const singleRootMatch = value.match(/^\{\{\s*([^\}]+)\s*\}\}$/);
              if (singleRootMatch) {
                const { found, value: resolved } = resolveWithContext(singleRootMatch[1], data, {
                  now,
                  sheet: worksheet.name,
                  row: newRowNumber,
                  col: colNumber,
                });
                if (found) {
                  value = resolved;
                }
              }
            }

            // Final string interpolation for remaining cases
            if (typeof value === 'string') {
              // Array interpolation: [[array.key]] or [[array]]
              value = value.replace(/\[\[\s*([^\].\s]+)(?:\.([^\]\s]+))?\s*\]\]/g, (_, arrKey, propPath) => {
                if (arrKey !== arrayKey) return propPath ? `[[${arrKey}.${propPath}]]` : `[[${arrKey}]]`;

                if (!propPath) {
                  return item == null ? '' : String(item);
                }

                if (propPath === '$index' || propPath === '$index0') return String(i);
                if (propPath === '$index1' || propPath === '$number') return String(i + 1);
                if (propPath === '$first') return String(i === 0);
                if (propPath === '$last') return String(i === array.length - 1);
                if (propPath === '$length') return String(array.length);
                if (propPath === '$even') return String((i + 1) % 2 === 0);
                if (propPath === '$odd') return String((i + 1) % 2 !== 0);
                if (propPath === '$row') return String(newRowNumber);
                if (propPath === '$col') return String(colNumber);

                const { found, value: resolved } = resolvePath(item, propPath);
                if (!found) return `[[${arrKey}.${propPath}]]`;
                return resolved == null ? '' : String(resolved);
              });

              // Root-level interpolation: {{key}}
              value = value.replace(/\{\{\s*([^\}]+)\s*\}\}/g, (_, path) => {
                const { found, value: resolved } = resolveWithContext(path, data, {
                  now,
                  sheet: worksheet.name,
                  row: newRowNumber,
                  col: colNumber,
                });
                if (!found) return `{{${path}}}`;
                return resolved == null ? '' : String(resolved);
              });
            }
          }

          if (value !== undefined) {
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

        // Replicate merges for this new row
        for (const range of templateRowMerges) {
          const parsed = parseMergeRange(range);
          if (!parsed) continue;
          const { startRow, endRow, startCol, endCol } = parsed;

          const rowOffset = newRowNumber - rowNumber;
          const newStartRow = startRow + rowOffset;
          const newEndRow = endRow + rowOffset;
          const newRange = `${startCol}${newStartRow}:${endCol}${newEndRow}`;

          worksheet.mergeCells(newRange);
        }
      }
    }

    // Second pass: interpolate root-level {{ }} markers in all cells
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        if (typeof cell.value !== 'string') return;

        let value: any = cell.value;

        // Check if it's a single {{ }} marker to preserve type
        const singleRootMatch = value.match(/^\{\{\s*([^\}]+)\s*\}\}$/);
        if (singleRootMatch) {
          const { found, value: resolved } = resolveWithContext(singleRootMatch[1], data, {
            now,
            sheet: worksheet.name,
            row: rowNumber,
            col: colNumber,
          });
          if (found) {
            value = resolved;
          }
        }

        if (typeof value === 'string') {
          value = value.replace(/\{\{\s*([^\}]+)\s*\}\}/g, (_, path) => {
            const { found, value: resolved } = resolveWithContext(path, data, {
              now,
              sheet: worksheet.name,
              row: rowNumber,
              col: colNumber,
            });
            if (!found) return `{{${path}}}`;
            return resolved == null ? '' : String(resolved);
          });
        }

        cell.value = value;
      });
    });
  }

  const result = await workbook.xlsx.writeBuffer();
  return result as any as Buffer;
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

// Utility helpers for working with merge ranges
function getWorksheetMergeRanges(worksheet: any): string[] {
  const merges = worksheet._merges;
  if (!merges) return [];

  if (typeof merges.keys === 'function') {
    return Array.from(merges.keys());
  }

  return Object.keys(merges);
}

function mergeRangeIncludesRow(range: string, row: number): boolean {
  const parsed = parseMergeRange(range);
  if (!parsed) return false;
  const { startRow, endRow } = parsed;
  return row >= startRow && row <= endRow;
}

function parseMergeRange(range: string):
  | { startRow: number; endRow: number; startCol: string; endCol: string }
  | null {
  const [startRef, endRef] = range.split(':');
  if (!startRef || !endRef) return null;

  const start = parseCellRef(startRef);
  const end = parseCellRef(endRef);
  if (!start || !end) return null;

  return {
    startRow: start.row,
    endRow: end.row,
    startCol: start.col,
    endCol: end.col,
  };
}

function parseCellRef(ref: string): { col: string; row: number } | null {
  const match = /^\$?([A-Z]+)(\d+)$/.exec(ref);
  if (!match) return null;
  const [, col, rowStr] = match;
  const row = Number(rowStr);
  if (Number.isNaN(row)) return null;
  return { col, row };
}