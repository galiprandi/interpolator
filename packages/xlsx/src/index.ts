import ExcelJS from 'exceljs';
import { resolvePath, applyTransforms } from '@interpolator/core';

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
  ctx: {
    now: Date;
    sheet?: string;
    sheetIndex?: number;
    totalSheets?: number;
    row?: number;
    col?: number;
    index?: number;
    length?: number;
  },
): { found: boolean; value: any } {
  const trimmed = path.trim();
  const parts = trimmed.split(/\s*\|\|\s*/);
  const mainPathWithTransforms = parts[0];
  const defaultValue = parts[1];

  const [mainPath, ...transforms] = mainPathWithTransforms.split(/\s*\|\s*/);

  let result = resolveInternal(mainPath, data, ctx);

  if ((!result.found || result.value == null) && defaultValue !== undefined) {
    // Try to resolve default value as a path first, if not found, use as literal
    const [defaultPath, ...defaultTransforms] = defaultValue.split(/\s*\|\s*/);
    const resolvedDefault = resolveInternal(defaultPath, data, ctx);
    const value = resolvedDefault.found ? resolvedDefault.value : defaultPath;
    return { found: true, value: applyTransforms(value, defaultTransforms) };
  }

  if (result.found && transforms.length > 0) {
    result.value = applyTransforms(result.value, transforms);
  }

  return result;
}

function resolveInternal(
  trimmedPath: string,
  data: any,
  ctx: {
    now: Date;
    sheet?: string;
    sheetIndex?: number;
    totalSheets?: number;
    row?: number;
    col?: number;
    index?: number;
    length?: number;
  },
): { found: boolean; value: any } {
  switch (trimmedPath) {
    case '$now': return { found: true, value: ctx.now };
    case '$year': return { found: true, value: ctx.now.getFullYear() };
    case '$month': return { found: true, value: ctx.now.getMonth() + 1 };
    case '$day': return { found: true, value: ctx.now.getDate() };
    case '$hour': return { found: true, value: ctx.now.getHours() };
    case '$minute': return { found: true, value: ctx.now.getMinutes() };
    case '$second': return { found: true, value: ctx.now.getSeconds() };
    case '$weekday': return { found: true, value: ctx.now.getDay() };
    case '$sheet':
    case '$sheetName': return { found: true, value: ctx.sheet };
    case '$sheetIndex': return { found: true, value: ctx.sheetIndex };
    case '$sheetNumber':
      return { found: true, value: ctx.sheetIndex !== undefined ? ctx.sheetIndex + 1 : undefined };
    case '$totalSheets': return { found: true, value: ctx.totalSheets };
    case '$isFirstSheet':
      return { found: true, value: ctx.sheetIndex === 0 };
    case '$isLastSheet':
      return {
        found: true,
        value: ctx.sheetIndex !== undefined && ctx.totalSheets !== undefined
            ? ctx.sheetIndex === ctx.totalSheets - 1
            : undefined,
      };
    case '$isEvenSheet':
      return { found: true, value: ctx.sheetIndex !== undefined ? (ctx.sheetIndex + 1) % 2 === 0 : undefined };
    case '$isOddSheet':
      return { found: true, value: ctx.sheetIndex !== undefined ? (ctx.sheetIndex + 1) % 2 !== 0 : undefined };
    case '$row':
    case '$rowNumber': return { found: true, value: ctx.row };
    case '$col':
    case '$colNumber': return { found: true, value: ctx.col };
    case '$rowIndex': return { found: true, value: ctx.row !== undefined ? ctx.row - 1 : undefined };
    case '$colIndex': return { found: true, value: ctx.col !== undefined ? ctx.col - 1 : undefined };
    case '$isEven':
    case '$even':
      if (ctx.index !== undefined) return { found: true, value: (ctx.index + 1) % 2 === 0 };
      if (ctx.row !== undefined) return { found: true, value: ctx.row % 2 === 0 };
      return { found: true, value: ctx.sheetIndex !== undefined ? (ctx.sheetIndex + 1) % 2 === 0 : undefined };
    case '$isEvenRow':
      return { found: true, value: ctx.row !== undefined ? ctx.row % 2 === 0 : undefined };
    case '$isOdd':
    case '$odd':
      if (ctx.index !== undefined) return { found: true, value: (ctx.index + 1) % 2 !== 0 };
      if (ctx.row !== undefined) return { found: true, value: ctx.row % 2 !== 0 };
      return { found: true, value: ctx.sheetIndex !== undefined ? (ctx.sheetIndex + 1) % 2 !== 0 : undefined };
    case '$isOddRow':
      return { found: true, value: ctx.row !== undefined ? ctx.row % 2 !== 0 : undefined };
    case '$isEvenCol':
      return { found: true, value: ctx.col !== undefined ? ctx.col % 2 === 0 : undefined };
    case '$isOddCol':
      return { found: true, value: ctx.col !== undefined ? ctx.col % 2 !== 0 : undefined };
    case '$colLetter':
    case '$columnLetter':
      return { found: true, value: ctx.col ? getColLetter(ctx.col) : undefined };
    case '$cell':
      return {
        found: true,
        value: ctx.row && ctx.col ? `${getColLetter(ctx.col)}${ctx.row}` : undefined,
      };
    case '$isHeader':
      return { found: true, value: ctx.row === 1 };
    case '$index':
    case '$index0':
      return { found: true, value: ctx.index };
    case '$index1':
    case '$number':
      return { found: true, value: ctx.index !== undefined ? ctx.index + 1 : undefined };
    case '$length':
      return { found: true, value: ctx.length };
    case '$first':
    case '$isFirst':
      if (ctx.index !== undefined) return { found: true, value: ctx.index === 0 };
      return { found: true, value: ctx.sheetIndex === 0 };
    case '$last':
    case '$isLast':
      if (ctx.index !== undefined && ctx.length !== undefined) {
        return { found: true, value: ctx.index === ctx.length - 1 };
      }
      return {
        found: true,
        value: ctx.sheetIndex !== undefined && ctx.totalSheets !== undefined
            ? ctx.sheetIndex === ctx.totalSheets - 1
            : undefined,
      };
    default:
      return resolvePath(data, trimmedPath);
  }
}

function findArrayInPath(
  fullPath: string,
  data: any,
  ctx: any,
): { arrayPath: string; propertyPath: string; array: any; found: boolean } | null {
  const parts = fullPath.trim().split('.');
  for (let i = 1; i <= parts.length; i++) {
    const arrayPath = parts.slice(0, i).join('.').trim();
    const { found, value } = resolveWithContext(arrayPath, data, ctx);
    if (found && (Array.isArray(value) || typeof value === 'boolean')) {
      return { arrayPath, propertyPath: parts.slice(i).join('.').trim(), array: value, found: true };
    }
  }

  // Fallback for backward compatibility: if the first part exists but is not an array,
  // we still return it so it can throw an error during expansion if it matches the marker.
  const firstPart = parts[0].trim();
  const { found, value } = resolveWithContext(firstPart, data, ctx);
  if (found) {
    return { arrayPath: firstPart, propertyPath: parts.slice(1).join('.').trim(), array: value, found: true };
  }

  return null;
}

export async function interpolateXlsx(options: InterpolateXlsxOptions): Promise<Buffer> {
  const { template, data } = options;
  const workbook = new Workbook();
  await workbook.xlsx.load(template as any);

  const now = new Date();
  const totalSheets = workbook.worksheets.length;

  for (let sIdx = 0; sIdx < totalSheets; sIdx++) {
    const worksheet = workbook.worksheets[sIdx];
    const sheetCtx = { now, sheet: worksheet.name, sheetIndex: sIdx, totalSheets };
    // Interpolate worksheet name
    worksheet.name = worksheet.name.replace(/\{\{\s*([^\}]+)\s*\}\}/g, (_, path) => {
      const { found, value: resolved } = resolveWithContext(path, data, sheetCtx);
      if (!found) return `{{${path}}}`;
      return resolved == null ? '' : String(resolved);
    });
    sheetCtx.sheet = worksheet.name;

    const rowsToExpand: { rowNumber: number; arrayKey: string }[] = [];

    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      let arrayKey: string | null = null;

      row.eachCell({ includeEmpty: true }, (cell) => {
        if (typeof cell.value === 'string') {
          const matches = cell.value.matchAll(/\[\[\s*([^\]]+)\s*\]\]/g);
          for (const match of matches) {
            const res = findArrayInPath(match[1], data, sheetCtx);
            if (res) {
              if (arrayKey && arrayKey !== res.arrayPath) {
                throw new Error(`Mixed array keys in row ${rowNumber}: ${arrayKey} vs ${res.arrayPath}`);
              }
              arrayKey = res.arrayPath;
            }
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
      const { found: arrayFound, value: resolvedArray } = resolveWithContext(arrayKey, data, sheetCtx);
      let array = resolvedArray;
      if (!arrayFound || array === undefined) {
        continue; // Leave markers untouched
      }

      // Support boolean conditional rows: true -> render once, false -> remove
      const isBooleanRow = typeof array === 'boolean';
      if (isBooleanRow) {
        array = array ? [{ __isConditional: true }] : [];
      }

      if (!Array.isArray(array)) {
        const sheetName = worksheet.name;
        throw new Error(
          `[[${arrayKey}.*]] requires "${arrayKey}" to be an array or boolean in worksheet "${sheetName}", row ${rowNumber}. Received: ${
            array === null ? 'null' : typeof array
          }`,
        );
      }

      const originalRow = worksheet.getRow(rowNumber);
      const rowValues: any[] = [];
      originalRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        rowValues[colNumber] = cell.value;
      });

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
        for (let colNumber = 1; colNumber < rowValues.length; colNumber++) {
          const originalCell = originalRow.getCell(colNumber);
          let value = rowValues[colNumber];
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
            const singleArMatch = value.match(/^\[\[\s*([^\]]+)\s*\]\]$/);
            if (singleArMatch) {
              const fullPath = singleArMatch[1].trim();
              if (fullPath === arrayKey || fullPath.startsWith(arrayKey + '.')) {
                const propPath = fullPath === arrayKey ? '' : fullPath.slice(arrayKey.length + 1);
                if (!propPath) {
                  // For boolean-based rows, resolve to empty string
                  value = (isBooleanRow && item && (item as any).__isConditional) ? '' : (item === undefined ? value : item);
                } else {
                  const { found, value: resolved } = resolveWithContext(propPath, item, {
                    ...sheetCtx,
                    row: newRowNumber,
                    col: colNumber,
                    index: i,
                    length: array.length,
                  });
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
                  ...sheetCtx,
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
              value = value.replace(/\[\[\s*([^\]]+)\s*\]\]/g, (full, fullPath) => {
                const trimmedPath = fullPath.trim();
                if (trimmedPath !== arrayKey && !trimmedPath.startsWith(arrayKey + '.')) return full;

                const propPath = trimmedPath === arrayKey ? '' : trimmedPath.slice(arrayKey.length + 1);

                if (!propPath) {
                  // For boolean-based rows, resolve to empty string
                  if (isBooleanRow && item && (item as any).__isConditional) return '';
                  return item == null ? '' : String(item);
                }

                const { found, value: resolved } = resolveWithContext(propPath, item, {
                  ...sheetCtx,
                  row: newRowNumber,
                  col: colNumber,
                  index: i,
                  length: array.length,
                });

                if (!found) return full;
                return resolved == null ? '' : String(resolved);
              });

              // Root-level interpolation: {{key}}
              value = value.replace(/\{\{\s*([^\}]+)\s*\}\}/g, (_: string, path: string) => {
                const { found, value: resolved } = resolveWithContext(path, data, {
                  ...sheetCtx,
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
        }

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
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (typeof cell.value !== 'string') return;

        let value: any = cell.value;

        // Check if it's a single {{ }} marker to preserve type
        const singleRootMatch = value.match(/^\{\{\s*([^\}]+)\s*\}\}$/);
        if (singleRootMatch) {
          const { found, value: resolved } = resolveWithContext(singleRootMatch[1], data, {
            ...sheetCtx,
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
              ...sheetCtx,
              row: rowNumber,
              col: colNumber,
            });
            if (!found) return `{{${path}}}`;
            return resolved == null ? '' : String(resolved);
          });
        }

        if (value !== undefined) {
          cell.value = value;
        }
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

function getColLetter(col: number): string {
  let letter = '';
  let n = col;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}