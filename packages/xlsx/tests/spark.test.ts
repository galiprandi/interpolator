import { describe, it, expect } from 'vitest';
import { Workbook } from 'exceljs';
import { interpolateXlsx } from '../src';

async function buildTemplateBuffer(build: (wb: Workbook) => void): Promise<Buffer> {
  const wb = new Workbook();
  build(wb);
  const arrayBuffer = await wb.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer as ArrayBuffer);
}

async function loadWorksheetFromResult(result: Buffer, sheetName: string) {
  const wb = new Workbook();
  await wb.xlsx.load(result as any);
  const ws = wb.getWorksheet(sheetName);
  if (!ws) throw new Error(`Worksheet ${sheetName} not found`);
  return ws;
}

describe('interpolateXlsx - Spark enhancements', () => {
  it('should interpolate primitive arrays with [[array]] markers', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'Value: [[items]]';
    });

    const data = {
      items: ['Apple', 'Banana', 'Cherry'],
    };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    expect(ws.getCell('A1').value).toBe('Value: Apple');
    expect(ws.getCell('A2').value).toBe('Value: Banana');
    expect(ws.getCell('A3').value).toBe('Value: Cherry');
  });

  it('should support special index markers $index, $index1, and $number', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = '[[items.$index]]';
      ws.getCell('B1').value = '[[items.$index1]]';
      ws.getCell('C1').value = '[[items.$number]]';
      ws.getCell('D1').value = '[[items.name]]';
    });

    const data = {
      items: [
        { name: 'First' },
        { name: 'Second' },
      ],
    };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    // Row 1
    expect(ws.getCell('A1').value).toBe('0');
    expect(ws.getCell('B1').value).toBe('1');
    expect(ws.getCell('C1').value).toBe('1');
    expect(ws.getCell('D1').value).toBe('First');

    // Row 2
    expect(ws.getCell('A2').value).toBe('1');
    expect(ws.getCell('B2').value).toBe('2');
    expect(ws.getCell('C2').value).toBe('2');
    expect(ws.getCell('D2').value).toBe('Second');
  });

  it('should combine primitive value and index in one row', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = '[[items.$number]]. [[items]]';
    });

    const data = {
      items: ['One', 'Two'],
    };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    expect(ws.getCell('A1').value).toBe('1. One');
    expect(ws.getCell('A2').value).toBe('2. Two');
  });
});
