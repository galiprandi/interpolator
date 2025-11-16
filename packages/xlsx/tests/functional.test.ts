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

describe('interpolateXlsx - functional', () => {
  it('should interpolate simple {{}} markers with root data', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'Hello {{user.name}}';
    });

    const data = { user: { name: 'Germán' } };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    expect(ws.getCell('A1').value).toBe('Hello Germán');
  });

  it('should expand array rows for [[ ]] markers and interpolate {{}} in the same sheet', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'Client: {{client.name}}';
      ws.getCell('A2').value = 'ID: [[items.id]]';
      ws.getCell('B2').value = 'Qty: [[items.qty]]';
    });

    const data = {
      client: { name: 'Germán' },
      items: [
        { id: '001', qty: 2 },
        { id: '002', qty: 1 },
      ],
    };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    // Header
    expect(ws.getCell('A1').value).toBe('Client: Germán');

    // Expanded rows
    expect(ws.getCell('A2').value).toBe('ID: 001');
    expect(ws.getCell('B2').value).toBe('Qty: 2');

    expect(ws.getCell('A3').value).toBe('ID: 002');
    expect(ws.getCell('B3').value).toBe('Qty: 1');

    // No leftover template row with markers
    expect(String(ws.getCell('A2').value)).not.toContain('[[items.id]]');
    expect(String(ws.getCell('B2').value)).not.toContain('[[items.qty]]');
  });

  it('should leave markers intact for missing keys and clear cells for null values', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'Name: {{user.name}}';
      ws.getCell('A2').value = 'Missing: {{user.missing}}';
    });

    const data = { user: { name: null } };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    // null -> empty string
    expect(ws.getCell('A1').value).toBe('Name: ');

    // missing key -> marker stays
    expect(ws.getCell('A2').value).toBe('Missing: {{user.missing}}');
  });

  it('should remove the template row when the array is empty', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'Header';
      ws.getCell('A2').value = 'ID: [[items.id]]';
    });

    const data = { items: [] as Array<{ id: string }> };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    // Only header row should remain
    expect(ws.getCell('A1').value).toBe('Header');
    expect(ws.getCell('A2').value).toBeNull();
    expect(ws.rowCount).toBeGreaterThanOrEqual(1);
  });
});
