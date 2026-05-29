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

  it('should leave {{}} markers intact when the root key does not exist', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'Hello {{profile.name}}';
    });

    const data = {}; // no profile key

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    expect(ws.getCell('A1').value).toBe('Hello {{profile.name}}');
  });

  it('should leave {{}} markers intact when an intermediate nested property is missing', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'Email: {{user.profile.email}}';
    });

    const data = { user: {} }; // user.profile is missing

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    expect(ws.getCell('A1').value).toBe('Email: {{user.profile.email}}');
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

  it('should leave [[array.prop]] markers intact when the item property does not exist', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A2').value = 'ID: [[items.id]] - Name: [[items.name]]';
    });

    const data = {
      items: [{ id: '001' }], // no name property
    };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    expect(ws.getCell('A2').value).toBe('ID: 001 - Name: [[items.name]]');
  });

  it('should render empty string when an item property is null or undefined', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      // Single template row that will be expanded into two rows
      ws.getCell('A2').value = 'Amount: [[payments.amount]]';
    });

    const data = {
      payments: [{ amount: null }, { amount: undefined }],
    };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    // Row for first item (null)
    expect(ws.getCell('A2').value).toBe('Amount: ');
    // Row for second item (undefined)
    expect(ws.getCell('A3').value).toBe('Amount: ');
  });

  it('should preserve formulas and adjust relative references per cloned row', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A2').value = 'ID: [[items.id]]';
      ws.getCell('B2').value = 'Qty: [[items.qty]]';
      ws.getCell('C2').value = 'Price: [[items.price]]';
      // Line total formula for the template row
      ws.getCell('D2').value = { formula: 'B2*C2' };
    });

    const data = {
      items: [
        { id: '001', qty: 2, price: 10 },
        { id: '002', qty: 3, price: 20 },
      ],
    };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    // Formulas should be preserved and references adjusted per row
    // First item row (template position)
    expect(ws.getCell('D2').type).toBe(6 /* formula */);
    expect((ws.getCell('D2').value as any).formula).toBe('B2*C2');

    // Second item row should have the formula adjusted to point to its own row
    expect(ws.getCell('D3').type).toBe(6 /* formula */);
    expect((ws.getCell('D3').value as any).formula).toBe('B3*C3');
  });

  it('should preserve basic cell styles when expanding array rows', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      const row = ws.getRow(2);
      row.getCell(1).value = 'ID: [[items.id]]';
      row.getCell(1).style = {
        font: { bold: true },
        alignment: { horizontal: 'center' },
      };
    });

    const data = {
      items: [{ id: '001' }, { id: '002' }],
    };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    const cell1 = ws.getCell('A2');
    const cell2 = ws.getCell('A3');

    expect(cell1.style.font?.bold).toBe(true);
    expect(cell2.style.font?.bold).toBe(true);
    expect(cell1.style.alignment?.horizontal).toBe('center');
    expect(cell2.style.alignment?.horizontal).toBe('center');
  });

  // NOTE: exceljs currently does not reliably round-trip dynamically added merges
  // through writeBuffer/load in our setup. This test captures the desired
  // behavior, but is skipped for v1 until we can investigate a robust approach
  // to merging cloned rows.
  it.skip('should replicate merged cells for each expanded array row', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A2').value = 'ID: [[items.id]]';
      ws.getCell('B2').value = 'Qty: [[items.qty]]';
      ws.getCell('C2').value = 'Price: [[items.price]]';
      ws.mergeCells('A2:C2');
    });

    const data = {
      items: [
        { id: '001', qty: 2, price: 10 },
        { id: '002', qty: 3, price: 20 },
      ],
    };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    // Expanded rows should be present
    expect(ws.getCell('A2').value).toBe('ID: 001');
    expect(ws.getCell('A3').value).toBe('ID: 002');

    // Merged regions should be replicated for each expanded row
    // We assert at the cell level using the public API
    // First item row: A2:C2 merged
    expect(ws.getCell('A2').isMerged).toBe(true);
    expect(ws.getCell('B2').isMerged).toBe(true);
    expect(ws.getCell('C2').isMerged).toBe(true);

    // Second item row: A3:C3 merged
    expect(ws.getCell('A3').isMerged).toBe(true);
    expect(ws.getCell('B3').isMerged).toBe(true);
    expect(ws.getCell('C3').isMerged).toBe(true);
  });

  it('should throw a clear error when array key exists but is not an array', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A2').value = 'ID: [[user.id]]';
    });

    const data = {
      user: { id: 'U1' }, // not an array
    };

    await expect(interpolateXlsx({ template, data })).rejects.toThrow(
      /\[\[user\.\*\]\] requires "user" to be an array in worksheet "Sheet1", row 2\. Received:/i,
    );
  });

  it('should keep other worksheets unchanged when only one contains markers', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws1 = wb.addWorksheet('WithMarkers');
      const ws2 = wb.addWorksheet('StaticSheet');

      ws1.getCell('A1').value = 'Client: {{client.name}}';
      ws1.getCell('A2').value = 'ID: [[items.id]]';

      ws2.getCell('A1').value = 'Static header';
      ws2.getCell('A2').value = 'Static value';
    });

    const data = {
      client: { name: 'Germán' },
      items: [{ id: '001' }],
    };

    const result = await interpolateXlsx({ template, data });
    const wb = new Workbook();
    await wb.xlsx.load(result as any);

    const ws1 = wb.getWorksheet('WithMarkers')!;
    const ws2 = wb.getWorksheet('StaticSheet')!;

    expect(ws1.getCell('A1').value).toBe('Client: Germán');
    expect(ws1.getCell('A2').value).toBe('ID: 001');

    expect(ws2.getCell('A1').value).toBe('Static header');
    expect(ws2.getCell('A2').value).toBe('Static value');
  });

  it('should leave markers untouched when array key is missing (undefined)', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A2').value = 'ID: [[payments.id]]';
    });

    const data = {};

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    expect(ws.getCell('A2').value).toBe('ID: [[payments.id]]');
  });

  it('should throw an error when a row mixes different array keys', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A2').value = 'ID: [[items.id]]';
      ws.getCell('B2').value = 'Payment: [[payments.id]]';
    });

    const data = {
      items: [{ id: 'I1' }],
      payments: [{ id: 'P1' }],
    };

    await expect(interpolateXlsx({ template, data })).rejects.toThrow(
      /Mixed array keys in row 2: items vs payments/i,
    );
  });

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
      items: [{ name: 'First' }, { name: 'Second' }],
    };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    // Row 1
    expect(ws.getCell('A1').value).toBe(0);
    expect(ws.getCell('B1').value).toBe(1);
    expect(ws.getCell('C1').value).toBe(1);
    expect(ws.getCell('D1').value).toBe('First');

    // Row 2
    expect(ws.getCell('A2').value).toBe(1);
    expect(ws.getCell('B2').value).toBe(2);
    expect(ws.getCell('C2').value).toBe(2);
    expect(ws.getCell('D2').value).toBe('Second');
  });

  it('should support special metadata markers $first, $last, and $length', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = '[[items.$first]]';
      ws.getCell('B1').value = '[[items.$last]]';
      ws.getCell('C1').value = '[[items.$length]]';
    });

    const data = {
      items: [{ name: 'First' }, { name: 'Second' }],
    };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    // Row 1
    expect(ws.getCell('A1').value).toBe(true);
    expect(ws.getCell('B1').value).toBe(false);
    expect(ws.getCell('C1').value).toBe(2);

    // Row 2
    expect(ws.getCell('A2').value).toBe(false);
    expect(ws.getCell('B2').value).toBe(true);
    expect(ws.getCell('C2').value).toBe(2);
  });

  it('should support {{array.length}} as it is a property of the array', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'Total items: {{items.length}}';
    });

    const data = {
      items: [1, 2, 3],
    };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    expect(ws.getCell('A1').value).toBe('Total items: 3');
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

  it('should preserve type (number) for single {{}} markers', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = '{{amount}}';
    });

    const data = { amount: 123.45 };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    expect(ws.getCell('A1').value).toBe(123.45);
    expect(typeof ws.getCell('A1').value).toBe('number');
  });

  it('should preserve Date values in expanded rows', async () => {
    const now = new Date();
    // Normalize date to avoid ms differences if any during serialization
    now.setMilliseconds(0);

    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'Date';
      ws.getCell('A2').value = now;
      ws.getCell('B2').value = '[[items.id]]';
    });

    const data = { items: [{ id: 1 }] };
    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    const cellValue = ws.getCell('A2').value;
    expect(cellValue).toBeInstanceOf(Date);
    expect((cellValue as Date).getTime()).toBe(now.getTime());
  });

  it('should interpolate sheet names', async () => {
    const template = await buildTemplateBuffer((wb) => {
      wb.addWorksheet('Report {{year}}');
    });

    const data = { year: 2024 };
    const result = await interpolateXlsx({ template, data });
    const wb = new Workbook();
    await wb.xlsx.load(result as any);

    expect(wb.getWorksheet('Report 2024')).toBeDefined();
    expect(wb.getWorksheet('Report {{year}}')).toBeUndefined();
  });

  it('should support $even and $odd metadata markers', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = '[[items.$even]]';
      ws.getCell('B1').value = '[[items.$odd]]';
      ws.getCell('C1').value = 'Is even: [[items.$even]]';
    });

    const data = {
      items: [{ name: 'First' }, { name: 'Second' }],
    };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    // Row 1 (Index 0, Number 1 - Odd)
    expect(ws.getCell('A1').value).toBe(false);
    expect(ws.getCell('B1').value).toBe(true);
    expect(ws.getCell('C1').value).toBe('Is even: false');

    // Row 2 (Index 1, Number 2 - Even)
    expect(ws.getCell('A2').value).toBe(true);
    expect(ws.getCell('B2').value).toBe(false);
    expect(ws.getCell('C2').value).toBe('Is even: true');
  });

  describe('Excel context markers', () => {
    it('should support {{$now}} for current date', async () => {
      const template = await buildTemplateBuffer((wb) => {
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = '{{$now}}';
      });

      const result = await interpolateXlsx({ template, data: {} });
      const ws = await loadWorksheetFromResult(result, 'Sheet1');

      const value = ws.getCell('A1').value;
      expect(value).toBeInstanceOf(Date);
      // It should be roughly now
      expect(Math.abs((value as Date).getTime() - Date.now())).toBeLessThan(10000);
    });

    it('should support {{$sheet}}, {{$row}}, {{$col}} root markers', async () => {
      const template = await buildTemplateBuffer((wb) => {
        const ws = wb.addWorksheet('MySheet');
        ws.getCell('A1').value = 'Sheet: {{$sheet}}';
        ws.getCell('B2').value = '{{$row}}:{{$col}}';
      });

      const result = await interpolateXlsx({ template, data: {} });
      const ws = await loadWorksheetFromResult(result, 'MySheet');

      expect(ws.getCell('A1').value).toBe('Sheet: MySheet');
      expect(ws.getCell('B2').value).toBe('2:2');
    });

    it('should support array expansion markers [[$row]], [[$col]], [[$index0]]', async () => {
      const template = await buildTemplateBuffer((wb) => {
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = '[[items.$index0]]';
        ws.getCell('B1').value = '[[items.$row]]';
        ws.getCell('C1').value = '[[items.$col]]';
      });

      const data = {
        items: [{ name: 'A' }, { name: 'B' }],
      };

      const result = await interpolateXlsx({ template, data });
      const ws = await loadWorksheetFromResult(result, 'Sheet1');

      // Row 1
      expect(ws.getCell('A1').value).toBe(0);
      expect(ws.getCell('B1').value).toBe(1);
      expect(ws.getCell('C1').value).toBe(3); // C is 3rd column

      // Row 2
      expect(ws.getCell('A2').value).toBe(1);
      expect(ws.getCell('B2').value).toBe(2);
      expect(ws.getCell('C2').value).toBe(3);
    });

    it('should support $row and $col in root interpolation within expanded rows', async () => {
      const template = await buildTemplateBuffer((wb) => {
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = 'Row {{ $row }} Col {{ $col }} for [[ items.name ]]';
      });

      const data = {
        items: [{ name: 'A' }, { name: 'B' }],
      };

      const result = await interpolateXlsx({ template, data });
      const ws = await loadWorksheetFromResult(result, 'Sheet1');

      expect(ws.getCell('A1').value).toBe('Row 1 Col 1 for A');
      expect(ws.getCell('A2').value).toBe('Row 2 Col 1 for B');
    });

    it('should support $colLetter and $cell markers', async () => {
      const template = await buildTemplateBuffer((wb) => {
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = 'Col: {{$colLetter}} Cell: {{$cell}}';
        ws.getCell('Z1').value = '{{$colLetter}}';
        ws.getCell('AA1').value = '{{$colLetter}}';
        ws.getCell('A2').value = '[[items.name]] at [[items.$cell]] ([[items.$colLetter]])';
      });

      const data = {
        items: [{ name: 'A' }, { name: 'B' }],
      };

      const result = await interpolateXlsx({ template, data });
      const ws = await loadWorksheetFromResult(result, 'Sheet1');

      expect(ws.getCell('A1').value).toBe('Col: A Cell: A1');
      expect(ws.getCell('Z1').value).toBe('Z');
      expect(ws.getCell('AA1').value).toBe('AA');

      expect(ws.getCell('A2').value).toBe('A at A2 (A)');
      expect(ws.getCell('A3').value).toBe('B at A3 (A)');
    });

    it('should support boolean aliases $isFirst, $isLast, $isEven, $isOdd', async () => {
      const template = await buildTemplateBuffer((wb) => {
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = '[[items.$isFirst]]';
        ws.getCell('B1').value = '[[items.$isLast]]';
        ws.getCell('C1').value = '[[items.$isEven]]';
        ws.getCell('D1').value = '[[items.$isOdd]]';
      });

      const data = {
        items: [{ name: 'A' }, { name: 'B' }],
      };

      const result = await interpolateXlsx({ template, data });
      const ws = await loadWorksheetFromResult(result, 'Sheet1');

      // Row 1
      expect(ws.getCell('A1').value).toBe(true);
      expect(ws.getCell('B1').value).toBe(false);
      expect(ws.getCell('C1').value).toBe(false);
      expect(ws.getCell('D1').value).toBe(true);

      // Row 2
      expect(ws.getCell('A2').value).toBe(false);
      expect(ws.getCell('B2').value).toBe(true);
      expect(ws.getCell('C2').value).toBe(true);
      expect(ws.getCell('D2').value).toBe(false);
    });
  });
});
