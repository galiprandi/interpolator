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

describe('interpolateXlsx - functional - Spark enhancements', () => {
  it('should support pipes in array expansion for transformations', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = '[[items | reverse]]';
    });

    const data = {
      items: ['A', 'B', 'C'],
    };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    expect(ws.getCell('A1').value).toBe('C');
    expect(ws.getCell('A2').value).toBe('B');
    expect(ws.getCell('A3').value).toBe('A');
  });

  it('should support dots in array expansion for nested paths', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = '[[order.items.name]]';
    });

    const data = {
      order: {
        items: [{ name: 'Apple' }, { name: 'Banana' }]
      }
    };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    expect(ws.getCell('A1').value).toBe('Apple');
    expect(ws.getCell('A2').value).toBe('Banana');
  });

  it('should support lines transform to render multi-line string as rows', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = '[[description | lines]]';
    });

    const data = {
      description: "Line 1\nLine 2\nLine 3"
    };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    expect(ws.getCell('A1').value).toBe('Line 1');
    expect(ws.getCell('A2').value).toBe('Line 2');
    expect(ws.getCell('A3').value).toBe('Line 3');
  });

  it('should support aliases for transformations', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = '{{name | camel}}';
      ws.getCell('A2').value = '{{name | pascal}}';
      ws.getCell('A3').value = '{{name | snake}}';
      ws.getCell('A4').value = '{{name | kebab}}';
      ws.getCell('A5').value = '{{name | title}}';
    });

    const data = { name: 'hello world' };

    const result = await interpolateXlsx({ template, data });
    const ws = await loadWorksheetFromResult(result, 'Sheet1');

    expect(ws.getCell('A1').value).toBe('helloWorld');
    expect(ws.getCell('A2').value).toBe('HelloWorld');
    expect(ws.getCell('A3').value).toBe('hello_world');
    expect(ws.getCell('A4').value).toBe('hello-world');
    expect(ws.getCell('A5').value).toBe('Hello World');
  });

  it('should support $isEvenSheet and $isOddSheet markers', async () => {
    const template = await buildTemplateBuffer((wb) => {
      const ws1 = wb.addWorksheet('Sheet 1');
      ws1.getCell('A1').value = '{{$isEvenSheet}}/{{$isOddSheet}}';
      const ws2 = wb.addWorksheet('Sheet 2');
      ws2.getCell('A1').value = '{{$isEvenSheet}}/{{$isOddSheet}}';
    });

    const result = await interpolateXlsx({ template, data: {} });
    const wb = new Workbook();
    await wb.xlsx.load(result as any);

    expect(wb.getWorksheet('Sheet 1')!.getCell('A1').value).toBe('false/true');
    expect(wb.getWorksheet('Sheet 2')!.getCell('A1').value).toBe('true/false');
  });

  it('should support $isEven and $isOdd in sheet name interpolation', async () => {
    const template = await buildTemplateBuffer((wb) => {
      wb.addWorksheet('Odd Sheet {{$isOdd}}');
      wb.addWorksheet('Even Sheet {{$isEven}}');
    });

    const result = await interpolateXlsx({ template, data: {} });
    const wb = new Workbook();
    await wb.xlsx.load(result as any);

    expect(wb.getWorksheet('Odd Sheet true')).toBeDefined();
    expect(wb.getWorksheet('Even Sheet true')).toBeDefined();
  });
});
