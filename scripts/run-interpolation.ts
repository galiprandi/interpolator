import { readFileSync, writeFileSync } from 'fs';
import { Workbook } from 'exceljs';
import { interpolateXlsx } from '../packages/xlsx/src';

async function run() {
  // 1) Read template
  const template = readFileSync('packages/xlsx/tests/fixtures/template.xlsx');

  // 2) Sample data
  const data = {
    name: 'Germán',
    date: new Date('2024-01-02T00:00:00Z'),
    items: [
      { id: '001', qty: 2, price: 120.5, date: new Date('2024-01-03T00:00:00Z') },
      { id: '002', qty: 0, price: null, date: null },
    ],
  };

  // 3) Interpolate XLSX
  const resultBuffer = await interpolateXlsx({ template, data });

  // 4) Save result for manual inspection in Excel/Numbers
  writeFileSync('packages/xlsx/tests/fixtures/result.xlsx', resultBuffer);

  // 5) Reload result with exceljs for quick inspection in console
  const wb = new Workbook();
  await wb.xlsx.load(resultBuffer as any);
  const ws = wb.getWorksheet('Sheet1');

  if (!ws) {
    console.error('Worksheet "Sheet1" not found');
    return;
  }

  console.log('--- Interpolated values ---');
  console.log('A1 (name):', ws.getCell('A1').value);
  console.log('B1 (root date):', ws.getCell('B1').value);
  console.log('A2/A3 (IDs):', ws.getCell('A2').value, ws.getCell('A3').value);
  console.log('B2/B3 (qty):', ws.getCell('B2').value, ws.getCell('B3').value);
  console.log('C2/C3 (prices):', ws.getCell('C2').value, ws.getCell('C3').value);
  console.log('D2/D3 (item dates):', ws.getCell('D2').value, ws.getCell('D3').value);
  console.log('E2/E3 (missing prop):', ws.getCell('E2').value, ws.getCell('E3').value);
  console.log('F2/F3 (line totals):', ws.getCell('F2').value, ws.getCell('F3').value);
}

run().catch((err) => {
  console.error(err);
  process.exit(1);
});
