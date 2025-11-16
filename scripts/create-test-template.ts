import { Workbook } from 'exceljs';
import * as fs from 'fs';

async function createTemplate() {
  const wb = new Workbook();
  const ws = wb.addWorksheet('Sheet1');

  ws.getCell('A1').value = 'Name: {{name}}';
  ws.getCell('B1').value = 'Date: {{date}}';

  ws.getCell('A2').value = 'ID: [[items.id]]';
  ws.getCell('B2').value = 'Qty: [[items.qty]]';
  ws.getCell('C2').value = 'Price: [[items.price]]';
  ws.getCell('D2').value = 'Item date: [[items.date]]';
  ws.getCell('E2').value = 'Missing: [[items.missingProp]]';
  ws.getCell('F2').value = { formula: 'B2*C2' };

  await wb.xlsx.writeFile('packages/xlsx/tests/fixtures/template.xlsx');
  console.log('Template created');
}

createTemplate();