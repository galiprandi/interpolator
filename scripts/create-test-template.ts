import { Workbook } from 'exceljs';
import * as fs from 'fs';

async function createTemplate() {
  const wb = new Workbook();
  const ws = wb.addWorksheet('Hoja1');

  ws.getCell('A1').value = 'Nombre: {{name}}';
  ws.getCell('A2').value = 'ID: [[items.id]]';
  ws.getCell('B2').value = 'Qty: [[items.qty]]';

  await wb.xlsx.writeFile('packages/xlsx/tests/fixtures/template.xlsx');
  console.log('Plantilla creada');
}

createTemplate();