import { readFileSync, writeFileSync } from 'fs';
import { Workbook } from 'exceljs';
import { interpolateXlsx } from '../packages/xlsx/src';
async function run() {
  // 1) Leer la plantilla
  const template = readFileSync('packages/xlsx/tests/fixtures/template.xlsx');

  // 2) Datos de ejemplo
  const data = {
    name: 'Germán',
    items: [
      { id: '001', qty: 2 },
      { id: '002', qty: 1 },
    ],
  };

  // 3) Interpolar XLSX
  const resultBuffer = await interpolateXlsx({ template, data });

  // 4) Guardar resultado para inspección manual en Excel/Numbers
  writeFileSync('packages/xlsx/tests/fixtures/result.xlsx', resultBuffer);

  // 5) Volver a cargar el resultado con exceljs para inspección rápida en consola
  const wb = new Workbook();
  await wb.xlsx.load(resultBuffer);
  const ws = wb.getWorksheet('Hoja1');

  if (!ws) {
    console.error('Worksheet "Hoja1" not found');
    return;
  }

  console.log('--- Valores interpolados ---');
  console.log('A1:', ws.getCell('A1').value);
  console.log('A2:', ws.getCell('A2').value);
  console.log('B2:', ws.getCell('B2').value);
  console.log('A3:', ws.getCell('A3').value);
  console.log('B3:', ws.getCell('B3').value);
}

run().catch((err) => {
  console.error(err);
  process.exit(1);
});
