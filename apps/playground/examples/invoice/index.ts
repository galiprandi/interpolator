import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

// Import locally from the workspace source to avoid relying on workspace linking
// Path: apps/playground/examples/invoice -> ../../../../packages/xlsx/src
import { interpolateXlsx } from '../../../../packages/xlsx/src';

export async function runExample(opts: { repoRoot: string }): Promise<void> {
  const { repoRoot } = opts;
  const baseDir = join(repoRoot, 'apps/playground/examples/invoice');

  const templatePath = join(baseDir, 'template.xlsx');
  const dataPath = join(baseDir, 'data.json');
  const outputPath = join(baseDir, 'output.xlsx');

  const template = readFileSync(templatePath);
  const data = JSON.parse(readFileSync(dataPath, 'utf8'));

  const result = await interpolateXlsx({ template, data });

  writeFileSync(outputPath, result);

  // eslint-disable-next-line no-console
  console.log('Invoice example generated at:', outputPath);
}
