import { argv } from 'node:process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

interface ExampleModule {
  runExample(opts: { repoRoot: string }): Promise<void>;
}

const examples: Record<string, () => Promise<ExampleModule>> = {
  invoice: () => import('../examples/invoice/index.js'),
  'invoice-visual': () => import('../examples/invoice-visual/index.js'),
  // Future examples can be added here, e.g.
  // 'edge-cases': () => import('../examples/edge-cases/index.js'),
};

async function main() {
  const [, , name = 'invoice'] = argv;

  if (!examples[name]) {
    // eslint-disable-next-line no-console
    console.error(`Unknown example "${name}". Available: ${Object.keys(examples).join(', ')}`);
    process.exit(1);
    return;
  }

  const loader = examples[name];
  const { runExample } = await loader();

  const __filename = fileURLToPath(import.meta.url);
  const __dirname = dirname(__filename);
  const repoRoot = join(__dirname, '..', '..', '..');

  await runExample({ repoRoot });
}

main().catch((err) => {
  // eslint-disable-next-line no-console
  console.error('Playground error:', err);
  process.exitCode = 1;
});
