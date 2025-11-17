import { argv } from 'node:process';

interface ExampleModule {
  runExample(opts: { repoRoot: string }): Promise<void>;
}

const examples: Record<string, () => Promise<ExampleModule>> = {
  invoice: () => import('../examples/invoice/index.js'),
  'invoice-2': () => import('../examples/invoice-2/index.js'),
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

  await runExample({ repoRoot: process.cwd() });
}

main().catch((err) => {
  // eslint-disable-next-line no-console
  console.error('Playground error:', err);
  process.exitCode = 1;
});
