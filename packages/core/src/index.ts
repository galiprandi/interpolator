// Aquí puedes poner lógica compartida como:
// - Parser de marcadores
// - Resolver de rutas anidadas
// - Tipos comunes
export function parseMarkers(template: string): Array<{ type: 'interpolation' | 'array'; path: string; full: string }> {
  const regex = /(\{\{|\[\[)\s*([^\]}]+)\s*(\}\}|\]\])/g;
  const matches: Array<{ type: 'interpolation' | 'array'; path: string; full: string }> = [];
  let match;

  while ((match = regex.exec(template)) !== null) {
    const [full, open, path, close] = match;
    matches.push({
      type: open === '{{' ? 'interpolation' : 'array',
      path,
      full
    });
  }

  return matches;
}

export function resolvePath(obj: any, path: string): { found: boolean; value: any } {
  const keys = path.split('.');
  let current = obj;

  for (const key of keys) {
    if (current == null || typeof current !== 'object') {
      return { found: false, value: undefined };
    }
    if (!(key in current)) {
      return { found: false, value: undefined };
    }
    current = current[key];
  }

  return { found: true, value: current };
}