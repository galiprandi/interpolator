// Shared logic lives here, for example:
// - Marker parser
// - Nested path resolver
// - Common types
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

export function applyTransforms(value: any, transforms: string[]): any {
  if (value == null || typeof value !== 'string') return value;

  let result = value;
  for (const t of transforms) {
    const transform = t.trim().toLowerCase();
    switch (transform) {
      case 'upper':
        result = result.toUpperCase();
        break;
      case 'lower':
        result = result.toLowerCase();
        break;
      case 'capitalize':
        result = result.charAt(0).toUpperCase() + result.slice(1);
        break;
      case 'trim':
        result = result.trim();
        break;
      case 'camelcase':
        result = result
          .trim()
          .replace(/[^a-zA-Z0-9]+(.)/g, (_, chr) => chr.toUpperCase())
          .replace(/[^a-zA-Z0-9]+$/, '')
          .replace(/^[A-Z]/, (chr) => chr.toLowerCase());
        break;
    }
  }
  return result;
}