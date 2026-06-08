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
  let result = value;
  for (const t of transforms) {
    const transform = t.trim().toLowerCase();
    switch (transform) {
      case 'upper':
      case 'uppercase':
        if (typeof result === 'string') result = result.toUpperCase();
        break;
      case 'lower':
      case 'lowercase':
        if (typeof result === 'string') result = result.toLowerCase();
        break;
      case 'capitalize':
        if (typeof result === 'string') result = result.charAt(0).toUpperCase() + result.slice(1);
        break;
      case 'trim':
        if (typeof result === 'string') result = result.trim();
        break;
      case 'trimstart':
        if (typeof result === 'string') result = result.trimStart();
        break;
      case 'trimend':
        if (typeof result === 'string') result = result.trimEnd();
        break;
      case 'camel':
      case 'camelcase':
        if (typeof result === 'string') {
          result = result
            .trim()
            .replace(/[^a-zA-Z0-9]+(.)/g, (_, chr) => chr.toUpperCase())
            .replace(/[^a-zA-Z0-9]+$/, '')
            .replace(/^[A-Z]/, (chr) => chr.toLowerCase());
        }
        break;
      case 'pascal':
      case 'pascalcase':
        if (typeof result === 'string') {
          result = result
            .trim()
            .replace(/[^a-zA-Z0-9]+(.)/g, (_, chr) => chr.toUpperCase())
            .replace(/[^a-zA-Z0-9]+$/, '')
            .replace(/^[a-z]/, (chr) => chr.toUpperCase());
        }
        break;
      case 'snake':
      case 'snakecase':
        if (typeof result === 'string') {
          result = result
            .trim()
            .replace(/([a-z])([A-Z])/g, '$1_$2')
            .replace(/[^a-zA-Z0-9]+/g, '_')
            .toLowerCase()
            .replace(/^_+|_+$/g, '');
        }
        break;
      case 'kebab':
      case 'kebabcase':
        if (typeof result === 'string') {
          result = result
            .trim()
            .replace(/([a-z])([A-Z])/g, '$1-$2')
            .replace(/[^a-zA-Z0-9]+/g, '-')
            .toLowerCase()
            .replace(/^-+|-+$/g, '');
        }
        break;
      case 'title':
      case 'titlecase':
        if (typeof result === 'string') {
          result = result
            .trim()
            .replace(/([a-z])([A-Z])/g, '$1 $2')
            .replace(/[^a-zA-Z0-9]+/g, ' ')
            .replace(/\b([a-z])/g, (_, chr) => chr.toUpperCase())
            .trim();
        }
        break;
      case 'initials':
        if (typeof result === 'string') {
          result = result
            .trim()
            .split(/\s+/)
            .filter((word) => word.length > 0)
            .map((word) => word.charAt(0).toUpperCase())
            .join('');
        }
        break;
      case 'json':
        result = JSON.stringify(result, null, 2);
        break;
      case 'lines':
        if (typeof result === 'string') result = result.split(/\r?\n/);
        break;
      case 'join':
        if (Array.isArray(result)) result = result.join(', ');
        break;
      case 'unique':
        if (Array.isArray(result)) result = Array.from(new Set(result));
        break;
      case 'first':
        if (Array.isArray(result) || typeof result === 'string') {
          result = result.length > 0 ? result[0] : undefined;
        }
        break;
      case 'last':
        if (Array.isArray(result) || typeof result === 'string') {
          result = result.length > 0 ? result[result.length - 1] : undefined;
        }
        break;
      case 'length':
        if (Array.isArray(result) || typeof result === 'string') {
          result = result.length;
        } else if (result != null && typeof result === 'object') {
          result = Object.keys(result).length;
        }
        break;
      case 'keys':
        if (result != null && typeof result === 'object' && !Array.isArray(result)) {
          result = Object.keys(result);
        }
        break;
      case 'values':
        if (result != null && typeof result === 'object' && !Array.isArray(result)) {
          result = Object.values(result);
        }
        break;
      case 'flat':
        if (Array.isArray(result)) {
          result = result.flat();
        }
        break;
      case 'plural':
        if (typeof result === 'number') result = result === 1 ? '' : 's';
        break;
      case 'round':
        if (typeof result === 'number') result = Math.round(result);
        break;
      case 'floor':
        if (typeof result === 'number') result = Math.floor(result);
        break;
      case 'ceil':
        if (typeof result === 'number') result = Math.ceil(result);
        break;
      case 'abs':
        if (typeof result === 'number') result = Math.abs(result);
        break;
      case 'reverse':
        if (Array.isArray(result)) {
          result = [...result].reverse();
        } else if (typeof result === 'string') {
          result = [...result].reverse().join('');
        }
        break;
      case 'sort':
        if (Array.isArray(result)) {
          result = [...result].sort();
        }
        break;
      case 'compact':
        if (Array.isArray(result)) {
          result = result.filter((item) => item !== null && item !== undefined && item !== '');
        }
        break;
      case 'sum':
        if (Array.isArray(result)) {
          result = result.reduce((acc, item) => {
            const num = Number(item);
            return Number.isNaN(num) ? acc : acc + num;
          }, 0);
        }
        break;
      case 'avg':
        if (Array.isArray(result)) {
          const nums = result.map(Number).filter((n) => !Number.isNaN(n));
          result = nums.length > 0 ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
        }
        break;
      case 'min':
        if (Array.isArray(result)) {
          const nums = result.map(Number).filter((n) => !Number.isNaN(n));
          result = nums.length > 0 ? Math.min(...nums) : undefined;
        }
        break;
      case 'max':
        if (Array.isArray(result)) {
          const nums = result.map(Number).filter((n) => !Number.isNaN(n));
          result = nums.length > 0 ? Math.max(...nums) : undefined;
        }
        break;
      case 'empty':
        result =
          result === null ||
          result === undefined ||
          result === '' ||
          (Array.isArray(result) && result.length === 0);
        break;
      case 'notempty':
        result = !(
          result === null ||
          result === undefined ||
          result === '' ||
          (Array.isArray(result) && result.length === 0)
        );
        break;
      case 'boolean':
        result = Boolean(result);
        break;
    }
  }
  return result;
}