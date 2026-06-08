import { describe, it, expect } from 'vitest';
import { applyTransforms } from '../src/index';

describe('applyTransforms', () => {
  it('should handle upper and uppercase', () => {
    expect(applyTransforms('hello', ['upper'])).toBe('HELLO');
    expect(applyTransforms('hello', ['uppercase'])).toBe('HELLO');
  });

  it('should handle lower and lowercase', () => {
    expect(applyTransforms('HELLO', ['lower'])).toBe('hello');
    expect(applyTransforms('HELLO', ['lowercase'])).toBe('hello');
  });

  it('should handle capitalize', () => {
    expect(applyTransforms('hello', ['capitalize'])).toBe('Hello');
  });

  it('should handle trim', () => {
    expect(applyTransforms('  hello  ', ['trim'])).toBe('hello');
  });

  it('should handle trimstart', () => {
    expect(applyTransforms('  hello  ', ['trimstart'])).toBe('hello  ');
  });

  it('should handle trimend', () => {
    expect(applyTransforms('  hello  ', ['trimend'])).toBe('  hello');
  });

  it('should handle camel and camelcase', () => {
    expect(applyTransforms('hello world', ['camel'])).toBe('helloWorld');
    expect(applyTransforms('hello world', ['camelcase'])).toBe('helloWorld');
    expect(applyTransforms('Hello World', ['camelcase'])).toBe('helloWorld');
    expect(applyTransforms('hello-world', ['camelcase'])).toBe('helloWorld');
    expect(applyTransforms('hello_world', ['camelcase'])).toBe('helloWorld');
  });

  it('should handle pascal and pascalcase', () => {
    expect(applyTransforms('hello world', ['pascal'])).toBe('HelloWorld');
    expect(applyTransforms('hello world', ['pascalcase'])).toBe('HelloWorld');
    expect(applyTransforms('hello-world', ['pascalcase'])).toBe('HelloWorld');
    expect(applyTransforms('hello_world', ['pascalcase'])).toBe('HelloWorld');
  });

  it('should handle snake and snakecase', () => {
    expect(applyTransforms('helloWorld', ['snake'])).toBe('hello_world');
    expect(applyTransforms('helloWorld', ['snakecase'])).toBe('hello_world');
    expect(applyTransforms('hello world', ['snakecase'])).toBe('hello_world');
    expect(applyTransforms('hello-world', ['snakecase'])).toBe('hello_world');
  });

  it('should handle kebab and kebabcase', () => {
    expect(applyTransforms('helloWorld', ['kebab'])).toBe('hello-world');
    expect(applyTransforms('helloWorld', ['kebabcase'])).toBe('hello-world');
    expect(applyTransforms('hello world', ['kebabcase'])).toBe('hello-world');
    expect(applyTransforms('hello_world', ['kebabcase'])).toBe('hello-world');
  });

  it('should handle title and titlecase', () => {
    expect(applyTransforms('hello world', ['title'])).toBe('Hello World');
    expect(applyTransforms('hello world', ['titlecase'])).toBe('Hello World');
    expect(applyTransforms('hello-world', ['titlecase'])).toBe('Hello World');
    expect(applyTransforms('hello_world', ['titlecase'])).toBe('Hello World');
    expect(applyTransforms('helloWorld', ['titlecase'])).toBe('Hello World');
    expect(applyTransforms('  hello   world  ', ['titlecase'])).toBe('Hello World');
  });

  it('should handle initials', () => {
    expect(applyTransforms('John Doe', ['initials'])).toBe('JD');
    expect(applyTransforms('spark agent', ['initials'])).toBe('SA');
    expect(applyTransforms('Single', ['initials'])).toBe('S');
    expect(applyTransforms('  multi   space  ', ['initials'])).toBe('MS');
    expect(applyTransforms('', ['initials'])).toBe('');
    expect(applyTransforms(null, ['initials'])).toBe(null);
    expect(applyTransforms(123, ['initials'])).toBe(123);
  });

  it('should handle chained transformations', () => {
    expect(applyTransforms('  hello world  ', ['trim', 'camelcase', 'capitalize'])).toBe('HelloWorld');
  });

  it('should return original value for non-strings when using string-only transforms', () => {
    expect(applyTransforms(123, ['upper'])).toBe(123);
    expect(applyTransforms(null, ['upper'])).toBe(null);
  });

  it('should handle json transformation', () => {
    const obj = { a: 1, b: 'hello' };
    expect(applyTransforms(obj, ['json'])).toBe(JSON.stringify(obj, null, 2));
    expect(applyTransforms([1, 2, 3], ['json'])).toBe(JSON.stringify([1, 2, 3], null, 2));
    expect(applyTransforms('hello', ['json'])).toBe('"hello"');
  });

  it('should handle lines transformation', () => {
    expect(applyTransforms("line1\nline2\r\nline3", ['lines'])).toEqual(['line1', 'line2', 'line3']);
    expect(applyTransforms("single line", ['lines'])).toEqual(['single line']);
    expect(applyTransforms(123, ['lines'])).toBe(123);
  });

  it('should handle chained transformations with json', () => {
    const obj = { name: 'spark' };
    // json -> uppercase (no effect as json returns string, but uppercase works on it)
    const jsonStr = JSON.stringify(obj, null, 2);
    expect(applyTransforms(obj, ['json', 'uppercase'])).toBe(jsonStr.toUpperCase());
  });

  it('should handle join transformation', () => {
    expect(applyTransforms(['a', 'b', 'c'], ['join'])).toBe('a, b, c');
    expect(applyTransforms([], ['join'])).toBe('');
    expect(applyTransforms('not an array', ['join'])).toBe('not an array');
  });

  it('should handle unique transformation', () => {
    expect(applyTransforms(['a', 'b', 'a', 'c', 'b'], ['unique'])).toEqual(['a', 'b', 'c']);
    expect(applyTransforms([1, 2, 1, 3, 2], ['unique'])).toEqual([1, 2, 3]);
    expect(applyTransforms([], ['unique'])).toEqual([]);
    expect(applyTransforms('not an array', ['unique'])).toBe('not an array');
  });

  it('should handle first transformation', () => {
    expect(applyTransforms(['a', 'b', 'c'], ['first'])).toBe('a');
    expect(applyTransforms('hello', ['first'])).toBe('h');
    expect(applyTransforms([], ['first'])).toBeUndefined();
    expect(applyTransforms('', ['first'])).toBeUndefined();
  });

  it('should handle last transformation', () => {
    expect(applyTransforms(['a', 'b', 'c'], ['last'])).toBe('c');
    expect(applyTransforms('hello', ['last'])).toBe('o');
    expect(applyTransforms([], ['last'])).toBeUndefined();
    expect(applyTransforms('', ['last'])).toBeUndefined();
  });

  it('should handle length transformation', () => {
    expect(applyTransforms(['a', 'b', 'c'], ['length'])).toBe(3);
    expect(applyTransforms('hello', ['length'])).toBe(5);
    expect(applyTransforms([], ['length'])).toBe(0);
    expect(applyTransforms('', ['length'])).toBe(0);
    expect(applyTransforms({ a: 1, b: 2 }, ['length'])).toBe(2);
    expect(applyTransforms({}, ['length'])).toBe(0);
  });

  it('should handle keys transformation', () => {
    expect(applyTransforms({ a: 1, b: 2 }, ['keys'])).toEqual(['a', 'b']);
    expect(applyTransforms({}, ['keys'])).toEqual([]);
    expect(applyTransforms(['a', 'b'], ['keys'])).toEqual(['a', 'b']); // should skip arrays
    expect(applyTransforms('not an object', ['keys'])).toBe('not an object');
  });

  it('should handle values transformation', () => {
    expect(applyTransforms({ a: 1, b: 2 }, ['values'])).toEqual([1, 2]);
    expect(applyTransforms({}, ['values'])).toEqual([]);
    expect(applyTransforms(['a', 'b'], ['values'])).toEqual(['a', 'b']); // should skip arrays
    expect(applyTransforms('not an object', ['values'])).toBe('not an object');
  });

  it('should handle flat transformation', () => {
    expect(applyTransforms([[1, 2], [3, 4]], ['flat'])).toEqual([1, 2, 3, 4]);
    expect(applyTransforms([1, [2, [3]]], ['flat'])).toEqual([1, 2, [3]]);
    expect(applyTransforms([], ['flat'])).toEqual([]);
    expect(applyTransforms('not an array', ['flat'])).toBe('not an array');
  });

  it('should handle plural transformation', () => {
    expect(applyTransforms(0, ['plural'])).toBe('s');
    expect(applyTransforms(1, ['plural'])).toBe('');
    expect(applyTransforms(2, ['plural'])).toBe('s');
    expect(applyTransforms(1.5, ['plural'])).toBe('s');
    expect(applyTransforms('not a number', ['plural'])).toBe('not a number');
  });

  it('should handle round transformation', () => {
    expect(applyTransforms(1.4, ['round'])).toBe(1);
    expect(applyTransforms(1.5, ['round'])).toBe(2);
    expect(applyTransforms(1.6, ['round'])).toBe(2);
    expect(applyTransforms(-1.5, ['round'])).toBe(-1); // Math.round(-1.5) is -1
    expect(applyTransforms('not a number', ['round'])).toBe('not a number');
  });

  it('should handle floor transformation', () => {
    expect(applyTransforms(1.9, ['floor'])).toBe(1);
    expect(applyTransforms(-1.1, ['floor'])).toBe(-2);
    expect(applyTransforms('not a number', ['floor'])).toBe('not a number');
  });

  it('should handle ceil transformation', () => {
    expect(applyTransforms(1.1, ['ceil'])).toBe(2);
    expect(applyTransforms(-1.9, ['ceil'])).toBe(-1);
    expect(applyTransforms('not a number', ['ceil'])).toBe('not a number');
  });

  it('should handle abs transformation', () => {
    expect(applyTransforms(-5, ['abs'])).toBe(5);
    expect(applyTransforms(5, ['abs'])).toBe(5);
    expect(applyTransforms('not a number', ['abs'])).toBe('not a number');
  });

  it('should handle reverse transformation', () => {
    expect(applyTransforms(['a', 'b', 'c'], ['reverse'])).toEqual(['c', 'b', 'a']);
    expect(applyTransforms('hello', ['reverse'])).toBe('olleh');
    expect(applyTransforms([], ['reverse'])).toEqual([]);
    expect(applyTransforms('', ['reverse'])).toBe('');
    expect(applyTransforms(123, ['reverse'])).toBe(123);
  });

  it('should handle sort transformation', () => {
    expect(applyTransforms(['c', 'a', 'b'], ['sort'])).toEqual(['a', 'b', 'c']);
    expect(applyTransforms([3, 1, 2], ['sort'])).toEqual([1, 2, 3]);
    expect(applyTransforms([], ['sort'])).toEqual([]);
    expect(applyTransforms('not an array', ['sort'])).toBe('not an array');
  });

  it('should handle compact transformation', () => {
    expect(applyTransforms(['a', '', null, 'b', undefined, 'c'], ['compact'])).toEqual(['a', 'b', 'c']);
    expect(applyTransforms([], ['compact'])).toEqual([]);
    expect(applyTransforms('not an array', ['compact'])).toBe('not an array');
  });

  it('should handle sum transformation', () => {
    expect(applyTransforms([1, 2, 3], ['sum'])).toBe(6);
    expect(applyTransforms(['1', '2', '3'], ['sum'])).toBe(6);
    expect(applyTransforms([1, 'invalid', 2], ['sum'])).toBe(3);
    expect(applyTransforms([], ['sum'])).toBe(0);
    expect(applyTransforms('not an array', ['sum'])).toBe('not an array');
  });

  it('should handle avg transformation', () => {
    expect(applyTransforms([2, 4, 6], ['avg'])).toBe(4);
    expect(applyTransforms(['2', '4', '6'], ['avg'])).toBe(4);
    expect(applyTransforms([2, 'invalid', 4], ['avg'])).toBe(3);
    expect(applyTransforms([], ['avg'])).toBe(0);
    expect(applyTransforms('not an array', ['avg'])).toBe('not an array');
  });

  it('should handle min transformation', () => {
    expect(applyTransforms([2, 4, 1, 6], ['min'])).toBe(1);
    expect(applyTransforms(['2', '4', '1', '6'], ['min'])).toBe(1);
    expect(applyTransforms([2, 'invalid', 1], ['min'])).toBe(1);
    expect(applyTransforms([], ['min'])).toBeUndefined();
    expect(applyTransforms(['invalid'], ['min'])).toBeUndefined();
    expect(applyTransforms('not an array', ['min'])).toBe('not an array');
  });

  it('should handle max transformation', () => {
    expect(applyTransforms([2, 4, 1, 6], ['max'])).toBe(6);
    expect(applyTransforms(['2', '4', '1', '6'], ['max'])).toBe(6);
    expect(applyTransforms([2, 'invalid', 6], ['max'])).toBe(6);
    expect(applyTransforms([], ['max'])).toBeUndefined();
    expect(applyTransforms(['invalid'], ['max'])).toBeUndefined();
    expect(applyTransforms('not an array', ['max'])).toBe('not an array');
  });

  it('should handle empty transformation', () => {
    expect(applyTransforms(null, ['empty'])).toBe(true);
    expect(applyTransforms(undefined, ['empty'])).toBe(true);
    expect(applyTransforms('', ['empty'])).toBe(true);
    expect(applyTransforms([], ['empty'])).toBe(true);
    expect(applyTransforms('hello', ['empty'])).toBe(false);
    expect(applyTransforms(['a'], ['empty'])).toBe(false);
    expect(applyTransforms(0, ['empty'])).toBe(false);
    expect(applyTransforms(false, ['empty'])).toBe(false);
  });

  it('should handle notempty transformation', () => {
    expect(applyTransforms(null, ['notempty'])).toBe(false);
    expect(applyTransforms(undefined, ['notempty'])).toBe(false);
    expect(applyTransforms('', ['notempty'])).toBe(false);
    expect(applyTransforms([], ['notempty'])).toBe(false);
    expect(applyTransforms('hello', ['notempty'])).toBe(true);
    expect(applyTransforms(['a'], ['notempty'])).toBe(true);
    expect(applyTransforms(0, ['notempty'])).toBe(true);
    expect(applyTransforms(false, ['notempty'])).toBe(true);
  });

  it('should handle boolean transformation', () => {
    expect(applyTransforms(1, ['boolean'])).toBe(true);
    expect(applyTransforms(0, ['boolean'])).toBe(false);
    expect(applyTransforms('hello', ['boolean'])).toBe(true);
    expect(applyTransforms('', ['boolean'])).toBe(false);
    expect(applyTransforms(null, ['boolean'])).toBe(false);
    expect(applyTransforms(undefined, ['boolean'])).toBe(false);
    expect(applyTransforms([], ['boolean'])).toBe(true); // empty array is truthy in JS
  });
});
