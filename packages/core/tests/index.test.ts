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

  it('should handle camelcase', () => {
    expect(applyTransforms('hello world', ['camelcase'])).toBe('helloWorld');
    expect(applyTransforms('Hello World', ['camelcase'])).toBe('helloWorld');
    expect(applyTransforms('hello-world', ['camelcase'])).toBe('helloWorld');
    expect(applyTransforms('hello_world', ['camelcase'])).toBe('helloWorld');
  });

  it('should handle pascalcase', () => {
    expect(applyTransforms('hello world', ['pascalcase'])).toBe('HelloWorld');
    expect(applyTransforms('hello-world', ['pascalcase'])).toBe('HelloWorld');
    expect(applyTransforms('hello_world', ['pascalcase'])).toBe('HelloWorld');
  });

  it('should handle snakecase', () => {
    expect(applyTransforms('helloWorld', ['snakecase'])).toBe('hello_world');
    expect(applyTransforms('hello world', ['snakecase'])).toBe('hello_world');
    expect(applyTransforms('hello-world', ['snakecase'])).toBe('hello_world');
  });

  it('should handle kebabcase', () => {
    expect(applyTransforms('helloWorld', ['kebabcase'])).toBe('hello-world');
    expect(applyTransforms('hello world', ['kebabcase'])).toBe('hello-world');
    expect(applyTransforms('hello_world', ['kebabcase'])).toBe('hello-world');
  });

  it('should handle titlecase', () => {
    expect(applyTransforms('hello world', ['titlecase'])).toBe('Hello World');
    expect(applyTransforms('hello-world', ['titlecase'])).toBe('Hello World');
    expect(applyTransforms('hello_world', ['titlecase'])).toBe('Hello World');
    expect(applyTransforms('helloWorld', ['titlecase'])).toBe('Hello World');
    expect(applyTransforms('  hello   world  ', ['titlecase'])).toBe('Hello World');
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
    expect(applyTransforms('hello', ['reverse'])).toBe('olleh');
    expect(applyTransforms('💩', ['reverse'])).toBe('💩'); // Unicode surrogate pairs (U+1F4A9)
    expect(applyTransforms(['a', 'b', 'c'], ['reverse'])).toEqual(['c', 'b', 'a']);
    expect(applyTransforms([], ['reverse'])).toEqual([]);
    expect(applyTransforms(123, ['reverse'])).toBe(123);
  });

  it('should handle sort transformation', () => {
    expect(applyTransforms(['c', 'a', 'b'], ['sort'])).toEqual(['a', 'b', 'c']);
    expect(applyTransforms([3, 1, 2], ['sort'])).toEqual([1, 2, 3]);
    expect(applyTransforms([], ['sort'])).toEqual([]);
    expect(applyTransforms('not an array', ['sort'])).toBe('not an array');
  });

  it('should handle compact transformation', () => {
    expect(applyTransforms(['a', null, 'b', undefined, '', 'c'], ['compact'])).toEqual(['a', 'b', 'c']);
    expect(applyTransforms([1, 0, false, null], ['compact'])).toEqual([1, 0, false]);
    expect(applyTransforms([], ['compact'])).toEqual([]);
  });

  it('should handle sum transformation', () => {
    expect(applyTransforms([1, 2, 3, 4], ['sum'])).toBe(10);
    expect(applyTransforms([1, 'a', 2, null, 3], ['sum'])).toBe(6);
    expect(applyTransforms([], ['sum'])).toBe(0);
    expect(applyTransforms(['a', 'b'], ['sum'])).toBe(0);
    expect(applyTransforms('not an array', ['sum'])).toBe('not an array');
  });

  it('should handle avg transformation', () => {
    expect(applyTransforms([1, 2, 3, 4], ['avg'])).toBe(2.5);
    expect(applyTransforms([1, 'a', 2, null, 3], ['avg'])).toBe(2); // (1+2+3)/3 = 2
    expect(applyTransforms([], ['avg'])).toBe(0);
    expect(applyTransforms(['a', 'b'], ['avg'])).toBe(0);
    expect(applyTransforms('not an array', ['avg'])).toBe('not an array');
  });
});
