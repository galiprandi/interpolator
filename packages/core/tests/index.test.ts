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
});
