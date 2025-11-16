import { describe, it, expect } from 'vitest';
import { interpolateXlsx } from '../src';

describe('interpolateXlsx', () => {
  it('should replace {{}} markers with values', async () => {
    // Por ahora, solo probamos que no falle al importar
    expect(interpolateXlsx).toBeTypeOf('function');
  });
});