// test/base/excel-table.test.ts
import { describe, expect, it } from 'vitest';
import { Excel } from '../../src/index.js';

// Vitest makes 'describe', 'it', 'expect' (and others) global by default
// if you set `globals: true` in vitest.config.ts.

describe('Excel - Table', () => {
    console.log('describe');
    it('basic test', async() => {
        console.log('b');
        const _x = new Excel();
        expect(true).toBe(true);
    });
});
