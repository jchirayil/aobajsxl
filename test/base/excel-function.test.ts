// test/base/excel-function.test.ts
import { expect } from 'chai';
import { Excel } from '../../src/index';
import fs from 'fs';
import path from 'path';

describe('Excel Functions', () => {
    let excelBase: Excel;
    beforeEach(() => {
        excelBase = new Excel();
    });

    it('should write incorrect formula as is (string)', async () => {
        const filePath = path.join(__dirname, '../data/test-error-formula.xlsx'); // Construct the file path
        excelBase['setData']('FormulaSheet', [
            {id: 1, operation: 'multiplication', total: '=*3'},
            {id: 2, operation: 'addition', total: '=5+2+4'}, 
            {id: 3, operation: 'subtraction', total: '=152-24'},
            {id: 4, operation: 'division', total: '=100/5'},
            {id: 5, operation: 'concatenate', total: '=CONCATENATE("Hello", " ", "World")'}
        ]);
        await excelBase.write(filePath);
        expect(fs.existsSync(filePath)).to.be.true;
        //fs.unlinkSync(filePath);
      });
});