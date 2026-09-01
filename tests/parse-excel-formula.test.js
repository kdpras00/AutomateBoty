const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const scriptPath = path.join(__dirname, '..', 'features-excel.js');
const source = fs.readFileSync(scriptPath, 'utf8');

function loadShiftFormulaRows() {
    const m = source.match(/function shiftFormulaRows\([\s\S]*?\n\}/);
    assert.ok(m, 'shiftFormulaRows harus ditemukan di features-excel.js');
    const sandbox = { parseInt };
    vm.createContext(sandbox);
    vm.runInContext(m[0], sandbox);
    return sandbox.shiftFormulaRows;
}

const shiftFormulaRows = loadShiftFormulaRows();

test('shiftFormulaRows: geser referensi baris pada SUM', () => {
    assert.strictEqual(shiftFormulaRows('=SUM(C2:G2)', 5), '=SUM(C7:G7)');
    assert.strictEqual(shiftFormulaRows('=SUM(B2:D2)', 0), '=SUM(B2:D2)');
});

test('shiftFormulaRows: rumus pembagian / rerata ikut digeser', () => {
    assert.strictEqual(shiftFormulaRows('=H2/SUM(C2:G2)*5', 3), '=H5/SUM(C5:G5)*5');
    assert.strictEqual(shiftFormulaRows('=AVERAGE(B:B)', 1), '=AVERAGE(B:B)');
});

test('shiftFormulaRows: bukan rumus ("=") dikembalikan apa adanya', () => {
    assert.strictEqual(shiftFormulaRows('Variabel 1', 0), 'Variabel 1');
    assert.strictEqual(shiftFormulaRows('', 9), '');
    assert.strictEqual(shiftFormulaRows(null, 5), null);
});

test('shiftFormulaRows: kolom banyak huruf tetap aman', () => {
    assert.strictEqual(shiftFormulaRows('=AA2+AB3', 4), '=AA6+AB7');
});
