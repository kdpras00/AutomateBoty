const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const scriptPath = path.join(__dirname, '..', 'script.js');
const source = fs.readFileSync(scriptPath, 'utf8');

function loadParseNumericCell() {
    const m = source.match(/function parseNumericCell\([\s\S]*?\n\}/);
    assert.ok(m, 'parseNumericCell harus ditemukan di script.js');
    const sandbox = { Number, isFinite, console };
    vm.createContext(sandbox);
    vm.runInContext(m[0], sandbox);
    return sandbox.parseNumericCell;
}

const parseNumericCell = loadParseNumericCell();

test('parseNumericCell: desimal titik (format US)', () => {
    assert.strictEqual(parseNumericCell('3.14'), 3.14);
    assert.strictEqual(parseNumericCell('0.5'), 0.5);
    assert.strictEqual(parseNumericCell('-2.75'), -2.75);
});

test('parseNumericCell: desimal koma (format Indonesia)', () => {
    assert.strictEqual(parseNumericCell('3,14'), 3.14);
    assert.strictEqual(parseNumericCell('0,5'), 0.5);
});

test('parseNumericCell: ribuan dengan koma (format US) & koma+nol', () => {
    assert.strictEqual(parseNumericCell('1,234'), 1234);
    assert.strictEqual(parseNumericCell('2,500'), 2500);
});

test('parseNumericCell: titik tunggal = desimal (bukan ribuan)', () => {
    // "12.345" ambigu; perlakuan aman = desimal (mencegah salah korupsi 3.14)
    assert.strictEqual(parseNumericCell('12.345'), 12.345);
});

test('parseNumericCell: ribuan + desimal gabungan', () => {
    assert.strictEqual(parseNumericCell('1.234,56'), 1234.56);
    assert.strictEqual(parseNumericCell('1,234.56'), 1234.56); // fallback titik-desimal
});

test('parseNumericCell: angka murni number sudah numeric', () => {
    assert.strictEqual(parseNumericCell(42), 42);
    assert.strictEqual(parseNumericCell(3.14), 3.14);
    assert.strictEqual(parseNumericCell(NaN), null);
    assert.strictEqual(parseNumericCell(Infinity), null);
});

test('parseNumericCell: nilai non-angka / kosong / formula → null atau dijaga', () => {
    assert.strictEqual(parseNumericCell(''), null);
    assert.strictEqual(parseNumericCell('   '), null);
    assert.strictEqual(parseNumericCell('abc'), null);
    assert.strictEqual(parseNumericCell(null), null);
    assert.strictEqual(parseNumericCell(undefined), null);
    // Formula dikembalikan apa adanya di pemanggil, di sini parseNumericCell tidak
    // boleh salah menilai "=SUM(...)" sebagai angka.
    assert.strictEqual(parseNumericCell('=SUM(A1)'), null);
});

test('parseNumericCell: teks berawalan angka namun bukan angka murni', () => {
    // "2026 OK" = bukan angka murni → null (dibiarkan sebagai teks)
    assert.strictEqual(parseNumericCell('2026 OK'), null);
});
