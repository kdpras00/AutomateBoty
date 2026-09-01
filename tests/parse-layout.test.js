const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const scriptPath = path.join(__dirname, '..', 'script.js');
const source = fs.readFileSync(scriptPath, 'utf8');

// Ekstrak fungsi parseLayoutJson (+ dependensinya) dari script.js dan jalankan isolation.
function loadParseLayoutJson() {
    const funcSource = (name) => {
        const m = source.match(new RegExp('function ' + name + '\\([\\s\\S]*?\\n\\}'));
        assert.ok(m, name + ' harus ditemukan di script.js');
        return m[0];
    };
    const sandbox = { JSON, console };
    vm.createContext(sandbox);
    vm.runInContext(`${funcSource('extractBareLayoutJson')}\n${funcSource('parseLayoutJson')}`, sandbox);
    return sandbox.parseLayoutJson;
}

const parseLayoutJson = loadParseLayoutJson();

test('parseLayoutJson: blok JSON layout fenced yang valid', () => {
    const text = '```json\n{"layout": {"paperSize": "A4", "alignment": "justified"}}\n```\n\nKonten dokumen di sini.';
    const r = parseLayoutJson(text);
    assert.strictEqual(r.hasLayout, true);
    assert.strictEqual(r.hadLayoutBlock, true);
    assert.strictEqual(r.layoutCmds.paperSize, 'A4');
    assert.strictEqual(r.layoutCmds.alignment, 'justified');
    assert.strictEqual(r.cleanText, 'Konten dokumen di sini.');
    assert.ok(!r.cleanText.includes('json'));
});

test('parseLayoutJson: blok JSON layout polos (tanpa fenced) yang valid', () => {
    const text = '{"layout": {"font": "Times New Roman", "margins": {"top": 85}}}\n\nIsi paragraf.';
    const r = parseLayoutJson(text);
    assert.strictEqual(r.hasLayout, true);
    assert.strictEqual(r.layoutCmds.font, 'Times New Roman');
    assert.strictEqual(r.layoutCmds.margins.top, 85);
    assert.strictEqual(r.cleanText, 'Isi paragraf.');
});

test('parseLayoutJson: tidak ada blok layout → cleanText sama, hasLayout false', () => {
    const text = 'Hanya teks biasa tanpa format JSON.';
    const r = parseLayoutJson(text);
    assert.strictEqual(r.hasLayout, false);
    assert.strictEqual(r.hadLayoutBlock, false);
    assert.strictEqual(r.layoutCmds, null);
    assert.strictEqual(r.cleanText, text);
});

test('parseLayoutJson: JSON layout rusak → hadLayoutBlock true, hasLayout false', () => {
    const text = '```json\n{"layout": {"paperSize": "A4", }}  // bukan JSON valid\n```\n\nKonten.';
    const r = parseLayoutJson(text);
    assert.strictEqual(r.hadLayoutBlock, true);
    assert.strictEqual(r.hasLayout, false);
    assert.strictEqual(r.layoutCmds, null);
});

test('parseLayoutJson: layout non-objek (mis. string) tidak diterima', () => {
    const text = '```json\n{"layout": "A4"}\n```\n\nKonten.';
    const r = parseLayoutJson(text);
    assert.strictEqual(r.hasLayout, false);
    assert.strictEqual(r.layoutCmds, null);
});

test('parseLayoutJson: input non-string / kosong', () => {
    assert.strictEqual(parseLayoutJson(null).layoutCmds, null);
    assert.strictEqual(parseLayoutJson(undefined).layoutCmds, null);
    assert.strictEqual(parseLayoutJson('').cleanText, '');
});

// ── requestedWordFormatting ──────────────────────────────────────────────────
function loadRequestedWordFormatting() {
    const m = source.match(/function requestedWordFormatting\([\s\S]*?\n\}/);
    assert.ok(m, 'requestedWordFormatting harus ditemukan di script.js');
    const sandbox = { console };
    vm.createContext(sandbox);
    vm.runInContext(m[0], sandbox);
    return sandbox.requestedWordFormatting;
}

const requestedWordFormatting = loadRequestedWordFormatting();

test('requestedWordFormatting: mendeteksi permintaan format dokumen', () => {
    assert.strictEqual(requestedWordFormatting('Buatkan dokumen format A4 dengan margin 3-3-2.5-2.5 dan justify'), true);
    assert.strictEqual(requestedWordFormatting('tolong format tulisan 2 kolom font Times New Roman'), true);
    assert.strictEqual(requestedWordFormatting('alignment justification dan kertas a4'), true);
});

test('requestedWordFormatting: tidak memicu untuk prompt biasa', () => {
    assert.strictEqual(requestedWordFormatting('Buatkan ringkasan skripsi saya'), false);
    assert.strictEqual(requestedWordFormatting('Parafrasekan paragraf ini'), false);
    assert.strictEqual(requestedWordFormatting(null), false);
    assert.strictEqual(requestedWordFormatting(''), false);
});
