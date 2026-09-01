const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const scriptPath = path.join(__dirname, '..', 'script.js');
const source = fs.readFileSync(scriptPath, 'utf8');

// Ekstrak fungsi parseSlides dari script.js dan jalankan dalam isolation.
function loadParseSlides() {
    const m = source.match(/function parseSlides\([\s\S]*?\n\}/);
    assert.ok(m, 'parseSlides harus ditemukan di script.js');
    const sandbox = { JSON, Array, String, console };
    vm.createContext(sandbox);
    vm.runInContext(m[0], sandbox);
    return sandbox.parseSlides;
}

const parseSlides = loadParseSlides();

// Array hasil dari vm sandbox punya prototype berbeda (cross-realm),
// jadi normalisasi lewat JSON agar perbandingan konsisten.
const norm = (v) => JSON.parse(JSON.stringify(v));

test('parseSlides: JSON Array slide dengan title/points/notes', () => {
    const text = '[{"title":"Judul 1","points":["A","B"],"notes":"Catatan 1"},{"title":"Judul 2","points":["C"],"notes":"Catatan 2"}]';
    const slides = norm(parseSlides(text));
    assert.strictEqual(slides.length, 2);
    assert.strictEqual(slides[0].title, 'Judul 1');
    assert.deepStrictEqual(slides[0].points, ['A', 'B']);
    assert.strictEqual(slides[0].notes, 'Catatan 1');
    assert.strictEqual(slides[1].title, 'Judul 2');
});

test('parseSlides: fallback TITLE: multi-slide', () => {
    const text = 'TITLE: Pendahuluan\n- Latar belakang\n- Rumusan masalah\n\nTITLE: Metode\n- Kualitatif\n- Wawancara';
    const slides = norm(parseSlides(text));
    assert.strictEqual(slides.length, 2);
    assert.strictEqual(slides[0].title, 'Pendahuluan');
    assert.deepStrictEqual(slides[0].points, ['Latar belakang', 'Rumusan masalah']);
    assert.strictEqual(slides[1].title, 'Metode');
    assert.deepStrictEqual(slides[1].points, ['Kualitatif', 'Wawancara']);
});

test('parseSlides: fallback heading markdown multi-slide', () => {
    const text = '## Pendahuluan\n- A\n- B\n\n## Kesimpulan\n- C';
    const slides = norm(parseSlides(text));
    assert.strictEqual(slides.length, 2);
    assert.strictEqual(slides[0].title, 'Pendahuluan');
    assert.deepStrictEqual(slides[0].points, ['A', 'B']);
    assert.strictEqual(slides[1].title, 'Kesimpulan');
});

test('parseSlides: fallback SLIDE n: marker', () => {
    const text = 'SLIDE 1: Judul Satu\n- poin satu\nSLIDE 2: Judul Dua\n- poin dua';
    const slides = norm(parseSlides(text));
    assert.strictEqual(slides.length, 2);
    assert.strictEqual(slides[0].title, 'Judul Satu');
    assert.strictEqual(slides[1].title, 'Judul Dua');
});

test('parseSlides: titik (bullets) dibersihkan dari "-"/"*"/"•"', () => {
    const text = 'TITLE: Intro\n• poin satu\n- poin dua\n* poin tiga';
    const slides = norm(parseSlides(text));
    assert.deepStrictEqual(slides[0].points, ['poin satu', 'poin dua', 'poin tiga']);
});

test('parseSlides: tanpa marker → satu slide (baris pertama judul)', () => {
    const text = 'Tren Market 2026\n- Pertumbuhan 20%\n- Go digital';
    const slides = norm(parseSlides(text));
    assert.strictEqual(slides.length, 1);
    assert.strictEqual(slides[0].title, 'Tren Market 2026');
    assert.deepStrictEqual(slides[0].points, ['Pertumbuhan 20%', 'Go digital']);
});

test('parseSlides: JSON rusak → jatuh ke fallback teks, tidak crash', () => {
    const text = '[{"title":"A","points":["1"]} malformed\nTITLE: Fallback\n- x';
    const slides = norm(parseSlides(text));
    assert.ok(slides.length >= 1);
    assert.strictEqual(slides[0].title, 'Fallback');
});

test('parseSlides: input kosong / null → array kosong', () => {
    assert.deepStrictEqual(norm(parseSlides('')), []);
    assert.deepStrictEqual(norm(parseSlides(null)), []);
});
