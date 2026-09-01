/*
 * tests/build.test.js — Verifikasi integritas build & konsistensi referensi.
 * Memastikan index.html hasil build mengandung semua fitur inti dan bahwa
 * script.js tidak merujuk elemen DOM yang hilang (bug laten).
 */
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');

function read(p) {
    return fs.readFileSync(path.join(root, p), 'utf8');
}

test('index.html menyertakan semua panel & fitur inti', () => {
    const html = read('index.html');
    const required = [
        'lang-toggle-btn',
        'folder-upload',
        'academic-panel',
        'data-panel',
        'presentation-panel',
        'templates-panel',
        'history-panel',
        'settings-panel',
        'user-input',
        'send-btn',
        'upload-btn',
        'onboarding-modal',
        'dompurify',
        'pdf.min.js',
        'mammoth',
        'jszip',
        'window.AB_CONFIG',
    ];
    for (const token of required) {
        assert.ok(html.includes(token), `index.html harus mengandung "${token}"`);
    }
});

test('index.html tidak merujuk commands yang dihapus', () => {
    const html = read('index.html');
    assert.ok(!html.includes('commands.html'), 'index.html tidak boleh mereferensikan commands.html');
});

test('script.js tidak merujuk elemen DOM yang sudah dihapus', () => {
    const script = read('script.js');
    const removedRefs = ['word-tools-panel', 'ppt-tools-panel', 'host-tools-btn'];
    for (const ref of removedRefs) {
        assert.ok(!script.includes(ref), `script.js tidak boleh merujuk elemen yang sudah dihapus: ${ref}`);
    }
});

test('script.js tidak lagi memanggil Gemini langsung dari browser', () => {
    const script = read('script.js');
    assert.ok(!script.includes('x-goog-api-key'), 'Client tidak boleh mengirim API key');
    assert.ok(!script.includes('generativelanguage.googleapis.com'), 'Client tidak boleh memanggil Google langsung');
    assert.ok(script.includes('/api/gemini/chat'), 'Client harus memanggil endpoint proxy');
});

test('script.js mendukung parsing .pptx (PowerPoint)', () => {
    const script = read('script.js');
    assert.ok(script.includes('ext === "pptx" || ext === "ppt"'), 'harus ada handler .pptx/.ppt');
    assert.ok(script.includes('JSZip'), 'parser .pptx harus memakai JSZip');
    assert.ok(script.includes('extractPptxText'), 'harus ada helper ekstraksi teks slide');
    assert.ok(script.includes('slideNum'), 'harus ada helper nomor slide');
});

test('manifest.xml tidak mereferensikan commands.html', () => {
    const manifest = read('manifest.xml');
    assert.ok(!manifest.includes('commands.html'), 'manifest tidak boleh menyertakan commands.html');
});
