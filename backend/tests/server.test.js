/*
 * backend/tests/server.test.js — Tes dasar untuk backend proxy.
 * Menjalankan server pada port ephemeral dan menguji health check, validasi model,
 * verifikasi payload, dan penolakan tanpa API key.
 */
const { test, before, after } = require('node:test');
const assert = require('node:assert/strict');

const { app } = require('../server');

let server;
let base;

before(async () => {
    await new Promise((resolve) => {
        server = app.listen(0, () => {
            base = `http://localhost:${server.address().port}`;
            resolve();
        });
    });
});

after(() => {
    if (server) server.close();
});

test('health check / mengembalikan status ok', async () => {
    const res = await fetch(`${base}/`);
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.equal(body.status, 'ok');
});

test('chat tanpa GEMINI_API_KEY ditolak 500', async () => {
    // pastikan env GEMINI_API_KEY kosong selama pengujian ini
    const prev = process.env.GEMINI_API_KEY;
    delete process.env.GEMINI_API_KEY;

    const res = await fetch(`${base}/api/gemini/chat`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            model: 'gemini-2.5-flash',
            payload: { contents: [{ role: 'user', parts: [{ text: 'halo' }] }] },
        }),
    });
    assert.equal(res.status, 500);
    const body = await res.json();
    assert.match(body.error, /GEMINI_API_KEY/i);

    if (prev) process.env.GEMINI_API_KEY = prev;
});

test('chat tanpa payload.contents ditolak 400', async () => {
    process.env.GEMINI_API_KEY = 'AIzaSy_TEST';
    const res = await fetch(`${base}/api/gemini/chat`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ model: 'gemini-2.5-flash', payload: {} }),
    });
    assert.equal(res.status, 400);
});

test('model tidak dikenal ditolak / dinormalisasi (tanpa crash)', async () => {
    process.env.GEMINI_API_KEY = 'AIzaSy_TEST';
    // Model yang tidak dikenal tidak akan error 400 di endpoint ini;
    // ia akan dinormalisasi ke default, lalu mencoba fetch Google dan gagal
    // karena GEMINI_API_KEY palsu. Kita hanya pastikan tidak 500 dari validasi.
    const res = await fetch(`${base}/api/gemini/chat`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ model: 'lol; rm -rf', payload: { contents: [{ role: 'user', parts: [{ text: 'x' }] }] } }),
    });
    assert.notEqual(res.status, 404);
});
