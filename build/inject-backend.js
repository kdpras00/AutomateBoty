/*
 * inject-backend.js — Menyuntikkan URL backend ke index.html yang sudah di-build.
 *
 * BACKEND_URL (env) adalah URL server proxy (mis. https://your-api.up.railway.app).
 * Nilai ini menggantikan blok window.AB_CONFIG sehingga panggilan AI diarahkan
 * ke backend (API key aman di server), bukan langsung ke Google.
 */
const fs = require('fs');
const path = require('path');

function injectBackendUrl() {
    const indexPath = path.join(__dirname, '..', 'index.html');
    const backendUrl = (process.env.BACKEND_URL || '').trim();

    if (!backendUrl) {
        console.log('⚠️  BACKEND_URL kosong. AB_CONFIG dibiarkan default (BACKEND_URL="").');
        return;
    }

    let html = fs.readFileSync(indexPath, 'utf8');
    const pattern = /window\.AB_CONFIG\s*=\s*\{[^}]*BACKEND_URL:[^}]*\};/;
    const replacement = `window.AB_CONFIG = {\n            BACKEND_URL: "${backendUrl}"\n        };`;

    if (!pattern.test(html)) {
        throw new Error('AB_CONFIG block tidak ditemukan di index.html. Jalankan `npm run build` dulu.');
    }

    html = html.replace(pattern, replacement);
    fs.writeFileSync(indexPath, html);
    console.log(`✅ BACKEND_URL disuntikkan: ${backendUrl}`);
}

injectBackendUrl();
