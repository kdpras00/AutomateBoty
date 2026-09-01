require('dotenv').config();
const express = require('express');
const cors = require('cors');
const rateLimit = require('express-rate-limit');

const app = express();
const PORT = process.env.PORT || 3000;

// Origin yang diizinkan — default origin Netlify, bisa di-override via env ALLOWED_ORIGIN.
const allowedOrigins = (process.env.ALLOWED_ORIGIN || 'https://automateboty.netlify.app')
    .split(',')
    .map((o) => o.trim())
    .filter(Boolean);

// Whitelist model yang boleh dipakai — cegah manipulasi URL di sisi client.
const SUPPORTED_MODELS = ['gemini-2.5-flash', 'gemini-1.5-pro', 'gemini-1.5-flash'];
const DEFAULT_MODEL = 'gemini-2.5-flash';

// Middleware: batasi sumber permintaan
app.use(cors({
    origin(origin, callback) {
        // Izinkan non-browser (mis. curl) dan origin dari daftar yang diizinkan.
        if (!origin || allowedOrigins.includes(origin)) return callback(null, true);
        return callback(new Error('Origin tidak diizinkan oleh CORS'));
    },
    methods: ['GET', 'POST', 'OPTIONS'],
}));
app.use(express.json({ limit: '10mb' })); // Batas payload (base64 dokumen besar)

// Rate limiting: cegah penyalahgunaan gateway sebagai proxy gratis.
const limiter = rateLimit({
    windowMs: 60 * 1000, // 1 menit
    max: Number(process.env.RATE_LIMIT_MAX) || 30, // maks 30 request/menit/IP
    standardHeaders: true,
    legacyHeaders: false,
    message: { error: 'Terlalu banyak permintaan. Coba lagi dalam 1 menit.' },
});
app.use('/api/', limiter);

// Health Check
app.get('/', (req, res) => {
    res.json({
        message: '✅ AutomateBoty Backend Proxy is running normally.',
        status: 'ok',
        time: new Date().toISOString(),
    });
});

// Endpoint Chat (satu-satunya jembatan ke Gemini, API key tetap di server)
app.post('/api/gemini/chat', async (req, res) => {
    try {
        const { model, payload } = req.body;
        const apiKey = process.env.GEMINI_API_KEY;

        if (!apiKey) {
            return res.status(500).json({ error: 'Server backend belum mengonfigurasi GEMINI_API_KEY di .env' });
        }

        // Validasi model (whitelist) untuk mencegah path-injection di URL Google.
        const selectedModel = SUPPORTED_MODELS.includes(model) ? model : DEFAULT_MODEL;
        if (model && !SUPPORTED_MODELS.includes(model)) {
            console.warn(`Model tidak dikenal ditolak: ${model}`);
        }

        if (!payload || !payload.contents) {
            return res.status(400).json({ error: 'Payload tidak valid: field "contents" wajib ada' });
        }

        const url = `https://generativelanguage.googleapis.com/v1beta/models/${selectedModel}:generateContent`;

        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'x-goog-api-key': apiKey, // KEY AMAN DI HEADER BACKEND
            },
            body: JSON.stringify(payload),
        });

        const data = await response.json();

        if (!response.ok) {
            console.error(`Gemini API error ${response.status}`);
            return res.status(response.status).json(data);
        }

        res.json(data);
    } catch (error) {
        console.error('Kesalahan Proxy API: ', error);
        res.status(500).json({ error: 'Terjadi kesalahan internal server.' });
    }
});

// Middleware error untuk menangani kesalahan CORS/rate-limit secara rapi.
// Catatan: signature 4 argumen (termasuk next) WAJIB agar Express mengenali
// ini sebagai error-handling middleware.
// eslint-disable-next-line no-unused-vars
app.use((err, req, res, next) => {
    if (err && err.message && err.message === 'Origin tidak diizinkan oleh CORS') {
        return res.status(403).json({ error: err.message });
    }
    console.error('Unhandled error:', err);
    res.status(500).json({ error: 'Terjadi kesalahan internal server.' });
});

// Hanya mulai server ketika dijalankan langsung (bukan ketika di-require untuk test).
if (require.main === module) {
    app.listen(PORT, () => {
        console.log(`✅ Backend AutomateBoty berjalan & melindungi API pada port http://localhost:${PORT}`);
        console.log(`   Allowed origins: ${allowedOrigins.join(', ')}`);
        console.log(`   Supported models: ${SUPPORTED_MODELS.join(', ')}`);
    });
}

module.exports = { app, SUPPORTED_MODELS, DEFAULT_MODEL };
