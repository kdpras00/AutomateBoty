require('dotenv').config();
const express = require('express');
const cors = require('cors');
const fetch = require('node-fetch');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors()); // Mencegah pemblokiran CORS dari Add-in Office
app.use(express.json({ limit: '10mb' })); // Batas payload kalau ada base64 dokumen besar

// Endpoint Rahasia untuk Chat
app.post('/api/gemini/chat', async (req, res) => {
    try {
        const { model, payload } = req.body;
        const apiKey = process.env.GEMINI_API_KEY;

        if (!apiKey || apiKey === "AIzaSy_MASUKKAN_KEY_RAHASIA_ANDA_DISINI") {
            return res.status(500).json({ error: "Server Backend belum mengonfigurasi API Key di .env" });
        }

        // Endpoint Google (hanya bisa diakses Backend) -> tidak mengekspos API Key ke user
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${model || 'gemini-2.5-flash'}:generateContent`;

        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'x-goog-api-key': apiKey // KEY AMAN DI HEADER BACKEND
            },
            body: JSON.stringify(payload) // Teruskan instruksi dan riwayat dari AutomateBoty Add-in
        });

        const data = await response.json();

        // Teruskan data jawaban murni ke pengguna
        if (!response.ok) {
            return res.status(response.status).json(data);
        }

        res.json(data);
    } catch (error) {
        console.error("Kesalahan Proxy API: ", error);
        res.status(500).json({ error: "Terjadi kesalahan internal server." });
    }
});

app.listen(PORT, () => {
    console.log(`✅ Backend AutomateBoty berjalan dan melindungi API pada port http://localhost:${PORT}`);
});
