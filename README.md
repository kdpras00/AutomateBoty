# 🚀 AutomateBoty

**Asisten AI Google Gemini untuk Microsoft Word, Excel, dan PowerPoint** — mempercepat penulisan akademik, analisis data, dan pembuatan presentasi langsung dari Office.

> **Status:** v7.0 · Referensi lengkap: [docs/PRD.md](docs/PRD.md)

---

## ✨ Fitur Utama

### Chat AI Terpadu (Word/Excel/PPT)
- 💬 Konsultasi AI dengan **konteks dokumen aktif** otomatis
- 📎 **Upload file**: gambar, `.docx`, `.pdf`, `.txt/.csv/.md/.json/.js/.py`
- 📂 **Upload folder**: baca banyak bab skripsi + **deteksi BAB otomatis** + saran lanjut BAB
- 🛡️ **Mode offline** (fallback cache), retry/backoff untuk 429/503, mode bahasa ID/EN

### Word Pro (Academic)
- ✂️ Parafrase anti-plagiarisme (ringan/sedang/ketat)
- 🔍 Proofreading mendalam (laporan ejaan, kalimat, kata tidak baku)
- 🎓 Mode **Bimbingan Skripsi** (AI sebagai dosen pembimbing)
- 📐 Outline Builder interaktif + ekspansi bab
- 📋 **6 template**: Jurnal ID/EN, Artikel, Skripsi 5 bab, Prosiding, Makalah
- 🔖 Sitasi otomatis: **APA / IEEE / Chicago / Vancouver**
- ⌨️ Auto-format dokumen (A4, margin, kolom, font, justify) via JSON blok

### Excel Pro (Data)
- 📊 Statistik deskriptif (N, Mean, Median, Std Dev, Min, Max, Range)
- 📈 **Regresi & Korelasi** (b₀, b₁, r, R²) + sisip tabel hasil + narasi Bab IV
- 📝 Interpretasi statistik ke narasi akademik 300–500 kata
- 📋 Template tabel: Frekuensi, Distribusi, Crosstab, Kuesioner, Rangkuman (dengan rumus)
- ✅ Validasi data + dropdown programmatis

### Slides Pro (PowerPoint)
- 🎯 **Slide dari file** (upload → JSON slide + speaker notes)
- 🎨 4 tema desain (Akademik Biru/Hijau, Dark Pro, Minimalis)
- 🎤 **Timer latihan** dengan kode warna + tips presentasi

### Ekstra
- 🗂️ Riwayat percakapan (max 10 sesi)
- ⭐ Prompt presets
- 👍 Rating & salin per pesan
- 🎓 Onboarding tour 5 langkah

---

## 🏗️ Arsitektur

```
┌───────────────────────┐        ┌──────────────────────┐        ┌─────────────────┐
│  Office Add-in        │  HTTPS │  Backend Proxy       │  HTTPS │  Google Gemini  │
│  index.html           │───────▶│  /api/gemini/chat    │───────▶│  v1beta          │
│  (Netlify static)     │        │  (Node/Express)      │        │  generateContent │
└───────────────────────┘        └──────────────────────┘        └─────────────────┘
        ▲                                 │
        │                                API key HANYA di sini
```

### Keamanan API Key 🔐
- **Sebelum (bug kritis):** API key dikirim langsung dari browser → terekspos.
- **Sesudah (perbaikan):** Semua panggilan AI melewati **backend proxy**. API key **hanya** tersimpan di server (`GEMINI_API_KEY`), tidak pernah di browser/devtools.
- Backend mem-*whitelist* model, membatasi CORS & rate limit.

---

## 🧑‍💻 Pengaturan Lokal

### Prasyarat
- Node.js ≥ 18

### 1. Pasang dependensi
```bash
npm install                 # root (tooling penuh)
cd backend && npm install   # backend (express, cors, rate-limit, dotenv)
```

### 2. Konfigurasi backend
```bash
cd backend
cp .env.example .env
# lalu isi GEMINI_API_KEY dengan kunci dari https://aistudio.google.com/app/apikey
```

### 3. Jalankan
```bash
# Terminal 1 — backend
npm run dev:backend         # dari root, atau: cd backend && npm run dev

# Terminal 2 — build frontend
npm run build               # menghasilkan index.html dari components/
```

### 4. Arahkan add-in ke backend
Saat develop lokal, set `window.AB_CONFIG.BACKEND_URL` di `index.html` menjadi URL backend Anda (mis. `http://localhost:3000`), atau inject saat build:
```bash
BACKEND_URL=http://localhost:3000 node build/inject-backend.js
```

### 5. Sideload ke Office
Gunakan **Sideloading** office add-in (lihat petunjuk resmi Microsoft) dengan `manifest.xml`, atau buka `index.html` di browser untuk mode debug (chat bekerja, fitur Office terbatas).

---

## 📦 Build & Deploy

### Script yang tersedia (root)
| Script | Fungsi |
|---|---|
| `npm run build` | Gabungkan komponen → `index.html` |
| `npm run build:watch` | Build otomatis saat file berubah |
| `npm run start:backend` | Jalankan backend production |
| `npm run dev:backend` | Jalankan backend (nodemon) |
| `npm test` | Jalankan seluruh tes |
| `npm run lint` | Lint seluruh JS |
| `npm run lint:fix` | Lint + perbaiki otomatis |
| `npm run format` | Format kode (Prettier) |
| `npm run check` | Lint + tes sekaligus |

### Deploy Frontend (Netlify)
File `netlify.toml` sudah disiapkan. Build command: `npm run build && node build/inject-backend.js`.
Set environment variable `BACKEND_URL` di Netlify ke URL backend Anda.

### Deploy Backend (Railway / Render / Netlify Functions)
Deploy folder `backend/` sebagai service Node standalone, lalu set env:
- `GEMINI_API_KEY` (wajib)
- `ALLOWED_ORIGIN` (origin frontend, default Netlify)
- `RATE_LIMIT_MAX` (opsional)

---

## 🧪 Pengujian & CI
- **Unit/integrasi:** `node --test` di `tests/` dan `backend/tests/`
- **CI:** `.github/workflows/ci.yml` menjalankan lint, build, dan test otomatis di setiap push/PR.

```bash
npm run lint   # 0 error, 0 warning
npm test       # 9 tes lulus
```

---

## 📁 Struktur Proyek
```
AutomateBoty/
├── index.html            # Output build (jangan diedit manual)
├── build.js              # Merangkai components/ → index.html
├── build/inject-backend.js  # Menyuntik BACKEND_URL ke index.html
├── components/           # Fragmen HTML (source of truth UI)
│   ├── Header.html / Navbar.html / ChatArea.html
│   └── Panels/           # Academic, Data, Presentation, Extras
├── script.js             # Logika inti: chat, API, insert dokumen
├── features-word.js      # Parafrase, proofreading, bimbingan, outline
├── features-excel.js     # Regresi, statistik, template tabel, validasi
├── features-ppt.js       # Slide, tema, timer
├── ui-extras.js          # Riwayat, preset, rating, onboarding, cache
├── styles.css            # Styling task pane
├── manifest.xml          # Office Add-in manifest
├── backend/              # Express proxy untuk keamanan API key
│   ├── server.js
│   ├── .env.example
│   └── tests/
├── tests/                # Build integrity & konsistensi
├── docs/PRD.md           # Product Requirements Document
├── netlify.toml          # Konfigurasi deploy Netlify
└── .github/workflows/ci.yml
```

---

## 🛡️ Keamanan
- API key Gemini **tidak pernah** di client / localStorage / git.
- Whitelist model, CORS terbatas, rate limiting, validasi payload di backend.
- Semua konten AI di-*sanitize* dengan DOMPurify (kecuali render outline yang kini di-*escape*).

---

## 📈 Roadmap
Lihat [docs/PRD.md](docs/PRD.md) §10 untuk roadmap (TypeScript, PWA/offline, dark mode, auth/billing).

---

## 📄 Lisensi
Proyek **belum berlisensi** (proprietary). Hubungi pemilik untuk penggunaan komersial. Jangan commit API key ke repositori.
```
