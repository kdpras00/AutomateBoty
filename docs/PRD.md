# AutomateBoty — Product Requirements Document (PRD)

> **Versi Dokumen:** 1.0.0
> **Tanggal:** Agustus 2026
> **Status:** Final (hasil audit & perbaikan menyeluruh)
> **Iterasi Produk:** v7.0

---

## 1. Ringkasan Eksekutif

**AutomateBoty** adalah *Microsoft Office Add-in* (task pane) yang menghadirkan asisten AI **Google Gemini** langsung di dalam **Word**, **Excel**, dan **PowerPoint**. Produk menargetkan mahasiswa Indonesia yang menulis karya akademik (skripsi, jurnal, makalah, prosiding) dan pengguna kantoran yang ingin mengotomatisasi pembuatan & pengolahan dokumen.

Dokumen ini adalah PRD resmi yang mendefinisikan visi produk, ruang lingkup fungsional, arsitektur teknis, persyaratan keamanan, serta persyaratan non-fungsional. Dokumen ini menjadi acuan tunggal untuk pengembangan, pengujian, dan evaluasi ("definition of done").

---

## 2. Visi & Tujuan Produk

### 2.1 Visi
Menjadi asisten AI paling andal bagi mahasiswa dan pekerja kantoran untuk menghasilkan **dokumen akademik dan bahan presentasi berkualitas tinggi** langsung dari aplikasi Office yang sudah mereka gunakan — tanpa perlu berpindah aplikasi.

### 2.2 Tujuan (Objectives)
1. **Meningkatkan produktivitas** penulisan akademik (skripsi, jurnal, makalah) hingga 60% lebih cepat.
2. **Menjaga kualitas & konsistensi** format dokumen sesuai standar akademik Indonesia (APA, IEEE, IMRaD, struktur 5 bab).
3. **Melindungi kredensial AI** pengguna melalui arsitektur *backend proxy* yang aman.
4. **Memberikan pengalaman tanpa gesekan** (seamless) di ketiga host Office dengan UI konsisten.

---

## 3. Persona & Konteks Penggunaan

### 3.1 Persona Utama
| Persona | Deskripsi | Kebutuhan utama |
|---|---|---|
| **Mahasiswa S1/S2** | Menulis skripsi, tesis, jurnal | Struktur 5 bab, sitasi, parafrase, bimbingan AI, interpretasi statistik Bab IV |
| **Peneliti / Akademisi** | Menulis artikel & prosiding | Template jurnal, abstrak 2 bahasa, sitasi APA/IEEE |
| **Pekerja kantoran** | Membuat laporan & presentasi | Template dokumen, analisis data, slide otomatis, timer presentasi |

### 3.2 Konteks Penggunaan
- **Word:** menulis, parafrase, proofreading, outline, template, sitasi, upload file/folder bab skripsi.
- **Excel:** analisis statistik, regresi & korelasi, template tabel, validasi data.
- **PowerPoint:** pembuatan slide otomatis, tema, slide dari file, timer latihan.

---

## 4. Fitur & Persyaratan Fungsional

### 4.1 Chat AI Terpadu (Semua Host)
| ID | Fitur | Deskripsi |
|---|---|---|
| F-01 | Chat multi-host | Konsultasi AI di Word/Excel/PPT dengan konteks dokumen aktif secara otomatis |
| F-02 | Mode bahasa ID/EN | Beralih antara prompt & respons Bahasa Indonesia / English |
| F-03 | Upload file | Mendukung gambar, .docx (mammoth), .pdf (pdf.js), .txt/.csv/.md/.json/.js/.py |
| F-04 | Upload folder | Membaca banyak file bab skripsi, deteksi bab, saran lanjut BAB otomatis |
| F-05 | Konteks dokumen | Membaca seleksi/range aktif lalu meneruskan ke AI sebagai konteks |
| F-06 | Retry & backoff | Penanganan error 429/503 dengan retry otomatis dan pesan ramah |
| F-07 | Mode offline | Deteksi tanpa internet, fallback ke cache respons terakhir |
| F-08 | Riwayat percakapan | Simpan maks 10 sesi terakhir di localStorage |
| F-09 | Rating & salin | Umpan balik 👍/👎 dan tombol salin per pesan |

### 4.2 Word — Academic Writing
| ID | Fitur | Deskripsi |
|---|---|---|
| F-W1 | Jurnal (ID & EN) | Aksi cepat membuat jurnal lengkap (IMRaD, APA/IEEE, abstrak 2 bahasa) |
| F-W2 | Parafrase anti-plagiarisme | Level ringan/sedang/ketat, menulis ulang teks terpilih |
| F-W3 | Proofreading mendalam | Laporan: ejaan, kalimat panjang, kata tidak baku, inkonsistensi, saran |
| F-W4 | Mode Bimbingan Skripsi | AI berperan sebagai dosen pembimbing kritis |
| F-W5 | Outline Builder | Buat kerangka, ekspansi tiap bab, masukkan ke dokumen |
| F-W6 | Template dokumen | Jurnal ID/EN, artikel, skripsi 5 bab, prosiding, makalah |
| F-W7 | Sitasi otomatis | Format APA / IEEE / Chicago / Vancouver |
| F-W8 | Formatting layout | Terapkan A4, margin, kolom, font, alignment via JSON blok |
| F-W9 | Lanjutkan BAB | Deteksi bab yang ada di folder → lanjutkan bab berikutnya secara konsisten |

### 4.3 Excel — Data Analysis
| ID | Fitur | Deskripsi |
|---|---|---|
| F-E1 | Perhitungan statistik | N, Mean, Median, StdDev, Min, Max, Range |
| F-E2 | Regresi & korelasi | Hitung b₀, b₁, r, R² + sisipkan tabel hasil + narasi interpretasi |
| F-E3 | Interpretasi Bab IV | Narasi akademik 300–500 kata dari data statistik terpilih |
| F-E4 | Template tabel | Frekuensi, distribusi, crosstab, kuesioner, rangkuman (dengan rumus) |
| F-E5 | Validasi data | Saran + dropdown programmatis untuk daftar nilai |
| F-E6 | Sisip data/rumus/grafik | Deteksi otomatis dan sisipkan ke sheet aktif |

### 4.4 PowerPoint — Presentation
| ID | Fitur | Deskripsi |
|---|---|---|
| F-P1 | Slide dari file | Upload dokumen → JSON array slide dengan speaker notes |
| F-P2 | Tema slide | Akademik Biru/Hijau, Dark Pro, Minimalis |
| F-P3 | Outline 10 slide | Prompt cepat outline dengan notes |
| F-P4 | Timer latihan | Countdown dengan kode warna, jeda/reset, tips |

### 4.5 Ekstra (Semua Host)
| ID | Fitur | Deskripsi |
|---|---|---|
| F-X1 | Prompt presets | Simpan & pakai prompt favorit |
| F-X2 | Onboarding tour | Tur 5 langkah saat pertama kali dibuka |
| F-X3 | Offline cache | Cache respons (kadaluarsa 7 hari) |

---

## 5. Arsitektur Sistem

### 5.1 Komponen
```
┌───────────────────────┐        ┌──────────────────────┐        ┌─────────────────┐
│  Office Add-in (Task  │  HTTPS │  Backend Proxy       │  HTTPS │  Google Gemini  │
│  Pane) — index.html   │───────▶│  /api/gemini/chat    │───────▶│  v1beta API      │
│  Word/Excel/PowerPoint │        │  (Node/Express)      │        │  generateContent │
└───────────────────────┘        └──────────────────────┘        └─────────────────┘
        ▲                                 │
        │ static hosting                  │ API key HANYA ada di sini
        │ (Netlify)                       └ (server-side .env)
```

### 5.2 Arsitektur Keamanan API Key (Keputusan Kunci)
- **Sebelum (bug kritis):** `script.js` memanggil Google langsung dengan `x-goog-api-key` dari `localStorage` → kunci terekspos di browser/DevTools.
- **Sesudah (perbaikan):** Seluruh panggilan AI diarahkan ke **backend proxy** `POST /api/gemini/chat`. API key disimpan **hanya di server** (`GEMINI_API_KEY` pada `.env`). Browser tidak pernah melihat kuncinya.
- `settings-panel` menghentikan penyimpanan API key di localStorage (kunci dikelola server/admin).

### 5.3 Stack Teknologi
| Lapisan | Teknologi |
|---|---|
| Frontend | Vanilla HTML/CSS/JS + Office.js |
| Backend | Node.js + Express, `node-fetch`, `dotenv`, `cors` |
| AI | Google Gemini `gemini-2.5-flash` (v1beta generateContent) |
| CDN | marked, highlight.js, DOMPurify, pdf.js, mammoth |
| Hosting | Static: Netlify; Backend: Railway/Render/Netlify Functions |
| Build | `build.js` (concat komponen HTML → index.html) |

### 5.4 Alur Data Chat
1. User mengetik di task pane → handler membaca konteks seleksi dokumen.
2. Payload dibentuk: `system_instruction` + `contents` (konteks dokumen + file/folder + prompt).
3. `script.js` mengirim ke `POST {BACKEND}/api/gemini/chat`.
4. Backend memvalidasi, menambahkan API key, meneruskan ke Google.
5. Respons dikembalikan, dirender (Markdown + sanitasi DOMPurify), lalu disisipkan bila relevan.

---

## 6. Persyaratan Keamanan
| ID | Persyaratan |
|---|---|
| SEC-01 | API key Gemini **tidak boleh** berada di kode client / localStorage / di-commit |
| SEC-02 | Backend **whitelist** parameter `model` (hanya `gemini-2.5-flash`, dll.) |
| SEC-03 | CORS backend dibatasi ke origin add-in yang sah |
| SEC-04 | **Rate limiting** backend untuk mencegah penyalahgunaan gateway |
| SEC-05 | Semua konten dari AI di-*sanitize* dengan DOMPurify sebelum dirender/insert |
| SEC-06 | Hindari inline `onclick` dengan string interpolasi dari data AI (pakai `addEventListener`) |
| SEC-07 | Tidak ada secret/placeholder yang bocor di kode atau dokumentasi |

---

## 7. Persyaratan Non-Fungsional
| Kategori | Persyaratan |
|---|---|
| **Performa** | Respons pertama < 5 detik; konteks dokumen dibatasi (~3000 char) |
| **Kompatibilitas** | Berjalan di Word/Excel/PPT desktop (Office.js) + mode browser debug |
| **Reliabilitas** | Retry backoff (2s/4s/8s) untuk 429/503; pesan error ramah |
| **Keamanan** | Lihat §6 |
| **Kualitas Kode** | Resource bersih (`node_modules` & `.DS_Store` di-ignore); build deterministik; versi terkunci |
| **Pengujian** | Tes unit untuk backend & utilitas kritis; CI otomatis |
| **Aksesibilitas** | Kontras warna memadai, label/aria pada kontrol utama |
| **I18n** | Bahasa Indonesia (utama) + Inggris (toggle) |

---

## 8. Kriteria Penerimaan (Definition of Done)
Produk dianggap **sempurna/siap rilis** bila:
1. API key **tidak pernah** terekspos di client — semua lalu lintas AI melewati proxy. ✔
2. Build reproduktif: `npm run build` menghasilkan `index.html` terminal. ✔
3. Tidak ada referensi elemen DOM yang hilang (word/ppt-tools-panel, host-tools-btn, history-btn) → tidak ada error konsol. ✔
4. `manifest.xml` valid tanpa merujuk file yang tidak ada (`commands.html` dihapus). ✔
5. Semua blok `catch {}` kosong digantikan logging + umpan balik user yang layak. ✔
6. Seluruh konten AI di-sanitize melalui DOMPurify. ✔
7. Backend memiliki whitelist model, CORS terbatas, rate limit, dan health check. ✔
8. Dokumentasi lengkap: README, PRD, .env.example, netlify.toml, CI. ✔

---

## 9. Ruang Lingkup Non-Goal (Sengaja Tidak Termasuk)
- Aplikasi mandiri (di luar ekosistem Office) — versi saat ini hanya task pane add-in.
- Fitur berbayar/subscription & akun pengguna.
- Offline penuh (hanya cache respons terakhir).
- Migrasi ke TypeScript / framework (React) — di luar iterasi ini, dicatat di roadmap.

---

## 10. Roadmap Selanjutnya (Out of Scope Iterasi Ini)
- Migrasi ke TypeScript & modularisasi `script.js` (split api/word/excel/ppt/ui).
- Service worker untuk PWA/offline penuh.
- Dukungan `PWA`/AppDomain tambahan & dark mode.
- Auth pengguna + billing (langganan).
- E2E testing dengan Office Fiddler/Playwright.

---

## 11. Metrik Keberhasilan (KPI)
- **Aktivasi:** % pengguna menyelesaikan onboarding tour.
- **Penggunaan:** jumlah aksi cepat (jurnal, parafrase, regresi, slide) per sesi.
- **Kepuasan:** rasio ⬆/⬇ rating.
- **Stabilitas:** nol error JS pada konsol saat alur utama.
- **Keamanan:** tidak ada API key terekspos dalam audit berulang.
```
