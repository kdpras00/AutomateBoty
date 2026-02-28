/*
 * features-word.js — AutomateBoty v7
 * Word-specific: Parafrase, Proofreading Mendalam, Bimbingan Skripsi, Outline Builder
 */

// ── STATE ─────────────────────────────────────────────────────────────────────
let bimbinganActive = false;
let outlineData = null;

// ── PARAFRASE ANTI-PLAGIARISME ─────────────────────────────────────────────────
window.parafraseTeks = async function(level = "sedang") {
    if (Office.context.host !== Office.HostType.Word) { showToast("⚠️ Hanya untuk Microsoft Word"); return; }
    showToast("⏳ Membaca teks terpilih...");
    let teks = "";
    try {
        teks = await new Promise((resolve) => {
            Word.run(async (ctx) => {
                const sel = ctx.document.getSelection();
                sel.load("text");
                await ctx.sync();
                resolve(sel.text);
            }).catch(() => resolve(""));
        });
    } catch {}

    if (!teks.trim()) { showToast("⚠️ Pilih teks terlebih dahulu!"); return; }

    const levelDesc = {
        "ringan": "Parafrase ringan: ubah sekitar 30% kata dengan sinonim, pertahankan struktur kalimat.",
        "sedang": "Parafrase sedang: ubah 50-60% kata, variasikan struktur kalimat, pertahankan makna.",
        "ketat": "Parafrase ketat: tulis ulang hampir seluruh kalimat dengan struktur berbeda, gunakan sudut pandang lain, pastikan 80%+ berbeda dari aslinya.",
    };

    const prompt = `${levelDesc[level] || levelDesc["sedang"]}\n\nTeks asli:\n"${teks}"\n\nOutput HANYA teks hasil parafrase tanpa penjelasan. Pertahankan bahasa (Indonesia/Inggris) yang sama. Jika ada teks Inggris, tetap tulis italic.`;

    addBotMessage(`⏳ Memparafrase teks (level: **${level}**)...`);
    try {
        const result = await callGeminiAPI(prompt);
        // Replace selected text
        await Word.run(async (ctx) => {
            const sel = ctx.document.getSelection();
            sel.insertText(result.trim(), Word.InsertLocation.replace);
            await ctx.sync();
        });
        addBotMessage(`✅ **Parafrase selesai!** Teks sudah diganti.\n\n**Hasil:**\n${result}`);
        saveToHistory(`Parafrase (${level}): ${teks.substring(0,80)}…`, result);
    } catch (e) {
        addBotMessage(`❌ Gagal parafrase: ${e.message}`);
    }
};

// ── PROOFREADING MENDALAM ─────────────────────────────────────────────────────
window.proofreadingMendalam = async function() {
    if (Office.context.host !== Office.HostType.Word) { showToast("⚠️ Hanya untuk Microsoft Word"); return; }
    showToast("⏳ Membaca dokumen...");
    let teks = "";
    try {
        teks = await new Promise((resolve) => {
            Word.run(async (ctx) => {
                const sel = ctx.document.getSelection();
                sel.load("text");
                await ctx.sync();
                if (sel.text?.trim()) return resolve(sel.text);
                const body = ctx.document.body;
                body.load("text");
                await ctx.sync();
                resolve(body.text.substring(0, 4000));
            }).catch(() => resolve(""));
        });
    } catch {}

    if (!teks.trim()) { showToast("⚠️ Tidak ada teks untuk diperiksa!"); return; }

    const prompt = `Lakukan proofreading mendalam terhadap teks akademik berikut. Format laporan Markdown:\n
## Laporan Proofreading Mendalam

### 1. Kesalahan Ejaan & Tata Bahasa
Daftarkan kata/kalimat yang salah dengan koreksinya.

### 2. Kalimat Terlalu Panjang
Identifikasi kalimat >30 kata. Sarankan cara memecahnya.

### 3. Kata Tidak Baku / Tidak Formal
Temukan kata gaul/informal/tidak baku. Berikan padanan bakunya.

### 4. Ketidakkonsistenan Penulisan
(Ejaan tidak konsisten, penggunaan istilah berbeda untuk hal sama, dll.)

### 5. Saran Perbaikan Keseluruhan
Nilai secara umum (1-10) dan sarankan perbaikan utama.

Teks:\n"${teks}"`;

    addBotMessage("⏳ Menganalisis tulisan...");
    try {
        const result = await callGeminiAPI(prompt);
        addBotMessage(result);
        saveToHistory("Proofreading mendalam", result);
    } catch (e) {
        addBotMessage(`❌ Gagal: ${e.message}`);
    }
};

// ── MODE BIMBINGAN SKRIPSI ────────────────────────────────────────────────────
window.toggleBimbinganSkripsi = function() {
    bimbinganActive = !bimbinganActive;
    const btn = document.getElementById("bimbingan-btn");
    if (btn) {
        btn.classList.toggle("active-mode", bimbinganActive);
        btn.textContent = bimbinganActive ? "🎓 Mode Bimbingan: ON" : "🎓 Mode Bimbingan";
    }
    if (bimbinganActive) {
        addBotMessage(`🎓 **Mode Bimbingan Skripsi Aktif!**\n\nSaya sekarang berperan sebagai **Pembimbing Akademik** Anda.\n\nCara menggunakan:\n- Paste bagian skripsi yang ingin di-review\n- Tanya: *"Review Bab I saya"* atau *"Apa kelemahan metodologi ini?"*\n- Minta saran: *"Bagaimana cara memperkuat argumen di bagian ini?"*\n\n✍️ Silakan mulai!`);
    } else {
        addBotMessage("Mode Bimbingan Skripsi dinonaktifkan.");
    }
};

// Intercept prompt jika bimbingan aktif
function getBimbinganPrefix() {
    return bimbinganActive
        ? `Kamu adalah Dosen Pembimbing Skripsi yang berpengalaman dan kritis. Berikan feedback mendalam, akurat, dan konstruktif seperti seorang pembimbing senior. Identifikasi kelemahan argumen, logika, metodologi, dan struktur. Sarankan perbaikan konkret. Gunakan bahasa formal akademik Indonesia.\n\n`
        : "";
}
window.getBimbinganPrefix = getBimbinganPrefix;

// ── OUTLINE BUILDER INTERAKTIF ────────────────────────────────────────────────
window.openOutlineBuilder = function() {
    const panel = document.getElementById("outline-builder-panel");
    if (panel) {
        panel.classList.toggle("hidden");
        document.getElementById("outline-btn").classList.toggle("active", !panel.classList.contains("hidden"));
    }
};

window.generateOutline = async function() {
    const topicInput = document.getElementById("outline-topic-input");
    const topic = topicInput ? topicInput.value.trim() : "";
    if (!topic) { showToast("⚠️ Masukkan topik terlebih dahulu!"); return; }

    const jenis = document.getElementById("outline-type-select").value;
    const prompt = `Buatkan outline/kerangka ${jenis} tentang: "${topic}"\n\nFormat JSON:\n{"title":"Judul ${jenis}","chapters":[{"num":"BAB I","title":"JUDUL BAB","sections":["1.1 Sub-bab","1.2 Sub-bab"]},...]}\n\nOutput HANYA JSON, tanpa penjelasan.`;

    const resultDiv = document.getElementById("outline-result");
    resultDiv.innerHTML = '<p class="empty-msg">⏳ Membuat outline...</p>';

    try {
        const raw = await callGeminiAPI(prompt);
        const match = raw.match(/\{[\s\S]*\}/);
        if (!match) throw new Error("Format tidak valid");
        outlineData = JSON.parse(match[0]);
        renderOutline(outlineData);
    } catch (e) {
        resultDiv.innerHTML = `<p class="empty-msg">❌ Gagal: ${e.message}</p>`;
    }
};

function renderOutline(data) {
    const resultDiv = document.getElementById("outline-result");
    if (!data) return;
    const html = data.chapters.map((ch, i) => `
        <div class="outline-chapter">
            <div class="outline-ch-header" onclick="toggleOutlineChapter(${i})">
                <span>${ch.num}: ${ch.title}</span>
                <span class="outline-expand-btn">▼</span>
            </div>
            <div class="outline-sections" id="outline-ch-${i}">
                ${ch.sections.map(s => `<div class="outline-section">• ${s}</div>`).join("")}
                <button class="btn-outline" style="margin-top:6px;font-size:10px;" onclick="expandChapter(${i}, '${ch.title.replace(/'/g,"\\'")}')">✍️ Tulis isi bab ini</button>
            </div>
        </div>`).join("");
    resultDiv.innerHTML = `<div class="outline-title">${data.title}</div>${html}
        <div style="margin-top:8px;display:flex;gap:6px;">
            <button class="btn-primary" style="font-size:11px;" onclick="insertOutlineToDoc()">📄 Masukkan ke Dokumen</button>
        </div>`;
}

window.toggleOutlineChapter = function(i) {
    const el = document.getElementById(`outline-ch-${i}`);
    if (el) el.style.display = el.style.display === "none" ? "block" : "none";
};

window.expandChapter = async function(i, title) {
    if (!outlineData) return;
    const ch = outlineData.chapters[i];
    const prompt = `Tulis isi lengkap ${ch.num}: "${title}" dengan sub-bab ${ch.sections.join(", ")} untuk ${outlineData.title}. Format Markdown. Bahasa Indonesia akademik. Teks bahasa Inggris ditulis italic.`;
    addBotMessage(`⏳ Menulis **${ch.num}: ${title}**...`);
    try {
        const result = await callGeminiAPI(prompt);
        addBotMessage(result);
        if (Office.context.host === Office.HostType.Word) {
            const html = marked.parse(result);
            Office.context.document.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html });
        }
    } catch (e) { addBotMessage(`❌ Gagal: ${e.message}`); }
};

window.insertOutlineToDoc = function() {
    if (!outlineData || Office.context.host !== Office.HostType.Word) { showToast("⚠️ Perlu Microsoft Word"); return; }
    const text = outlineData.chapters.map(ch => `${ch.num}: ${ch.title}\n${ch.sections.map(s => `    ${s}`).join("\n")}`).join("\n\n");
    const html = marked.parse(`# ${outlineData.title}\n\n${outlineData.chapters.map(ch => `## ${ch.num}: ${ch.title}\n${ch.sections.map(s => `- ${s}`).join("\n")}`).join("\n\n")}`);
    Office.context.document.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, () => showToast("✅ Outline dimasukkan ke dokumen!"));
};
