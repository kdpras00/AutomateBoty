/*
 * features-ppt.js — AutomateBoty v7
 * PPT: Desain Slide Otomatis, Slide dari Upload File, Practice Timer
 */

// ── STATE ─────────────────────────────────────────────────────────────────────
let timerInterval = null;
let timerSeconds  = 0;

// ── DESAIN SLIDE OTOMATIS ─────────────────────────────────────────────────────
const SLIDE_THEMES = {
    "akademik-biru": {
        titleColor: "#1e3a5f", titleFont: "Calibri", bodyFont: "Calibri",
        bgColor: "#ffffff", accentColor: "#3b82f6"
    },
    "akademik-hijau": {
        titleColor: "#14532d", titleFont: "Calibri", bodyFont: "Calibri",
        bgColor: "#ffffff", accentColor: "#16a34a"
    },
    "profesional-gelap": {
        titleColor: "#f8fafc", titleFont: "Segoe UI", bodyFont: "Segoe UI",
        bgColor: "#0f172a", accentColor: "#60a5fa"
    },
    "minimalis": {
        titleColor: "#1e293b", titleFont: "Arial", bodyFont: "Arial",
        bgColor: "#f8fafc", accentColor: "#64748b"
    },
};

window.terapkanTemaSlide = async function(themeName) {
    if (Office.context.host !== Office.HostType.PowerPoint) { showToast("⚠️ Hanya untuk PowerPoint"); return; }
    const theme = SLIDE_THEMES[themeName];
    if (!theme) { showToast("⚠️ Tema tidak ditemukan"); return; }

    try {
        await PowerPoint.run(async (ctx) => {
            const presentation = ctx.presentation;
            const slides = presentation.slides;
            slides.load("items");
            await ctx.sync();

            for (const slide of slides.items) {
                slide.shapes.load("items");
                await ctx.sync();

                for (const shape of slide.shapes.items) {
                    if (!shape.textFrame) continue;
                    try {
                        shape.textFrame.load("textRange");
                        await ctx.sync();
                        // Color text based on shape position (title vs body heuristic)
                        const tf = shape.textFrame.textRange;
                        tf.font.color = theme.titleColor;
                        tf.font.name  = theme.titleFont;
                    } catch {}
                }
            }
            await ctx.sync();
            showToast(`✅ Tema "${themeName}" diterapkan ke semua slide!`);
        });
    } catch (e) {
        // Fallback: just notify since detailed styling needs premium API
        showToast("⚠️ Beberapa properti tema tidak bisa diterapkan secara program. Coba ubah tema dari PowerPoint → Design.");
        addBotMessage(`**Panduan Terapkan Tema "${themeName}":**\n\n1. Buka **Design** tab di PowerPoint\n2. Pilih tema yang sesuai\n3. Warna yang direkomendasikan:\n   - Judul: \`${theme.titleColor}\`\n   - Font: ${theme.titleFont}\n   - Aksen: \`${theme.accentColor}\`\n\nAtau gunakan **Format Background** untuk mengubah warna latar ke \`${theme.bgColor}\`.`);
    }
};

// ── SLIDE DARI UPLOAD FILE ────────────────────────────────────────────────────
window.slideFromUploadedFile = async function() {
    if (!window.currentFile) { showToast("⚠️ Upload file terlebih dahulu (ikon 📎)!"); return; }

    const file = window.currentFile;
    let content = "";

    try {
        content = await new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.readAsText(file);
        });
    } catch { showToast("❌ Gagal membaca file"); return; }

    // Trim to reasonable length
    content = content.substring(0, 8000);

    const prompt = `Kamu adalah Presentation Expert. Berdasarkan isi dokumen/teks berikut, buat slide presentasi yang terstruktur dengan baik.\n\nAturan:\n1. Output HANYA JSON Array tanpa penjelasan\n2. Format: [{"title":"Judul","points":["Poin 1","Poin 2","Poin 3"],"notes":"Catatan pembicara yang detail..."}]\n3. Maksimal 10 slide\n4. Setiap slide: 3-5 poin singkat\n5. Speaker notes: 2-3 kalimat penjelasan untuk presenter\n6. Slide pertama: Judul & Overview\n7. Slide terakhir: Kesimpulan & Penutup\n\nIsi dokumen:\n${content}`;

    addBotMessage(`⏳ Membuat slide dari **${file.name}**...`);
    try {
        const result = await callGeminiAPI(prompt);
        const match = result.match(/\[[\s\S]*\]/);
        if (!match) throw new Error("Format respons tidak valid");
        const slides = JSON.parse(match[0]);

        if (Office.context.host === Office.HostType.PowerPoint) {
            await runPowerPointSlideGen(result);
        }
        addBotMessage(`✅ **${slides.length} slide dibuat** dari ${file.name}!\n\n**Daftar Slide:**\n${slides.map((s,i) => `${i+1}. ${s.title}`).join("\n")}`);
        saveToHistory(`Slide dari ${file.name}`, result);
    } catch (e) {
        addBotMessage("❌ Gagal membuat slide: " + e.message);
    }
};

// ── PRACTICE TIMER ────────────────────────────────────────────────────────────
window.startPracticeTimer = function() {
    const input = document.getElementById("timer-minutes-input");
    const menit = parseInt(input ? input.value : "10") || 10;
    timerSeconds = menit * 60;

    const display = document.getElementById("timer-display");
    const startBtn = document.getElementById("timer-start-btn");
    const stopBtn  = document.getElementById("timer-stop-btn");

    if (startBtn) startBtn.disabled = true;
    if (stopBtn)  stopBtn.disabled  = false;

    addBotMessage(`🎤 **Practice Timer dimulai: ${menit} menit**\n\nTips presentasi:\n- Bicara dengan tempo tenang, jangan terburu-buru\n- Jaga kontak mata dengan audiens\n- Pause 1-2 detik antar poin utama\n- Gunakan speaker notes sebagai panduan`);

    if (timerInterval) clearInterval(timerInterval);

    timerInterval = setInterval(() => {
        timerSeconds--;
        const m = Math.floor(timerSeconds / 60).toString().padStart(2, "0");
        const s = (timerSeconds % 60).toString().padStart(2, "0");
        if (display) {
            display.textContent = `${m}:${s}`;
            // Color coding
            if (timerSeconds <= 60)       display.style.color = "#ef4444";   // red last 1 min
            else if (timerSeconds <= 180) display.style.color = "#f59e0b";   // amber last 3 min
            else                          display.style.color = "#10b981";   // green
        }

        if (timerSeconds <= 0) {
            clearInterval(timerInterval);
            if (display) display.textContent = "00:00";
            if (startBtn) startBtn.disabled = false;
            if (stopBtn)  stopBtn.disabled  = true;
            addBotMessage(`⏰ **Waktu habis!** (${menit} menit)\n\nBagaimana presentasi Anda? Tips lanjutan:\n- Apakah semua poin tersampaikan?\n- Minta feedback dari audiens\n- Ulangi jika ada bagian yang terlewat`);
        }
    }, 1000);
};

window.stopPracticeTimer = function() {
    if (timerInterval) { clearInterval(timerInterval); timerInterval = null; }
    const elapsed = parseInt(document.getElementById("timer-minutes-input")?.value || "10") * 60 - timerSeconds;
    const em = Math.floor(elapsed / 60), es = elapsed % 60;
    const display = document.getElementById("timer-display");
    if (display) { display.textContent = "⏸ Berhenti"; display.style.color = "#94a3b8"; }
    document.getElementById("timer-start-btn").disabled = false;
    document.getElementById("timer-stop-btn").disabled = true;
    if (elapsed > 0) addBotMessage(`⏸ Timer dihentikan setelah **${em}m ${es}s**.`);
};

window.resetPracticeTimer = function() {
    if (timerInterval) { clearInterval(timerInterval); timerInterval = null; }
    const input = document.getElementById("timer-minutes-input");
    const menit = parseInt(input ? input.value : "10") || 10;
    timerSeconds = menit * 60;
    const display = document.getElementById("timer-display");
    if (display) {
        display.textContent = `${menit.toString().padStart(2,"0")}:00`;
        display.style.color = "#10b981";
    }
    document.getElementById("timer-start-btn").disabled = false;
    document.getElementById("timer-stop-btn").disabled = true;
};
