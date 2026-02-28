/*
 * features-ppt.js — AutomateBoty v8
 * PPT: Desain Slide Otomatis, Slide dari Upload File, Practice Timer
 * + executePPTAction() for AI-commanded slide creation via Action JSON
 */

// ── STATE ─────────────────────────────────────────────────────────────────────
let timerInterval = null;
let timerSeconds  = 0;

// ── PPT ACTION EXECUTOR ────────────────────────────────────────────────────────
window.executePPTAction = async function(action, params) {
    switch (action) {
        case "PPT_DYNAMIC_FORMAT":return await pptDynamicFormat(params);
        case "PPT_CREATE_SLIDES": return await pptCreateSlides(params);
        case "PPT_ADD_SLIDE":     return await pptAddSlide(params);
        case "PPT_APPLY_THEME":   return await pptApplyTheme(params);
        case "PPT_FORMAT_TEXT":   return await pptFormatText(params);
        default:
            console.warn(`⚠️ PPT action tidak dikenal: ${action}`);
            throw new Error("UNSUPPORTED_ACTION");
    }
};

// ── DYNAMIC OMNI-FORMATTER ────────────────────────────────────────────────────
async function pptDynamicFormat(p) {
    // Expected p: { tasks: [ { target: "font"|"background"|"slide", props: { color: "#f00", name: "Arial"... } } ] }
    if (!p || !p.tasks || !Array.isArray(p.tasks)) return;

    await PowerPoint.run(async (ctx) => {
        const slide = ctx.presentation.slides.getItemAt(0);
        slide.shapes.load("items");
        await ctx.sync();

        for (const task of p.tasks) {
            const tgtName = task.target || "font";
            const props   = task.props || {};
            
            if (tgtName === "font" || tgtName === "text") {
                for (const shape of slide.shapes.items) {
                    try {
                        const tf = shape.textFrame.textRange;
                        if (props.color) tf.font.color = props.color;
                        if (props.name)  tf.font.name  = props.name;
                        if (props.size)  tf.font.size  = parseFloat(props.size);
                        if (props.bold !== undefined) tf.font.bold = props.bold;
                        if (props.italic !== undefined) tf.font.italic = props.italic;
                    } catch {}
                }
            } 
        }
        await ctx.sync();
    });
    showToast("✅ Format presentasi dinamis diterapkan!");
}

// ── Create Multiple Slides ────────────────────────────────────────────────────
async function pptCreateSlides(p) {
    // p: { slides: [{ title, points: [], notes, layout? }] }
    const slides = p.slides || [];
    if (!slides.length) { showToast("⚠️ Tidak ada data slide"); return; }

    await PowerPoint.run(async (ctx) => {
        for (const data of slides) {
            ctx.presentation.slides.add();
        }
        await ctx.sync();

        // Load slides and fill content
        const pptSlides = ctx.presentation.slides;
        pptSlides.load("items");
        await ctx.sync();

        const existing = pptSlides.items;
        const offset = existing.length - slides.length;

        for (let i = 0; i < slides.length; i++) {
            const slide = existing[offset + i];
            const data = slides[i];
            slide.shapes.load("items");
            await ctx.sync();

            if (slide.shapes.items.length > 0) {
                try {
                    slide.shapes.items[0].textFrame.textRange.text = data.title || "";
                } catch {}
            }
            if (slide.shapes.items.length > 1) {
                const points = Array.isArray(data.points) ? data.points.join("\n") : (data.points || "");
                try {
                    slide.shapes.items[1].textFrame.textRange.text = points;
                } catch {}
            }
        }
        await ctx.sync();
        showToast(`✅ ${slides.length} slide dibuat!`);
    }).catch(e => {
        showToast("⚠️ Sebagian slide mungkin tidak terformat. Cek hasilnya di PPT.");
        console.error("PPT error:", e);
    });
}

// ── Add Single Slide ──────────────────────────────────────────────────────────
async function pptAddSlide(p) {
    // p: { title, points: [], notes }
    await pptCreateSlides({ slides: [p] });
}

// ── Apply Theme ───────────────────────────────────────────────────────────────
async function pptApplyTheme(p) {
    // p: { titleColor, bodyColor, fontName, bgColor }
    const theme = SLIDE_THEMES[p.themeName] || p;

    try {
        await PowerPoint.run(async (ctx) => {
            const slides = ctx.presentation.slides;
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
                        shape.textFrame.textRange.font.color = theme.titleColor || p.titleColor || "#1e293b";
                        if (theme.titleFont || p.fontName) {
                            shape.textFrame.textRange.font.name = theme.titleFont || p.fontName;
                        }
                    } catch {}
                }
            }
            await ctx.sync();
            showToast(`✅ Tema diterapkan ke semua slide!`);
        });
    } catch (e) {
        showToast("⚠️ Tema sebagian diterapkan. Gunakan Design tab untuk tema lanjutan.");
    }
}

// ── Format Text on Active Slide ───────────────────────────────────────────────
async function pptFormatText(p) {
    // p: { color, fontName, fontSize, bold }
    await PowerPoint.run(async (ctx) => {
        const slide = ctx.presentation.slides.getItemAt(0);
        slide.shapes.load("items");
        await ctx.sync();
        for (const shape of slide.shapes.items) {
            try {
                const tf = shape.textFrame.textRange;
                if (p.color)    tf.font.color = p.color;
                if (p.fontName) tf.font.name  = p.fontName;
                if (p.fontSize) tf.font.size  = parseFloat(p.fontSize);
                if (p.bold !== undefined) tf.font.bold = p.bold;
            } catch {}
        }
        await ctx.sync();
        showToast("✅ Format teks slide diterapkan!");
    }).catch(e => showToast("❌ Error: " + e.message));
}

// ── DESAIN SLIDE OTOMATIS ─────────────────────────────────────────────────────
const SLIDE_THEMES = {
    "akademik-biru":    { titleColor: "#1e3a5f", titleFont: "Calibri",  bodyFont: "Calibri",   bgColor: "#ffffff", accentColor: "#3b82f6" },
    "akademik-hijau":   { titleColor: "#14532d", titleFont: "Calibri",  bodyFont: "Calibri",   bgColor: "#ffffff", accentColor: "#16a34a" },
    "profesional-gelap":{ titleColor: "#f8fafc", titleFont: "Segoe UI", bodyFont: "Segoe UI",  bgColor: "#0f172a", accentColor: "#60a5fa" },
    "minimalis":        { titleColor: "#1e293b", titleFont: "Arial",    bodyFont: "Arial",     bgColor: "#f8fafc", accentColor: "#64748b" },
};

window.terapkanTemaSlide = async function(themeName) {
    if (Office.context.host !== Office.HostType.PowerPoint) { showToast("⚠️ Hanya untuk PowerPoint"); return; }
    const theme = SLIDE_THEMES[themeName];
    if (!theme) { showToast("⚠️ Tema tidak ditemukan"); return; }
    await pptApplyTheme({ ...theme, themeName });
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

    content = content.substring(0, 8000);
    const prompt = `Kamu adalah Presentation Expert. Berdasarkan dokumen berikut, buat slide presentasi terstruktur.\n\nAturan:\n1. Output HANYA JSON Array tanpa penjelasan\n2. Format: [{"title":"Judul","points":["Poin 1","Poin 2","Poin 3"],"notes":"Catatan..."}]\n3. Maksimal 10 slide\n4. Setiap slide: 3-5 poin singkat\n5. Slide pertama: Judul & Overview, terakhir: Kesimpulan\n\nIsi dokumen:\n${content}`;

    addBotMessage(`⏳ Membuat slide dari **${file.name}**...`);
    try {
        const result = await callGeminiAPI(prompt);
        const match = result.match(/\[[\s\S]*\]/);
        if (!match) throw new Error("Format respons tidak valid");
        const slides = JSON.parse(match[0]);
        if (Office.context.host === Office.HostType.PowerPoint) {
            await pptCreateSlides({ slides });
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
    addBotMessage(`🎤 **Practice Timer dimulai: ${menit} menit**\n\nTips: Bicara tenang, jaga kontak mata, pause antar poin utama.`);
    if (timerInterval) clearInterval(timerInterval);
    timerInterval = setInterval(() => {
        timerSeconds--;
        const m = Math.floor(timerSeconds / 60).toString().padStart(2, "0");
        const s = (timerSeconds % 60).toString().padStart(2, "0");
        if (display) {
            display.textContent = `${m}:${s}`;
            if (timerSeconds <= 60)       display.style.color = "#ef4444";
            else if (timerSeconds <= 180) display.style.color = "#f59e0b";
            else                          display.style.color = "#10b981";
        }
        if (timerSeconds <= 0) {
            clearInterval(timerInterval);
            if (display)  display.textContent = "00:00";
            if (startBtn) startBtn.disabled = false;
            if (stopBtn)  stopBtn.disabled  = true;
            addBotMessage(`⏰ **Waktu habis!** (${menit} menit)\n\nApakah semua poin tersampaikan? Minta feedback dari audiens!`);
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
    if (display) { display.textContent = `${menit.toString().padStart(2,"0")}:00`; display.style.color = "#10b981"; }
    document.getElementById("timer-start-btn").disabled = false;
    document.getElementById("timer-stop-btn").disabled = true;
};
