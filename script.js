/*
 * AutomateBoty v7 — Core script.js
 * Integrates: History, Presets, Rating, Bimbingan, Word/Excel/PPT tools
 */

// ── CONFIG ────────────────────────────────────────────────────────────────────
const DEFAULT_API_KEY = "AIzaSyDM4fjZTT1F6Ux2D8anrQMe7SSShKVUgrQ";
const INVALID_KEYS = ["AIzaSyCmSlRCCPgC1ph4vuco9hwLsTaDtnBPcSA","AIzaSyAmsulrYYqrxuWnlqwrn1UzHsPdTSedyR0"];
let apiKey = localStorage.getItem("gemini_api_key") || DEFAULT_API_KEY;
let currentLang = localStorage.getItem("ab_lang") || "ID";

// ── DOM ───────────────────────────────────────────────────────────────────────
const chatContainer = document.getElementById("chat-container");
const userInput     = document.getElementById("user-input");
const sendBtn       = document.getElementById("send-btn");

// ── OFFICE INIT ───────────────────────────────────────────────────────────────
Office.onReady((info) => {
    const host = info.host;
    updateHostBadge(host);
    setupQuickActions(host);
    setupEventListeners();
    showHostTools(host);

    // API Key
    const saved = localStorage.getItem("gemini_api_key");
    const apiInput = document.getElementById("api-key-input");
    if (saved && !INVALID_KEYS.includes(saved)) { apiInput.value = saved; apiKey = saved; }
    else { if (INVALID_KEYS.includes(saved)) localStorage.removeItem("gemini_api_key"); apiInput.value = DEFAULT_API_KEY; apiKey = DEFAULT_API_KEY; }

    // Language
    document.getElementById("lang-toggle-btn").textContent = currentLang === "ID" ? "🇮🇩" : "🇬🇧";

    // Network
    updateNetworkStatus();
    window.addEventListener("online",  updateNetworkStatus);
    window.addEventListener("offline", updateNetworkStatus);

    // Init ui-extras
    if (typeof setupExtraPanels === "function") setupExtraPanels();

    // Version
    const v = document.createElement("div");
    v.className = "version-info";
    v.textContent = "v7.0 · AutomateBoty · 17 Fitur Baru";
    document.querySelector(".app-container").appendChild(v);
});

function updateHostBadge(host) {
    const badge = document.getElementById("host-badge");
    if (!badge) return;
    if (!host)                                  { badge.textContent = "Browser"; }
    else if (host === Office.HostType.Word)     { badge.textContent = "Word";  badge.className = "host-badge word"; }
    else if (host === Office.HostType.Excel)    { badge.textContent = "Excel"; badge.className = "host-badge excel"; }
    else if (host === Office.HostType.PowerPoint){ badge.textContent = "PPT";  badge.className = "host-badge ppt"; }
}

function showHostTools(host) {
    const navAcademic = document.getElementById("nav-academic");
    const navData = document.getElementById("nav-data");
    const navPresentation = document.getElementById("nav-presentation");
    const navTemplates = document.getElementById("nav-templates");

    if (host === Office.HostType.Word) {
        if (navData) navData.style.display = "none";
        if (navPresentation) navPresentation.style.display = "none";
    } else if (host === Office.HostType.Excel) {
        if (navAcademic) navAcademic.style.display = "none";
        if (navPresentation) navPresentation.style.display = "none";
        if (navTemplates) navTemplates.style.display = "none";
    } else if (host === Office.HostType.PowerPoint) {
        if (navAcademic) navAcademic.style.display = "none";
        if (navData) navData.style.display = "none";
        if (navTemplates) navTemplates.style.display = "none";
    }

    let initialMode = "chat";
    if (host === Office.HostType.Word) initialMode = "academic";
    else if (host === Office.HostType.Excel) initialMode = "data";
    else if (host === Office.HostType.PowerPoint) initialMode = "presentation";
    
    if (typeof switchMode === "function") switchMode(initialMode);
}

window.switchMode = function(mode) {
    // Hide all mode panels
    document.querySelectorAll(".mode-panel").forEach(p => p.classList.add("hidden"));
    // Disable active state
    document.querySelectorAll(".nav-btn").forEach(b => b.classList.remove("active"));
    
    // Activate target
    const panel = document.getElementById(mode + "-panel");
    const btn = document.getElementById("nav-" + mode);
    
    if (panel) panel.classList.remove("hidden");
    if (btn) btn.classList.add("active");
    
    // Triggers
    if (mode === "history" && typeof renderHistory === "function") renderHistory();
    if (mode === "settings" && typeof renderPresets === "function") renderPresets();
};

function updateNetworkStatus() {
    if (!navigator.onLine) {
        addBotMessage("⚠️ **Tidak Ada Koneksi**\nSaya butuh internet untuk terhubung ke Gemini AI.");
        sendBtn.disabled = true;
    } else { sendBtn.disabled = false; }
}

// ── EVENT LISTENERS ───────────────────────────────────────────────────────────
function setupEventListeners() {
    sendBtn.addEventListener("click", handleSendMessage);
    userInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSendMessage(); }
    });
    userInput.addEventListener("input", function() {
        this.style.height = "auto";
        this.style.height = Math.min(this.scrollHeight, 150) + "px";
    });

    // Settings
    // Settings Input Save
    const saveBtn = document.getElementById("save-settings");
    if (saveBtn) {
        saveBtn.addEventListener("click", () => {
            const k = document.getElementById("api-key-input").value.trim();
            if (k) { 
                apiKey = k; 
                localStorage.setItem("gemini_api_key", k); 
                showToast("✅ API Key disimpan"); 
            }
        });
    }

    // Main menu toggle (simulate lang toggle functionality or settings)
    const menuBtn = document.getElementById("menu-btn");
    if (menuBtn) {
        menuBtn.addEventListener("click", () => {
            currentLang = currentLang === "ID" ? "EN" : "ID";
            localStorage.setItem("ab_lang", currentLang);
            showToast(currentLang === "ID" ? "🇮🇩 Mode Indonesia aktif" : "🇬🇧 English mode active");
        });
    }

    // File upload
    const uploadBtn = document.getElementById("upload-btn");
    const fileInput = document.getElementById("file-upload");
    if (uploadBtn && fileInput) {
        uploadBtn.addEventListener("click", () => fileInput.click());
        fileInput.addEventListener("change", (e) => {
            const file = e.target.files[0];
            if (!file) return;
            window.currentFile = file;
            const ex = document.getElementById("file-indicator");
            if (ex) ex.remove();
            const ind = document.createElement("div");
            ind.id = "file-indicator";
            ind.innerHTML = `<span>📎 <strong>${file.name}</strong></span><button id="remove-file-btn" style="background:none;border:none;cursor:pointer;color:#ef4444;font-weight:bold;">✕</button>`;
            document.querySelector(".input-area").parentElement.insertBefore(ind, document.querySelector(".input-area"));
            document.getElementById("remove-file-btn").addEventListener("click", () => { window.currentFile = null; ind.remove(); fileInput.value = ""; });
        });
    }
}

// ── MESSAGE HANDLING ──────────────────────────────────────────────────────────
async function handleSendMessage() {
    const text = userInput.value.trim();
    if (!text) return;
    addUserMessage(text);
    userInput.value = "";
    userInput.style.height = "auto";
    sendBtn.disabled = true;

    try {
        const loadingId = addLoadingMessage();
        const response = await callGeminiAPI(text);
        removeMessage(loadingId);
        const msgId = "msg-" + Date.now();
        addBotMessage(response, msgId);
        if (response && !response.startsWith("❌")) insertIntoDocument(response);

        // Save to history
        if (typeof saveToHistory === "function") saveToHistory(text, response);
        // Cache for offline
        if (typeof cacheOffline === "function") cacheOffline("last_response", { q: text, a: response });
    } catch (err) {
        removeMessage("loading-msg");
        addBotMessage(`❌ **Error**: ${err.message}\n\nPeriksa koneksi atau API Key.`);
    } finally {
        window.currentFile = null;
        const ind = document.getElementById("file-indicator");
        if (ind) ind.remove();
        const fi = document.getElementById("file-upload");
        if (fi) fi.value = "";
        sendBtn.disabled = false;
        userInput.focus();
    }
}

// ── GEMINI API ────────────────────────────────────────────────────────────────
async function callGeminiAPI(prompt) {
    if (!apiKey) throw new Error("API Key kosong. Buka Settings ⚙️");

    // Offline fallback
    if (!navigator.onLine) {
        if (typeof getOfflineCache === "function") {
            const cached = getOfflineCache("last_response");
            if (cached) return `*(Mode Offline — jawaban dari cache)*\n\n${cached.a}`;
        }
        throw new Error("Tidak ada koneksi internet.");
    }

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=${apiKey}`;
    const systemRole = buildSystemPrompt(Office.context.host, currentLang);

    // Bimbingan mode prefix
    const bimbinganPrefix = typeof getBimbinganPrefix === "function" ? getBimbinganPrefix() : "";

    let docContext = "";
    try { docContext = await getDocumentContext(); } catch {}

    let filePart = null, fileText = "";
    if (window.currentFile) {
        const file = window.currentFile;
        if (file.type.startsWith("image/")) {
            const b64 = await fileToBase64(file);
            filePart = { inlineData: { mimeType: file.type, data: b64 } };
        } else {
            fileText = await fileToText(file);
            fileText = `\n\n[File: ${file.name}]\n${fileText.substring(0, 6000)}\n[/File]\n`;
        }
    }

    const fullPrompt = bimbinganPrefix + systemRole + "\n\nKonteks:\n" + docContext + fileText + "\n\nPermintaan: " + prompt;
    const textPart = { text: fullPrompt };
    const parts = filePart ? [textPart, filePart] : [textPart];
    const payload = { contents: [{ role: "user", parts }] };

    const res = await fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
    if (!res.ok) { const err = await res.text(); throw new Error(`API Error (${res.status}): ${err.substring(0,200)}`); }
    const data = await res.json();
    if (!data.candidates?.length) return "Gemini tidak menghasilkan respons.";
    return data.candidates[0]?.content?.parts?.[0]?.text || "Tidak ada respons.";
}

function buildSystemPrompt(host, lang) {
    const isEN = lang === "EN";
    if (host === Office.HostType.Word) {
        return isEN
            ? `You are an expert academic writer. Rules: 1) Use Markdown. 2) English text ALWAYS in *italic*. 3) For journals: Title, Abstract (EN+ID), Introduction, Literature Review, Methodology, Results & Discussion, Conclusion, References (IEEE). 4) Citations: IEEE format. 5) Output ONLY document content. 6) If user requests document formatting like A4 paper, two columns, or specific font, output a JSON block AT THE VERY BEGINNING like: \`\`\`json\n{"layout": {"paperSize": "A4", "columns": 2, "font": "Times New Roman"}}\n\`\`\` before the text.`
            : `Kamu adalah penulis akademis ahli. Aturan: 1) Gunakan Markdown. 2) Teks Inggris SELALU *italic*. 3) Untuk jurnal: Judul, Abstrak (ID+EN), Pendahuluan, Tinjauan Pustaka, Metodologi, Hasil & Pembahasan, Kesimpulan, Daftar Pustaka (APA). 4) Sitasi: format APA. 5) Output HANYA konten dokumen. 6) Jika user meminta format dokumen seperti kertas A4, dua kolom, atau font khusus, keluarkan blok JSON DI PALING ATAS seperti: \`\`\`json\n{"layout": {"paperSize": "A4", "columns": 2, "font": "Times New Roman"}}\n\`\`\` sebelum menyajikan teks.`;
    } else if (host === Office.HostType.Excel) {
        return `Kamu adalah Excel Expert. Aturan: 1) Formula: output HANYA rumus dimulai =. 2) Data/tabel: CSV atau Markdown table. 3) Statistik: hitung N,Mean,Median,StdDev,Min,Max,Range dalam format tabel. 4) Interpretasi: narasi profesional. 5) Tanpa filler.`;
    } else if (host === Office.HostType.PowerPoint) {
        return `Kamu adalah Presentation Expert. Aturan: 1) Multi-slide: JSON Array [{"title":"...","points":["..."],"notes":"..."}]. 2) Single: TITLE:[Judul]\\n- Poin. 3) Notes: 2-3 kalimat informatif per slide. 4) Tanpa Markdown.`;
    }
    return "Kamu adalah asisten AI yang membantu mahasiswa. Jawab dengan ringkas dan akurat dalam Bahasa Indonesia.";
}

// ── DOCUMENT CONTEXT ──────────────────────────────────────────────────────────
async function getDocumentContext() {
    return new Promise((resolve) => {
        const host = Office.context.host;
        if (host === Office.HostType.Word) {
            Word.run(async (ctx) => {
                const sel = ctx.document.getSelection(); sel.load("text"); await ctx.sync();
                if (sel.text?.trim()) return resolve(sel.text);
                const body = ctx.document.body; body.load("text"); await ctx.sync();
                resolve(body.text.substring(0, 5000));
            }).catch(() => resolve(""));
        } else if (host === Office.HostType.Excel) {
            Excel.run(async (ctx) => {
                const range = ctx.workbook.getSelectedRange(); range.load("text"); await ctx.sync();
                resolve(range.text.map(r => r.join(", ")).join("\n"));
            }).catch(() => resolve(""));
        } else {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (res) => resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value : ""));
        }
    });
}

function fileToBase64(file) {
    return new Promise((resolve) => { const r = new FileReader(); r.onloadend = () => resolve(r.result.split(",")[1]); r.readAsDataURL(file); });
}
function fileToText(file) {
    return new Promise((resolve) => { const r = new FileReader(); r.onload = (e) => resolve(e.target.result); r.readAsText(file); });
}

// ── UI HELPERS ────────────────────────────────────────────────────────────────
function addUserMessage(text) {
    const div = document.createElement("div");
    div.className = "message user-message";
    div.innerHTML = `<div class="message-content">${escapeHtml(text).replace(/\n/g,"<br>")}</div>`;
    chatContainer.appendChild(div); scrollToBottom();
}

function addLoadingMessage() {
    const id = "loading-" + Date.now();
    const div = document.createElement("div"); div.id = id; div.className = "message bot-message";
    div.innerHTML = `<div class="message-content"><span style="animation:fadeIn 1s infinite"></span> Memproses...</div>`;
    chatContainer.appendChild(div); scrollToBottom(); return id;
}

function removeMessage(id) { const el = document.getElementById(id); if (el) el.remove(); }

function addBotMessage(text, msgId) {
    const div = document.createElement("div");
    div.className = "message bot-message";
    if (msgId) div.id = msgId;
    div.innerHTML = `<div class="message-content">${marked.parse(text)}</div>`;
    div.querySelectorAll("pre code").forEach((b) => hljs.highlightElement(b));

    // Add rating buttons
    if (typeof addRatingButtons === "function" && msgId) addRatingButtons(div, msgId);

    chatContainer.appendChild(div); scrollToBottom();
}

function scrollToBottom() { chatContainer.scrollTop = chatContainer.scrollHeight; }

function escapeHtml(t) { return t.replace(/[&<>"']/g, m => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[m])); }

function showToast(msg) {
    const t = document.createElement("div"); t.className = "toast"; t.textContent = msg;
    document.body.appendChild(t); setTimeout(() => t.remove(), 2300);
}

// ── QUICK ACTIONS ─────────────────────────────────────────────────────────────
function setupQuickActions(host) {
    const container = document.createElement("div");
    container.className = "quick-actions";
    let actions = [];

    if (host === Office.HostType.Word) {
        actions = [
            { label: "📄 Jurnal ID",     prompt: "Buatkan jurnal ilmiah lengkap Bahasa Indonesia tentang {topic}. Sertakan Judul, Abstrak (ID+EN), Pendahuluan, Tinjauan Pustaka, Metodologi, Hasil & Pembahasan, Kesimpulan, Daftar Pustaka APA. Teks Inggris ditulis italic." },
            { label: "📜 Journal EN",    prompt: "Write a complete English academic journal about {topic}. Include Title, Abstract (EN+ID), Introduction, Literature Review, Methodology, Results, Conclusion, References (IEEE). English text in italic." },
            { label: "✂️ Parafrase",    prompt: "PARAFRASE" },
            { label: "🔍 Proofreading", prompt: "PROOFREADING" },
            { label: "📐 Outline",      prompt: "OUTLINE" },
            { label: "🎓 Bimbingan",    prompt: "BIMBINGAN" },
            { label: "🔖 Sitasi APA",   prompt: "Buatkan daftar pustaka format APA untuk: {topic}" },
            { label: "🔖 Sitasi IEEE",  prompt: "Buatkan referensi format IEEE untuk: {topic}" },
            { label: "🌐 Ke Inggris",   prompt: "Translate teks yang dipilih ke Bahasa Inggris akademik. Hasil ditulis italic." },
            { label: "🇮🇩 Ke Indo",     prompt: "Terjemahkan teks terpilih ke Bahasa Indonesia baku dan formal." },
            { label: "✅ Grammar",       prompt: "Perbaiki tata bahasa, ejaan, dan gaya penulisan teks yang dipilih." },
            { label: "📋 Abstrak 2Bhs", prompt: "Buatkan abstrak Bahasa Indonesia dan Bahasa Inggris (italic) untuk {topic}. 150-250 kata + kata kunci." },
        ];
    } else if (host === Office.HostType.Excel) {
        actions = [
            { label: "🧮 Rumus",          prompt: "Buatkan rumus Excel untuk: " },
            { label: "📊 Statistik",       prompt: "Hitung N, Mean, Median, Std Dev, Min, Max, Range dari data terpilih. Format tabel." },
            { label: "📈 Regresi",         prompt: "REGRESI" },
            { label: "📝 Interpretasi",    prompt: "INTERPRETASI" },
            { label: "📋 Tabel Frekuensi", prompt: "TABEL:frekuensi" },
            { label: "📋 Tabel Kuesioner", prompt: "TABEL:kuesioner" },
            { label: "🔍 Outlier",         prompt: "Analisis data terpilih: identifikasi outlier dan anomali. Berikan rekomendasi." },
            { label: "📝 Laporan",         prompt: "Buat narasi laporan singkat profesional dari data terpilih." },
            { label: "📈 Grafik",          prompt: "Buatkan grafik dari data terpilih." },
            { label: "🧹 Cek Kosong",      prompt: "Analisis data terpilih: temukan sel kosong, duplikat, format tidak konsisten." },
        ];
    } else if (host === Office.HostType.PowerPoint) {
        actions = [
            { label: "🎯 PPT dari File",   prompt: "SLIDE_FROM_FILE" },
            { label: "📑 Outline 10 Slide",prompt: "Buatkan outline presentasi 10 slide dengan speaker notes tentang {topic}. Format JSON Array." },
            { label: "🎤 Slide + Notes",   prompt: "Buatkan slide presentasi dengan catatan pembicara tentang {topic}. Format JSON Array." },
            { label: "🎤 Timer Latihan",  prompt: "TIMER" },
            { label: "🌐 Translate Slide", prompt: "Translate konten slide ini ke Bahasa Indonesia: " },
            { label: "✨ Perbaiki Teks",   prompt: "Perbaiki grammar dan profesionalitas teks slide ini: " },
        ];
    } else {
        actions = [
            { label: "📝 Rangkum", prompt: "Rangkum ini secara singkat." },
            { label: "✅ Grammar", prompt: "Cek tata bahasa dan gaya penulisan." },
        ];
    }

    actions.forEach((action) => {
        const btn = document.createElement("button");
        btn.className = "action-pill";
        btn.textContent = action.label;
        btn.onclick = () => handleActionPill(action.prompt);
        container.appendChild(btn);
    });

    const inputArea = document.querySelector(".input-area");
    inputArea.parentElement.insertBefore(container, inputArea);
}

function handleActionPill(prompt) {
    // Special commands
    if (prompt === "PARAFRASE")     { const p = document.getElementById("word-tools-panel"); if(p){p.classList.remove("hidden"); document.getElementById("host-tools-btn")?.classList.add("active");} showToast("Pilih teks lalu klik level parafrase di Word Tools ✍️"); return; }
    if (prompt === "PROOFREADING")  { proofreadingMendalam(); return; }
    if (prompt === "OUTLINE")       { openOutlineBuilder(); const p = document.getElementById("word-tools-panel"); if(p) p.classList.remove("hidden"); return; }
    if (prompt === "BIMBINGAN")     { toggleBimbinganSkripsi(); return; }
    if (prompt === "REGRESI")       { analisisRegresi(); return; }
    if (prompt === "INTERPRETASI")  { interpretasiStatistik(); return; }
    if (prompt === "SLIDE_FROM_FILE") { slideFromUploadedFile(); return; }
    if (prompt === "TIMER")         { const p = document.getElementById("ppt-tools-panel"); if(p) p.classList.remove("hidden"); showToast("Timer ada di PPT Tools 🎤"); return; }
    if (prompt.startsWith("TABEL:"))  { insertTemplateTabel(prompt.split(":")[1]); return; }

    const curr = userInput.value.trim();
    if (prompt.includes("{topic}")) {
        userInput.value = prompt.replace("{topic}", curr ? curr : "[TOPIK]"); 
    } else if (prompt.endsWith(": ")) {
        userInput.value = prompt + (curr ? curr : ""); 
    } else {
        // If there's already text and we click a generic action, maybe just replace or append.
        // For simplicity, we just replace it as requested.
        userInput.value = prompt; 
    }
    
    // Terapkan penyesuaian tinggi (auto-resize) untuk SEMUA tombol
    userInput.dispatchEvent(new Event('input', { bubbles: true }));
    userInput.focus(); 
}

// ── INSERT INTO DOCUMENT ──────────────────────────────────────────────────────
function insertIntoDocument(text) {
    const host = Office.context.host;
    if (host === Office.HostType.Word) {
        // Attempt to parse formatting layout JSON blocks
        let cleanText = text;
        let layoutCmds = null;
        try {
            const layoutMatch = text.match(/```json\s*\n(\{\s*"layout".*?\})\n```/is) || text.match(/(\{\s*"layout".*?\})/is);
            if (layoutMatch) {
                const parsed = JSON.parse(layoutMatch[1]);
                if (parsed.layout) {
                    layoutCmds = parsed.layout;
                    cleanText = text.replace(layoutMatch[0], "").trim();
                }
            }
        } catch (e) {
            console.log("No layout JSON found or invalid format.");
        }

        Word.run(async (ctx) => {
            if (layoutCmds) {
                const sections = ctx.document.sections;
                sections.load("items");
                await ctx.sync();
                const section = sections.items[0];
                
                if (layoutCmds.paperSize) {
                    // Try to set page size if supported, A4 is generally default but we can set dimensions
                    // Note: direct page size string like 'A4' might not be fully supported in all APIs without specific dimensions
                    // So we do a generic approach or page setup
                    try {
                        // Page setup properties
                        if (layoutCmds.paperSize.toLowerCase() === 'a4') {
                            section.getPageSetup().pageHeight = 842; // Points (29.7cm)
                            section.getPageSetup().pageWidth = 595;  // Points (21cm)
                        }
                    } catch (e) { console.warn("PageSetup sizing failed", e); }
                }
                
                // Set columns if requested
                if (layoutCmds.columns && layoutCmds.columns > 1) {
                    // Not all versions of Word API support setting columns directly on section yet,
                    // but we will attempt to set it if available or log.
                    console.log("Requested columns: " + layoutCmds.columns);
                }
            }

            const html = marked.parse(cleanText);
            const selection = ctx.document.getSelection();
            selection.insertHtml(html, Word.InsertLocation.after);
            
            // Try to format text
            if (layoutCmds && layoutCmds.font) {
                // Apply font specifically
                const body = ctx.document.body;
                body.font.name = layoutCmds.font;
                if (layoutCmds.fontSize) {
                   body.font.size = layoutCmds.fontSize;
                }
            }
            
            await ctx.sync();
            showToast("✅ Berhasil disisipkan" + (layoutCmds ? " & format diterapkan" : ""));
        }).catch(err => {
            console.error(err);
            // Fallback
            Office.context.document.setSelectedDataAsync(marked.parse(cleanText), { coercionType: Office.CoercionType.Html }, (res) => {
                if (res.status === Office.AsyncResultStatus.Failed)
                    Office.context.document.setSelectedDataAsync(cleanText, { coercionType: Office.CoercionType.Text });
            });
        });
    } else if (host === Office.HostType.Excel) {
        const low = text.toLowerCase();
        if (low.includes("chart") || low.includes("grafik")) runExcelChartGen(text);
        else if (low.includes("statistik") || low.includes("mean") || low.includes("median") || low.includes("regresi")) runExcelStatInsert(text);
        else runExcelDataGen(text);
    } else if (host === Office.HostType.PowerPoint) {
        Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (res) => {
            if (res.status === Office.AsyncResultStatus.Failed) runPowerPointSlideGen(text);
        });
    }
}

// ── EXCEL CORE HANDLERS ───────────────────────────────────────────────────────
async function runExcelDataGen(text) {
    const trimmed = text.trim();
    if (trimmed.startsWith("=")) {
        await Excel.run(async (ctx) => { ctx.workbook.getSelectedRange().formulas = [[trimmed]]; await ctx.sync(); showToast("✅ Formula disisipkan!"); }).catch(console.error);
        return;
    }
    let rows = text.split("\n").filter(r => r.trim());
    if (rows[0] && /^(sure|here|berikut)/i.test(rows[0])) rows.shift();
    if (!rows.length) return;
    let delim = ",";
    if (rows[0].includes("|")) { delim = "|"; rows = rows.filter(r => !r.includes("---")); }
    const matrix = rows.map(r => r.split(delim).map(c => c.trim().replace(/^\||\|$/g, "")));
    if (!matrix.length) return;
    await Excel.run(async (ctx) => {
        const tgt = ctx.workbook.getSelectedRange().getResizedRange(matrix.length - 1, matrix[0].length - 1);
        tgt.values = matrix; tgt.format.autofitColumns(); await ctx.sync(); showToast(`✅ ${matrix.length} baris disisipkan!`);
    }).catch(console.error);
}

async function runExcelStatInsert(text) {
    let rows = text.split("\n").filter(r => r.includes("|") && !r.includes("---"));
    if (!rows.length) { Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }); return; }
    const matrix = rows.map(r => r.split("|").map(c => c.trim()).filter(c => c !== ""));
    await Excel.run(async (ctx) => {
        const tgt = ctx.workbook.getSelectedRange().getResizedRange(matrix.length - 1, matrix[0].length - 1);
        tgt.values = matrix; tgt.format.autofitColumns(); await ctx.sync(); showToast("✅ Statistik disisipkan!");
    }).catch(console.error);
}

async function runExcelChartGen(text) {
    let type = "ColumnClustered";
    if (/pie|lingkaran/i.test(text)) type = "Pie";
    else if (/line|garis/i.test(text)) type = "Line";
    else if (/bar|batang/i.test(text)) type = "BarClustered";
    await Excel.run(async (ctx) => {
        const chart = ctx.workbook.worksheets.getActiveWorksheet().charts.add(type, ctx.workbook.getSelectedRange());
        chart.title.text = "Generated Chart"; await ctx.sync(); showToast("✅ Grafik dibuat!");
    }).catch(console.error);
}

// ── PPT CORE HANDLER ──────────────────────────────────────────────────────────
async function runPowerPointSlideGen(text) {
    let slidesData = [];
    try { const m = text.match(/\[[\s\S]*\]/); if (m) slidesData = JSON.parse(m[0]); } catch {}
    if (!slidesData.length) {
        const lines = text.split("\n").filter(l => l.trim());
        if (lines.length) slidesData.push({ title: lines[0].replace(/^(TITLE:|[#*]+)\s*/i,"").trim(), points: lines.slice(1).map(l => l.replace(/^[-*•]\s*/,"")), notes: "" });
    }
    if (!slidesData.length) return;
    await PowerPoint.run(async (ctx) => {
        for (const data of slidesData) {
            const slide = ctx.presentation.slides.add();
            slide.shapes.getItemAt(0).textFrame.textRange.text = data.title || "Slide";
            slide.shapes.getItemAt(1).textFrame.textRange.text = Array.isArray(data.points) ? data.points.join("\n") : (data.points || "");
        }
        await ctx.sync(); showToast(`✅ ${slidesData.length} slide dibuat!`);
    }).catch((err) => { console.error(err); Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }); });
}

// ── TEMPLATE MANAGER (Word templates) ────────────────────────────────────────
window.applyBuiltinTemplate = function(type) {
    if (Office.context.host !== Office.HostType.Word) { showToast("⚠️ Template hanya untuk Microsoft Word"); return; }
    const templates = {
        "jurnal-indonesia": `# [JUDUL JURNAL]\n\n**Penulis:** [Nama]¹  \n**Afiliasi:** ¹[Universitas]  \n**Email:** [email@domain.com]\n\n---\n\n## Abstrak\n\n[Abstrak Bahasa Indonesia, 150-250 kata: latar belakang, tujuan, metode, hasil, kesimpulan.]\n\n**Kata Kunci:** kata1; kata2; kata3\n\n---\n\n## *Abstract*\n\n*[Abstract in English, 150-250 words.]*\n\n***Keywords:*** *keyword1; keyword2; keyword3*\n\n---\n\n## 1. Pendahuluan\n\n## 2. Tinjauan Pustaka\n\n## 3. Metodologi Penelitian\n\n### 3.1 Jenis Penelitian\n\n### 3.2 Teknik Pengumpulan Data\n\n### 3.3 Teknik Analisis Data\n\n## 4. Hasil dan Pembahasan\n\n### 4.1 Hasil\n\n### 4.2 Pembahasan\n\n## 5. Kesimpulan\n\n## Daftar Pustaka\n\n[Penulis, A. (Tahun). *Judul*. Jurnal, Vol(No), Hal.]\n`,
        "jurnal-inggris": `# [JOURNAL TITLE]\n\n**Authors:** [Name]¹  \n**Affiliation:** ¹[University]  \n**Email:** [email@domain.com]\n\n---\n\n## *Abstract*\n\n*[Abstract in English, 150-250 words.]*\n\n***Keywords:*** *keyword1; keyword2; keyword3*\n\n---\n\n## Abstrak\n\n[Abstrak Bahasa Indonesia, 150-250 kata.]\n\n**Kata Kunci:** kata1; kata2; kata3\n\n---\n\n## 1. *Introduction*\n\n## 2. *Literature Review*\n\n## 3. *Methodology*\n\n## 4. *Results and Discussion*\n\n## 5. *Conclusion*\n\n## *References*\n\n*[1] A. Author, "Title," Journal, vol. X, pp. Y-Z, Year.*\n`,
        "artikel-ilmiah": `# [JUDUL ARTIKEL]\n\n**Penulis:** [Nama] — [Institusi]\n**Tanggal:** [Tanggal]\n\n---\n\n## Abstrak\n\n[100-150 kata.]\n\n**Kata Kunci:** kata1, kata2, kata3\n\n---\n\n## 1. Pendahuluan\n\n## 2. Pembahasan\n\n### 2.1 [Sub-topik 1]\n\n### 2.2 [Sub-topik 2]\n\n## 3. Penutup\n\n## Daftar Pustaka\n`,
        "skripsi": `# JUDUL SKRIPSI\n\n**Nama:** [Nama]  \n**NIM:** [NIM]  \n**Program Studi:** [Prodi]  \n**Universitas:** [Universitas]  \n**Tahun:** [Tahun]\n\n---\n\n# BAB I — PENDAHULUAN\n\n## 1.1 Latar Belakang\n\n## 1.2 Rumusan Masalah\n\n## 1.3 Tujuan\n\n## 1.4 Manfaat\n\n## 1.5 Batasan\n\n---\n\n# BAB II — TINJAUAN PUSTAKA\n\n## 2.1 [Teori Utama]\n\n## 2.2 Penelitian Terdahulu\n\n---\n\n# BAB III — METODOLOGI\n\n## 3.1 Jenis Penelitian\n\n## 3.2 Populasi dan Sampel\n\n## 3.3 Pengumpulan Data\n\n## 3.4 Analisis Data\n\n---\n\n# BAB IV — HASIL DAN PEMBAHASAN\n\n## 4.1 Hasil\n\n## 4.2 Pembahasan\n\n---\n\n# BAB V — PENUTUP\n\n## 5.1 Kesimpulan\n\n## 5.2 Saran\n\n---\n\n# DAFTAR PUSTAKA\n\n---\n\n# LAMPIRAN\n`,
        "prosiding": `# [JUDUL PROSIDING]\n\n**Penulis:** [Nama]¹  \n**Konferensi:** [Nama Konferensi, Tahun]\n\n---\n\n## *Abstract*\n*[150-200 words]*\n\n***Keywords:*** *keyword1, keyword2*\n\n## 1. Pendahuluan\n\n## 2. Metode\n\n## 3. Hasil\n\n## 4. Kesimpulan\n\n## Referensi\n*[1] A. Author, "Title," Proc. Conf., Year.*\n`,
        "makalah": `# [JUDUL MAKALAH]\n\n**Disusun oleh:** [Nama]  \n**Mata Kuliah:** [MK]  \n**Dosen:** [Nama Dosen]  \n**Tanggal:** [Tanggal]\n\n---\n\n## BAB I — PENDAHULUAN\n\n### 1.1 Latar Belakang\n\n### 1.2 Rumusan Masalah\n\n### 1.3 Tujuan\n\n## BAB II — PEMBAHASAN\n\n### 2.1 [Sub-topik]\n\n## BAB III — PENUTUP\n\n### 3.1 Kesimpulan\n\n### 3.2 Saran\n\n## DAFTAR PUSTAKA\n`
    };
    const content = templates[type];
    if (!content) return;
    Word.run(async (ctx) => {
        ctx.document.body.clear();
        ctx.document.body.insertText(content, Word.InsertLocation.start);
        await ctx.sync();
        showToast("✅ Template diterapkan!");
    }).catch(() => showToast("❌ Gagal menerapkan template"));
    
    const panel = document.getElementById("templates-panel");
    const btn = document.getElementById("nav-templates");
    if (panel) panel.classList.add("hidden");
    if (btn) btn.classList.remove("active");
};

// ── CITATION ──────────────────────────────────────────────────────────────────
window.insertCitation = async function(style) {
    const input = document.getElementById("citation-input").value.trim();
    if (!input) { showToast("⚠️ Masukkan info referensi!"); return; }
    showToast("⏳ Memformat sitasi...");
    try {
        const result = await callGeminiAPI(`Format referensi berikut ke ${style}:\n${input}\nOutput HANYA sitasi, tanpa penjelasan.`);
        if (Office.context.host === Office.HostType.Word) {
            Word.run(async (ctx) => { ctx.document.getSelection().insertText("\n" + result.trim() + "\n", Word.InsertLocation.after); await ctx.sync(); showToast(`✅ Sitasi ${style} disisipkan!`); });
        } else { addBotMessage(`**Sitasi ${style}:**\n\n${result}`); }
        document.getElementById("citation-input").value = "";
    } catch (e) { showToast("❌ Error: " + e.message); }
};
