/*
 * features-word.js — AutomateBoty v8
 * Word-specific: Parafrase, Proofreading, Bimbingan Skripsi, Outline Builder
 * + executeWordAction() for AI-commanded formatting via Action JSON
 */

// ── STATE ─────────────────────────────────────────────────────────────────────
let bimbinganActive = false;
let outlineData = null;

// ── WORD ACTION EXECUTOR ──────────────────────────────────────────────────────
/**
 * Executes a structured Word formatting command received from the AI.
 * Called by executor.js dispatchAction().
 */
window.executeWordAction = async function(action, params) {
    switch (action) {
        case "WORD_DYNAMIC_FORMAT":  return await wordDynamicFormat(params);
        case "WORD_FORMAT_PAGE":     return await wordFormatPage(params);
        case "WORD_FORMAT_FONT":     return await wordFormatFont(params);
        case "WORD_HEADING_STYLES":  return await wordApplyHeadingStyles(params);
        case "WORD_INSERT_TABLE":    return await wordInsertTable(params);
        case "WORD_ADD_CONTENT":     return await wordAddContent(params);
        case "WORD_SET_SPACING":     return await wordSetSpacing(params);
        default:
            console.warn(`⚠️ Word action tidak dikenal: ${action}`);
            throw new Error("UNSUPPORTED_ACTION");
    }
};

// ── DYNAMIC OMNI-FORMATTER ────────────────────────────────────────────────────
async function wordDynamicFormat(p) {
    // Expected p: { tasks: [ { target: "font"|"paragraph"|"page"|"selection", props: { bold: true, size: 14... } } ] }
    if (!p || !p.tasks || !Array.isArray(p.tasks)) return;

    await Word.run(async (ctx) => {
        const body = ctx.document.body;
        const sel  = ctx.document.getSelection();

        for (const task of p.tasks) {
            const tgtName = task.target || "selection";
            const props   = task.props || {};
            
            let applyTo = null;
            if (tgtName === "selection") applyTo = sel;
            else if (tgtName === "font" || tgtName === "body") applyTo = body;
            
            if (applyTo) {
                if (props.font) {
                    if (props.font.name) applyTo.font.name = props.font.name;
                    if (props.font.size) applyTo.font.size = parseFloat(props.font.size);
                    if (props.font.bold !== undefined) applyTo.font.bold = props.font.bold;
                    if (props.font.italic !== undefined) applyTo.font.italic = props.font.italic;
                    if (props.font.color) applyTo.font.color = props.font.color;
                }
                if (tgtName === "paragraph" && applyTo.paragraphs) {
                    const paras = applyTo.paragraphs;
                    paras.load("items");
                    await ctx.sync();
                    for (const para of paras.items) {
                        if (props.alignment) para.alignment = props.alignment; // "Left", "Centered", "Right", "Justified"
                        if (props.lineSpacing) para.lineSpacing = parseFloat(props.lineSpacing) * 12; // Approximation
                        if (props.spaceBefore) para.spaceBefore = parseFloat(props.spaceBefore);
                        if (props.spaceAfter)  para.spaceAfter  = parseFloat(props.spaceAfter);
                        if (props.firstLineIndent) para.firstLineIndent = parseFloat(props.firstLineIndent);
                    }
                } else if (tgtName === "page" && ctx.document.sections) {
                    const sections = ctx.document.sections;
                    sections.load("items");
                    await ctx.sync();
                    for (const section of sections.items) {
                        const pg = section.pageSetup;
                        if (props.marginTop) pg.topMargin = parseFloat(props.marginTop) * 28.35;
                        if (props.marginBottom) pg.bottomMargin = parseFloat(props.marginBottom) * 28.35;
                        if (props.marginLeft) pg.leftMargin = parseFloat(props.marginLeft) * 28.35;
                        if (props.marginRight) pg.rightMargin = parseFloat(props.marginRight) * 28.35;
                        const ori = (props.orientation || "").toLowerCase();
                        if (ori === "landscape") pg.orientation = Word.PageOrientation.landscape;
                        else if (ori === "portrait") pg.orientation = Word.PageOrientation.portrait;

                        if (props.paperSize) {
                            const size = (props.paperSize || "").toLowerCase();
                            if (size === "a4") pg.paperSize = Word.PaperType.a4;
                            else if (size === "letter") pg.paperSize = Word.PaperType.letter;
                            else if (size === "legal") pg.paperSize = Word.PaperType.legal;
                        }
                    }
                }
            }
        }
        await ctx.sync();
    });
    showToast("✅ Format dokumen dinamis diterapkan!");
}

// ── Page Layout ───────────────────────────────────────────────────────────────
async function wordFormatPage(p) {
    // p: { marginTop, marginBottom, marginLeft, marginRight, orientation, columns }
    // All margins in cm (will be converted to points: 1cm = 28.35pt)
    const cmToPt = (cm) => (parseFloat(cm) || 2.54) * 28.35;

    await Word.run(async (ctx) => {
        const body = ctx.document.body;
        body.load("style");
        await ctx.sync();

        // Page setup via document sections (requires Word 2019+ API)
        const sections = ctx.document.sections;
        sections.load("items");
        await ctx.sync();

        for (const section of sections.items) {
            const pg = section.pageSetup;
            if (p.marginTop    !== undefined) pg.topMargin    = cmToPt(p.marginTop);
            if (p.marginBottom !== undefined) pg.bottomMargin = cmToPt(p.marginBottom);
            if (p.marginLeft   !== undefined) pg.leftMargin   = cmToPt(p.marginLeft);
            if (p.marginRight  !== undefined) pg.rightMargin  = cmToPt(p.marginRight);
            if (p.orientation === "landscape") pg.orientation = Word.PageOrientation.landscape;
            else if (p.orientation === "portrait")  pg.orientation = Word.PageOrientation.portrait;
        }
        await ctx.sync();
    });

    // Handle column layout via content insertion if columns specified
    if (p.columns && parseInt(p.columns) > 1) {
        addBotMessage(`✅ **Layout halaman diatur!**\n\nMargin & orientasi sudah diterapkan.\n\n> ⚠️ **Layout ${p.columns} kolom**: Untuk dua kolom, buka **Layout → Columns** di Word dan pilih **Two**. Office.js tidak mendukung pengaturan kolom secara langsung.`);
    } else {
        showToast("✅ Layout halaman diterapkan!");
    }
}

// ── Font Formatting ───────────────────────────────────────────────────────────
async function wordFormatFont(p) {
    // p: { name, sizeBody, sizeHeading, sizeTitle }
    const fontName = p.name || "Times New Roman";
    const sizeBody = parseFloat(p.sizeBody) || 12;

    await Word.run(async (ctx) => {
        const body = ctx.document.body;
        const paras = body.paragraphs;
        paras.load("items,style,font");
        await ctx.sync();

        for (const para of paras.items) {
            para.font.name = fontName;
            // Set size based on style
            const style = (para.style || "").toLowerCase();
            if (style.includes("heading 1") || style.includes("title")) {
                para.font.size = parseFloat(p.sizeTitle) || parseFloat(p.sizeHeading) || 14;
            } else if (style.includes("heading")) {
                para.font.size = parseFloat(p.sizeHeading) || 12;
            } else {
                para.font.size = sizeBody;
            }
        }
        await ctx.sync();
    });
    showToast(`✅ Font "${fontName}" ${sizeBody}pt diterapkan!`);
}

// ── Heading Styles ────────────────────────────────────────────────────────────
async function wordApplyHeadingStyles(p) {
    // p: { uppercase: true/false, bold: true/false }
    await Word.run(async (ctx) => {
        const body = ctx.document.body;
        const paras = body.paragraphs;
        paras.load("items,style,text");
        await ctx.sync();

        for (const para of paras.items) {
            const style = (para.style || "").toLowerCase();
            if (style.includes("heading")) {
                para.font.bold = p.bold !== false;
                if (p.uppercase) {
                    para.load("text");
                    await ctx.sync();
                    if (para.text && !para.text.startsWith("=")) {
                        para.insertText(para.text.toUpperCase(), Word.InsertLocation.replace);
                    }
                }
            }
        }
        await ctx.sync();
    });
    showToast("✅ Heading styles diterapkan!");
}

// ── Insert Table ──────────────────────────────────────────────────────────────
async function wordInsertTable(p) {
    // p: { rows, cols, headers: [], data: [[]], style }
    const rows = parseInt(p.rows) || 3;
    const cols = parseInt(p.cols) || (p.headers?.length) || 3;
    const headers = p.headers || [];
    const data = p.data || [];

    await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        const totalRows = headers.length > 0 ? rows + 1 : rows;
        const table = sel.insertTable(totalRows, cols, Word.InsertLocation.after, []);
        table.styleBuiltIn = Word.Style.tableGrid;
        table.styleBandedRows = true;
        table.load("rows");
        await ctx.sync();

        // Fill headers
        if (headers.length > 0) {
            const hRow = table.rows.items[0];
            hRow.load("cells");
            await ctx.sync();
            headers.forEach((h, i) => {
                if (hRow.cells.items[i]) {
                    hRow.cells.items[i].body.insertText(String(h), Word.InsertLocation.replace);
                    hRow.cells.items[i].body.paragraphs.getLast().font.bold = true;
                }
            });
        }

        // Fill data rows
        const dataOffset = headers.length > 0 ? 1 : 0;
        data.forEach((row, ri) => {
            const tRow = table.rows.items[ri + dataOffset];
            if (!tRow) return;
            tRow.load("cells");
        });
        await ctx.sync();
        data.forEach((row, ri) => {
            const tRow = table.rows.items[ri + dataOffset];
            if (!tRow) return;
            row.forEach((cell, ci) => {
                if (tRow.cells.items[ci]) {
                    tRow.cells.items[ci].body.insertText(String(cell || ""), Word.InsertLocation.replace);
                }
            });
        });

        await ctx.sync();
    });
    showToast(`✅ Tabel ${rows}×${cols} dibuat!`);
}

// ── Add Content (HTML) ────────────────────────────────────────────────────────
async function wordAddContent(p) {
    // p: { html, text, location: "start"|"end"|"replace" }
    const loc = {
        "start": Word.InsertLocation.start,
        "end": Word.InsertLocation.end,
        "replace": Word.InsertLocation.replace,
    }[p.location] || Word.InsertLocation.end;

    const html = p.html || (p.text ? `<p>${p.text}</p>` : "");
    if (!html) return;

    await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.insertHtml(html, loc);
        await ctx.sync();
    });
    showToast("✅ Konten disisipkan!");
}

// ── Line/Paragraph Spacing ────────────────────────────────────────────────────
async function wordSetSpacing(p) {
    // p: { lineSpacing: "single"|"1.5"|"double", spaceBefore, spaceAfter }
    const spacingMap = { "single": 12, "1.5": 18, "double": 24 };
    const lineSpacingPt = spacingMap[p.lineSpacing] || parseFloat(p.lineSpacing) || 12;

    await Word.run(async (ctx) => {
        const body = ctx.document.body;
        const paras = body.paragraphs;
        paras.load("items");
        await ctx.sync();
        for (const para of paras.items) {
            para.lineSpacing = lineSpacingPt;
            if (p.spaceBefore !== undefined) para.spaceBeforeParagraph = parseFloat(p.spaceBefore);
            if (p.spaceAfter  !== undefined) para.spaceAfterParagraph  = parseFloat(p.spaceAfter);
        }
        await ctx.sync();
    });
    showToast(`✅ Spasi "${p.lineSpacing || 'custom'}" diterapkan!`);
}

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
        "ketat":  "Parafrase ketat: tulis ulang hampir seluruh kalimat dengan struktur berbeda, gunakan sudut pandang lain, pastikan 80%+ berbeda dari aslinya.",
    };

    const prompt = `${levelDesc[level] || levelDesc["sedang"]}\n\nTeks asli:\n"${teks}"\n\nOutput HANYA teks hasil parafrase tanpa penjelasan. Pertahankan bahasa (Indonesia/Inggris) yang sama. Jika ada teks Inggris, tetap tulis italic.`;

    addBotMessage(`⏳ Memparafrase teks (level: **${level}**)...`);
    try {
        const result = await callGeminiAPI(prompt);
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
        const btn = document.getElementById("outline-btn");
        if(btn) btn.classList.toggle("active", !panel.classList.contains("hidden"));
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
    const html = marked.parse(`# ${outlineData.title}\n\n${outlineData.chapters.map(ch => `## ${ch.num}: ${ch.title}\n${ch.sections.map(s => `- ${s}`).join("\n")}`).join("\n\n")}`);
    Office.context.document.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, () => showToast("✅ Outline dimasukkan ke dokumen!"));
};
