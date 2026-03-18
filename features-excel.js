/*
 * features-excel.js — AutomateBoty v7
 * Excel: Regresi & Korelasi, Validasi Data, Template Tabel, Interpretasi Statistik
 */

// ── ANALISIS REGRESI & KORELASI ───────────────────────────────────────────────
window.analisisRegresi = async function() {
    let data = [];
    try {
        data = await new Promise((resolve) => {
            Excel.run(async (ctx) => {
                const range = ctx.workbook.getSelectedRange();
                range.load("text,values,rowCount,columnCount");
                await ctx.sync();
                resolve(range.values);
            }).catch(() => resolve([]));
        });
    } catch {}

    if (!data.length || data[0].length < 2) {
        showToast("⚠️ Pilih minimal 2 kolom data numerik (X dan Y)!");
        return;
    }

    const colX = data.map(r => parseFloat(r[0])).filter(v => !isNaN(v));
    const colY = data.map(r => parseFloat(r[1])).filter(v => !isNaN(v));
    const n = Math.min(colX.length, colY.length);

    if (n < 3) { showToast("⚠️ Butuh minimal 3 data untuk regresi!"); return; }

    const xArr = colX.slice(0, n), yArr = colY.slice(0, n);
    const xMean = xArr.reduce((a,b) => a+b, 0) / n;
    const yMean = yArr.reduce((a,b) => a+b, 0) / n;

    let ssXY = 0, ssXX = 0, ssYY = 0;
    for (let i = 0; i < n; i++) {
        ssXY += (xArr[i] - xMean) * (yArr[i] - yMean);
        ssXX += (xArr[i] - xMean) ** 2;
        ssYY += (yArr[i] - yMean) ** 2;
    }

    const b1 = ssXY / ssXX;
    const b0 = yMean - b1 * xMean;
    const r  = ssXY / Math.sqrt(ssXX * ssYY);
    const r2 = r ** 2;

    const stats = { n, b0: b0.toFixed(4), b1: b1.toFixed(4), r: r.toFixed(4), r2: r2.toFixed(4), xMean: xMean.toFixed(3), yMean: yMean.toFixed(3) };

    // Insert result table below selection
    const resultMatrix = [
        ["Statistik Regresi Linear", "Nilai"],
        ["N (jumlah data)", n],
        ["Intercept (b₀)", stats.b0],
        ["Slope/Koefisien (b₁)", stats.b1],
        ["Persamaan Regresi", `Y = ${stats.b0} + ${stats.b1}X`],
        ["Korelasi Pearson (r)", stats.r],
        ["Koefisien Determinasi (R²)", stats.r2],
        ["Rata-rata X", stats.xMean],
        ["Rata-rata Y", stats.yMean],
    ];

    await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getActiveWorksheet();
        const sel = ctx.workbook.getSelectedRange();
        sel.load("rowIndex,columnIndex");
        await ctx.sync();
        const startRow = sel.rowIndex + data.length + 2;
        const startCol = sel.columnIndex;
        const tgt = ws.getRangeByIndexes(startRow, startCol, resultMatrix.length, 2);
        tgt.values = resultMatrix;
        tgt.format.autofitColumns();
        tgt.getRow(0).format.fill.color = "#1e3a5f";
        tgt.getRow(0).format.font.color = "white";
        tgt.getRow(0).format.font.bold = true;
        await ctx.sync();
        showToast("✅ Hasil regresi disisipkan!");
    }).catch(e => showToast("❌ Error: " + e.message));

    // Get AI narration
    const narasiPrompt = `Interpretasikan hasil regresi linear berikut dalam Bahasa Indonesia akademis untuk laporan skripsi/penelitian:\n- N=${n}, b₀=${stats.b0}, b₁=${stats.b1}, r=${stats.r}, R²=${stats.r2}\n- Persamaan: Y = ${stats.b0} + ${stats.b1}X\nDua paragraf: (1) penjelasan persamaan, (2) interpretasi korelasi dan R². Formal, untuk Bab IV skripsi.`;
    addBotMessage("⏳ Membuat narasi interpretasi...");
    try {
        const narasi = await callGeminiAPI(narasiPrompt);
        addBotMessage(`## Interpretasi Regresi\n\n${narasi}`);
        saveToHistory("Analisis Regresi", narasi);
    } catch {}
};

// ── INTERPRETASI STATISTIK (NARASI BAB IV) ────────────────────────────────────
window.interpretasiStatistik = async function() {
    let rawData = "";
    try {
        rawData = await new Promise((resolve) => {
            Excel.run(async (ctx) => {
                const range = ctx.workbook.getSelectedRange();
                range.load("text");
                await ctx.sync();
                resolve(range.text.map(r => r.join("\t")).join("\n"));
            }).catch(() => resolve(""));
        });
    } catch {}

    if (!rawData.trim()) { showToast("⚠️ Pilih data statistik yang ingin diinterpretasikan!"); return; }

    const prompt = `Kamu adalah penulis akademik untuk Bab IV skripsi penelitian. Berdasarkan data statistik berikut, tulis interpretasi naratif dalam Bahasa Indonesia yang formal, mendalam, dan siap masuk ke laporan Bab IV:\n\n${rawData}\n\nStruktur:\n1. Deskripsi hasil statistik\n2. Analisis dan Pembahasan\n3. Kaitan dengan teori/hipotesis\n\nBeri format paragraf, bukan list. Panjang 300-500 kata.`;

    addBotMessage("⏳ Menulis narasi interpretasi untuk Bab IV...");
    try {
        const result = await callGeminiAPI(prompt);
        addBotMessage(result);
        if (Office.context.host === Office.HostType.Word) {
            Office.context.document.setSelectedDataAsync(marked.parse(result), { coercionType: Office.CoercionType.Html });
        }
        saveToHistory("Interpretasi Statistik", result);
    } catch (e) { addBotMessage("❌ Gagal: " + e.message); }
};

// ── TEMPLATE TABEL PENELITIAN ─────────────────────────────────────────────────
const EXCEL_TEMPLATES = {
    "frekuensi": [
        ["No", "Kategori/Interval", "Frekuensi (f)", "Frekuensi Relatif (%)", "Frekuensi Kumulatif"],
        ["1", "", "", "", ""], ["2", "", "", "", ""], ["3", "", "", "", ""],
        ["4", "", "", "", ""], ["5", "", "", "", ""],
        ["", "TOTAL", "=SUM(C2:C6)", "=SUM(D2:D6)", ""],
    ],
    "distribusi": [
        ["No", "Kelas Interval", "Batas Bawah", "Batas Atas", "Titik Tengah", "Frekuensi", "Persentase"],
        ["1","","","","","",""],["2","","","","","",""],["3","","","","","",""],
        ["4","","","","","",""],["5","","","","","",""],
        ["","Total","","","","=SUM(F2:F6)","=SUM(G2:G6)"],
    ],
    "crosstab": [
        ["Variabel X \\ Variabel Y", "Kategori Y1", "Kategori Y2", "Kategori Y3", "Total"],
        ["Kategori X1","","","","=SUM(B2:D2)"],
        ["Kategori X2","","","","=SUM(B3:D3)"],
        ["Kategori X3","","","","=SUM(B4:D4)"],
        ["Total","=SUM(B2:B4)","=SUM(C2:C4)","=SUM(D2:D4)","=SUM(E2:E4)"],
    ],
    "kuesioner": [
        ["No","Pernyataan","SS (5)","S (4)","N (3)","TS (2)","STS (1)","Total","Rata-rata","Ket."],
        ["1","Pernyataan 1","","","","","","=SUM(C2:G2)","=H2/SUM(C2:G2)*5",""],
        ["2","Pernyataan 2","","","","","","=SUM(C3:G3)","=H3/SUM(C3:G3)*5",""],
        ["3","Pernyataan 3","","","","","","=SUM(C4:G4)","=H4/SUM(C4:G4)*5",""],
    ],
    "rangkuman": [
        ["Variabel","N","Mean","Median","Std Dev","Min","Max","Range"],
        ["Variabel 1","","=AVERAGE(B:B)","=MEDIAN(B:B)","=STDEV(B:B)","=MIN(B:B)","=MAX(B:B)",""],
        ["Variabel 2","","","","","","",""],
    ],
};

window.insertTemplateTabel = async function(type) {
    const template = EXCEL_TEMPLATES[type];
    if (!template) { showToast("⚠️ Template tidak ditemukan"); return; }

    await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getActiveWorksheet();
        const sel = ctx.workbook.getSelectedRange();
        sel.load("rowIndex,columnIndex");
        await ctx.sync();

        const tgt = ws.getRangeByIndexes(sel.rowIndex, sel.columnIndex, template.length, template[0].length);
        tgt.values = template;
        tgt.format.autofitColumns();

        // Style header row
        const header = ws.getRangeByIndexes(sel.rowIndex, sel.columnIndex, 1, template[0].length);
        header.format.fill.color = "#1e3a5f";
        header.format.font.color = "white";
        header.format.font.bold = true;
        header.format.borders.getItem("EdgeBottom").style = "Continuous";

        // Add borders
        tgt.format.borders.getItem("InsideHorizontal").style = "Continuous";
        tgt.format.borders.getItem("InsideVertical").style = "Continuous";
        tgt.format.borders.getItem("EdgeTop").style = "Continuous";
        tgt.format.borders.getItem("EdgeBottom").style = "Continuous";
        tgt.format.borders.getItem("EdgeLeft").style = "Continuous";
        tgt.format.borders.getItem("EdgeRight").style = "Continuous";

        await ctx.sync();
        showToast(`✅ Template tabel ${type} disisipkan!`);
    }).catch(e => showToast("❌ Error: " + e.message));
};

// ── VALIDASI DATA & CONDITIONAL FORMATTING ────────────────────────────────────
window.terapkanValidasiData = async function() {
    const input = document.getElementById("validation-input")?.value.trim();
    if (!input) { showToast("⚠️ Masukkan spesifikasi validasi di kotak!"); return; }

    const prompt = `Berikan instruksi langkah demi langkah cara menerapkan Data Validation di Excel untuk: "${input}". Sertakan juga formula Conditional Formatting yang sesuai. Format Markdown. Bahasa Indonesia.`;
    addBotMessage("⏳ Membuat saran validasi data...");
    try {
        const result = await callGeminiAPI(prompt);
        addBotMessage(result);
    } catch (e) { addBotMessage("❌ Gagal: " + e.message); }

    // Also try programmatic dropdown if simple comma-separated list
    const items = input.includes(",") ? input.split(",").map(i => i.trim()) : null;
    if (items && items.length > 1 && items.length <= 20) {
        await Excel.run(async (ctx) => {
            const range = ctx.workbook.getSelectedRange();
            range.dataValidation.rule = {
                list: { inCellDropDown: true, source: items.join(",") }
            };
            await ctx.sync();
            showToast("✅ Dropdown diterapkan ke sel terpilih!");
        }).catch(() => {/* ignore if no selection */});
    }
};
