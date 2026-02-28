/*
 * features-excel.js — AutomateBoty v8
 * Excel: Regresi & Korelasi, Validasi Data, Template Tabel, Interpretasi Statistik
 * + executeExcelAction() for AI-commanded operations via Action JSON
 */

// ── EXCEL ACTION EXECUTOR ─────────────────────────────────────────────────────
window.executeExcelAction = async function(action, params) {
    switch (action) {
        case "EXCEL_DYNAMIC_FORMAT":       return await excelDynamicFormat(params);
        case "EXCEL_INSERT_DATA":          return await excelInsertData(params);
        case "EXCEL_CREATE_CHART":         return await excelCreateChart(params);
        case "EXCEL_FORMAT_CELLS":         return await excelFormatCells(params);
        case "EXCEL_INSERT_FORMULA":       return await excelInsertFormula(params);
        case "EXCEL_CONDITIONAL_FORMAT":   return await excelConditionalFormat(params);
        case "EXCEL_SORT":                 return await excelSort(params);
        default:
            console.warn(`⚠️ Excel action tidak dikenal: ${action}`);
            throw new Error("UNSUPPORTED_ACTION");
    }
};

// ── DYNAMIC OMNI-FORMATTER ────────────────────────────────────────────────────
async function excelDynamicFormat(p) {
    // Expected p: { tasks: [ { target: "selection"|"A1:B2"|"sheet", props: { fill: "#f00", font: { bold: true }, autoFit: true... } } ] }
    if (!p || !p.tasks || !Array.isArray(p.tasks)) return;

    await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getActiveWorksheet();
        const sel = ctx.workbook.getSelectedRange();

        for (const task of p.tasks) {
            const tgtName = task.target || "selection";
            const props   = task.props || {};
            
            let applyTo = null;
            if (tgtName === "selection") applyTo = sel;
            else if (tgtName === "sheet") applyTo = ws.getRange();
            else applyTo = ws.getRange(tgtName); // Assume address like "A1:C10"

            if (applyTo) {
                if (props.fill) applyTo.format.fill.color = props.fill;
                
                if (props.font) {
                    if (props.font.name) applyTo.format.font.name = props.font.name;
                    if (props.font.size) applyTo.format.font.size = parseFloat(props.font.size);
                    if (props.font.bold !== undefined) applyTo.format.font.bold = props.font.bold;
                    if (props.font.italic !== undefined) applyTo.format.font.italic = props.font.italic;
                    if (props.font.color) applyTo.format.font.color = props.font.color;
                }

                if (props.alignment) {
                    if (props.alignment.horizontal) applyTo.format.horizontalAlignment = props.alignment.horizontal;
                    if (props.alignment.vertical) applyTo.format.verticalAlignment = props.alignment.vertical;
                }

                if (props.numberFormat) {
                    applyTo.numberFormat = [[props.numberFormat]];
                }

                if (props.wrapText !== undefined) applyTo.format.wrapText = props.wrapText;

                if (props.borders) {
                    // Expecting { style: "Continuous", weight: "Thin", color: "#000" }
                    const b = applyTo.format.borders;
                    const items = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"];
                    for (const item of items) {
                        b.getItem(item).style = props.borders.style || "Continuous";
                        if (props.borders.weight) b.getItem(item).weight = props.borders.weight;
                        if (props.borders.color) b.getItem(item).color = props.borders.color;
                    }
                }

                if (props.autoFit) applyTo.format.autofitColumns();
            }
        }
        await ctx.sync();
    });
    showToast("✅ Format sel dinamis diterapkan!");
}

// ── Insert Structured Data ────────────────────────────────────────────────────
async function excelInsertData(p) {
    // p: { headers: [], rows: [[]], startRow?, startCol?, styleHeader? }
    const headers = p.headers || [];
    const rows = p.rows || [];
    const allRows = headers.length > 0 ? [headers, ...rows] : rows;

    if (!allRows.length) { showToast("⚠️ Tidak ada data untuk disisipkan"); return; }

    await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getActiveWorksheet();
        const sel = ctx.workbook.getSelectedRange();
        sel.load("rowIndex,columnIndex");
        await ctx.sync();

        const startRow = p.startRow !== undefined ? p.startRow : sel.rowIndex;
        const startCol = p.startCol !== undefined ? p.startCol : sel.columnIndex;
        const maxCols = Math.max(...allRows.map(r => r.length));
        const tgt = ws.getRangeByIndexes(startRow, startCol, allRows.length, maxCols);
        tgt.values = allRows.map(r => { while (r.length < maxCols) r.push(""); return r; });
        tgt.format.autofitColumns();

        if (headers.length > 0 && p.styleHeader !== false) {
            const hdr = ws.getRangeByIndexes(startRow, startCol, 1, maxCols);
            hdr.format.fill.color = "#1e3a5f";
            hdr.format.font.color = "white";
            hdr.format.font.bold = true;
        }
        await ctx.sync();
        showToast(`✅ ${allRows.length} baris data disisipkan!`);
    }).catch(e => showToast("❌ Error: " + e.message));
}

// ── Create Chart ──────────────────────────────────────────────────────────────
async function excelCreateChart(p) {
    // p: { type: "Bar"|"Column"|"Line"|"Pie"|"Area", title, useSelection }
    const typeMap = {
        "bar":     "BarClustered", "column": "ColumnClustered",
        "line":    "Line",         "pie":    "Pie",
        "area":    "Area",         "scatter":"XYScatter",
    };
    const chartType = typeMap[(p.type || "column").toLowerCase()] || "ColumnClustered";

    await Excel.run(async (ctx) => {
        const ws  = ctx.workbook.worksheets.getActiveWorksheet();
        const src = ctx.workbook.getSelectedRange();
        const chart = ws.charts.add(chartType, src, "Auto");
        if (p.title) chart.title.text = p.title;
        chart.legend.visible = true;
        await ctx.sync();
        showToast(`✅ Grafik "${p.type || "kolom"}" dibuat!`);
    }).catch(e => showToast("❌ Error grafik: " + e.message));
}

// ── Format Cells ──────────────────────────────────────────────────────────────
async function excelFormatCells(p) {
    // p: { bgColor, fontColor, bold, italic, fontSize, numberFormat, range? }
    await Excel.run(async (ctx) => {
        let range;
        if (p.range) {
            range = ctx.workbook.worksheets.getActiveWorksheet().getRange(p.range);
        } else {
            range = ctx.workbook.getSelectedRange();
        }
        range.load("address");
        await ctx.sync();

        if (p.bgColor)       range.format.fill.color       = p.bgColor;
        if (p.fontColor)     range.format.font.color       = p.fontColor;
        if (p.bold !== undefined)   range.format.font.bold = p.bold;
        if (p.italic !== undefined) range.format.font.italic = p.italic;
        if (p.fontSize)      range.format.font.size        = parseFloat(p.fontSize);
        if (p.numberFormat)  range.numberFormat            = [[p.numberFormat]];
        if (p.wrapText !== undefined) range.format.wrapText = p.wrapText;
        if (p.horizontalAlignment) range.format.horizontalAlignment = p.horizontalAlignment;
        await ctx.sync();
        showToast(`✅ Format sel diterapkan!`);
    }).catch(e => showToast("❌ Error format: " + e.message));
}

// ── Insert Formula ────────────────────────────────────────────────────────────
async function excelInsertFormula(p) {
    // p: { formula } — must start with =
    const formula = (p.formula || "").trim();
    if (!formula) { showToast("⚠️ Formula tidak diberikan"); return; }
    const f = formula.startsWith("=") ? formula : "=" + formula;

    await Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.formulas = [[f]];
        await ctx.sync();
        showToast(`✅ Formula disisipkan: ${f}`);
    }).catch(e => showToast("❌ Error formula: " + e.message));
}

// ── Conditional Format ────────────────────────────────────────────────────────
async function excelConditionalFormat(p) {
    // p: { type: "dataBar"|"colorScale"|"topBottom"|"cellValue", rules }
    await Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load("address");
        await ctx.sync();

        if (p.type === "dataBar") {
            const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
            cf.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight;
        } else if (p.type === "colorScale" || !p.type) {
            const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
            const cs = cf.colorScale;
            cs.criteria = [
                { type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "#FF0000" },
                { type: Excel.ConditionalFormatColorCriterionType.midpoint,    color: "#FFFF00", formula: "50" },
                { type: Excel.ConditionalFormatColorCriterionType.highestValue,color: "#00FF00" },
            ];
        } else if (p.type === "topBottom" && p.rules) {
            const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
            cf.topBottom.rule = { rank: p.rules.rank || 10, percent: p.rules.percent !== false, type: "TopItems" };
            cf.topBottom.format.fill.color = p.rules.color || "#FFD700";
        }
        await ctx.sync();
        showToast("✅ Conditional formatting diterapkan!");
    }).catch(e => showToast("❌ Error formatting: " + e.message));
}

// ── Sort Range ────────────────────────────────────────────────────────────────
async function excelSort(p) {
    // p: { column: 0-based index, ascending: true/false }
    await Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.sort.apply([{
            key: p.column || 0,
            ascending: p.ascending !== false,
        }]);
        await ctx.sync();
        showToast(`✅ Data diurutkan ${p.ascending !== false ? "A→Z" : "Z→A"}!`);
    }).catch(e => showToast("❌ Error sort: " + e.message));
}

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
    const b1 = ssXY / ssXX, b0 = yMean - b1 * xMean;
    const r = ssXY / Math.sqrt(ssXX * ssYY), r2 = r ** 2;
    const stats = { n, b0: b0.toFixed(4), b1: b1.toFixed(4), r: r.toFixed(4), r2: r2.toFixed(4), xMean: xMean.toFixed(3), yMean: yMean.toFixed(3) };

    const resultMatrix = [
        ["Statistik Regresi Linear", "Nilai"],
        ["N (jumlah data)", n],
        ["Intercept (b₀)", stats.b0], ["Slope/Koefisien (b₁)", stats.b1],
        ["Persamaan Regresi", `Y = ${stats.b0} + ${stats.b1}X`],
        ["Korelasi Pearson (r)", stats.r], ["Koefisien Determinasi (R²)", stats.r2],
        ["Rata-rata X", stats.xMean], ["Rata-rata Y", stats.yMean],
    ];

    await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getActiveWorksheet();
        const sel = ctx.workbook.getSelectedRange();
        sel.load("rowIndex,columnIndex");
        await ctx.sync();
        const startRow = sel.rowIndex + data.length + 2;
        const tgt = ws.getRangeByIndexes(startRow, sel.columnIndex, resultMatrix.length, 2);
        tgt.values = resultMatrix;
        tgt.format.autofitColumns();
        tgt.getRow(0).format.fill.color = "#1e3a5f";
        tgt.getRow(0).format.font.color = "white";
        tgt.getRow(0).format.font.bold = true;
        await ctx.sync();
        showToast("✅ Hasil regresi disisipkan!");
    }).catch(e => showToast("❌ Error: " + e.message));

    addBotMessage("⏳ Membuat narasi interpretasi...");
    try {
        const narasiPrompt = `Interpretasikan hasil regresi linear berikut dalam Bahasa Indonesia akademis:\n- N=${n}, b₀=${stats.b0}, b₁=${stats.b1}, r=${stats.r}, R²=${stats.r2}\n- Persamaan: Y = ${stats.b0} + ${stats.b1}X\nDua paragraf formal untuk Bab IV skripsi.`;
        const narasi = await callGeminiAPI(narasiPrompt);
        addBotMessage(`## Interpretasi Regresi\n\n${narasi}`);
        saveToHistory("Analisis Regresi", narasi);
    } catch {}
};

// ── INTERPRETASI STATISTIK ────────────────────────────────────────────────────
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

    const prompt = `Kamu adalah penulis akademik untuk Bab IV skripsi. Berdasarkan data statistik berikut, tulis interpretasi naratif dalam Bahasa Indonesia yang formal, mendalam, dan siap masuk ke laporan:\n\n${rawData}\n\nStruktur:\n1. Deskripsi hasil\n2. Analisis dan Pembahasan\n3. Kaitan dengan teori/hipotesis\n\nFormat paragraf, bukan list. 300-500 kata.`;
    addBotMessage("⏳ Menulis narasi interpretasi untuk Bab IV...");
    try {
        const result = await callGeminiAPI(prompt);
        addBotMessage(result);
        saveToHistory("Interpretasi Statistik", result);
    } catch (e) { addBotMessage("❌ Gagal: " + e.message); }
};

// ── TEMPLATE TABEL PENELITIAN ─────────────────────────────────────────────────
const EXCEL_TEMPLATES = {
    "frekuensi": [
        ["No","Kategori/Interval","Frekuensi (f)","Frekuensi Relatif (%)","Frekuensi Kumulatif"],
        ["1","","","",""],["2","","","",""],["3","","","",""],["4","","","",""],["5","","","",""],
        ["","TOTAL","=SUM(C2:C6)","=SUM(D2:D6)",""],
    ],
    "distribusi": [
        ["No","Kelas Interval","Batas Bawah","Batas Atas","Titik Tengah","Frekuensi","Persentase"],
        ["1","","","","","",""],["2","","","","","",""],["3","","","","","",""],
        ["4","","","","","",""],["5","","","","","",""],
        ["","Total","","","","=SUM(F2:F6)","=SUM(G2:G6)"],
    ],
    "crosstab": [
        ["Variabel X \\ Variabel Y","Kategori Y1","Kategori Y2","Kategori Y3","Total"],
        ["Kategori X1","","","","=SUM(B2:D2)"],["Kategori X2","","","","=SUM(B3:D3)"],
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
        const header = ws.getRangeByIndexes(sel.rowIndex, sel.columnIndex, 1, template[0].length);
        header.format.fill.color = "#1e3a5f";
        header.format.font.color = "white";
        header.format.font.bold = true;
        header.format.borders.getItem("EdgeBottom").style = "Continuous";
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

// ── VALIDASI DATA ─────────────────────────────────────────────────────────────
window.terapkanValidasiData = async function() {
    const input = document.getElementById("validation-input")?.value.trim();
    if (!input) { showToast("⚠️ Masukkan spesifikasi validasi di kotak!"); return; }

    const prompt = `Berikan instruksi langkah demi langkah cara menerapkan Data Validation di Excel untuk: "${input}". Sertakan formula Conditional Formatting yang sesuai. Format Markdown. Bahasa Indonesia.`;
    addBotMessage("⏳ Membuat saran validasi data...");
    try {
        const result = await callGeminiAPI(prompt);
        addBotMessage(result);
    } catch (e) { addBotMessage("❌ Gagal: " + e.message); }

    const items = input.includes(",") ? input.split(",").map(i => i.trim()) : null;
    if (items && items.length > 1 && items.length <= 20) {
        await Excel.run(async (ctx) => {
            const range = ctx.workbook.getSelectedRange();
            range.dataValidation.rule = { list: { inCellDropDown: true, source: items.join(",") } };
            await ctx.sync();
            showToast("✅ Dropdown diterapkan ke sel terpilih!");
        }).catch(() => {});
    }
};
