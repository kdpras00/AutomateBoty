/*
 * executor.js — AutomateBoty v8
 * Central AI-to-Office command execution dispatcher.
 * Parses Action JSON from AI responses and routes to host-specific executors.
 *
 * Action JSON format embedded in AI responses:
 * ```action
 * {"action":"WORD_FORMAT_PAGE","params":{...}}
 * ```
 */

// ── ACTION JSON PARSER ────────────────────────────────────────────────────────
/**
 * Detects and extracts Action JSON block from raw AI response text.
 * Returns { actionData, cleanText } where cleanText has the action block removed.
 */
function parseAIResponse(rawText) {
    // Match ```action ... ``` block
    const actionMatch = rawText.match(/```action\s*([\s\S]*?)```/i);
    if (actionMatch) {
        try {
            const actionData = JSON.parse(actionMatch[1].trim());
            const cleanText = rawText.replace(actionMatch[0], "").trim();
            return { actionData, cleanText };
        } catch {}
    }

    // Also try inline JSON with "action" key (fallback)
    const inlineMatch = rawText.match(/\{"action"\s*:\s*"[A-Z_]+"/);
    if (inlineMatch) {
        try {
            // Find the full JSON object
            const startIdx = rawText.indexOf(inlineMatch[0]);
            let depth = 0, endIdx = startIdx;
            for (let i = startIdx; i < rawText.length; i++) {
                if (rawText[i] === '{') depth++;
                else if (rawText[i] === '}') { depth--; if (depth === 0) { endIdx = i; break; } }
            }
            const jsonStr = rawText.substring(startIdx, endIdx + 1);
            const actionData = JSON.parse(jsonStr);
            const cleanText = (rawText.substring(0, startIdx) + rawText.substring(endIdx + 1)).trim();
            return { actionData, cleanText };
        } catch {}
    }

    return { actionData: null, cleanText: rawText };
}

// ── DISPATCHER ───────────────────────────────────────────────────────────────
/**
 * Main dispatcher: routes action to the right host executor.
 */
async function dispatchAction(actionData) {
    if (!actionData || !actionData.action) return false;

    const host = Office.context.host;
    const action = actionData.action;
    const params = actionData.params || {};

    showActionToast(action);

    try {
        // Raw SCRIPT EXECUTION (Ultimate Fallback/Omni Tool)
        if (action === "SCRIPT_EXECUTION") {
            return await executeRawScript(params.script, host);
        }

        // WORD actions
        if (action.startsWith("WORD_") && host === Office.HostType.Word) {
            if (typeof executeWordAction === "function") {
                await executeWordAction(action, params);
                return true;
            }
        }
        // EXCEL actions
        else if (action.startsWith("EXCEL_") && host === Office.HostType.Excel) {
            if (typeof executeExcelAction === "function") {
                await executeExcelAction(action, params);
                return true;
            }
        }
        // PPT actions
        else if (action.startsWith("PPT_") && host === Office.HostType.PowerPoint) {
            if (typeof executePPTAction === "function") {
                await executePPTAction(action, params);
                return true;
            }
        }
        // Unknown or wrong host
        else {
            console.warn(`Action "${action}" tidak didukung di host ini.`);
            if (typeof addBotMessage === "function") {
                addBotMessage(`⚠️ Maaf, perintah format/aksi **${action}** belum didukung atau tidak sesuai untuk aplikasi Office ini.`);
            }
            return false;
        }
    } catch (e) {
        console.error("dispatchAction error:", e);
        if (e.message === "UNSUPPORTED_ACTION") {
            if (typeof addBotMessage === "function") addBotMessage(`⚠️ Maaf, perintah format **${action}** belum didukung saat ini.`);
        } else {
            if (typeof showToast === "function") showToast(`❌ Gagal mengeksekusi aksi: ${e.message || action}`);
            if (typeof addBotMessage === "function") addBotMessage(`❌ Gagal menjalankan perintah format ke dokumen:\n\`${e.message}\``);
        }
    }
    return false;
}

function showActionToast(action) {
    const labels = {
        WORD_FORMAT_PAGE: "📄 Mengatur layout halaman...",
        WORD_FORMAT_FONT: "🔤 Menerapkan font...",
        WORD_HEADING_STYLES: "📝 Menerapkan heading styles...",
        WORD_INSERT_TABLE: "📊 Membuat tabel...",
        WORD_ADD_CONTENT: "✍️ Menyisipkan konten...",
        WORD_SET_SPACING: "↕️ Mengatur spasi...",
        EXCEL_INSERT_DATA: "📋 Menyisipkan data...",
        EXCEL_CREATE_CHART: "📈 Membuat grafik...",
        EXCEL_FORMAT_CELLS: "🎨 Memformat sel...",
        EXCEL_INSERT_FORMULA: "🧮 Menyisipkan formula...",
        EXCEL_CONDITIONAL_FORMAT: "🎨 Conditional formatting...",
        PPT_CREATE_SLIDES: "🎞️ Membuat slides...",
        PPT_APPLY_THEME: "🎨 Menerapkan tema...",
        PPT_ADD_SLIDE: "➕ Menambah slide...",
        SCRIPT_EXECUTION: "⚡ Mengeksekusi Office.js script...",
    };
    if (typeof showToast === "function") {
        showToast(labels[action] || `⚙️ Menjalankan: ${action}`);
    }
}

// ── RAW SCRIPT EXECUTOR (OMNI) ────────────────────────────────────────────────
/**
 * Executes a raw Office.js script string directly via `eval` using a new Function.
 * This grants the AI 100% unrestricted access to the Office.js API.
 */
async function executeRawScript(scriptString, host) {
    if (!scriptString) return false;
    
    try {
        if (host === Office.HostType.Word) {
            await Word.run(async (context) => {
                // Wrap in async IIFE and execute
                const runScript = new Function("context", `return (async () => { ${scriptString} })();`);
                await runScript(context);
                await context.sync();
            });
        } else if (host === Office.HostType.Excel) {
            await Excel.run(async (context) => {
                const runScript = new Function("context", `return (async () => { ${scriptString} })();`);
                await runScript(context);
                await context.sync();
            });
        } else if (host === Office.HostType.PowerPoint) {
            await PowerPoint.run(async (context) => {
                const runScript = new Function("context", `return (async () => { ${scriptString} })();`);
                await runScript(context);
                await context.sync();
            });
        } else {
            throw new Error("Unsupported host for script execution.");
        }
        showToast("✅ Script dieksekusi dengan sukses!");
        return true;
    } catch (e) {
        console.error("executeRawScript Error:", e);
        throw new Error("Script Error: " + (e.message || String(e)));
    }
}

// ── EXPOSE GLOBALS ────────────────────────────────────────────────────────────
window.parseAIResponse = parseAIResponse;
window.dispatchAction  = dispatchAction;
