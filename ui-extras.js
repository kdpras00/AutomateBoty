/*
 * ui-extras.js — AutomateBoty v7
 * Features: Chat History, Prompt Presets, Rating, Onboarding Tour, Offline Cache
 */

// ── CONSTANTS ─────────────────────────────────────────────────────────────────
const MAX_HISTORY = 10;
const HISTORY_KEY  = "ab_history";
const PRESETS_KEY  = "ab_presets";
const RATINGS_KEY  = "ab_ratings";
const ONBOARD_KEY  = "ab_onboarded";

// ── CHAT HISTORY ──────────────────────────────────────────────────────────────
function saveToHistory(question, answer) {
    let history = getHistory();
    history.unshift({ q: question, a: answer, ts: Date.now() });
    if (history.length > MAX_HISTORY) history = history.slice(0, MAX_HISTORY);
    localStorage.setItem(HISTORY_KEY, JSON.stringify(history));
}

function getHistory() {
    try { return JSON.parse(localStorage.getItem(HISTORY_KEY) || "[]"); } catch { return []; }
}

function renderHistory() {
    const panel = document.getElementById("history-panel");
    if (!panel) return;
    const history = getHistory();
    if (!history.length) {
        panel.querySelector(".history-list").innerHTML = '<p class="empty-msg">Belum ada riwayat percakapan.</p>';
        return;
    }
    panel.querySelector(".history-list").innerHTML = history.map((item, i) => {
        const date = new Date(item.ts).toLocaleString("id-ID", { day:"2-digit", month:"short", hour:"2-digit", minute:"2-digit" });
        const preview = item.q.length > 60 ? item.q.substring(0,60)+"…" : item.q;
        return `<div class="history-item" onclick="loadHistoryItem(${i})">
            <div class="history-preview">${escapeHtml(preview)}</div>
            <div class="history-meta">${date}</div>
        </div>`;
    }).join("");
}

window.loadHistoryItem = function(index) {
    const history = getHistory();
    const item = history[index];
    if (!item) return;
    addUserMessage(item.q);
    addBotMessage(item.a);
    document.getElementById("history-panel").classList.add("hidden");
    document.getElementById("history-btn").classList.remove("active");
};

window.clearHistory = function() {
    localStorage.removeItem(HISTORY_KEY);
    renderHistory();
    showToast("🗑️ Riwayat dihapus");
};

// ── PROMPT PRESETS ────────────────────────────────────────────────────────────
function getPresets() {
    try { return JSON.parse(localStorage.getItem(PRESETS_KEY) || "[]"); } catch { return []; }
}

function savePreset(name, prompt) {
    const presets = getPresets();
    const exists = presets.findIndex(p => p.name === name);
    if (exists >= 0) presets[exists] = { name, prompt };
    else presets.push({ name, prompt });
    localStorage.setItem(PRESETS_KEY, JSON.stringify(presets));
    renderPresets();
    showToast(`⭐ Preset "${name}" disimpan`);
}

function renderPresets() {
    const list = document.getElementById("presets-list");
    if (!list) return;
    const presets = getPresets();
    if (!presets.length) {
        list.innerHTML = '<p class="empty-msg">Belum ada preset. Simpan prompt favorit Anda!</p>';
        return;
    }
    list.innerHTML = presets.map((p, i) => `
        <div class="preset-item">
            <button class="preset-use-btn" onclick="usePreset(${i})">${escapeHtml(p.name)}</button>
            <button class="preset-del-btn" onclick="deletePreset(${i})" title="Hapus">✕</button>
        </div>`).join("");
}

window.usePreset = function(index) {
    const p = getPresets()[index];
    if (!p) return;
    const userInput = document.getElementById("user-input");
    userInput.value = p.prompt;
    userInput.dispatchEvent(new Event('input', { bubbles: true }));
    userInput.focus();
    document.getElementById("settings-panel").classList.add("hidden");
    const settingsBtn = document.getElementById("nav-settings");
    if(settingsBtn) settingsBtn.classList.remove("active");
};

window.deletePreset = function(index) {
    const presets = getPresets();
    presets.splice(index, 1);
    localStorage.setItem(PRESETS_KEY, JSON.stringify(presets));
    renderPresets();
    showToast("🗑️ Preset dihapus");
};

window.saveCurrentAsPreset = function() {
    const input = document.getElementById("user-input").value.trim();
    const nameInput = document.getElementById("preset-name-input");
    const name = nameInput ? nameInput.value.trim() : "";
    if (!input) { showToast("⚠️ Ketik prompt dulu di kotak chat!"); return; }
    const finalName = name || input.substring(0, 30) + (input.length > 30 ? "…" : "");
    savePreset(finalName, input);
    if (nameInput) nameInput.value = "";
};

// ── RATING / FEEDBACK ─────────────────────────────────────────────────────────
function addRatingButtons(msgDiv, msgId) {
    const ratingBar = document.createElement("div");
    ratingBar.className = "rating-bar";
    ratingBar.innerHTML = `
        <button class="rating-btn" onclick="rateResponse('${msgId}', true, this)" title="Bagus">👍</button>
        <button class="rating-btn" onclick="rateResponse('${msgId}', false, this)" title="Kurang bagus">👎</button>
        <button class="rating-btn copy-btn" onclick="copyMessageText(this)" title="Salin">📋</button>`;
    msgDiv.appendChild(ratingBar);
}

window.rateResponse = function(id, isUp, btn) {
    const ratings = JSON.parse(localStorage.getItem(RATINGS_KEY) || "{}");
    ratings[id] = isUp ? "up" : "down";
    localStorage.setItem(RATINGS_KEY, JSON.stringify(ratings));
    btn.parentElement.querySelectorAll(".rating-btn").forEach(b => b.classList.remove("active-rate"));
    btn.classList.add("active-rate");
    showToast(isUp ? "👍 Terima kasih!" : "👎 Masukan dicatat");
};

window.copyMessageText = function(btn) {
    const content = btn.closest(".message").querySelector(".message-content");
    if (!content) return;
    navigator.clipboard.writeText(content.innerText).then(() => showToast("📋 Teks disalin!"));
};

// ── OFFLINE CACHE ─────────────────────────────────────────────────────────────
function cacheOffline(key, value) {
    try { localStorage.setItem("ab_cache_" + key, JSON.stringify({ v: value, ts: Date.now() })); } catch {}
}

function getOfflineCache(key) {
    try {
        const raw = localStorage.getItem("ab_cache_" + key);
        if (!raw) return null;
        const parsed = JSON.parse(raw);
        // Expire after 7 days
        if (Date.now() - parsed.ts > 7 * 86400 * 1000) { localStorage.removeItem("ab_cache_" + key); return null; }
        return parsed.v;
    } catch { return null; }
}

// ── ONBOARDING TOUR ───────────────────────────────────────────────────────────
const TOUR_STEPS = [
    { icon: "✨", title: "Selamat datang di AutomateBoty v7!", desc: "Asisten AI canggih untuk Word, Excel, dan PowerPoint. Mari kenali fitur-fitur utamanya!" },
    { icon: "📋", title: "Template Manager", desc: "Klik ikon 📋 di header untuk memilih template jurnal, skripsi, artikel, atau prosiding. Template langsung diterapkan ke dokumen Word Anda." },
    { icon: "⭐", title: "Prompt Presets", desc: "Simpan prompt favorit Anda! Klik ikon ⭐ di header, lalu simpan prompt yang sering digunakan agar tidak perlu mengetik ulang." },
    { icon: "💬", title: "Riwayat Percakapan", desc: "Semua percakapan tersimpan otomatis. Klik ikon 💬 untuk membuka riwayat dan melanjutkan percakapan sebelumnya." },
    { icon: "🚀", title: "Quick Actions", desc: "Gunakan tombol-tombol di bawah untuk aksi cepat! Di Word: buat jurnal, parafrase, proofreading. Di Excel: analisis statistik, regresi. Di PPT: buat slide dari file!" },
];
let tourStep = 0;

function startOnboarding() {
    tourStep = 0;
    const modal = document.getElementById("onboarding-modal");
    if (modal) { modal.classList.remove("hidden"); updateTourStep(); }
}

function updateTourStep() {
    const step = TOUR_STEPS[tourStep];
    if (!step) return;
    const modal = document.getElementById("onboarding-modal");
    modal.querySelector(".tour-icon").textContent = step.icon;
    modal.querySelector(".tour-title").textContent = step.title;
    modal.querySelector(".tour-desc").textContent = step.desc;
    modal.querySelector(".tour-progress").textContent = `${tourStep + 1} / ${TOUR_STEPS.length}`;
    modal.querySelector(".tour-prev").disabled = tourStep === 0;
    modal.querySelector(".tour-next").textContent = tourStep === TOUR_STEPS.length - 1 ? "Mulai!" : "Lanjut →";
}

window.tourNext = function() {
    if (tourStep < TOUR_STEPS.length - 1) { tourStep++; updateTourStep(); }
    else { closeTour(); }
};
window.tourPrev = function() { if (tourStep > 0) { tourStep--; updateTourStep(); } };
window.closeTour = function() {
    document.getElementById("onboarding-modal").classList.add("hidden");
    localStorage.setItem(ONBOARD_KEY, "1");
};

// ── PANEL TOGGLE SETUP ────────────────────────────────────────────────────────
function setupExtraPanels() {
    // Help / Re-tour (fallback if exists)
    const helpBtn = document.getElementById("help-btn");
    if (helpBtn) helpBtn.addEventListener("click", startOnboarding);

    // Onboarding first time
    if (!localStorage.getItem(ONBOARD_KEY)) setTimeout(startOnboarding, 800);
}

// ── EXPOSE GLOBALS ────────────────────────────────────────────────────────────
window.saveToHistory    = saveToHistory;
window.getOfflineCache  = getOfflineCache;
window.cacheOffline     = cacheOffline;
window.addRatingButtons = addRatingButtons;
window.setupExtraPanels = setupExtraPanels;
