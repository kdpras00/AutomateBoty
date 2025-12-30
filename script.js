/*
 * Gemini AI Office Add-in
 * Logic for handling chat and interacting with Office apps.
 */

// Core Configuration
const DEFAULT_API_KEY = "AIzaSyCmSlRCCPgC1ph4vuco9hwLsTaDtnBPcSA"; // User provided
let apiKey = localStorage.getItem("gemini_api_key") || DEFAULT_API_KEY;

// DOM Elements
const chatContainer = document.getElementById("chat-container");
const userInput = document.getElementById("user-input");
const sendBtn = document.getElementById("send-btn");
const settingsBtn = document.getElementById("settings-btn");
const settingsPanel = document.getElementById("settings-panel");
const apiKeyInput = document.getElementById("api-key-input");
const saveSettingsBtn = document.getElementById("save-settings");

// Initialize Office.js
Office.onReady((info) => {
    console.log("Office.js ready. Host:", info.host);
    
    // Always setup listeners so it works in Browser too for testing
    setupEventListeners();
    
    // Check for saved API key
    if (localStorage.getItem("gemini_api_key")) {
        apiKeyInput.value = localStorage.getItem("gemini_api_key");
    } else {
        apiKeyInput.value = DEFAULT_API_KEY;
    }

    // Network Status Check
    updateNetworkStatus();
    window.addEventListener('online', updateNetworkStatus);
    window.addEventListener('offline', updateNetworkStatus);
});

function updateNetworkStatus() {
    if (!navigator.onLine) {
        addBotMessage("⚠️ **No Internet Connection**\nI need an internet connection to talk to Gemini AI. Please check your network.");
        sendBtn.disabled = true;
    } else {
        sendBtn.disabled = false;
    }
}

function setupEventListeners() {
    // Send Message Logic
    sendBtn.addEventListener("click", handleSendMessage);
    userInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault();
            handleSendMessage();
        }
    });

    // Auto-resize textarea
    userInput.addEventListener("input", function() {
        this.style.height = "auto";
        this.style.height = (this.scrollHeight) + "px";
        if (this.value === "") this.style.height = "auto";
    });

    // Settings Toggle
    settingsBtn.addEventListener("click", () => {
        settingsPanel.classList.toggle("hidden");
    });

    // Save Settings
    saveSettingsBtn.addEventListener("click", () => {
        const newKey = apiKeyInput.value.trim();
        if (newKey) {
            apiKey = newKey;
            localStorage.setItem("gemini_api_key", apiKey);
            settingsPanel.classList.add("hidden");
            addBotMessage("API Key updated successfully.");
        }
    });
}

async function handleSendMessage() {
    const text = userInput.value.trim();
    if (!text) return;

    // 1. Display User Message
    addUserMessage(text);
    userInput.value = "";
    userInput.style.height = "auto";
    sendBtn.disabled = true;

    // 2. Call Gemini API
    try {
        const loadingId = addLoadingMessage();
        console.log("Calling Gemini API...");
        const responseText = await callGeminiAPI(text);
        
        // 3. Remove Loading & Display Bot Response
        removeMessage(loadingId);
        addBotMessage(responseText, true); // true = enable actions
    } catch (error) {
        console.error("Full Error Object:", error);
        removeMessage("loading-msg"); // Ensure loading is removed
        // Show the actual error message to the user for debugging
        addBotMessage(`❌ **Error**: ${error.message}\n\nPlease check your internet connection or API Key.`);
    } finally {
        sendBtn.disabled = false;
        userInput.focus();
    }
}

async function callGeminiAPI(prompt) {
    if (!apiKey) {
        throw new Error("API Key is missing. Please check settings.");
    }
    
    // Check if we are online before fetching
    if (!navigator.onLine) {
        throw new Error("No Internet connection.");
    }

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
    
    const systemContext = `You are a helpful assistant living inside Microsoft ${Office.context.host || 'Office'}. 
    Keep answers concise and relevant to document creation. 
    If the user asks to generate text, table, or content, provide it clearly.
    Format usage: Markdown.`;

    const payload = {
        contents: [{
            role: "user",
            parts: [{ text: systemContext + "\n\nUser Question: " + prompt }]
        }]
    };

    console.log("Sending fetch request to Gemini...");
    const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
    });

    console.log("Response status:", response.status);

    if (!response.ok) {
        const errText = await response.text();
        console.error("API Error Response:", errText);
        throw new Error(`API Error (${response.status}): ${response.statusText}`);
    }

    const data = await response.json();
    console.log("API Data received:", data);
    
    if (!data.candidates || data.candidates.length === 0) {
        return "I received an empty response from Gemini.";
    }

    return data.candidates?.[0]?.content?.parts?.[0]?.text || "No response generated.";
}

// UI Helpers
function addUserMessage(text) {
    const div = document.createElement("div");
    div.className = "message user-message";
    div.innerHTML = `<div class="message-content">${escapeHtml(text).replace(/\n/g, "<br>")}</div>`;
    chatContainer.appendChild(div);
    scrollToBottom();
}

function addLoadingMessage() {
    const id = "loading-" + Date.now();
    const div = document.createElement("div");
    div.id = id;
    div.className = "message bot-message";
    div.innerHTML = `<div class="message-content">Thinking...</div>`;
    chatContainer.appendChild(div);
    scrollToBottom();
    return id;
}

function removeMessage(id) {
    const el = document.getElementById(id);
    if (el) el.remove();
}

function addBotMessage(text, withActions = false) {
    const div = document.createElement("div");
    div.className = "message bot-message";
    
    // Parse Markdown
    const htmlContent = marked.parse(text);
    
    let actionButtons = "";
    if (withActions) {
        actionButtons = `
            <div class="message-actions" style="margin-top: 8px; display: flex; gap: 8px;">
                <button onclick="insertTextFromMessage(this)" class="action-btn" title="Insert response into document">Insert</button>
                <button onclick="copyToClipboard(this)" class="action-btn" title="Copy to clipboard">Copy</button>
            </div>
        `;
    }

    div.innerHTML = `
        <div class="message-content">${htmlContent}</div>
        ${actionButtons}
    `;
    
    // Highlight code blocks
    div.querySelectorAll('pre code').forEach((block) => {
        hljs.highlightElement(block);
    });

    chatContainer.appendChild(div);
    scrollToBottom();
}

function scrollToBottom() {
    chatContainer.scrollTop = chatContainer.scrollHeight;
}

function escapeHtml(text) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return text.replace(/[&<>"']/g, (m) => map[m]);
}

// Office Interaction Logic
window.insertTextFromMessage = function(btn) {
    // Traverse up to find the message content
    const messageContent = btn.parentElement.previousElementSibling.innerText;
    insertIntoDocument(messageContent);
};

window.copyToClipboard = function(btn) {
    const messageContent = btn.parentElement.previousElementSibling.innerText;
    navigator.clipboard.writeText(messageContent).then(() => {
        const originalText = btn.innerText;
        btn.innerText = "Copied!";
        setTimeout(() => btn.innerText = originalText, 2000);
    });
};

function insertIntoDocument(text) {
    // Generic insertion using setSelectedDataAsync
    // This works for Word (text), Excel (cell), PPT (width depends on selection)
    
    // Check if it looks like a table (markdown table)
    const isTable = text.includes("|") && text.includes("---");
    
    if (Office.context.host === Office.HostType.Excel) {
        // Excel specific logic could go here (e.g., parsing CSV/Table)
        // For now, simple text insertion or matrix insertion would be ideal
        // But matrix requires [[],[]] format for coercion.Table
        Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Action failed: " + asyncResult.error.message);
            }
        });
    } else if (Office.context.host === Office.HostType.Word) {
        // Word can handle HTML or Text. Markdown parsed to HTML is good.
        // Let's try inserting as HTML if we have it, else text.
        // We only have innerText passed here. Let's re-generate HTML ? 
        // No, let's just insert Ooxml or Text.
        // For simplicity in V1: Insert Text.
        Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (asyncResult) => {
             if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                 // Try inserting as HTML if Text fails or just log
                 console.error("Action failed: " + asyncResult.error.message);
             }
        });
    } else {
        // PowerPoint / Other
        Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Action failed: " + asyncResult.error.message);
            }
        });
    }
}
