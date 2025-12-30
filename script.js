/*
 * Gemini AI Office Add-in
 * Logic for handling chat and interacting with Office apps.
 */

// Core Configuration
const DEFAULT_API_KEY = "AIzaSyDM4fjZTT1F6Ux2D8anrQMe7SSShKVUgrQ"; // User provided
let apiKey = localStorage.getItem("gemini_api_key") || DEFAULT_API_KEY;

// DOM Elements
const chatContainer = document.getElementById("chat-container");
const userInput = document.getElementById("user-input");
const sendBtn = document.getElementById("send-btn");
const settingsBtn = document.getElementById("settings-btn");
const settingsPanel = document.getElementById("settings-panel");
const apiKeyInput = document.getElementById("api-key-input");
const saveSettingsBtn = document.getElementById("save-settings");
let quickActionsContainer = null; // Container for buttons

// Initialize Office.js
Office.onReady((info) => {
    console.log("Office.js ready. Host:", info.host);
    
    // Create UI for Quick Actions
    setupQuickActions(info.host);

    // Always setup listeners so it works in Browser too for testing
    setupEventListeners();
    
    // Check for saved API key
    // Check for saved API key
    const currentSavedKey = localStorage.getItem("gemini_api_key");
    // List of invalid/blocked keys to force clear
    const INVALID_KEYS = [
        "AIzaSyCmSlRCCPgC1ph4vuco9hwLsTaDtnBPcSA",
        "AIzaSyAmsulrYYqrxuWnlqwrn1UzHsPdTSedyR0" // Blocked key
    ];

    if (currentSavedKey && !INVALID_KEYS.includes(currentSavedKey)) {
        apiKeyInput.value = currentSavedKey;
        apiKey = currentSavedKey;
    } else {
        // If no key, or we found an invalid key, reset to new default
        if (INVALID_KEYS.includes(currentSavedKey)) {
            console.log("Removing invalid API key from storage");
            localStorage.removeItem("gemini_api_key");
        }
        apiKeyInput.value = DEFAULT_API_KEY;
        apiKey = DEFAULT_API_KEY;
    }

    // Network Status Check
    updateNetworkStatus();
    window.addEventListener('online', updateNetworkStatus);
    window.addEventListener('offline', updateNetworkStatus);
    
    // Add version indicator
    const versionDiv = document.createElement("div");
    versionDiv.style.fontSize = "10px";
    versionDiv.style.color = "#888";
    versionDiv.style.textAlign = "center";
    versionDiv.style.marginTop = "5px";
    versionDiv.innerText = "v5.0 - Gemini Flash Latest";
    document.querySelector(".app-container").appendChild(versionDiv);
    console.log("Gemini Add-in v5 loaded");
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

async function handleSendMessage(autoSend = false) {
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
        
        // 4. AUTO-INSERT (The "Powerful" Feature)
        // We only auto-insert if it successfully generated and isn't an error message
        if (responseText && !responseText.startsWith("❌") && !responseText.includes("Api Error")) {
             insertIntoDocument(responseText); 
             // We also show a small toast or log?
             console.log("Auto-inserted content.");
        }

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

    // Use v1beta and gemini-flash-latest (Aliases to best available flash model)
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=${apiKey}`;
    
    // Dynamic System Context based on Host
    let systemRole = "You are a helpful assistant.";
    const host = Office.context.host;
    
    if (host === Office.HostType.Word) {
        systemRole = "You are an expert academic writer and editor. Generate high-quality, professional, and detailed content suitable for Thesis (Skripsi), Journals, and formal documents. Use strictly formatted Markdown (headers, lists, bold). Do NOT chat. DIRECTLY output the document content requested.";
    } else if (host === Office.HostType.Excel) {
        systemRole = "You are an Excel Expert. If asked for a formula, output ONLY the formula starting with =. If asked for data, output a CSV or Markdown table. Do not include explanations unless asked.";
    } else if (host === Office.HostType.PowerPoint) {
        systemRole = "You are a Presentation Expert. Output content in bullet points suitable for slides. Keep it concise and impactful.";
    }

    const payload = {
        contents: [{
            role: "user",
            parts: [{ text: systemRole + "\n\nUser Request: " + prompt }]
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

function setupQuickActions(host) {
    // Insert Quick Actions div before the chat container
    quickActionsContainer = document.createElement("div");
    quickActionsContainer.className = "quick-actions";
    
    // Define Actions based on Host
    let actions = [];
    
    // default actions if browser/unknown
    actions = [
        { label: "Summarize", prompt: "Summarize this document concisely." },
        { label: "Fix Grammar", prompt: "Check grammar and style." }
    ];

    if (host === Office.HostType.Word) {
        actions = [
            { label: "Draft Thesis", prompt: "Write a detailed outline for a Thesis (Skripsi) on {topic}. Structure it with Chapters 1-5." },
            { label: "Create Journal", prompt: "Write an academic journal abstract and introduction about {topic}." },
            { label: "Fix Grammar", prompt: "Rewrite the selected text to be more academic and grammatically correct." },
            { label: "Expand", prompt: "Expand heavily on this topic with detailed explanations and examples." }
        ];
    } else if (host === Office.HostType.Excel) {
        actions = [
            { label: "Gen Formula", prompt: "Write an Excel formula to: " },
            { label: "Create Table", prompt: "Create a dummy dataset table for: " },
            { label: "Analyze Data", prompt: "Analyze this data and give insights: " },
            { label: "Format", prompt: "Suggest conditional formatting rules for: " }
        ];
    } else if (host === Office.HostType.PowerPoint) {
        actions = [
            { label: "New Slide", prompt: "Create a slide content about: " },
            { label: "Outline", prompt: "Create a 10-slide presentation outline for: " },
            { label: "Speaker Notes", prompt: "Write speaker notes for a slide about: " }
        ];
    }

    // Render Buttons
    actions.forEach(action => {
        const btn = document.createElement("button");
        btn.className = "action-pill";
        btn.innerText = action.label;
        btn.onclick = () => {
             // If prompt needs input (ends with :) or placeholder {topic}, put in input box
             if (action.prompt.includes("{topic}") || action.prompt.endsWith(": ")) {
                 const currentInput = userInput.value;
                 let newText = action.prompt;
                 if (currentInput) {
                     // If user typed something, replace {topic} or append
                     if(newText.includes("{topic}")) newText = newText.replace("{topic}", currentInput);
                     else newText = newText + currentInput;
                     
                     userInput.value = newText;
                     handleSendMessage(true); // Auto send if we combined it
                 } else {
                     // Just prep the prompt for them to fill in
                     userInput.value = action.prompt.replace("{topic}", "[TOPIC]");
                     userInput.focus();
                 }
             } else {
                 // Direct action (e.g. Fix Grammar of selection)
                 userInput.value = action.prompt;
                 handleSendMessage(true);
             }
        };
        quickActionsContainer.appendChild(btn);
    });

    const appContainer = document.querySelector(".app-container");
    const chatContainer = document.getElementById("chat-container");
    appContainer.insertBefore(quickActionsContainer, chatContainer.nextSibling); // Insert above input area? No, above chat input? 
    // Let's put it ABOVE the input area (footer)
    const inputArea = document.querySelector(".input-area");
    appContainer.insertBefore(quickActionsContainer, inputArea);
}

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
        // Convert Markdown to HTML for Word
        const htmlContent = marked.parse(text);
        
        Office.context.document.setSelectedDataAsync(htmlContent, { coercionType: Office.CoercionType.Html }, (asyncResult) => {
             if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                 console.error("Action failed: " + asyncResult.error.message);
                 // Fallback to text if HTML fails
                 Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text });
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
