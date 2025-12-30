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

    // File Upload Handlers (New)
    const uploadBtn = document.getElementById("upload-btn");
    const fileInput = document.getElementById("file-upload");

    if (uploadBtn && fileInput) {
        uploadBtn.addEventListener("click", () => fileInput.click());
        
        fileInput.addEventListener("change", (e) => {
            const file = e.target.files[0];
            if (file) {
                 // Show pill or indicator
                 let previewText = `📎 ${file.name} selected`;
                 // Just append to input as a visual cue? 
                 // Or better, store it in a global variable
                 window.currentFile = file;
                 
                 // Add visual indicator near input
                 const existingIndicator = document.getElementById("file-indicator");
                 if(existingIndicator) existingIndicator.remove();
                 
                 const indicator = document.createElement("div");
                 indicator.id = "file-indicator";
                 indicator.style.fontSize = "11px";
                 indicator.style.padding = "4px 8px";
                 indicator.style.backgroundColor = "#e0e7ff";
                 indicator.style.color = "#333";
                 indicator.style.borderRadius = "12px";
                 indicator.style.marginBottom = "5px";
                 indicator.style.display = "inline-block";
                 indicator.innerText = previewText;
                 
                 const inputArea = document.querySelector(".input-area");
                 inputArea.insertBefore(indicator, userInput);
                 
                 // If it's an image, read strictly for base64 now
                 // Implementation in handleSendMessage will read it.
            }
        });
    }
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
        // Clear file upload state
        window.currentFile = null;
        const indicator = document.getElementById("file-indicator");
        if(indicator) indicator.remove();
        if(document.getElementById("file-upload")) document.getElementById("file-upload").value = "";
        
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
        systemRole = "You are an Excel Expert. If asked for a formula, output ONLY the formula starting with =. If asked for data, output a clean CSV format (comma separated) or a Markdown Table. Do NOT include conversational filler.";
    } else if (host === Office.HostType.PowerPoint) {
        systemRole = "You are a Presentation Expert. If asked for multiple slides, output a JSON Array: [{\"title\": \"Title\", \"points\": [\"Bullet1\", \"Bullet2\"]}]. If asked for one slide, output regular text: TITLE: [Title]\n- Points. Do NOT use Markdown.";
    }

    // 0. Get Document Context (Read text from file)
    let docContext = "";
    try {
        docContext = await getDocumentContext();
    } catch (e) {
        console.log("Could not read document context:", e);
    }
    
    // 0b. Process User Uploaded File (if any)
    let filePart = null;
    let fileTextContent = "";
    
    if (window.currentFile) {
        const file = window.currentFile;
        console.log("Processing uploaded file:", file.name, file.type);
        
        if (file.type.startsWith("image/")) {
            // Convert to Base64 for Inline Data
            const base64Data = await new Promise((resolve) => {
                const reader = new FileReader();
                reader.onloadend = () => {
                    // result is data:image/jpeg;base64,....
                    const content = reader.result.toString().split(',')[1];
                    resolve(content);
                };
                reader.readAsDataURL(file);
            });
            
            filePart = {
                inlineData: {
                    mimeType: file.type,
                    data: base64Data
                }
            };
        } else {
            // Assume text-based (txt, csv, md, json, js, etc)
            fileTextContent = await new Promise((resolve) => {
                 const reader = new FileReader();
                 reader.onload = (e) => resolve(e.target.result);
                 reader.readAsText(file);
            });
            fileTextContent = `\n\n[Attached File: ${file.name}]\n${fileTextContent}\n[End Attached File]\n`;
        }
        
        // Clear file after processing? or keep until sent?
        // Let's clear visual indicator and var after successful send (in handleSendMessage), 
        // but for now we just prepare the payload.
    }

    // Construct Parts
    const textPart = { text: systemRole + "\n\nDocument Context:\n" + docContext + fileTextContent + "\n\nUser Request: " + prompt };
    const partsVal = filePart ? [textPart, filePart] : [textPart];

    const payload = {
        contents: [{
            role: "user",
            parts: partsVal
        }]
    };

    console.log("Sending fetch request to Gemini with payload parts:", partsVal.length);
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
    
    // Clear file selection after successful call setup (actually better to clear in handleSendMessage)
    
    if (!data.candidates || data.candidates.length === 0) {
        return "I received an empty response from Gemini.";
    }

    return data.candidates?.[0]?.content?.parts?.[0]?.text || "No response generated.";
}

async function getDocumentContext() {
    // Attempt to read selection or body text
    return new Promise((resolve) => {
        if (Office.context.host === Office.HostType.Word) {
            Word.run(async (context) => {
                const selection = context.document.getSelection();
                selection.load("text");
                await context.sync();
                if (selection.text && selection.text.trim().length > 0) {
                     resolve(selection.text);
                } else {
                     // If selection empty, read body (first 2000 chars to avoid overload)
                     const body = context.document.body;
                     body.load("text");
                     await context.sync();
                     resolve(body.text.substring(0, 5000)); // Read up to 5k chars
                }
            }).catch(() => resolve(""));
        } else if (Office.context.host === Office.HostType.Excel) {
             Excel.run(async (context) => {
                 const range = context.workbook.getSelectedRange();
                 range.load("text");
                 await context.sync();
                 // range.text is a 2D array
                 const textStr = range.text.map(row => row.join(", ")).join("\n");
                 resolve(textStr);
             }).catch(() => resolve(""));
        } else {
            // PowerPoint or others: Use generic getSelectedData
             Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
                 if (result.status === Office.AsyncResultStatus.Succeeded) {
                     resolve(result.value);
                 } else {
                     resolve("");
                 }
             });
        }
    });
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
    
    div.innerHTML = `
        <div class="message-content">${htmlContent}</div>
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
            { label: "Buat Skripsi", prompt: "Buatkan kerangka Bab 1 Skripsi tentang {topic}. Struktur lengkap dengan 5 Bab." },
            { label: "Buat Jurnal", prompt: "Buatkan abstrak dan pendahuluan jurnal akademik tentang {topic}." },
            { label: "Ke B.Inggris", prompt: "Translate teks yang dipilih ke Bahasa Inggris akademik." },
            { label: "Ke B.Indo", prompt: "Translate teks yang dipilih ke Bahasa Indonesia baku." },
            { label: "Tulis Ulang", prompt: "Tulis ulang teks ini agar lebih profesional, ringkas, dan berimpact." },
            { label: "Cek Grammar", prompt: "Perbaiki tata bahasa dan ejaan teks yang dipilih." }
        ];
    } else if (host === Office.HostType.Excel) {
        actions = [
            { label: "Bikin Rumus", prompt: "Buatkan rumus Excel untuk: " },
            { label: "Bikin Tabel", prompt: "Buatkan tabel dataset dummy untuk: " },
            { label: "Analisa Data", prompt: "Analisa data yang dipilih dan berikan insight: " },
            { label: "Format", prompt: "Berikan aturan Conditional Formatting untuk: " },
            { label: "Buat Grafik", prompt: "Buatkan grafik dari data yang dipilih." }
        ];
    } else if (host === Office.HostType.PowerPoint) {
        actions = [
            { label: "Slide Baru", prompt: "Buatkan konten slide tentang: " },
            { label: "Outline PPT", prompt: "Buatkan outline presentasi 10 slide tentang: " },
            { label: "Catatan Pembicara", prompt: "Buatkan catatan pembicara untuk slide tentang: " },
            { label: "5 Slide Langsung", prompt: "Buatkan 5 slide lengkap tentang: " }
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
    // Generic insertion logic
    
    if (Office.context.host === Office.HostType.Word) {
        // Word: Convert Markdown to HTML
        const htmlContent = marked.parse(text);
        Office.context.document.setSelectedDataAsync(htmlContent, { coercionType: Office.CoercionType.Html }, (asyncResult) => {
             if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                 console.error("Action failed: " + asyncResult.error.message);
                 // Fallback to text if HTML fails
                 Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text });
             }
        });

    } else if (Office.context.host === Office.HostType.Excel) {
        // Excel: Check for Chart/Grafik keyword
        if (text.toLowerCase().includes("chart") || text.toLowerCase().includes("grafik")) {
             runExcelChartGen(text);
        } else {
             // Else data/table or formula
             runExcelDataGen(text);
        }

    } else {
        // PowerPoint: Try to create a Smart Slide
        runPowerPointSlideGen(text);
    }
}

async function runExcelDataGen(text) {
    // 1. Check if it's a Formula (starts with =)
    const trimmed = text.trim();
    if (trimmed.startsWith("=")) {
        // Insert formula into selected cell
        Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.formulas = [[trimmed]];
            await context.sync();
            console.log("Formula inserted.");
        }).catch(err => console.error(err));
        return;
    }

    // 2. CSV / Data Table parser
    // Heuristic: Does it look like CSV? (Comma or Tab separated, multiple lines)
    // Or just a list.
    // We will parse standard CSV format (e.g. "Name,Age\nJohn,25")
    
    // Simple CSV parser (handles commas)
    let rows = text.split('\n').filter(r => r.trim() !== '');
    
    // Remove "Sure" lines
     if (rows.length > 0 && (rows[0].toLowerCase().startsWith("sure") || rows[0].toLowerCase().startsWith("here"))) {
        rows.shift();
    }
    
    if (rows.length === 0) return;

    // Detect delimiter (Comma or Pipe or Tab)
    // We'll assume the prompt asked for CSV or Table.
    // Let's try to detect | (Markdown table) vs , (CSV)
    let delimiter = ",";
    if (rows[0].includes("|")) {
        delimiter = "|";
        // Filter out separator lines like "|---|---|"
        rows = rows.filter(r => !r.includes("---"));
    }

    const matrix = rows.map(row => {
        // Split by delimiter
        let cols = row.split(delimiter);
        // Clean whitespace
        return cols.map(c => c.trim().replace(/^\||\|$/g, '')); // Remove leading/trailing pipes if markdown
    });

    if (matrix.length === 0) return;

    // 3. Insert Matrix into Excel
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const activeRange = context.workbook.getSelectedRange();
        // Determine target range based on matrix size
        const numRows = matrix.length;
        const numCols = matrix[0].length;
        
        // We can't easily "resize" from activeRange without loading props, 
        // but we can use getResizedRange relative to the top-left of selection.
        const targetRange = activeRange.getResizedRange(numRows - 1, numCols - 1);
        
        console.log(`Writing ${numRows}x${numCols} matrix to Excel.`);
        targetRange.values = matrix;
        targetRange.format.autofitColumns();
        
        await context.sync();
    }).catch(error => {
        console.error("Excel Error: " + error);
        // Fallback
        Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text });
    });
}

async function runExcelChartGen(text) {
    // Basic heuristic: Detect chart type logic or just assume "ColumnClustered"
    // Ideally prompt would return JSON {"type": "Line", "data": ...}
    // But for "Crazy" mode, let's just make it Smart.
    
    // If text contains "Pie", use Pie.
    // If text contains "Line", use Line.
    // Default Column.
    
    let chartType = "ColumnClustered";
    if (text.toLowerCase().includes("pie")) chartType = "Pie";
    else if (text.toLowerCase().includes("line") || text.toLowerCase().includes("garis")) chartType = "Line";
    else if (text.toLowerCase().includes("bar") || text.toLowerCase().includes("batang")) chartType = "BarClustered";
    
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = context.workbook.getSelectedRange(); // Assume user selected data
        
        // If selection is single cell, try used range? No, dangerous.
        // Let's assume user selected data.
        
        const chart = sheet.charts.add(chartType, range);
        chart.title.text = "Generated Chart";
        
        await context.sync();
        console.log("Chart created.");
    }).catch(err => {
        console.error("Chart Error: " + err);
    });
}

async function runPowerPointSlideGen(text) {
    // 1. Try to Parse as JSON (for Multi-Slide Decks)
    let slidesData = [];
    
    try {
        // Find JSON Array in text
        const jsonMatch = text.match(/\[\s*\{.*\}\s*\]/s);
        if (jsonMatch) {
            slidesData = JSON.parse(jsonMatch[0]);
        }
    } catch (e) {
        console.log("Not JSON, falling back to single slide parser.");
    }

    // 2. Fallback: Parse Single Slide Format (Title\nBullets)
    if (slidesData.length === 0) {
         let lines = text.split('\n').filter(line => line.trim() !== '');
         // Cleanup "Sure"
         if (lines.length > 0 && (lines[0].toLowerCase().startsWith("sure") || lines[0].toLowerCase().startsWith("here"))) lines.shift();
         
         if (lines.length > 0) {
             let title = lines[0].replace(/^TITLE:\s*/i, '').replace(/^[#*]+\s*/, '').trim();
             lines.shift();
             let points = lines.map(line => line.replace(/^[-*•]\s*/, '').trim());
             slidesData.push({ title: title, points: points });
         }
    }
    
    if (slidesData.length === 0) return;

    // 3. Generate Slides Loop
    await PowerPoint.run(async (context) => {
        const presentation = context.presentation;
        const slides = presentation.slides;

        // Loop through each slide data
        for (const data of slidesData) {
            const slide = slides.add();
            // Assuming Standard Layout (Title + Content)
            const titleShape = slide.shapes.getItemAt(0); 
            const bodyShape = slide.shapes.getItemAt(1);
            
            titleShape.textFrame.textRange.text = data.title || "No Title";
            
            // Body
            const bodyText = Array.isArray(data.points) ? data.points.join('\n') : (data.points || "");
            bodyShape.textFrame.textRange.text = bodyText;
        }

        await context.sync();
        console.log(`Created ${slidesData.length} slides.`);
    }).catch(error => {
        console.error("PPT Error: " + error);
        // Fallback
        Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text });
    });
}
