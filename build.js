const fs = require('fs');
const path = require('path');

function buildIndex() {
    console.log("Building index.html from components...");
    
    const header = fs.readFileSync(path.join(__dirname, 'components/Header.html'), 'utf8');
    const navbar = fs.readFileSync(path.join(__dirname, 'components/Navbar.html'), 'utf8');
    const chatArea = fs.readFileSync(path.join(__dirname, 'components/ChatArea.html'), 'utf8');
    const academicPanel = fs.readFileSync(path.join(__dirname, 'components/Panels/AcademicPanel.html'), 'utf8');
    const dataPanel = fs.readFileSync(path.join(__dirname, 'components/Panels/DataPanel.html'), 'utf8');
    const presentationPanel = fs.readFileSync(path.join(__dirname, 'components/Panels/PresentationPanel.html'), 'utf8');
    const extrasPanel = fs.readFileSync(path.join(__dirname, 'components/Panels/ExtrasPanel.html'), 'utf8');

    const finalHtml = `<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AutomateBoty Workspace</title>
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="styles.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/styles/github-light.min.css">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/highlight.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"></script>
</head>
<body>
<div class="layout">
    <main class="main-content">
        <!-- HEADER -->
        ${header}
        
        <!-- TOP NAVBAR -->
        ${navbar}
        
        <!-- DYNAMIC PANELS (Absolute or overlapping) -->
        <div class="panels-container">
            ${academicPanel}
            ${dataPanel}
            ${presentationPanel}
            ${extrasPanel}
        </div>

        <!-- CHAT & INPUT AREA -->
        <div class="chat-wrapper">
            ${chatArea}
        </div>
    </main>
</div>

<!-- ONBOARDING MODAL -->
<div id="onboarding-modal" class="modal-overlay hidden">
    <div class="modal-box">
        <div class="tour-icon">✨</div>
        <div class="tour-title">Selamat Datang di Workspace!</div>
        <div class="tour-desc">Antarmuka baru yang lebih bersih dan profesional.</div>
        <div class="tour-progress">1 / 5</div>
        <div class="tour-nav">
            <button class="btn-outline tour-prev" onclick="tourPrev()">← Kembali</button>
            <button class="btn-primary tour-next" onclick="tourNext()">Lanjut →</button>
        </div>
        <button class="modal-close" onclick="closeTour()">✕</button>
    </div>
</div>

<script src="ui-extras.js"></script>
<script src="features-word.js"></script>
<script src="features-excel.js"></script>
<script src="features-ppt.js"></script>
<script src="script.js?v=9"></script>
</body>
</html>`;

    fs.writeFileSync(path.join(__dirname, 'index.html'), finalHtml);
    console.log("Successfully built index.html!");
}

buildIndex();
