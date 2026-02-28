// loader.js - Dynamically loads HTML components into the main index.html
async function loadComponent(url, containerId) {
    try {
        const response = await fetch(url);
        if (!response.ok) throw new Error(`Failed to load ${url}: ${response.statusText}`);
        const html = await response.text();
        document.getElementById(containerId).innerHTML = html;
        console.log(`Loaded ${url} into #${containerId}`);
    } catch (error) {
        console.error('Component loading error:', error);
        document.getElementById(containerId).innerHTML = `<div style="color:red; padding:10px;">Error loading ${url}</div>`;
    }
}

async function loadAllComponents() {
    // Load components in parallel
    await Promise.all([
        loadComponent('components/Header.html', 'header-container'),
        loadComponent('components/Navbar.html', 'navbar-container'),
        loadComponent('components/ChatArea.html', 'chat-area-container'),
        loadComponent('components/Panels/AcademicPanel.html', 'academic-panel-container'),
        loadComponent('components/Panels/DataPanel.html', 'data-panel-container'),
        loadComponent('components/Panels/PresentationPanel.html', 'presentation-panel-container'),
        loadComponent('components/Panels/ExtrasPanel.html', 'extras-panel-container')
    ]);
    
    // Dispatch a custom event to notify script.js that the DOM is ready for event listeners
    document.dispatchEvent(new Event('componentsLoaded'));
}

// Start loading immediately as soon as loader.js is parsed
loadAllComponents();
