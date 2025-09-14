// MacroFlow AI - VSCode Style Interface with Real VBA Integration
Office.onReady(() => {
    
    // Global state
    const state = {
        modules: [],
        activeModule: null,
        editor: null,
        tabs: new Map(),
        contextMenu: null,
        // Chat state
        conversation: [],
        chatCollapsed: false
    };
    
    // State management for ribbon communication
    function saveStateToStorage() {
        try {
            const stateToSave = {
                activeModule: state.activeModule ? {
                    name: state.activeModule.name,
                    content: state.activeModule.content,
                    isModified: state.activeModule.isModified
                } : null,
                modules: state.modules.map(module => ({
                    name: module.name,
                    content: module.content,
                    isModified: module.isModified
                })),
                timestamp: Date.now()
            };
            localStorage.setItem('macroflow-state', JSON.stringify(stateToSave));
        } catch (error) {
            console.warn('Failed to save state to localStorage:', error);
        }
    }
    
    function loadStateFromStorage() {
        try {
            const savedState = localStorage.getItem('macroflow-state');
            if (savedState) {
                const parsed = JSON.parse(savedState);
                // Only load if it's recent (within 1 hour)
                if (Date.now() - parsed.timestamp < 3600000) {
                    return parsed;
                }
            }
        } catch (error) {
            console.warn('Failed to load state from localStorage:', error);
        }
        return null;
    }
    
    // DOM elements
    const elements = {
        moduleList: document.getElementById("moduleList"),
        editorTabs: document.getElementById("editorTabs"),
        monacoEditor: document.getElementById("monacoEditor"),
        welcomeScreen: document.getElementById("welcomeScreen"),
        prompt: document.getElementById("prompt"),
        generateBtn: document.getElementById("generateBtn"),
        welcomeGenerate: document.getElementById("welcomeGenerate"),
        addModuleBtn: document.getElementById("addModuleBtn"),
        moduleNameModal: document.getElementById("moduleNameModal"),
        moduleNameInput: document.getElementById("moduleNameInput"),
        createModuleOk: document.getElementById("createModuleOk"),
        createModuleCancel: document.getElementById("createModuleCancel"),
        renameModuleModal: document.getElementById("renameModuleModal"),
        renameModuleInput: document.getElementById("renameModuleInput"),
        renameModuleOk: document.getElementById("renameModuleOk"),
        renameModuleCancel: document.getElementById("renameModuleCancel"),
        // Line indicator element
        lineIndicator: document.getElementById("lineIndicator"),
        // Sidebar elements
        searchMacros: document.getElementById("searchMacros"),
        toggleSidebarBtn: document.getElementById("toggleSidebarBtn"),
        sidebar: document.getElementById("sidebar"),
        // Chat panel elements
        chatPanel: document.getElementById("chatPanel"),
        chatMessages: document.getElementById("chatMessages"),
        toggleChatBtn: document.getElementById("toggleChatBtn"),
        // Resize handle (no longer needed - resize handled directly on chat panel)
    };
    
    // Initialize Monaco Editor
    function initializeMonacoEditor() {
        require.config({ paths: { vs: 'https://cdn.jsdelivr.net/npm/monaco-editor@0.45.0/min/vs' } });
        
        require(['vs/editor/editor.main'], function () {
            state.editor = monaco.editor.create(elements.monacoEditor, {
                value: '// Select a module to start editing VBA code',
                language: 'vb',
                theme: 'vs-dark',
                automaticLayout: true,
                minimap: { enabled: false },
                fontSize: 13,
                fontFamily: 'Consolas, Monaco, monospace',
                lineNumbers: 'on',
                roundedSelection: false,
                scrollBeyondLastLine: false,
                readOnly: false,
                cursorStyle: 'line',
                matchBrackets: 'always',
                wordWrap: 'on',
                contextmenu: true,
                selectOnLineNumbers: true,
                glyphMargin: true,  // Enable glyph margin for run buttons
                lineNumbersMinChars: 2,  // Minimize line number width even more
                glyphMarginWidth: 18  // Set narrower glyph margin width
            });
            
            // Listen for content changes
            let updateTimeout = null;
            state.editor.onDidChangeModelContent(() => {
                updateActionButtons();
                if (state.activeModule) {
                    state.activeModule.content = state.editor.getValue();
                    state.activeModule.isModified = true;
                    updateTabModifiedState(state.activeModule.name);
                    // Auto-save after 2 seconds of inactivity
                    clearTimeout(state.autoSaveTimer);
                    state.autoSaveTimer = setTimeout(() => {
                        if (state.activeModule && state.activeModule.isModified) {
                            saveModuleToExcel(state.activeModule);
                        }
                    }, 2000);

                    // Update subroutines with debouncing
                    clearTimeout(updateTimeout);
                    updateTimeout = setTimeout(() => {
                        const metadata = extractVBAMetadata(state.activeModule.content);
                        state.activeModule.subroutines = metadata.subroutines;
                        renderModuleList();
                        updateRunButtons();
                    }, 500);
                }
            });

            // Listen for selection changes to update line indicator
            state.editor.onDidChangeCursorSelection((e) => {
                updateLineIndicator();
            });

            // Listen for cursor position changes
            state.editor.onDidChangeCursorPosition((e) => {
                updateLineIndicator();
            });
            
        });
    }
    
    // Chat Functions
    function addMessage(type, content) {
        const message = {
            id: Date.now(),
            type: type, // 'user' or 'assistant'
            content: content,
            timestamp: new Date()
        };
        
        state.conversation.push(message);
        renderMessage(message);
        scrollToBottom();
    }
    
    function renderMessage(message) {
        const messageDiv = document.createElement('div');
        messageDiv.className = `message ${message.type}`;
        messageDiv.dataset.messageId = message.id;
        
        // Render message content with proper formatting
        const content = message.type === 'assistant' ? 
            renderMarkdown(message.content) : 
            escapeHtml(message.content);
            
        const messageHTML = `
            <div class="message-bubble">
                ${content}
            </div>
            <div class="message-timestamp">
                ${formatTimestamp(message.timestamp)}
            </div>
        `;
        
        messageDiv.innerHTML = messageHTML;
        
        // Remove welcome message if it exists
        const welcomeMsg = elements.chatMessages.querySelector('.chat-welcome');
        if (welcomeMsg) {
            welcomeMsg.remove();
        }
        
        elements.chatMessages.appendChild(messageDiv);
    }
    
    function formatTimestamp(date) {
        return date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    }
    
    function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
    
    function renderMarkdown(text) {
        // Simple markdown rendering for AI responses
        let html = escapeHtml(text);
        
        // Convert **bold** to <strong>
        html = html.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
        
        // Convert • bullet points to proper list items
        const lines = html.split('\n');
        let result = [];
        let inList = false;
        
        for (let line of lines) {
            const trimmedLine = line.trim();
            
            if (trimmedLine.startsWith('•') || trimmedLine.startsWith('-')) {
                if (!inList) {
                    result.push('<ul>');
                    inList = true;
                }
                const content = trimmedLine.substring(1).trim();
                result.push(`<li>${content}</li>`);
            } else if (trimmedLine.startsWith('**') && trimmedLine.endsWith('**')) {
                // Handle standalone bold headers
                if (inList) {
                    result.push('</ul>');
                    inList = false;
                }
                result.push(`<div class="section-header">${trimmedLine}</div>`);
            } else {
                if (inList) {
                    result.push('</ul>');
                    inList = false;
                }
                if (trimmedLine) {
                    result.push(`<p>${trimmedLine}</p>`);
                }
            }
        }
        
        if (inList) {
            result.push('</ul>');
        }
        
        return result.join('');
    }
    
    function scrollToBottom() {
        // Use requestAnimationFrame for better timing with DOM updates
        requestAnimationFrame(() => {
            // Small delay to ensure DOM content is fully rendered
            setTimeout(() => {
                if (elements.chatMessages) {
                    elements.chatMessages.scrollTop = elements.chatMessages.scrollHeight;
                }
            }, 10);
        });
    }
    
    function toggleChatPanel() {
        state.chatCollapsed = !state.chatCollapsed;
        elements.chatPanel.classList.toggle('collapsed', state.chatCollapsed);
        elements.toggleChatBtn.title = state.chatCollapsed ? 'Expand Chat' : 'Collapse Chat';
        
        // If expanding, restore the previous width or default to 300px
        if (!state.chatCollapsed) {
            const storedWidth = elements.chatPanel.dataset.lastWidth || '300px';
            elements.chatPanel.style.width = storedWidth;
        } else {
            // Store current width before collapsing
            elements.chatPanel.dataset.lastWidth = elements.chatPanel.style.width || '300px';
        }
    }
    
    // Chat panel resize functionality
    function initializeChatResize() {
        let isResizing = false;
        
        // Add resize functionality to the chat panel's right edge
        elements.chatPanel.addEventListener('mousedown', (e) => {
            // Only trigger if clicking on the right edge (expanded detection area)
            const rect = elements.chatPanel.getBoundingClientRect();
            const isRightEdge = e.clientX >= rect.right - 8 && e.clientX <= rect.right + 8;
            
            if (!isRightEdge || state.chatCollapsed) return;
            
            e.preventDefault();
            isResizing = true;
            document.body.style.cursor = 'col-resize';
            document.body.style.userSelect = 'none';
            
            const startX = e.clientX;
            const startWidth = elements.chatPanel.offsetWidth;
            
            function handleMouseMove(e) {
                if (!isResizing) return;
                
                const deltaX = e.clientX - startX;
                const newWidth = startWidth + deltaX;
                
                // Apply min/max constraints
                const minWidth = 200;
                const maxWidth = 500;
                const constrainedWidth = Math.min(Math.max(newWidth, minWidth), maxWidth);
                
                elements.chatPanel.style.width = constrainedWidth + 'px';
            }
            
            function handleMouseUp() {
                isResizing = false;
                document.body.style.cursor = '';
                document.body.style.userSelect = '';
                document.removeEventListener('mousemove', handleMouseMove);
                document.removeEventListener('mouseup', handleMouseUp);
            }
            
            document.addEventListener('mousemove', handleMouseMove);
            document.addEventListener('mouseup', handleMouseUp);
        });
        
        // Add mousemove listener for cursor change when hovering over resize area
        elements.chatPanel.addEventListener('mousemove', (e) => {
            if (isResizing || state.chatCollapsed) return;
            
            const rect = elements.chatPanel.getBoundingClientRect();
            const isRightEdge = e.clientX >= rect.right - 8 && e.clientX <= rect.right + 8;
            
            if (isRightEdge) {
                elements.chatPanel.style.cursor = 'col-resize';
            } else {
                elements.chatPanel.style.cursor = '';
            }
        });
        
        // Reset cursor when leaving chat panel
        elements.chatPanel.addEventListener('mouseleave', () => {
            if (!isResizing) {
                elements.chatPanel.style.cursor = '';
            }
        });
    }
    
    // Sidebar resize functionality
    function initializeSidebarResize() {
        let isResizing = false;
        
        // Add resize functionality to the sidebar's left edge
        elements.sidebar.addEventListener('mousedown', (e) => {
            // Only trigger if clicking on the left edge (expanded detection area)
            const rect = elements.sidebar.getBoundingClientRect();
            const isLeftEdge = e.clientX >= rect.left - 8 && e.clientX <= rect.left + 8;
            
            if (!isLeftEdge || elements.sidebar.classList.contains('collapsed')) return;
            
            e.preventDefault();
            isResizing = true;
            document.body.style.cursor = 'col-resize';
            document.body.style.userSelect = 'none';
            
            const startX = e.clientX;
            const startWidth = elements.sidebar.offsetWidth;
            
            function handleMouseMove(e) {
                if (!isResizing) return;
                
                const deltaX = startX - e.clientX; // Inverted for left edge
                const newWidth = startWidth + deltaX;
                
                // Apply min/max constraints
                const minWidth = 180;
                const maxWidth = 400;
                const constrainedWidth = Math.min(Math.max(newWidth, minWidth), maxWidth);
                
                elements.sidebar.style.width = constrainedWidth + 'px';
            }
            
            function handleMouseUp() {
                isResizing = false;
                document.body.style.cursor = '';
                document.body.style.userSelect = '';
                document.removeEventListener('mousemove', handleMouseMove);
                document.removeEventListener('mouseup', handleMouseUp);
            }
            
            document.addEventListener('mousemove', handleMouseMove);
            document.addEventListener('mouseup', handleMouseUp);
        });
        
        // Add mousemove listener for cursor change when hovering over resize area
        elements.sidebar.addEventListener('mousemove', (e) => {
            if (isResizing || elements.sidebar.classList.contains('collapsed')) return;
            
            const rect = elements.sidebar.getBoundingClientRect();
            const isLeftEdge = e.clientX >= rect.left - 8 && e.clientX <= rect.left + 8;
            
            if (isLeftEdge) {
                elements.sidebar.style.cursor = 'col-resize';
            } else {
                elements.sidebar.style.cursor = '';
            }
        });
        
        // Reset cursor when leaving sidebar
        elements.sidebar.addEventListener('mouseleave', () => {
            if (!isResizing) {
                elements.sidebar.style.cursor = '';
            }
        });
    }
    
    // Helper function to generate next module name
    function generateNextModuleName() {
        // Find all existing module names that match the pattern ModuleN
        const moduleNumbers = state.modules
            .map(module => {
                const match = module.name.match(/^Module(\d+)$/);
                return match ? parseInt(match[1]) : 0;
            })
            .filter(num => num > 0);
        
        // Find the highest number, default to 0 if no modules exist
        const maxNumber = moduleNumbers.length > 0 ? Math.max(...moduleNumbers) : 0;
        
        // Return the next number
        return `Module${maxNumber + 1}`;
    }
    
    // Replace a specific subroutine in the full code (for cursor-based modifications)
    function replaceSubroutineInCode(fullCode, currentSubroutine, newCode) {
        if (!currentSubroutine || !currentSubroutine.startLine || !currentSubroutine.endLine) {
            console.warn('Invalid subroutine data for replacement:', currentSubroutine);
            return smartFunctionReplacement(fullCode, newCode);
        }

        const lines = fullCode.split('\n');

        // Replace the lines from startLine to endLine with the new code
        const beforeLines = lines.slice(0, currentSubroutine.startLine - 1);
        const afterLines = lines.slice(currentSubroutine.endLine);

        // Split new code into lines and clean up
        const newCodeLines = newCode.split('\n');

        // Combine: before + new code + after
        const result = [...beforeLines, ...newCodeLines, ...afterLines].join('\n');

        // Subroutine replaced successfully
        return result;
    }

    // Smart function replacement for code modifications (legacy function)
    function smartFunctionReplacement(originalCode, newCode) {
        // Extract function name from the new code
        const newFunctionMatch = newCode.match(/(?:Sub|Function)\s+(\w+)/i);
        if (!newFunctionMatch) {
            // If no function detected, append to existing code
            return originalCode + "\n\n" + newCode;
        }
        
        const functionName = newFunctionMatch[1];
        
        // Find and replace the existing function in original code
        const functionPattern = new RegExp(
            `(?:Private\\s+|Public\\s+)?(?:Sub|Function)\\s+${functionName}\\b.*?\\n(?:End\\s+(?:Sub|Function)\\b.*?\\n)`,
            'gis'
        );
        
        if (functionPattern.test(originalCode)) {
            // Replace the existing function
            return originalCode.replace(functionPattern, newCode + '\n');
        } else {
            // Function not found, append to existing code
            return originalCode + "\n\n" + newCode;
        }
    }

    // Automatic VBA code insertion function
    async function insertVBACode(vbaCode, context, isModification = false) {
        try {
            // If no active module, create a new one
            if (!state.activeModule) {
                const moduleName = generateNextModuleName();
                try {
                    const createResponse = await fetch("http://localhost:5000/create-module", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({
                            name: moduleName,
                            content: vbaCode
                        })
                    });
                    
                    if (createResponse.ok) {
                        const newModule = {
                            name: moduleName,
                            content: vbaCode,
                            isModified: false
                        };
                        state.modules.push(newModule);
                        renderModuleList();
                        openModule(newModule);
                        return { success: true, moduleName: moduleName };
                    } else {
                        return { success: false, message: "Failed to create new module for the code." };
                    }
                } catch (createError) {
                    return { success: false, message: "Could not create new module. Please select or create a module first." };
                }
            }
            
            // Insert into existing active module with enhanced logic
            if (state.editor) {
                if (isModification && context.currentSubroutine) {
                    // PRIORITY 1: Cursor inside subroutine - replace the entire subroutine
                    const currentSub = context.currentSubroutine;
                    const modifiedCode = replaceSubroutineInCode(context.fullCode, currentSub, vbaCode);
                    state.editor.setValue(modifiedCode);

                } else if (isModification && context.hasSelection && context.selectionRange) {
                    // PRIORITY 2: User selected specific text to replace
                    const selection = new monaco.Selection(
                        context.selectionRange.startLine,
                        context.selectionRange.startColumn,
                        context.selectionRange.endLine,
                        context.selectionRange.endColumn
                    );
                    state.editor.executeEdits("auto-vba-insertion", [{
                        range: selection,
                        text: vbaCode
                    }]);

                } else if (isModification && context.fullCode) {
                    // PRIORITY 3: General modification - smart function replacement
                    const modifiedCode = smartFunctionReplacement(context.fullCode, vbaCode);
                    state.editor.setValue(modifiedCode);

                } else if (context.currentSubroutine) {
                    // AVOID: Never insert new code inside existing subroutine
                    // Instead, append after the current subroutine
                    const currentSub = context.currentSubroutine;
                    const insertLine = currentSub.endLine + 1;
                    state.editor.executeEdits("auto-vba-insertion", [{
                        range: new monaco.Range(insertLine, 1, insertLine, 1),
                        text: "\n" + vbaCode + "\n"
                    }]);

                } else if (context.cursorPosition) {
                    // Safe cursor position insertion (only when not inside existing function)
                    state.editor.executeEdits("auto-vba-insertion", [{
                        range: new monaco.Range(
                            context.cursorPosition.lineNumber,
                            context.cursorPosition.column,
                            context.cursorPosition.lineNumber,
                            context.cursorPosition.column
                        ),
                        text: "\n" + vbaCode + "\n"
                    }]);

                } else {
                    // Fallback: Replace entire content
                    state.editor.setValue(vbaCode);
                }
                
                // Update module content and save
                state.activeModule.content = state.editor.getValue();
                state.activeModule.isModified = true;
                updateTabModifiedState(state.activeModule.name);
                await saveModuleToExcel(state.activeModule);
                
                return { success: true, moduleName: state.activeModule.name };
            } else {
                // No editor available, update module content directly
                state.activeModule.content = vbaCode;
                state.activeModule.isModified = true;
                updateTabModifiedState(state.activeModule.name);
                await saveModuleToExcel(state.activeModule);
                
                return { success: true, moduleName: state.activeModule.name };
            }
            
        } catch (error) {
            console.error("VBA insertion error:", error);
            return { success: false, message: `Failed to insert code: ${error.message}` };
        }
    }
    
    
    // Load modules from Excel
    async function loadModules() {
        try {
            
            const response = await fetch("http://localhost:5000/list-modules");
            if (!response.ok) {
                const errorData = await response.json().catch(() => ({}));
                throw new Error(errorData.error || `HTTP ${response.status}`);
            }
            
            const data = await response.json();
            state.modules = data.modules || [];
            saveStateToStorage();
            
            renderModuleList();
            
            
        } catch (error) {
            console.error("Error loading modules:", error);
            
            if (error.message.includes("fetch")) {
                console.error("Cannot connect to backend. Is Flask server running?");
            } else if (error.message.includes("VBA project not accessible")) {
                console.error("Enable 'Trust access to VBA project object model' in Excel Trust Center");
            } else if (error.message.includes("No active workbook")) {
                console.error("Please open an Excel workbook first");
            } else {
                console.error(`Error: ${error.message}`);
            }
            
            // Clear modules on error
            state.modules = [];
            renderModuleList();
        }
    }
    
    // Save module content to Excel
    async function saveModuleToExcel(module) {
        try {
            const response = await fetch("http://localhost:5000/update-module", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    name: module.name,
                    content: module.content
                })
            });
            
            if (response.ok) {
                module.isModified = false;
                updateTabModifiedState(module.name);
            }
        } catch (error) {
            console.error("Error auto-saving module:", error);
        }
    }
    
    // Render module list in sidebar
    function renderModuleList() {
        elements.moduleList.innerHTML = "";
        
        if (state.modules.length === 0) {
            elements.moduleList.innerHTML = '<div class="module-item loading">No modules found</div>';
            return;
        }
        
        state.modules.forEach(module => {
            const moduleContainer = document.createElement("div");
            
            const moduleItem = document.createElement("div");
            moduleItem.className = `module-item ${module === state.activeModule ? "active" : ""} ${module.expanded ? "expanded" : ""}`;
            
            const hasSubroutines = module.subroutines && module.subroutines.length > 0;
            
            moduleItem.innerHTML = `
                <div class="module-name">
                    <div class="expand-icon" ${hasSubroutines ? '' : 'style="visibility: hidden"'}>
                        <svg width="16" height="16" viewBox="0 0 16 16">
                            <path fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" d="M6 4l4 4-4 4"/>
                        </svg>
                    </div>
                    <span>${module.name}</span>
                </div>
            `;
            
            // Click to open module (but not on expand icon)
            moduleItem.addEventListener("click", (e) => {
                if (!e.target.closest(".expand-icon")) {
                    openModule(module);
                }
            });
            
            // Expand/collapse functionality
            const expandIcon = moduleItem.querySelector(".expand-icon");
            if (hasSubroutines) {
                expandIcon.addEventListener("click", (e) => {
                    e.stopPropagation();
                    module.expanded = !module.expanded;
                    renderModuleList(); // Re-render to update expansion state
                });
            }
            
            // Context menu
            moduleItem.addEventListener("contextmenu", (e) => {
                e.preventDefault();
                showContextMenu(e, module);
            });
            
            moduleContainer.appendChild(moduleItem);
            
            // Add subroutines if expanded
            if (module.expanded && hasSubroutines) {
                
                const subroutineList = document.createElement("div");
                subroutineList.className = "subroutine-list";
                
                module.subroutines.forEach(subroutine => {
                    
                    const subItem = document.createElement("div");
                    subItem.className = "subroutine-item";
                    subItem.innerHTML = `
                        <span>${subroutine.name}</span>
                    `;
                    
                    // Click to jump to subroutine in editor
                    subItem.addEventListener("click", (e) => {
                        e.stopPropagation();
                        openModule(module);
                        jumpToSubroutine(subroutine);
                    });
                    
                    subroutineList.appendChild(subItem);
                });
                
                moduleContainer.appendChild(subroutineList);
            }
            
            elements.moduleList.appendChild(moduleContainer);
        });
    }
    
    // Jump to a specific subroutine in the editor
    function jumpToSubroutine(subroutine) {
        if (!state.editor) return;
        
        const model = state.editor.getModel();
        if (!model) return;
        
        const content = model.getValue();
        const lines = content.split('\n');
        
        // Find the line with the subroutine
        for (let i = 0; i < lines.length; i++) {
            if (lines[i].trim().includes(subroutine.name) && 
                (lines[i].includes('Sub ') || lines[i].includes('Function '))) {
                // Jump to that line
                state.editor.setPosition({ lineNumber: i + 1, column: 1 });
                state.editor.revealLineInCenter(i + 1);
                state.editor.focus();
                break;
            }
        }
    }
    
    
    
    // Open module in editor
    function openModule(module) {
        if (state.activeModule === module) return;

        state.activeModule = module;
        saveStateToStorage();

        // Hide welcome screen
        elements.welcomeScreen.classList.add("hidden");

        // Update editor content
        if (state.editor) {
            state.editor.setValue(module.content || "");
        }

        // Create or show tab
        if (!state.tabs.has(module.name)) {
            createTab(module);
        }

        // Activate tab
        activateTab(module.name);

        // Update sidebar
        renderModuleList();

        // Update buttons
        updateActionButtons();

        // Update line indicator
        updateLineIndicator();

        // Update run buttons in glyph margin
        setTimeout(() => {
            updateRunButtons();
        }, 100);
    }
    
    // Create editor tab
    function createTab(module) {
        const tab = document.createElement("div");
        tab.className = "tab";
        tab.dataset.tab = module.name;
        tab.draggable = true;
        tab.innerHTML = `
            <span>${module.name}</span>
            <button class="tab-close" title="Close">×</button>
        `;
        
        // Click to activate
        tab.addEventListener("click", (e) => {
            if (!e.target.classList.contains("tab-close")) {
                const targetModule = state.modules.find(m => m.name === module.name);
                if (targetModule) openModule(targetModule);
            }
        });
        
        // Drag and drop functionality
        tab.addEventListener("dragstart", (e) => {
            e.dataTransfer.setData("text/plain", module.name);
            tab.classList.add("dragging");
        });
        
        tab.addEventListener("dragend", (e) => {
            tab.classList.remove("dragging");
        });
        
        tab.addEventListener("dragover", (e) => {
            e.preventDefault();
        });
        
        tab.addEventListener("drop", (e) => {
            e.preventDefault();
            const draggedModuleName = e.dataTransfer.getData("text/plain");
            const targetModuleName = module.name;
            
            if (draggedModuleName !== targetModuleName) {
                reorderTabs(draggedModuleName, targetModuleName);
            }
        });
        
        // Close tab
        tab.querySelector(".tab-close").addEventListener("click", (e) => {
            e.stopPropagation();
            closeTab(module.name);
        });
        
        elements.editorTabs.appendChild(tab);
        state.tabs.set(module.name, tab);
    }
    
    // Activate tab
    function activateTab(moduleName) {
        // Remove active from all tabs
        document.querySelectorAll(".tab").forEach(tab => tab.classList.remove("active"));
        
        // Add active to current tab
        const tab = state.tabs.get(moduleName);
        if (tab) {
            tab.classList.add("active");
        }
    }
    
    // Close tab
    function closeTab(moduleName) {
        const tab = state.tabs.get(moduleName);
        if (tab) {
            tab.remove();
            state.tabs.delete(moduleName);
        }
        
        // If this was the active module, switch to another or show welcome
        if (state.activeModule && state.activeModule.name === moduleName) {
            const remainingTabs = Array.from(state.tabs.keys());
            if (remainingTabs.length > 0) {
                const nextModule = state.modules.find(m => m.name === remainingTabs[0]);
                if (nextModule) openModule(nextModule);
            } else {
                state.activeModule = null;
                elements.welcomeScreen.classList.remove("hidden");
                if (state.editor) state.editor.setValue("");
                renderModuleList();
                updateActionButtons();
            }
        }
    }
    
    // Reorder tabs by moving dragged tab to target position
    function reorderTabs(draggedModuleName, targetModuleName) {
        const tabsContainer = elements.editorTabs;
        const draggedTab = state.tabs.get(draggedModuleName);
        const targetTab = state.tabs.get(targetModuleName);
        
        if (!draggedTab || !targetTab) return;
        
        // Get all tabs in current order
        const allTabs = Array.from(tabsContainer.children);
        const draggedIndex = allTabs.indexOf(draggedTab);
        const targetIndex = allTabs.indexOf(targetTab);
        
        // Remove dragged tab and insert it at target position
        tabsContainer.removeChild(draggedTab);
        
        if (targetIndex < draggedIndex) {
            // Insert before target
            tabsContainer.insertBefore(draggedTab, targetTab);
        } else {
            // Insert after target
            tabsContainer.insertBefore(draggedTab, targetTab.nextSibling);
        }
        
        // Update tab order state if needed (for persistence)
        // Could be extended to save tab order preferences
    }
    
    // Update tab modified state
    function updateTabModifiedState(moduleName) {
        const tab = state.tabs.get(moduleName);
        const module = state.modules.find(m => m.name === moduleName);
        if (tab && module) {
            const span = tab.querySelector("span");
            span.textContent = module.name;
        }
        renderModuleList(); // Update sidebar too
    }
    
    // Update action buttons state (placeholder for future use)
    function updateActionButtons() {
        // No action buttons to update anymore
    }
    
    // Show context menu
    function showContextMenu(event, module) {
        // Remove existing context menu
        if (state.contextMenu) {
            state.contextMenu.remove();
        }
        
        const contextMenu = document.createElement("div");
        contextMenu.className = "context-menu";
        contextMenu.innerHTML = `
            <div class="context-menu-item" data-action="rename">
                <span>Rename</span>
            </div>
        `;
        
        contextMenu.style.left = `${event.clientX}px`;
        contextMenu.style.top = `${event.clientY}px`;
        
        // Handle context menu clicks
        contextMenu.addEventListener("click", (e) => {
            const action = e.target.closest(".context-menu-item")?.dataset.action;
            if (action === "rename") {
                showRenameModuleModal(module);
            }
            contextMenu.remove();
            state.contextMenu = null;
        });
        
        // Close context menu on outside click
        setTimeout(() => {
            document.addEventListener("click", function closeContextMenu() {
                if (contextMenu.parentNode) {
                    contextMenu.remove();
                }
                state.contextMenu = null;
                document.removeEventListener("click", closeContextMenu);
            });
        }, 100);
        
        document.body.appendChild(contextMenu);
        state.contextMenu = contextMenu;
    }
    
    // Show modal for new module name
    function showNewModuleModal() {
        
        // Set default name using the same logic as automatic module creation
        elements.moduleNameInput.value = generateNextModuleName();
        
        // Show modal
        elements.moduleNameModal.classList.remove("hidden");
        
        // Focus input
        setTimeout(() => elements.moduleNameInput.focus(), 100);
    }
    
    // Hide modal
    function hideNewModuleModal() {
        elements.moduleNameModal.classList.add("hidden");
        elements.moduleNameInput.value = "";
    }
    
    // Create new module in Excel
    async function createNewModule() {
        const name = elements.moduleNameInput.value.trim();
        
        if (!name) {
            console.error("Please enter a module name");
            return;
        }
        
        
        // Check for duplicate names
        if (state.modules.some(m => m.name === name)) {
            console.error("Module name already exists");
            return;
        }
        
        // Hide modal first
        hideNewModuleModal();
        
        try {
            
            const initialContent = `Sub ${name}Macro()
    ' Start writing your VBA code here
    ' Use the chat on the left for help

End Sub`;
            
            const response = await fetch("http://localhost:5000/create-module", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    name: name,
                    content: initialContent
                })
            });
            
            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.error || `HTTP ${response.status}`);
            }
            
            // Add to local state
            const newModule = {
                name: name,
                content: initialContent,
                isModified: false
            };
            
            state.modules.push(newModule);
            renderModuleList();
            openModule(newModule);
            
            
        } catch (error) {
            console.error("Error creating module:", error);
            console.error(`Failed to create module: ${error.message}`);
        }
    }
    
    // Show rename module modal
    function showRenameModuleModal(module) {
        
        // Store module reference for renaming
        state.moduleToRename = module;
        
        // Set current name
        elements.renameModuleInput.value = module.name;
        
        // Show modal
        elements.renameModuleModal.classList.remove("hidden");
        
        // Focus input and select text
        setTimeout(() => {
            elements.renameModuleInput.focus();
            elements.renameModuleInput.select();
        }, 100);
    }
    
    // Hide rename modal
    function hideRenameModuleModal() {
        elements.renameModuleModal.classList.add("hidden");
        elements.renameModuleInput.value = "";
        state.moduleToRename = null;
    }
    
    // Rename module by creating new one with new name and copying content
    async function renameModule() {
        const newName = elements.renameModuleInput.value.trim();
        const oldModule = state.moduleToRename;
        
        if (!newName) {
            console.error("Please enter a module name");
            return;
        }
        
        if (newName === oldModule.name) {
            hideRenameModuleModal();
            return;
        }
        
        
        // Check for duplicate names
        if (state.modules.some(m => m.name === newName && m !== oldModule)) {
            console.error("Module name already exists");
            return;
        }
        
        // Hide modal first
        hideRenameModuleModal();
        
        try {
            
            // Create new module with the new name and same content
            const createResponse = await fetch("http://localhost:5000/create-module", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    name: newName,
                    content: oldModule.content
                })
            });
            
            if (!createResponse.ok) {
                const errorData = await createResponse.json();
                throw new Error(errorData.error || `HTTP ${createResponse.status}`);
            }
            
            // Delete the old module
            const deleteResponse = await fetch("http://localhost:5000/delete-module", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ name: oldModule.name })
            });
            
            if (!deleteResponse.ok) {
                console.warn("Failed to delete old module, but new one was created");
            }
            
            // Update local state
            const oldIndex = state.modules.indexOf(oldModule);
            if (oldIndex > -1) {
                // Replace old module with new one
                const newModule = {
                    name: newName,
                    content: oldModule.content,
                    isModified: oldModule.isModified
                };
                
                state.modules[oldIndex] = newModule;
                
                // Update tab if open
                if (state.tabs.has(oldModule.name)) {
                    const tab = state.tabs.get(oldModule.name);
                    tab.dataset.tab = newName;
                    tab.querySelector("span").textContent = newName;
                    state.tabs.delete(oldModule.name);
                    state.tabs.set(newName, tab);
                }
                
                // Update active module reference
                if (state.activeModule === oldModule) {
                    state.activeModule = newModule;
                }
                
                renderModuleList();
            }
            
        } catch (error) {
            console.error("Error renaming module:", error);
            console.error(`Failed to rename module: ${error.message}`);
        }
    }
    
    // Get editor selection and context
    // Extract VBA metadata from code
    function extractVBAMetadata(code) {
        const lines = code.split('\n');
        const subroutines = [];
        const functions = [];
        const variables = [];
        const comments = [];
        
        lines.forEach((line, index) => {
            const trimmed = line.trim();
            
            // Extract Subroutines (Sub declarations)
            const subMatch = trimmed.match(/^(Private\s+|Public\s+)?Sub\s+(\w+)/i);
            if (subMatch) {
                subroutines.push({
                    name: subMatch[2],
                    lineNumber: index + 1,
                    visibility: subMatch[1] ? subMatch[1].trim() : 'Public',
                    type: 'subroutine'
                });
            }
            
            // Extract Functions (Function declarations)
            const funcMatch = trimmed.match(/^(Private\s+|Public\s+)?Function\s+(\w+)/i);
            if (funcMatch) {
                functions.push({
                    name: funcMatch[2],
                    lineNumber: index + 1,
                    visibility: funcMatch[1] ? funcMatch[1].trim() : 'Public',
                    type: 'function'
                });
            }
            
            // Extract Dim/Variable declarations
            const varMatch = trimmed.match(/^(Dim|Private|Public)\s+(\w+)\s+As\s+(\w+)/i);
            if (varMatch) {
                variables.push({
                    name: varMatch[2],
                    dataType: varMatch[3],
                    scope: varMatch[1],
                    lineNumber: index + 1
                });
            }
            
            // Extract meaningful comments (not just apostrophes)
            if (trimmed.startsWith("'") && trimmed.length > 5) {
                comments.push({
                    text: trimmed.substring(1).trim(),
                    lineNumber: index + 1
                });
            }
        });
        
        return {
            subroutines: subroutines,
            functions: functions,
            procedures: [...subroutines, ...functions], // Combined for easy access
            variables: variables,
            comments: comments,
            totalLines: lines.length,
            currentSubroutine: null // Will be set based on cursor position if needed
        };
    }

    // Get information about all available modules
    function getModulesList() {
        return state.modules.map(module => ({
            name: module.name,
            isActive: module === state.activeModule,
            hasContent: module.content && module.content.trim().length > 0,
            contentLength: module.content ? module.content.length : 0,
            // Basic metadata without full parsing (for performance)
            preview: module.content ? module.content.split('\n').slice(0, 3).join('\n') : ''
        }));
    }

    // Detect which subroutine the cursor is currently inside
    function getCurrentSubroutine() {
        if (!state.editor || !state.activeModule) return null;

        const model = state.editor.getModel();
        const position = state.editor.getPosition();
        const currentLine = position.lineNumber;
        const fullCode = model.getValue();
        const lines = fullCode.split('\n');

        let currentSub = null;
        let subStartLine = -1;
        let subEndLine = -1;
        let indentLevel = 0;

        // Scan from top to find the subroutine containing current line
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            const lineNumber = i + 1;

            // Check for Sub/Function start
            const subMatch = line.match(/^(Private\s+|Public\s+)?(Sub|Function)\s+(\w+)/i);
            if (subMatch) {
                // If we were already in a sub and haven't hit End yet, this is a nested situation
                if (currentSub && lineNumber <= currentLine) {
                    // Update to this sub as we're deeper nested
                    currentSub = {
                        name: subMatch[3],
                        type: subMatch[2].toLowerCase(),
                        startLine: lineNumber,
                        endLine: -1, // Will be set when we find End Sub/Function
                        content: ''
                    };
                    subStartLine = lineNumber;
                    indentLevel++;
                } else if (lineNumber <= currentLine) {
                    // First sub we've encountered
                    currentSub = {
                        name: subMatch[3],
                        type: subMatch[2].toLowerCase(),
                        startLine: lineNumber,
                        endLine: -1,
                        content: ''
                    };
                    subStartLine = lineNumber;
                    indentLevel = 1;
                }
            }

            // Check for End Sub/Function
            if (line.match(/^End\s+(Sub|Function)$/i)) {
                if (currentSub && lineNumber >= currentLine && subEndLine === -1) {
                    // This End belongs to our current sub
                    currentSub.endLine = lineNumber;
                    subEndLine = lineNumber;
                    break;
                }
                indentLevel = Math.max(0, indentLevel - 1);
            }
        }

        // If we found a current sub, extract its full content
        if (currentSub && subStartLine > 0) {
            // Find the actual end line if not found yet
            if (currentSub.endLine === -1) {
                // Look forward to find End Sub/Function
                for (let i = subStartLine; i < lines.length; i++) {
                    if (lines[i].trim().match(/^End\s+(Sub|Function)$/i)) {
                        currentSub.endLine = i + 1;
                        break;
                    }
                }
            }

            // Extract the full subroutine content
            if (currentSub.endLine > 0) {
                currentSub.content = lines.slice(subStartLine - 1, currentSub.endLine).join('\n');
            } else {
                // No end found, take from start to end of file
                currentSub.content = lines.slice(subStartLine - 1).join('\n');
                currentSub.endLine = lines.length;
            }
        }

        return currentSub;
    }

    function getEditorContext() {
        if (!state.editor || !state.activeModule) {
            return {
                hasSelection: false,
                selectedText: "",
                cursorPosition: null,
                surroundingCode: "",
                currentSubroutine: null,
                fullCode: ""
            };
        }

        const selection = state.editor.getSelection();
        const model = state.editor.getModel();
        const position = state.editor.getPosition();
        
        // Get selected text
        const selectedText = selection && !selection.isEmpty() ? 
            model.getValueInRange(selection) : "";
        
        // Get surrounding context (5 lines before and after cursor/selection)
        const startLine = Math.max(1, (selection?.startLineNumber || position?.lineNumber || 1) - 5);
        const endLine = Math.min(model.getLineCount(), (selection?.endLineNumber || position?.lineNumber || 1) + 5);
        const surroundingCode = model.getValueInRange({
            startLineNumber: startLine,
            startColumn: 1,
            endLineNumber: endLine,
            endColumn: model.getLineMaxColumn(endLine)
        });

        const fullCode = model.getValue();
        const moduleMetadata = extractVBAMetadata(fullCode);
        const currentSubroutine = getCurrentSubroutine();

        // Determine if cursor is inside an existing subroutine
        const isInsideSubroutine = currentSubroutine !== null;

        // Enhanced surrounding code: if inside subroutine, use entire subroutine as context
        let enhancedSurroundingCode = surroundingCode;
        if (currentSubroutine) {
            enhancedSurroundingCode = currentSubroutine.content;
        }

        return {
            hasSelection: selectedText.length > 0,
            selectedText: selectedText,
            cursorPosition: position,
            surroundingCode: enhancedSurroundingCode,
            currentSubroutine: currentSubroutine,
            isInsideSubroutine: isInsideSubroutine,
            fullCode: fullCode,
            moduleMetadata: moduleMetadata,
            selectionRange: selection && !selection.isEmpty() ? {
                startLine: selection.startLineNumber,
                endLine: selection.endLineNumber,
                startColumn: selection.startColumn,
                endColumn: selection.endColumn
            } : null,
            // Additional context for better modification detection
            contextType: isInsideSubroutine ? 'modification' : (selectedText.length > 0 ? 'selection' : 'generation'),
            targetFunction: currentSubroutine ? currentSubroutine.name : null
        };
    }

    // Update line indicator display (ChatGPT style)
    function updateLineIndicator() {
        if (!state.editor || !state.activeModule) {
            hideLineIndicator();
            return;
        }

        const context = getEditorContext();
        
        if (context.hasSelection) {
            showLineIndicator(context);
        } else {
            hideLineIndicator();
        }
    }

    // Show simple line indicator
    function showLineIndicator(context) {
        const selectionRange = context.selectionRange;
        const lineCount = selectionRange.endLine - selectionRange.startLine + 1;
        
        // Simple text like ChatGPT: "Focused on lines 7-30" or "Focused on line 15"
        if (lineCount > 1) {
            elements.lineIndicator.textContent = `Focused on lines ${selectionRange.startLine}-${selectionRange.endLine}`;
        } else {
            elements.lineIndicator.textContent = `Focused on line ${selectionRange.startLine}`;
        }
        
        // Show indicator next to run button
        elements.lineIndicator.classList.remove("hidden");
    }

    // Hide line indicator
    function hideLineIndicator() {
        elements.lineIndicator.classList.add("hidden");
    }

    // Update run buttons in glyph margin
    function updateRunButtons() {
        if (!state.editor || !state.activeModule) return;

        // Clear existing decorations
        if (state.runButtonDecorations) {
            state.editor.deltaDecorations(state.runButtonDecorations, []);
        }

        const content = state.editor.getValue();
        const lines = content.split('\n');
        const decorations = [];

        // Find all Sub and Function declarations
        lines.forEach((line, index) => {
            const trimmed = line.trim();
            const subMatch = trimmed.match(/^(Private\s+|Public\s+)?Sub\s+(\w+)/i);
            const funcMatch = trimmed.match(/^(Private\s+|Public\s+)?Function\s+(\w+)/i);

            if (subMatch || funcMatch) {
                const name = subMatch ? subMatch[2] : funcMatch[2];
                const lineNumber = index + 1;

                decorations.push({
                    range: new monaco.Range(lineNumber, 1, lineNumber, 1),
                    options: {
                        isWholeLine: false,
                        glyphMarginClassName: 'run-button-glyph',
                        glyphMarginHoverMessage: { value: `Run ${name}` }
                    }
                });
            }
        });

        // Apply decorations
        state.runButtonDecorations = state.editor.deltaDecorations([], decorations);

        // Add click handler for glyph margin
        if (!state.glyphMarginClickHandler) {
            state.glyphMarginClickHandler = state.editor.onMouseDown((e) => {
                if (e.target.type === monaco.editor.MouseTargetType.GUTTER_GLYPH_MARGIN) {
                    const lineNumber = e.target.position.lineNumber;
                    const line = state.editor.getModel().getLineContent(lineNumber);
                    const subMatch = line.trim().match(/^(Private\s+|Public\s+)?Sub\s+(\w+)/i);
                    const funcMatch = line.trim().match(/^(Private\s+|Public\s+)?Function\s+(\w+)/i);

                    if (subMatch) {
                        runSubroutine(subMatch[2]);
                    } else if (funcMatch) {
                        // Functions can't be run directly, show a message
                        showChatMessage('system', `Functions cannot be run directly. Call "${funcMatch[2]}" from a Sub to execute it.`);
                    } else {
                        console.log('No Sub or Function found on this line');
                    }
                }
            });
        }
    }

    // Show message in chat panel
    function showChatMessage(type, content) {
        const message = {
            id: 'run-' + Date.now(),
            role: type === 'system' ? 'assistant' : 'user',
            content: content,
            timestamp: new Date().toISOString(),
            type: type
        };
        addMessage(message.role, content);
    }

    // Run a specific subroutine
    async function runSubroutine(subroutineName) {
        try {
            // Check for duplicate module names first
            console.log('Active module name:', state.activeModule.name);
            console.log('All module names:', state.modules.map(m => m.name));

            const duplicateModules = state.modules.filter(module =>
                module.name === state.activeModule.name
            );

            console.log('Found duplicates:', duplicateModules.length, 'modules with name:', state.activeModule.name);

            if (duplicateModules.length > 1) {
                showChatMessage('error', `Cannot run ${subroutineName}: There are ${duplicateModules.length} modules named "${state.activeModule.name}". Please rename the duplicate modules to have unique names.`);
                return;
            }

            // Save current module first
            if (state.activeModule && state.activeModule.isModified) {
                await saveModuleToExcel(state.activeModule);
            }


            const response = await fetch("http://localhost:5000/run-subroutine", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    module: state.activeModule.name,
                    subroutine: subroutineName
                })
            });

            const data = await response.json();

            if (response.ok && data.success) {
                showChatMessage('system', `✅ ${subroutineName} executed successfully`);
            } else {
                let errorMsg = data.message || data.error || `HTTP ${response.status}: ${response.statusText}`;

                // Check for common Excel macro errors and provide friendly messages
                if (errorMsg.includes("Cannot run the macro") && errorMsg.includes("The macro may not be available")) {
                    // Check if there are duplicate module/subroutine names
                    const duplicateModules = state.modules.filter(m => m.name === state.activeModule.name).length > 1;
                    const duplicateSubroutines = state.modules.some(m =>
                        m.subroutines && m.subroutines.filter(s => s.name === subroutineName).length > 0 &&
                        m !== state.activeModule &&
                        m.subroutines.some(s => s.name === subroutineName)
                    );

                    if (duplicateModules || duplicateSubroutines) {
                        errorMsg = `Multiple subroutines named "${subroutineName}" found. Please rename them to have unique names.`;
                    } else {
                        errorMsg = `Cannot run ${subroutineName}: The macro was not found. Make sure the subroutine exists and macros are enabled in Excel.`;
                    }
                } else if (errorMsg.includes("all macros may be disabled")) {
                    errorMsg = `Cannot run ${subroutineName}: Macros appear to be disabled in Excel. Please enable macros in Excel's Trust Center settings.`;
                } else if (errorMsg.includes("VBA project not accessible")) {
                    errorMsg = `Cannot run ${subroutineName}: VBA project access is disabled. Please enable "Trust access to VBA project object model" in Excel's Trust Center settings.`;
                }

                showChatMessage('error', errorMsg);
            }
        } catch (error) {
            console.error("Error running subroutine:", error);
            showChatMessage('error', `Error running ${subroutineName}: ${error.message}`);
        }
    }

    // Fetch sheet context from Excel
    async function getSheetContext() {
        try {
            const response = await fetch("http://localhost:5000/get-sheet-context");
            if (!response.ok) {
                console.warn("Could not fetch sheet context:", response.status);
                return null;
            }
            return await response.json();
        } catch (error) {
            console.warn("Error fetching sheet context:", error);
            return null;
        }
    }

    // Generate macro with AI - Updated for chat system
    async function generateMacro() {
        const prompt = elements.prompt.value.trim();
        
        if (!prompt) {
            addMessage('assistant', 'Please enter a message to get started.');
            elements.prompt.focus();
            return;
        }

        // Add user message to chat
        addMessage('user', prompt);
        
        // Clear input immediately for better UX
        elements.prompt.value = "";

        // Get editor context for enhanced AI generation
        const context = getEditorContext();
        
        // Loading state
        elements.generateBtn.disabled = true;
        
        // Add loading message
        const loadingMessage = {
            id: Date.now(),
            type: 'assistant',
            content: 'Generating VBA code...',
            timestamp: new Date()
        };
        renderMessage(loadingMessage);
        
        try {
            // Fetch sheet context for AI awareness
            const sheetContext = await getSheetContext();
            
            const requestBody = {
                prompt,
                context: {
                    hasSelection: context.hasSelection,
                    selectedText: context.selectedText,
                    surroundingCode: context.surroundingCode,
                    fullCode: context.fullCode,
                    moduleMetadata: context.moduleMetadata,
                    currentSubroutine: context.currentSubroutine,
                    selectionRange: context.selectionRange,
                    activeModule: state.activeModule?.name || null,
                    availableModules: getModulesList(),
                    sheetContext: sheetContext
                },
                conversation_history: state.conversation || []
            };

            const response = await fetch("http://localhost:5000/generate", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(requestBody)
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}`);
            }
            
            const data = await response.json();
            
            // Remove loading message
            const loadingDiv = document.querySelector(`[data-message-id="${loadingMessage.id}"]`);
            if (loadingDiv) {
                console.log('🗑️ Removing loading message ID:', loadingMessage.id);
                console.log('📊 DOM children before loading removal:', elements.chatMessages.children.length);
                console.log('📊 User messages before loading removal:', elements.chatMessages.querySelectorAll('.message.user').length);
                loadingDiv.remove();
                console.log('📊 DOM children after loading removal:', elements.chatMessages.children.length);
                console.log('📊 User messages after loading removal:', elements.chatMessages.querySelectorAll('.message.user').length);
                // Ensure chat stays scrolled to bottom after loading message removal
                scrollToBottom();
            }
            
            // Handle dual-agent responses
            console.log('📨 Processing AI response, type:', data.type);
            if (data.type === 'vba' && data.has_vba && data.vba_code) {
                // VBA Agent response with code
                const vbaCode = data.vba_code;
                const explanation = data.explanation;
                // Enhanced modification detection: cursor in subroutine OR selection OR AI classified as modification
                const isModification =
                    context.isInsideSubroutine ||  // Cursor is inside existing subroutine
                    context.contextType === 'modification' || // Context suggests modification
                    (context.hasSelection && data.context?.generationType === 'vba_modification') || // Original logic
                    data.context?.generationType === 'vba_modification'; // AI classified as modification

                // Show explanation in chat if present
                if (explanation && explanation.trim()) {
                    console.log('💬 Adding VBA explanation message');
                    addMessage('assistant', explanation);
                }

                // Automatically insert VBA code
                const insertionResult = await insertVBACode(vbaCode, context, isModification);
                
                // Show insertion status in chat
                if (insertionResult.success) {
                    addMessage('assistant', `✓ Code ${isModification ? 'modified' : 'inserted'} into ${insertionResult.moduleName}`);
                } else {
                    addMessage('assistant', `⚠ ${insertionResult.message}`);
                }
                
            } else if (data.type === 'conversation' || data.content) {
                // Chat Agent response (conversational)
                console.log('💬 Adding conversational response');
                const content = data.content || data.explanation;
                if (content && content.trim()) {
                    addMessage('assistant', content);
                } else {
                    addMessage('assistant', 'I understood your message, but I\'m not sure how to respond. Could you try rephrasing?');
                }
                
            } else if (data.explanation && data.explanation.trim()) {
                // Fallback: Legacy format support
                addMessage('assistant', data.explanation);
            } else {
                throw new Error("No response received");
            }
            
        } catch (error) {
            console.error("Generation error:", error);
            
            // Remove loading message
            const loadingDiv = document.querySelector(`[data-message-id="${loadingMessage.id}"]`);
            if (loadingDiv) {
                loadingDiv.remove();
                // Ensure chat stays scrolled to bottom after loading message removal
                scrollToBottom();
            }
            
            // Show error in chat
            let errorMessage = "Sorry, I encountered an error while generating the code.";
            if (error.message.includes("fetch")) {
                errorMessage = "Cannot connect to the backend server. Please make sure the Flask server is running.";
            } else {
                errorMessage = `Error: ${error.message}`;
            }
            
            addMessage('assistant', errorMessage);
        } finally {
            elements.generateBtn.disabled = false;
            console.log('✅ Request complete. Final DOM state:');
            console.log('📊 Total messages in DOM:', elements.chatMessages.children.length);
            console.log('📊 User messages in DOM:', elements.chatMessages.querySelectorAll('.message.user').length);
            console.log('📊 Assistant messages in DOM:', elements.chatMessages.querySelectorAll('.message.assistant').length);
        }
    }
    
    
    // Event listeners - with safety checks
    if (elements.generateBtn) {
        elements.generateBtn.addEventListener("click", generateMacro);
    } else {
        console.error("Generate button not found!");
    }
    
    if (elements.welcomeGenerate) {
        elements.welcomeGenerate.addEventListener("click", showNewModuleModal);
    } else {
        console.error("Welcome generate button not found!");
    }
    
    if (elements.addModuleBtn) {
        elements.addModuleBtn.addEventListener("click", showNewModuleModal);
    } else {
        console.error("Add module button not found!");
    }
    
    // Modal event listeners - with safety checks
    if (elements.createModuleOk) {
        elements.createModuleOk.addEventListener("click", createNewModule);
    } else {
        console.error("Create module OK button not found!");
    }
    
    if (elements.createModuleCancel) {
        elements.createModuleCancel.addEventListener("click", hideNewModuleModal);
    } else {
        console.error("Create module Cancel button not found!");
    }
    
    // Rename modal event listeners
    if (elements.renameModuleOk) {
        elements.renameModuleOk.addEventListener("click", renameModule);
    } else {
        console.error("Rename module OK button not found!");
    }
    
    if (elements.renameModuleCancel) {
        elements.renameModuleCancel.addEventListener("click", hideRenameModuleModal);
    } else {
        console.error("Rename module Cancel button not found!");
    }
    
    // Close modal when clicking outside - with safety checks
    if (elements.moduleNameModal) {
        elements.moduleNameModal.addEventListener("click", (e) => {
            if (e.target === elements.moduleNameModal) {
                hideNewModuleModal();
            }
        });
    }
    
    if (elements.renameModuleModal) {
        elements.renameModuleModal.addEventListener("click", (e) => {
            if (e.target === elements.renameModuleModal) {
                hideRenameModuleModal();
            }
        });
    }
    
    // Enter key in modal inputs - with safety checks
    if (elements.moduleNameInput) {
        elements.moduleNameInput.addEventListener("keydown", (e) => {
            if (e.key === "Enter") {
                createNewModule();
            } else if (e.key === "Escape") {
                hideNewModuleModal();
            }
        });
    }
    
    if (elements.renameModuleInput) {
        elements.renameModuleInput.addEventListener("keydown", (e) => {
            if (e.key === "Enter") {
                renameModule();
            } else if (e.key === "Escape") {
                hideRenameModuleModal();
            }
        });
    }
    
    // Auto-resize textarea function
    function autoResizeTextarea(textarea) {
        textarea.style.height = 'auto';
        
        // Calculate max height for 8 rows - text stops at bottom row boundary
        const lineHeight = 1.4 * 13; // 1.4em * 13px font-size = 18.2px per line
        const topPadding = 8; // Top padding
        const bottomPadding = 40; // Keep original total height, but text stops at bottom row
        const maxHeight = (lineHeight * 8) + topPadding + bottomPadding; // 8 full rows
        
        const newHeight = Math.min(textarea.scrollHeight, maxHeight);
        textarea.style.height = newHeight + 'px';
        
        // Enable scrolling when at max height (row 9+)
        if (textarea.scrollHeight > maxHeight) {
            textarea.style.overflowY = 'auto';
        } else {
            textarea.style.overflowY = 'hidden';
        }
    }

    // Auto-resize on input - with safety check
    if (elements.prompt) {
        elements.prompt.addEventListener('input', (e) => {
            autoResizeTextarea(e.target);
        });

        // Enter key in prompt - ChatGPT style behavior
        elements.prompt.addEventListener("keydown", (e) => {
            if (e.key === "Enter") {
                if (e.shiftKey) {
                    // Shift+Enter = new line (default behavior)
                    // Let it add the line, then resize
                    setTimeout(() => autoResizeTextarea(e.target), 0);
                    return;
                } else {
                    // Enter = send message
                    e.preventDefault();
                    generateMacro();
                    // Reset height after sending
                    setTimeout(() => {
                        e.target.style.height = 'auto';
                    }, 0);
                }
            }
        });
    } else {
        console.error("Prompt textarea not found!");
    }
    
    // Refresh modules button (hidden, for debugging)
    document.addEventListener("keydown", (e) => {
        if (e.key === "F5" && e.ctrlKey) {
            e.preventDefault();
            loadModules();
        }
    });
    
    // Search functionality - enhanced to search through subroutines and auto-expand
    function filterModules(searchTerm) {
        const searchLower = searchTerm.toLowerCase();

        state.modules.forEach(module => {
            const moduleElements = document.querySelectorAll('.module-item');
            let moduleElement = null;

            // Find the module element for this module
            moduleElements.forEach(el => {
                const nameSpan = el.querySelector('.module-name span:nth-child(2)');
                if (nameSpan && nameSpan.textContent === module.name) {
                    moduleElement = el;
                }
            });

            if (!moduleElement) return;

            const moduleName = module.name.toLowerCase();
            const moduleContainer = moduleElement.parentElement;
            const subroutineList = moduleContainer?.querySelector('.subroutine-list');

            // Check if module name matches
            let moduleMatches = moduleName.includes(searchLower);

            // Check if any subroutine matches
            let hasMatchingSubroutine = false;
            const matchingSubroutines = [];
            if (module.subroutines && module.subroutines.length > 0) {
                module.subroutines.forEach(sub => {
                    if (sub.name.toLowerCase().includes(searchLower)) {
                        hasMatchingSubroutine = true;
                        matchingSubroutines.push(sub.name);
                    }
                });
            }

            // Show module if it matches or has matching subroutines
            if (searchLower === '' || moduleMatches || hasMatchingSubroutine) {
                moduleElement.style.display = '';

                // Auto-expand if subroutine matches but module doesn't
                if (hasMatchingSubroutine && !moduleMatches && !module.expanded && searchLower !== '') {
                    module.expanded = true;
                    renderModuleList();
                    setTimeout(() => filterModules(searchTerm), 10); // Re-apply filter after re-render
                    return;
                }

                // Show/hide individual subroutines based on search
                if (subroutineList) {
                    const subItems = subroutineList.querySelectorAll('.subroutine-item');
                    subItems.forEach((subItem) => {
                        const subName = subItem.querySelector('span')?.textContent || '';
                        if (searchLower === '' || moduleMatches || subName.toLowerCase().includes(searchLower)) {
                            subItem.style.display = '';
                        } else {
                            subItem.style.display = 'none';
                        }
                    });
                }
            } else {
                moduleElement.style.display = 'none';
                if (subroutineList) {
                    const subItems = subroutineList.querySelectorAll('.subroutine-item');
                    subItems.forEach(subItem => subItem.style.display = 'none');
                }
            }
        });
    }
    
    // Sidebar toggle functionality
    function toggleSidebar() {
        const sidebar = elements.sidebar;
        const isCollapsed = sidebar.classList.toggle('collapsed');
        
        // Keep the same ChatGPT-style icon in both states
        elements.toggleSidebarBtn.title = isCollapsed ? 'Expand Explorer' : 'Collapse Explorer';
    }
    
    // Search input event listener
    if (elements.searchMacros) {
        elements.searchMacros.addEventListener('input', (e) => {
            filterModules(e.target.value);
        });
        
        elements.searchMacros.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                e.target.value = '';
                filterModules('');
            }
        });
    }
    
    // Search icon click handler for collapsed state
    document.addEventListener('click', (e) => {
        if (e.target.closest('.search-icon') && elements.sidebar.classList.contains('collapsed')) {
            // Expand sidebar and focus search input
            toggleSidebar();
            setTimeout(() => {
                elements.searchMacros.focus();
            }, 100);
        }
    });
    
    // Sidebar toggle event listener
    if (elements.toggleSidebarBtn) {
        elements.toggleSidebarBtn.addEventListener('click', toggleSidebar);
    }
    
    // Chat panel toggle event listener
    if (elements.toggleChatBtn) {
        elements.toggleChatBtn.addEventListener('click', toggleChatPanel);
    }
    
    // Initialize chat resize functionality
    initializeChatResize();
    
    // Initialize sidebar resize functionality
    initializeSidebarResize();
    
    // Initialize
    initializeMonacoEditor();
    loadModules();
    
});

// Run Active Macro function (called from ribbon button)
// This function must be in global scope for Office ExecuteFunction to access it
window.runActiveMacro = async function(event) {
    try {
        // Check if there's an active module
        if (!state.activeModule || !state.activeModule.content.trim()) {
            // Show notification
            await Excel.run(async (context) => {
                context.application.showNotification("No Active Macro", "Please select a module with VBA code to run.");
                await context.sync();
            });
            event.completed();
            return;
        }

        // Extract the first subroutine from the active module
        const vbaCode = state.activeModule.content;
        const subMatch = vbaCode.match(/Sub\s+(\w+)/i);
        
        if (!subMatch) {
            await Excel.run(async (context) => {
                context.application.showNotification("No Subroutine Found", "The active module doesn't contain a valid Sub routine to execute.");
                await context.sync();
            });
            event.completed();
            return;
        }

        const subroutineName = subMatch[1];

        // Run the macro using Excel's VBA execution
        await Excel.run(async (context) => {
            try {
                // First ensure the module is saved to Excel
                await saveModuleToExcel(state.activeModule);
                
                // Run the macro by calling the subroutine
                const workbook = context.workbook;
                workbook.evaluate(`Application.Run("${state.activeModule.name}.${subroutineName}")`);
                await context.sync();
                
                // Show success notification
                context.application.showNotification("Macro Executed", `Successfully ran ${subroutineName} from ${state.activeModule.name}`);
                await context.sync();
                
            } catch (runError) {
                console.error("Macro execution error:", runError);
                context.application.showNotification("Execution Error", `Failed to run macro: ${runError.message || "Unknown error"}`);
                await context.sync();
            }
        });

    } catch (error) {
        console.error("runActiveMacro error:", error);
        await Excel.run(async (context) => {
            context.application.showNotification("Error", `An error occurred: ${error.message || "Unknown error"}`);
            await context.sync();
        });
    }
    
    // Signal that the function has completed
    event.completed();
};