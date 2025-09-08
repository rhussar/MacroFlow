// MacroFlow AI - VSCode Style Interface with Real VBA Integration
Office.onReady(() => {
    
    // Global state
    const state = {
        modules: [],
        activeModule: null,
        editor: null,
        tabs: new Map(),
        contextMenu: null
    };
    
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
        runBtn: document.getElementById("runBtn"),
        moduleNameModal: document.getElementById("moduleNameModal"),
        moduleNameInput: document.getElementById("moduleNameInput"),
        createModuleOk: document.getElementById("createModuleOk"),
        createModuleCancel: document.getElementById("createModuleCancel"),
        renameModuleModal: document.getElementById("renameModuleModal"),
        renameModuleInput: document.getElementById("renameModuleInput"),
        renameModuleOk: document.getElementById("renameModuleOk"),
        renameModuleCancel: document.getElementById("renameModuleCancel"),
        // Line indicator element
        lineIndicator: document.getElementById("lineIndicator")
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
                selectOnLineNumbers: true
            });
            
            // Listen for content changes
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
                    <div class="expand-icon" ${hasSubroutines ? '' : 'style="visibility: hidden"'}>▶</div>
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
    
    // Find which subroutine the cursor is currently in
    function getCurrentSubroutine() {
        if (!state.editor || !state.activeModule) return null;
        
        const position = state.editor.getPosition();
        if (!position) return null;
        
        const currentLine = position.lineNumber;
        const content = state.editor.getValue();
        const lines = content.split('\n');
        
        let currentSub = null;
        
        // Look backwards from current line to find the subroutine we're in
        for (let i = currentLine - 1; i >= 0; i--) {
            const line = lines[i].trim().toLowerCase();
            
            // If we hit an End Sub/Function, we're not in a subroutine
            if (line.includes('end sub') || line.includes('end function')) {
                break;
            }
            
            // If we find a Sub or Function declaration
            if ((line.includes('sub ') || line.includes('function ')) && 
                !line.startsWith("'")) {
                // Extract the subroutine name
                const words = lines[i].trim().split(/\s+/);
                for (let j = 0; j < words.length; j++) {
                    if (words[j].toLowerCase() === 'sub' || words[j].toLowerCase() === 'function') {
                        if (j + 1 < words.length) {
                            const subName = words[j + 1].split('(')[0];
                            currentSub = subName;
                            break;
                        }
                    }
                }
                break;
            }
        }
        
        return currentSub;
    }
    
    // Run the current subroutine
    async function runCurrentSubroutine() {
        if (!state.activeModule) {
            console.error("Please open a module first");
            return;
        }
        
        const currentSub = getCurrentSubroutine();
        if (!currentSub) {
            console.error("Place cursor inside a subroutine to run it");
            return;
        }
        
        try {
            elements.runBtn.disabled = true;
            
            const response = await fetch("http://localhost:5000/run-subroutine", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    module: state.activeModule.name,
                    subroutine: currentSub
                })
            });
            
            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.error || `HTTP ${response.status}`);
            }
            
            const data = await response.json();
            
            if (data.success) {
            } else {
                throw new Error(data.error || "Execution failed");
            }
            
        } catch (error) {
            console.error("Error running subroutine:", error);
            if (error.message.includes("fetch")) {
                console.error("Cannot connect to backend");
            } else {
                console.error(`Error: ${error.message}`);
            }
        } finally {
            elements.runBtn.disabled = false;
        }
    }
    
    // Open module in editor
    function openModule(module) {
        if (state.activeModule === module) return;
        
        state.activeModule = module;
        
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
        
    }
    
    // Create editor tab
    function createTab(module) {
        const tab = document.createElement("div");
        tab.className = "tab";
        tab.dataset.tab = module.name;
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
    
    // Update tab modified state
    function updateTabModifiedState(moduleName) {
        const tab = state.tabs.get(moduleName);
        const module = state.modules.find(m => m.name === moduleName);
        if (tab && module) {
            const span = tab.querySelector("span");
            span.textContent = module.name + (module.isModified ? " ●" : "");
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
        
        // Set default name
        elements.moduleNameInput.value = `Module${state.modules.length + 1}`;
        
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
    ' Generated by MacroFlow AI
    ' Add your VBA code here
    
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
                    tab.querySelector("span").textContent = newName + (newModule.isModified ? " ●" : "");
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

        return {
            hasSelection: selectedText.length > 0,
            selectedText: selectedText,
            cursorPosition: position,
            surroundingCode: surroundingCode,
            currentSubroutine: getCurrentSubroutine(),
            fullCode: model.getValue(),
            selectionRange: selection && !selection.isEmpty() ? {
                startLine: selection.startLineNumber,
                endLine: selection.endLineNumber,
                startColumn: selection.startColumn,
                endColumn: selection.endColumn
            } : null
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

    // Generate macro with AI
    async function generateMacro() {
        const prompt = elements.prompt.value.trim();
        
        if (!prompt) {
            console.error("Please enter a prompt");
            elements.prompt.focus();
            return;
        }

        // Get editor context for enhanced AI generation
        const context = getEditorContext();
        
        // Loading state
        elements.generateBtn.disabled = true;
        
        // Log context for debugging
        
        try {
            const requestBody = {
                prompt,
                context: {
                    hasSelection: context.hasSelection,
                    selectedText: context.selectedText,
                    surroundingCode: context.surroundingCode,
                    currentSubroutine: context.currentSubroutine,
                    selectionRange: context.selectionRange,
                    activeModule: state.activeModule?.name || null
                }
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
            
            if (data.macro) {
                const generatedCode = data.macro;
                const isModification = context.hasSelection && data.context?.generationType === "modification";

                // If no active module, create a new one for the generated code
                if (!state.activeModule) {
                    const moduleName = `Generated${Date.now()}`;
                    try {
                        const createResponse = await fetch("http://localhost:5000/create-module", {
                            method: "POST",
                            headers: { "Content-Type": "application/json" },
                            body: JSON.stringify({
                                name: moduleName,
                                content: generatedCode
                            })
                        });
                        
                        if (createResponse.ok) {
                            const newModule = {
                                name: moduleName,
                                content: generatedCode,
                                isModified: false
                            };
                            state.modules.push(newModule);
                            renderModuleList();
                            openModule(newModule);
                            elements.prompt.value = ""; // Clear prompt
                            return;
                        } else {
                            throw new Error("Failed to create module for generated code");
                        }
                    } catch (createError) {
                        console.error("Please select or create a module first");
                        return;
                    }
                }

                // Handle code insertion based on context
                if (state.editor) {
                    if (isModification && context.selectionRange) {
                        // Replace selected text with generated code
                        const selection = new monaco.Selection(
                            context.selectionRange.startLine,
                            context.selectionRange.startColumn,
                            context.selectionRange.endLine,
                            context.selectionRange.endColumn
                        );
                        state.editor.executeEdits("ai-generation", [{
                            range: selection,
                            text: generatedCode
                        }]);
                        
                        // Update module content
                        state.activeModule.content = state.editor.getValue();
                        
                    } else if (context.cursorPosition) {
                        // Insert generated code at cursor position
                        state.editor.executeEdits("ai-generation", [{
                            range: new monaco.Range(
                                context.cursorPosition.lineNumber,
                                context.cursorPosition.column,
                                context.cursorPosition.lineNumber,
                                context.cursorPosition.column
                            ),
                            text: "\n" + generatedCode + "\n"
                        }]);
                        
                        // Update module content
                        state.activeModule.content = state.editor.getValue();
                        
                    } else {
                        // Fallback: Replace entire content
                        state.editor.setValue(generatedCode);
                        state.activeModule.content = generatedCode;
                    }
                } else {
                    // No editor available, update module content directly
                    state.activeModule.content = generatedCode;
                }
                
                state.activeModule.isModified = true;
                
                // Update tab and sidebar
                updateTabModifiedState(state.activeModule.name);
                
                // Auto-save to Excel
                await saveModuleToExcel(state.activeModule);
                
                const actionType = isModification ? "modified" : "generated";
                elements.prompt.value = ""; // Clear prompt
            } else {
                throw new Error("No macro received");
            }
            
        } catch (error) {
            console.error("Generation error:", error);
            
            if (error.message.includes("fetch")) {
                console.error("Cannot connect to backend. Is Flask server running?");
            } else {
                console.error(`Error: ${error.message}`);
            }
        } finally {
            elements.generateBtn.disabled = false;
        }
    }
    
    
    // Event listeners
    elements.generateBtn.addEventListener("click", generateMacro);
    elements.welcomeGenerate.addEventListener("click", generateMacro);
    
    // Run button event listener
    if (elements.runBtn) {
        elements.runBtn.addEventListener("click", runCurrentSubroutine);
    } else {
        console.error("Run button not found!");
    }

    
    // Debug: Check if addModuleBtn exists
    if (elements.addModuleBtn) {
        elements.addModuleBtn.addEventListener("click", showNewModuleModal);
    } else {
        console.error("Add module button not found!");
    }
    
    // Modal event listeners
    elements.createModuleOk.addEventListener("click", createNewModule);
    elements.createModuleCancel.addEventListener("click", hideNewModuleModal);
    
    // Rename modal event listeners
    elements.renameModuleOk.addEventListener("click", renameModule);
    elements.renameModuleCancel.addEventListener("click", hideRenameModuleModal);
    
    // Close modal when clicking outside
    elements.moduleNameModal.addEventListener("click", (e) => {
        if (e.target === elements.moduleNameModal) {
            hideNewModuleModal();
        }
    });
    
    elements.renameModuleModal.addEventListener("click", (e) => {
        if (e.target === elements.renameModuleModal) {
            hideRenameModuleModal();
        }
    });
    
    // Enter key in modal inputs
    elements.moduleNameInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
            createNewModule();
        } else if (e.key === "Escape") {
            hideNewModuleModal();
        }
    });
    
    elements.renameModuleInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
            renameModule();
        } else if (e.key === "Escape") {
            hideRenameModuleModal();
        }
    });
    
    // Enter key in prompt
    elements.prompt.addEventListener("keydown", (e) => {
        if (e.key === "Enter" && e.ctrlKey) {
            generateMacro();
        }
    });
    
    // Refresh modules button (hidden, for debugging)
    document.addEventListener("keydown", (e) => {
        if (e.key === "F5" && e.ctrlKey) {
            e.preventDefault();
            loadModules();
        }
    });
    
    // Initialize
    initializeMonacoEditor();
    loadModules();
    
});