Office.onReady(() => {
    console.log("MacroFlow AI loaded");
    
    // Get DOM elements
    const promptTextarea = document.getElementById("prompt");
    const generateButton = document.getElementById("generateMacro");
    const copyButton = document.getElementById("copyMacro");
    const insertButton = document.getElementById("insertMacro");
    const macroPreview = document.getElementById("macroPreview");
    const statusMessage = document.getElementById("statusMessage");
    
    // Status message helper
    function showStatus(message, type = "info") {
        statusMessage.textContent = message;
        statusMessage.className = `status-message ${type}`;
    }
    
    // Generate Macro
    generateButton.onclick = async () => {
        const prompt = promptTextarea.value.trim();
        
        if (!prompt) {
            showStatus("Please enter a prompt", "error");
            return;
        }
        
        // Loading state
        macroPreview.value = "";
        copyButton.disabled = true;
        insertButton.disabled = true;
        generateButton.disabled = true;
        generateButton.textContent = "Generating...";
        showStatus("Generating macro...", "info");
        
        try {
            const response = await fetch("http://localhost:5000/generate", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ prompt })
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}`);
            }
            
            const data = await response.json();
            
            if (data.macro) {
                macroPreview.value = data.macro;
                copyButton.disabled = false;
                insertButton.disabled = false;
                showStatus("Macro generated successfully!", "success");
            } else {
                throw new Error("No macro received");
            }
            
        } catch (error) {
            console.error("Generation error:", error);
            
            if (error.message.includes("fetch")) {
                showStatus("Cannot connect to backend. Is Flask server running?", "error");
            } else {
                showStatus(`Error: ${error.message}`, "error");
            }
            
            macroPreview.value = "// Error generating macro";
        } finally {
            generateButton.disabled = false;
            generateButton.textContent = "Generate";
        }
    };
    
    // Copy Macro
    copyButton.onclick = async () => {
        const macroText = macroPreview.value.trim();
        
        if (!macroText) {
            showStatus("No macro to copy", "error");
            return;
        }
        
        try {
            await navigator.clipboard.writeText(macroText);
            showStatus("Copied to clipboard!", "success");
        } catch (error) {
            showStatus("Failed to copy", "error");
        }
    };
    
    // Insert Macro
    insertButton.onclick = async () => {
        const macroText = macroPreview.value.trim();
        
        if (!macroText) {
            showStatus("No macro to insert", "error");
            return;
        }
        
        insertButton.disabled = true;
        insertButton.textContent = "Inserting...";
        showStatus("Inserting into Excel...", "info");
        
        try {
            const response = await fetch("http://localhost:5000/inject", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ macro: macroText })
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}`);
            }
            
            const data = await response.json();
            
            if (data.success) {
                showStatus(data.message, "success");
            } else {
                throw new Error(data.error || "Injection failed");
            }
            
        } catch (error) {
            console.error("Injection error:", error);
            
            if (error.message.includes("fetch")) {
                showStatus("Cannot connect to backend", "error");
            } else {
                showStatus(`Error: ${error.message}`, "error");
            }
        } finally {
            insertButton.disabled = false;
            insertButton.textContent = "Insert into Excel";
        }
    };
    
    // Enable buttons when macro preview has content
    macroPreview.addEventListener("input", () => {
        const hasContent = macroPreview.value.trim();
        copyButton.disabled = !hasContent;
        insertButton.disabled = !hasContent;
    });
    
    console.log("MacroFlow AI ready");
});