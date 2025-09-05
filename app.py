from flask import Flask, request, jsonify
from flask_cors import CORS
from openai import OpenAI
import csv
import datetime
import os
import win32com.client
import pythoncom

app = Flask(__name__)
CORS(app)  # ✅ Allow requests from localhost:3000 (Add-in)

# OpenAI client
api_key = os.getenv("OPENAI_API_KEY") or "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"
client = OpenAI(api_key=api_key)

# Paths - Make logging optional and safe
BASE_FOLDER = os.path.join(os.getcwd(), "logs")
LOG_FILE = os.path.join(BASE_FOLDER, "macro_log.csv")

def log_macro(prompt, macro_code):
    """Log macro generation - fails silently if directory doesn't exist"""
    try:
        # Create logs directory if it doesn't exist
        os.makedirs(BASE_FOLDER, exist_ok=True)
        
        with open(LOG_FILE, mode='a', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow([datetime.datetime.now(), prompt, macro_code])
        print(f"Logged macro to: {LOG_FILE}")
    except Exception as e:
        print(f"Warning: Could not log macro (continuing anyway): {e}")

def generate_vba(prompt):
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "You are an assistant that writes clean VBA macros for Excel. Return ONLY the VBA code without any explanations, markdown formatting, or additional text. Start directly with 'Sub' and end with 'End Sub'."},
            {"role": "user", "content": prompt}
        ],
        temperature=0
    )
    return response.choices[0].message.content

def extract_vba_code(text):
    """Extract only VBA code from the response, removing any explanatory text."""
    lines = text.split('\n')
    vba_lines = []
    in_vba_block = False
    
    for line in lines:
        # Start collecting when we see 'Sub' or 'Function'
        if line.strip().startswith(('Sub ', 'Function ', 'Private Sub ', 'Public Sub ', 'Private Function ', 'Public Function ')):
            in_vba_block = True
        
        if in_vba_block:
            vba_lines.append(line)
            
        # Stop collecting when we see 'End Sub' or 'End Function'
        if line.strip() in ['End Sub', 'End Function']:
            break
    
    # If we found VBA code, return it; otherwise return the original text
    if vba_lines:
        return '\n'.join(vba_lines)
    return text

def inject_vba_to_excel(macro_code):
    """Inject VBA code into the active Excel workbook."""
    try:
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
        print(f"Attempting to inject macro: {len(macro_code)} characters")
        
        # Connect to Excel - try GetObject first, then Dispatch
        try:
            excel = win32com.client.GetObject(None, "Excel.Application")
            print("Connected to existing Excel application")
        except:
            excel = win32com.client.Dispatch("Excel.Application")
            print("Created new Excel application")
        
        # Make Excel visible and bring to front
        excel.Visible = True
        excel.WindowState = 1  # xlNormal
        print(f"Excel is visible: {excel.Visible}")
        
        # Get active workbook or create new one
        wb = excel.ActiveWorkbook
        if not wb:
            print("No active workbook, creating new one")
            wb = excel.Workbooks.Add()
        else:
            print(f"Using active workbook: {wb.Name}")
        
        # Ensure the workbook is properly initialized for VBA
        try:
            if not wb.Path:
                print("Workbook has no path - this may cause VBA injection issues")
                print("Please save your workbook as .xlsm (Excel Macro-Enabled Workbook) for best results")
        except Exception as save_error:
            print(f"Could not check workbook path: {save_error}")
        
        # Check if VBA project is accessible
        try:
            vbproj = wb.VBProject
            print("Successfully accessed VBProject")
        except Exception as vb_error:
            return False, f"VBA project not accessible. Make sure 'Trust access to the VBA project object model' is enabled in Excel Trust Center. Error: {str(vb_error)}"
        
        # Add new module
        module = vbproj.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        print(f"Created new module: {module.Name}")
        
        # Add the macro code to the module
        module.CodeModule.AddFromString(macro_code)
        print("Successfully added macro code to module")
        
        # Verify the code was actually added
        actual_code = module.CodeModule.Lines(1, module.CodeModule.CountOfLines)
        print(f"Module now contains {module.CodeModule.CountOfLines} lines of code")
        print(f"First few lines: {actual_code[:200]}...")
        
        # List all modules to verify
        print("Current modules in workbook:")
        for i in range(1, vbproj.VBComponents.Count + 1):
            comp = vbproj.VBComponents(i)
            if comp.Type == 1:  # Standard module
                print(f"  - {comp.Name} (Type: Standard Module, Lines: {comp.CodeModule.CountOfLines})")
        
        return True, f"Macro injected successfully into module '{module.Name}' in workbook '{wb.Name}'. Check the VBA editor (Alt+F11) and look for the 'Modules' folder."
        
    except Exception as e:
        print(f"Error in inject_vba_to_excel: {str(e)}")
        return False, f"Failed to inject macro: {str(e)}"
    finally:
        # Clean up COM
        try:
            pythoncom.CoUninitialize()
        except:
            pass

@app.route("/generate", methods=["POST"])
def generate():
    data = request.json
    prompt = data.get("prompt", "")
    if not prompt:
        return jsonify({"error": "No prompt provided"}), 400

    macro = generate_vba(prompt)
    # Clean up the response to extract only VBA code
    clean_macro = extract_vba_code(macro)
    log_macro(prompt, clean_macro)
    return jsonify({"macro": clean_macro})

@app.route("/inject", methods=["POST"])
def inject():
    """Inject VBA code into Excel."""
    try:
        data = request.json
        if not data:
            return jsonify({"error": "No JSON data provided"}), 400
            
        macro_code = data.get("macro", "")
        
        if not macro_code:
            return jsonify({"error": "No macro code provided"}), 400
        
        print(f"Received inject request with macro length: {len(macro_code)}")
        success, message = inject_vba_to_excel(macro_code)
        
        if success:
            return jsonify({"success": True, "message": message})
        else:
            return jsonify({"success": False, "error": message}), 500
            
    except Exception as e:
        print(f"Error in inject endpoint: {str(e)}")
        return jsonify({"success": False, "error": f"Server error: {str(e)}"}), 500

@app.route("/list-modules", methods=["GET"])
def list_modules():
    """List all modules in the active workbook."""
    try:
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
        try:
            excel = win32com.client.GetObject(None, "Excel.Application")
        except:
            excel = win32com.client.Dispatch("Excel.Application")
        
        wb = excel.ActiveWorkbook
        if not wb:
            return jsonify({"error": "No active workbook"}), 400
        
        vbproj = wb.VBProject
        modules = []
        
        for i in range(1, vbproj.VBComponents.Count + 1):
            comp = vbproj.VBComponents(i)
            if comp.Type == 1:  # Standard module
                modules.append({
                    "name": comp.Name,
                    "type": "Standard Module",
                    "line_count": comp.CodeModule.CountOfLines
                })
        
        return jsonify({"workbook": wb.Name, "modules": modules})
        
    except Exception as e:
        return jsonify({"error": f"Failed to list modules: {str(e)}"}), 500
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

if __name__ == "__main__":
    app.run(host="localhost", port=5000, debug=True)


