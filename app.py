from flask import Flask, request, jsonify
from flask_cors import CORS
from openai import OpenAI
import os
import win32com.client
import pythoncom

app = Flask(__name__)
CORS(app)  # ✅ Allow requests from localhost:3000 (Add-in)

# OpenAI client
api_key = os.getenv("OPENAI_API_KEY") or "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"
client = OpenAI(api_key=api_key)


def generate_vba(prompt, context=None):
    # Build enhanced system prompt based on context
    system_prompt = "You are an expert VBA assistant for Excel. Return ONLY the VBA code without any explanations, markdown formatting, or additional text."
    
    # Build user message with context
    user_message = prompt
    
    if context:
        if context.get("hasSelection") and context.get("selectedText"):
            # User has selected code - focus on modification/enhancement
            system_prompt += " The user has selected specific code that they want you to work with. You can modify, enhance, or replace the selected code based on their request. Pay attention to the existing code structure and style."
            user_message += f"\n\nSELECTED CODE TO WORK WITH:\n{context['selectedText']}"
            
            if context.get("surroundingCode"):
                user_message += f"\n\nSURROUNDING CONTEXT (for reference):\n{context['surroundingCode']}"
        
        elif context.get("surroundingCode"):
            # User is in a module with existing code - be context-aware
            system_prompt += " The user is working in an existing VBA module. Generate code that integrates well with the existing code structure, follows the same naming conventions, and doesn't conflict with existing subroutines."
            user_message += f"\n\nEXISTING CODE CONTEXT:\n{context['surroundingCode']}"
        
        if context.get("currentSubroutine"):
            user_message += f"\n\nCURRENT SUBROUTINE: {context['currentSubroutine']}"
        
        if context.get("activeModule"):
            user_message += f"\n\nACTIVE MODULE: {context['activeModule']}"
    
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message}
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

def parse_subroutines_from_vba(content):
    """Parse subroutines and functions from VBA code."""
    if not content:
        return []
    
    subroutines = []
    lines = content.split('\n')
    
    for i, line in enumerate(lines):
        stripped = line.strip()
        
        # Skip comments and empty lines
        if not stripped or stripped.startswith("'"):
            continue
            
        # Look for Sub or Function declarations
        if any(keyword in stripped.upper() for keyword in ['SUB ', 'FUNCTION ']):
            # More flexible parsing
            upper_line = stripped.upper()
            
            # Determine type
            if 'FUNCTION' in upper_line:
                sub_type = "Function"
            elif 'SUB' in upper_line:
                sub_type = "Sub"
            else:
                continue
            
            try:
                # Extract name - look for the word after SUB or FUNCTION
                words = stripped.split()
                name_found = False
                sub_name = ""
                
                for word in words:
                    if name_found:
                        # This should be the subroutine name
                        sub_name = word.split('(')[0]  # Remove parameters
                        break
                    if word.upper() in ['SUB', 'FUNCTION']:
                        name_found = True
                
                if sub_name and sub_name.upper() not in ['SUB', 'FUNCTION', 'PRIVATE', 'PUBLIC']:
                    subroutines.append({
                        "name": sub_name,
                        "type": sub_type,
                        "line": stripped,
                        "line_number": i + 1
                    })
            except Exception as e:
                continue
    
    return subroutines

def inject_vba_to_excel(macro_code):
    """Inject VBA code into the active Excel workbook."""
    try:
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
        
        # Connect to Excel - try GetObject first, then Dispatch
        try:
            excel = win32com.client.GetObject(None, "Excel.Application")
        except:
            excel = win32com.client.Dispatch("Excel.Application")
        
        # Make Excel visible and bring to front
        excel.Visible = True
        excel.WindowState = 1  # xlNormal
        
        # Get active workbook or create new one
        wb = excel.ActiveWorkbook
        if not wb:
            wb = excel.Workbooks.Add()
        
        # Ensure the workbook is properly initialized for VBA
        try:
            if not wb.Path:
                pass  # Workbook is not saved, continue anyway
        except Exception as save_error:
            pass  # Continue even if we can't check the path
        
        # Check if VBA project is accessible
        try:
            vbproj = wb.VBProject
        except Exception as vb_error:
            return False, f"VBA project not accessible. Make sure 'Trust access to the VBA project object model' is enabled in Excel Trust Center. Error: {str(vb_error)}"
        
        # Add new module
        module = vbproj.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        
        # Add the macro code to the module
        module.CodeModule.AddFromString(macro_code)
        
        # Verify the code was actually added
        actual_code = module.CodeModule.Lines(1, module.CodeModule.CountOfLines)
        
        
        return True, f"Macro injected successfully into module '{module.Name}' in workbook '{wb.Name}'. Check the VBA editor (Alt+F11) and look for the 'Modules' folder."
        
    except Exception as e:
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
    context = data.get("context", None)
    
    if not prompt:
        return jsonify({"error": "No prompt provided"}), 400

    
    # Generate VBA with context awareness
    macro = generate_vba(prompt, context)
    
    # Clean up the response to extract only VBA code
    clean_macro = extract_vba_code(macro)
    
    
    # Return enhanced response with context information
    response_data = {
        "macro": clean_macro,
        "context": {
            "hasSelection": context.get("hasSelection", False) if context else False,
            "generationType": "modification" if (context and context.get("hasSelection")) else "generation"
        }
    }
    
    return jsonify(response_data)

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
        
        success, message = inject_vba_to_excel(macro_code)
        
        if success:
            return jsonify({"success": True, "message": message})
        else:
            return jsonify({"success": False, "error": message}), 500
            
    except Exception as e:
        return jsonify({"success": False, "error": f"Server error: {str(e)}"}), 500

@app.route("/list-modules", methods=["GET"])
def list_modules():
    """List all modules in the active workbook with their content."""
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
        
        # Check if VBA project is accessible
        try:
            vbproj = wb.VBProject
        except Exception as vb_error:
            return jsonify({"error": f"VBA project not accessible. Make sure 'Trust access to the VBA project object model' is enabled in Excel Trust Center. Error: {str(vb_error)}"}), 500
        
        modules = []
        
        for i in range(1, vbproj.VBComponents.Count + 1):
            comp = vbproj.VBComponents(i)
            if comp.Type == 1:  # Standard module
                try:
                    # Get the content of the module
                    if comp.CodeModule.CountOfLines > 0:
                        content = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
                    else:
                        content = ""
                    
                    # Parse subroutines from the content
                    subroutines = parse_subroutines_from_vba(content)
                    
                    modules.append({
                        "name": comp.Name,
                        "content": content,
                        "line_count": comp.CodeModule.CountOfLines,
                        "subroutines": subroutines,
                        "isModified": False  # Frontend will track modifications
                    })
                except Exception as module_error:
                    modules.append({
                        "name": comp.Name,
                        "content": f"' Error reading module: {str(module_error)}",
                        "line_count": 0,
                        "subroutines": [],
                        "isModified": False
                    })
        
        return jsonify({"workbook": wb.Name, "modules": modules})
        
    except Exception as e:
        return jsonify({"error": f"Failed to list modules: {str(e)}"}), 500
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

@app.route("/create-module", methods=["POST"])
def create_module():
    """Create a new VBA module in the active workbook."""
    try:
        data = request.json
        if not data:
            return jsonify({"error": "No JSON data provided"}), 400
            
        module_name = data.get("name", "")
        initial_content = data.get("content", "")
        
        if not module_name:
            return jsonify({"error": "Module name is required"}), 400
        
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
        try:
            excel = win32com.client.GetObject(None, "Excel.Application")
        except:
            excel = win32com.client.Dispatch("Excel.Application")
        
        wb = excel.ActiveWorkbook
        if not wb:
            return jsonify({"error": "No active workbook"}), 400
        
        # Check if VBA project is accessible
        try:
            vbproj = wb.VBProject
        except Exception as vb_error:
            return jsonify({"error": f"VBA project not accessible. Make sure 'Trust access to the VBA project object model' is enabled in Excel Trust Center. Error: {str(vb_error)}"}), 500
        
        # Check if module name already exists
        for i in range(1, vbproj.VBComponents.Count + 1):
            comp = vbproj.VBComponents(i)
            if comp.Name == module_name:
                return jsonify({"error": f"Module '{module_name}' already exists"}), 400
        
        # Create new module
        new_module = vbproj.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        new_module.Name = module_name
        
        # Add initial content if provided
        if initial_content:
            new_module.CodeModule.AddFromString(initial_content)
        
        return jsonify({"success": True, "message": f"Module '{module_name}' created successfully"})
        
    except Exception as e:
        return jsonify({"error": f"Failed to create module: {str(e)}"}), 500
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

@app.route("/update-module", methods=["POST"])
def update_module():
    """Update the content of an existing VBA module."""
    try:
        data = request.json
        if not data:
            return jsonify({"error": "No JSON data provided"}), 400
            
        module_name = data.get("name", "")
        new_content = data.get("content", "")
        
        if not module_name:
            return jsonify({"error": "Module name is required"}), 400
        
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
        try:
            excel = win32com.client.GetObject(None, "Excel.Application")
        except:
            excel = win32com.client.Dispatch("Excel.Application")
        
        wb = excel.ActiveWorkbook
        if not wb:
            return jsonify({"error": "No active workbook"}), 400
        
        # Check if VBA project is accessible
        try:
            vbproj = wb.VBProject
        except Exception as vb_error:
            return jsonify({"error": f"VBA project not accessible. Make sure 'Trust access to the VBA project object model' is enabled in Excel Trust Center. Error: {str(vb_error)}"}), 500
        
        # Find the module
        target_module = None
        for i in range(1, vbproj.VBComponents.Count + 1):
            comp = vbproj.VBComponents(i)
            if comp.Name == module_name and comp.Type == 1:  # Standard module
                target_module = comp
                break
        
        if not target_module:
            return jsonify({"error": f"Module '{module_name}' not found"}), 404
        
        # Replace the module content
        code_module = target_module.CodeModule
        
        # Clear existing content
        if code_module.CountOfLines > 0:
            code_module.DeleteLines(1, code_module.CountOfLines)
        
        # Add new content
        if new_content:
            code_module.AddFromString(new_content)
        
        return jsonify({"success": True, "message": f"Module '{module_name}' updated successfully"})
        
    except Exception as e:
        return jsonify({"error": f"Failed to update module: {str(e)}"}), 500
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

@app.route("/delete-module", methods=["POST"])
def delete_module():
    """Delete a VBA module from the active workbook."""
    try:
        data = request.json
        if not data:
            return jsonify({"error": "No JSON data provided"}), 400
            
        module_name = data.get("name", "")
        
        if not module_name:
            return jsonify({"error": "Module name is required"}), 400
        
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
        try:
            excel = win32com.client.GetObject(None, "Excel.Application")
        except:
            excel = win32com.client.Dispatch("Excel.Application")
        
        wb = excel.ActiveWorkbook
        if not wb:
            return jsonify({"error": "No active workbook"}), 400
        
        # Check if VBA project is accessible
        try:
            vbproj = wb.VBProject
        except Exception as vb_error:
            return jsonify({"error": f"VBA project not accessible. Make sure 'Trust access to the VBA project object model' is enabled in Excel Trust Center. Error: {str(vb_error)}"}), 500
        
        # Find and delete the module
        for i in range(1, vbproj.VBComponents.Count + 1):
            comp = vbproj.VBComponents(i)
            if comp.Name == module_name and comp.Type == 1:  # Standard module
                vbproj.VBComponents.Remove(comp)
                return jsonify({"success": True, "message": f"Module '{module_name}' deleted successfully"})
        
        return jsonify({"error": f"Module '{module_name}' not found"}), 404
        
    except Exception as e:
        return jsonify({"error": f"Failed to delete module: {str(e)}"}), 500
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

@app.route("/run-subroutine", methods=["POST"])
def run_subroutine():
    """Run a specific VBA subroutine in Excel."""
    try:
        data = request.json
        if not data:
            return jsonify({"error": "No JSON data provided"}), 400
            
        module_name = data.get("module", "")
        subroutine_name = data.get("subroutine", "")
        
        if not module_name or not subroutine_name:
            return jsonify({"error": "Module name and subroutine name are required"}), 400
        
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
        try:
            excel = win32com.client.GetObject(None, "Excel.Application")
        except:
            excel = win32com.client.Dispatch("Excel.Application")
        
        wb = excel.ActiveWorkbook
        if not wb:
            return jsonify({"error": "No active workbook"}), 400
        
        # Check if VBA project is accessible
        try:
            vbproj = wb.VBProject
        except Exception as vb_error:
            return jsonify({"error": f"VBA project not accessible. Make sure 'Trust access to the VBA project object model' is enabled in Excel Trust Center. Error: {str(vb_error)}"}), 500
        
        # Verify the module exists
        module_found = False
        for i in range(1, vbproj.VBComponents.Count + 1):
            comp = vbproj.VBComponents(i)
            if comp.Name == module_name and comp.Type == 1:  # Standard module
                module_found = True
                break
        
        if not module_found:
            return jsonify({"error": f"Module '{module_name}' not found"}), 404
        
        # Run the subroutine
        try:
            # Construct the full subroutine call
            subroutine_call = f"{module_name}.{subroutine_name}"
            
            # Execute the subroutine
            excel.Run(subroutine_call)
            
            return jsonify({"success": True, "message": f"Subroutine '{subroutine_name}' executed successfully"})
            
        except Exception as run_error:
            return jsonify({"error": f"Failed to run subroutine '{subroutine_name}': {str(run_error)}"}), 500
        
    except Exception as e:
        return jsonify({"error": f"Server error: {str(e)}"}), 500
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

if __name__ == "__main__":
    app.run(host="localhost", port=5000, debug=True)


