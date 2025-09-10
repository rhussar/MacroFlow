from flask import Flask, request, jsonify
from flask_cors import CORS
from openai import OpenAI
import os
import win32com.client
import pythoncom
import time

app = Flask(__name__)
CORS(app)  # ✅ Allow requests from localhost:3000 (Add-in)

# OpenAI client
api_key = os.getenv("OPENAI_API_KEY") or "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"
client = OpenAI(api_key=api_key)

# Simple cache for sheet context (to avoid repeated reads)
sheet_context_cache = {
    "data": None,
    "timestamp": 0,
    "cache_duration": 30  # Cache for 30 seconds
}


def generate_vba(prompt, context=None):
    # Build enhanced system prompt based on context
    system_prompt = "You are an expert VBA assistant for Excel. Return ONLY the VBA code without any explanations, markdown formatting, or additional text."
    
    # Build user message with context
    user_message = prompt
    
    if context:
        # VBA Code Context
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
        
        # Sheet Context (NEW)
        if context.get("sheetContext") and not context["sheetContext"].get("error"):
            sheet_ctx = context["sheetContext"]
            system_prompt += " You have access to the current Excel sheet's data and structure. Use this information to generate more relevant and data-aware VBA code."
            
            # Add sheet information
            user_message += f"\n\n=== CURRENT EXCEL SHEET CONTEXT ==="
            user_message += f"\nWorkbook: {sheet_ctx.get('workbook_name', 'Unknown')}"
            user_message += f"\nActive Sheet: {sheet_ctx.get('sheet_name', 'Unknown')}"
            user_message += f"\nSheet Type: {sheet_ctx.get('sheet_type', 'Unknown')}"
            
            # Add used range information
            if sheet_ctx.get("used_range"):
                used_range = sheet_ctx["used_range"]
                user_message += f"\nUsed Range: {used_range['address']} ({used_range['rows']} rows × {used_range['columns']} columns)"
            
            # Add column headers
            if sheet_ctx.get("column_headers"):
                headers = sheet_ctx["column_headers"][:10]  # Limit to first 10
                user_message += f"\nColumn Headers: {', '.join(headers)}"
                if len(sheet_ctx["column_headers"]) > 10:
                    user_message += f" (and {len(sheet_ctx['column_headers']) - 10} more...)"
            
            # Add data structure information
            if sheet_ctx.get("data_structure"):
                user_message += f"\n\nData Structure:"
                for col_name, col_info in list(sheet_ctx["data_structure"].items())[:5]:  # Limit to 5 columns
                    user_message += f"\n  - {col_name}: {col_info['type']}"
                    if col_info.get("sample_values"):
                        sample_str = ", ".join(str(v) for v in col_info["sample_values"][:2])
                        user_message += f" (samples: {sample_str})"
            
            # Add named ranges
            if sheet_ctx.get("named_ranges"):
                ranges = [nr["name"] for nr in sheet_ctx["named_ranges"][:5]]
                if ranges:
                    user_message += f"\nNamed Ranges: {', '.join(ranges)}"
            
            # Add chart objects
            if sheet_ctx.get("chart_objects"):
                charts = [c["name"] for c in sheet_ctx["chart_objects"][:3]]
                if charts:
                    user_message += f"\nChart Objects: {', '.join(charts)}"
            
            # Add sample data (first few rows)
            if sheet_ctx.get("data_sample") and len(sheet_ctx["data_sample"]) > 1:
                user_message += f"\n\nSample Data (first few rows):"
                headers = sheet_ctx.get("column_headers", [])
                sample_rows = sheet_ctx["data_sample"][:4]  # Header + 3 data rows
                
                for i, row in enumerate(sample_rows):
                    row_data = row[:5]  # First 5 columns only
                    if i == 0 and headers:
                        user_message += f"\n  Headers: {' | '.join(row_data)}"
                    else:
                        user_message += f"\n  Row {i}: {' | '.join(row_data)}"
            
            user_message += f"\n=== END SHEET CONTEXT ===\n"
    
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

def separate_vba_and_text(ai_response):
    """Separate VBA code from explanatory text in AI response."""
    lines = ai_response.split('\n')
    vba_lines = []
    text_lines = []
    in_vba_block = False
    vba_found = False
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        
        # Check for VBA code block start
        if line.startswith(('Sub ', 'Function ', 'Private Sub ', 'Public Sub ', 
                          'Private Function ', 'Public Function ')):
            in_vba_block = True
            vba_found = True
            vba_lines.append(lines[i])
            
        elif in_vba_block:
            vba_lines.append(lines[i])
            # Check for VBA code block end
            if line in ['End Sub', 'End Function']:
                in_vba_block = False
                
        else:
            # This is explanatory text
            # Skip empty lines at the beginning
            if text_lines or line:
                text_lines.append(lines[i])
        
        i += 1
    
    # Clean up text lines (remove excessive empty lines)
    while text_lines and not text_lines[0].strip():
        text_lines.pop(0)
    while text_lines and not text_lines[-1].strip():
        text_lines.pop()
    
    vba_code = '\n'.join(vba_lines) if vba_lines else ""
    explanation = '\n'.join(text_lines) if text_lines else ""
    
    return {
        "has_vba": vba_found,
        "vba_code": vba_code,
        "explanation": explanation
    }

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

def get_sheet_context():
    """Extract context from the active Excel sheet with caching and error handling."""
    # Check cache first
    current_time = time.time()
    if (sheet_context_cache["data"] and 
        current_time - sheet_context_cache["timestamp"] < sheet_context_cache["cache_duration"]):
        return sheet_context_cache["data"]
    
    try:
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
        try:
            excel = win32com.client.GetObject(None, "Excel.Application")
        except:
            try:
                excel = win32com.client.Dispatch("Excel.Application")
            except Exception as excel_error:
                return {"error": f"Could not connect to Excel: {str(excel_error)}"}
        
        wb = excel.ActiveWorkbook
        if not wb:
            return {"error": "No active workbook found. Please open an Excel file first."}
        
        ws = excel.ActiveSheet
        if not ws:
            return {"error": "No active worksheet found."}
        
        # Basic sheet information
        sheet_info = {
            "workbook_name": wb.Name,
            "sheet_name": ws.Name,
            "sheet_type": "Worksheet" if ws.Type == -4167 else "Chart" if ws.Type == 3 else "Unknown",
            "used_range": None,
            "data_sample": [],
            "column_headers": [],
            "data_structure": {},
            "named_ranges": [],
            "chart_objects": []
        }
        
        # Get used range information
        try:
            used_range = ws.UsedRange
            if used_range:
                sheet_info["used_range"] = {
                    "address": used_range.Address,
                    "rows": used_range.Rows.Count,
                    "columns": used_range.Columns.Count,
                    "first_row": used_range.Row,
                    "first_col": used_range.Column
                }
                
                # Sample data from used range (limit to first 50 rows, 20 columns)
                max_rows = min(50, used_range.Rows.Count)
                max_cols = min(20, used_range.Columns.Count)
                
                sample_data = []
                for row in range(1, max_rows + 1):
                    row_data = []
                    for col in range(1, max_cols + 1):
                        try:
                            cell_value = used_range.Cells(row, col).Value
                            # Convert to string, handle None values
                            row_data.append(str(cell_value) if cell_value is not None else "")
                        except:
                            row_data.append("")
                    sample_data.append(row_data)
                
                sheet_info["data_sample"] = sample_data
                
                # Extract column headers (assume first row contains headers)
                if max_rows > 0:
                    headers = []
                    for col in range(1, max_cols + 1):
                        try:
                            header = used_range.Cells(1, col).Value
                            headers.append(str(header) if header is not None else f"Column{col}")
                        except:
                            headers.append(f"Column{col}")
                    sheet_info["column_headers"] = headers
                
                # Analyze data structure
                if len(sample_data) > 1:  # Skip header row
                    data_structure = {}
                    for col_idx, header in enumerate(headers):
                        if col_idx < len(sample_data[1]):  # Check if data exists
                            # Analyze data types in this column (sample first few non-empty values)
                            sample_values = []
                            for row_idx in range(1, min(10, len(sample_data))):
                                if col_idx < len(sample_data[row_idx]):
                                    val = sample_data[row_idx][col_idx]
                                    if val and str(val).strip():
                                        sample_values.append(val)
                            
                            # Determine data type
                            data_type = "text"
                            if sample_values:
                                # Check if numeric
                                try:
                                    float(sample_values[0])
                                    data_type = "number"
                                except:
                                    # Check if date-like
                                    if any(char in str(sample_values[0]) for char in ['/', '-', ':']):
                                        data_type = "date_or_text"
                            
                            data_structure[header] = {
                                "type": data_type,
                                "sample_values": sample_values[:3]  # First 3 samples
                            }
                    
                    sheet_info["data_structure"] = data_structure
        except Exception as e:
            sheet_info["used_range_error"] = str(e)
        
        # Get named ranges
        try:
            for name in wb.Names:
                try:
                    sheet_info["named_ranges"].append({
                        "name": name.Name,
                        "refers_to": name.RefersTo,
                        "scope": "Workbook"
                    })
                except:
                    continue
        except Exception as e:
            sheet_info["named_ranges_error"] = str(e)
        
        # Get chart objects
        try:
            for chart in ws.ChartObjects():
                try:
                    sheet_info["chart_objects"].append({
                        "name": chart.Name,
                        "chart_type": chart.Chart.ChartType,
                        "position": {
                            "left": chart.Left,
                            "top": chart.Top,
                            "width": chart.Width,
                            "height": chart.Height
                        }
                    })
                except:
                    continue
        except Exception as e:
            sheet_info["chart_objects_error"] = str(e)
        
        # Cache the result
        sheet_context_cache["data"] = sheet_info
        sheet_context_cache["timestamp"] = current_time
        
        return sheet_info
        
    except Exception as e:
        error_result = {"error": f"Failed to get sheet context: {str(e)}"}
        # Don't cache errors, but return them
        return error_result
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

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

@app.route("/get-sheet-context", methods=["GET"])
def get_sheet_context_endpoint():
    """Get context from the active Excel sheet."""
    try:
        sheet_context = get_sheet_context()
        return jsonify(sheet_context)
    except Exception as e:
        return jsonify({"error": f"Failed to get sheet context: {str(e)}"}), 500

@app.route("/generate", methods=["POST"])
def generate():
    data = request.json
    prompt = data.get("prompt", "")
    context = data.get("context", None)
    
    if not prompt:
        return jsonify({"error": "No prompt provided"}), 400

    
    # Generate VBA with context awareness
    ai_response = generate_vba(prompt, context)
    
    # Separate VBA code from explanatory text
    separated = separate_vba_and_text(ai_response)
    
    
    # Return enhanced response with separated VBA and explanation
    response_data = {
        "has_vba": separated["has_vba"],
        "vba_code": separated["vba_code"],
        "explanation": separated["explanation"],
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


