"""
VBA Agent
Handles VBA code generation and modification requests
"""

from typing import Dict, Any, List, Optional
from .base_agent import BaseAgent
import re


class VBAAgent(BaseAgent):
    """Handles VBA code generation and modification"""
    
    def __init__(self):
        super().__init__()
    
    def _build_vba_system_prompt(self, context: Optional[Dict[str, Any]] = None, 
                                intent: str = 'vba_generation') -> str:
        """
        Build system prompt for VBA generation based on context and intent
        
        Args:
            context: Editor/sheet context
            intent: Specific intent (vba_generation, vba_modification)
        
        Returns:
            Customized system prompt
        """
        base_prompt = "You are an expert VBA assistant for Excel. Return ONLY clean, well-formatted VBA code without any explanations, markdown formatting, or additional text."
        
        if not context:
            return base_prompt
        
        # Add context-specific instructions
        if context.get("hasSelection") and context.get("selectedText"):
            if intent == 'vba_modification':
                base_prompt += " The user has selected specific code that they want you to modify, enhance, or fix. Pay attention to the existing code structure and style. Maintain compatibility with the surrounding code."
            else:
                base_prompt += " The user has selected code as a reference. You can modify, enhance, or replace the selected code based on their request. Pay attention to the existing code structure and style."
        
        elif context.get("surroundingCode"):
            base_prompt += " The user is working in an existing VBA module. Generate code that integrates well with the existing code structure, follows the same naming conventions, and doesn't conflict with existing subroutines."
        
        # Add sheet context awareness
        if context.get("sheetContext") and not context["sheetContext"].get("error"):
            base_prompt += " You have access to the current Excel sheet's data and structure. Use this information to generate more relevant and data-aware VBA code that works with the actual data present."
        
        return base_prompt
    
    def _build_modification_instructions(self, context: Dict[str, Any], 
                                      conversation_history: Optional[List[Dict]] = None) -> str:
        """
        Build specific instructions for code modification tasks
        
        Args:
            context: Context with full code and metadata
            conversation_history: Recent conversation for targeting
            
        Returns:
            Additional system prompt instructions for modifications
        """
        instructions = "\n\nMODIFICATION INSTRUCTIONS:"
        instructions += "\n- You MUST modify existing code, not add new functions"
        instructions += "\n- Replace the ENTIRE existing function that needs modification"
        instructions += "\n- Do NOT add new functions alongside existing ones"
        instructions += "\n- Preserve all other functions exactly as they are"
        
        # Try to identify which function needs modification from conversation
        target_function = None
        if conversation_history:
            recent_messages = conversation_history[-5:]  # Look at last 5 messages
            for msg in recent_messages:
                content = msg.get('content', '').lower()
                # Look for function mentions
                if 'customizesheet' in content:
                    target_function = 'CustomizeSheet'
                    break
                # Look for other function patterns
                import re
                func_match = re.search(r'(function|sub)\s+(\w+)', content, re.IGNORECASE)
                if func_match:
                    target_function = func_match.group(2)
                    break
        
        # Add function-specific targeting
        if target_function:
            instructions += f"\n- FOCUS ON MODIFYING THE {target_function} function specifically"
            instructions += f"\n- Return ONLY the modified {target_function} function, not the entire module"
        
        # If we have metadata, list available functions
        if context.get('moduleMetadata') and context['moduleMetadata'].get('procedures'):
            procs = [p['name'] for p in context['moduleMetadata']['procedures'][:5]]
            instructions += f"\n- Available functions to modify: {', '.join(procs)}"
        
        return instructions
    
    def _build_vba_user_message(self, prompt: str, context: Optional[Dict[str, Any]] = None,
                              conversation_history: Optional[List[Dict]] = None) -> str:
        """
        Build user message with context information
        
        Args:
            prompt: Original user prompt
            context: Editor/sheet context
            conversation_history: Previous conversation for context
        
        Returns:
            Enhanced user message with context
        """
        user_message = prompt
        
        if not context:
            return user_message
        
        # Add code context - now with full module visibility
        if context.get("hasSelection") and context.get("selectedText"):
            user_message += f"\n\nSELECTED CODE TO WORK WITH:\n{context['selectedText']}"
            
            if context.get("surroundingCode"):
                user_message += f"\n\nSURROUNDING CONTEXT (for reference):\n{context['surroundingCode']}"
        
        elif context.get("surroundingCode"):
            user_message += f"\n\nSURROUNDING CONTEXT:\n{context['surroundingCode']}"
        
        # Add full module code for comprehensive analysis (with size management)
        if context.get("fullCode") and context["fullCode"].strip():
            full_code = context["fullCode"]
            # Basic size management - truncate very large modules
            if len(full_code) > 8000:  # About 2000 tokens
                lines = full_code.split('\n')
                # Keep first 100 lines and add truncation note
                truncated_code = '\n'.join(lines[:100])
                user_message += f"\n\nFULL MODULE CODE (truncated - showing first 100 lines):\n{truncated_code}\n\n[Note: Module has {len(lines)} total lines, showing first 100 for analysis]"
            else:
                user_message += f"\n\nFULL MODULE CODE:\n{full_code}"
        
        # Add module metadata for better understanding
        if context.get("moduleMetadata"):
            metadata = context["moduleMetadata"]
            if metadata.get("procedures"):
                procedures_info = []
                for proc in metadata["procedures"]:
                    procedures_info.append(f"- {proc['type'].title()}: {proc['name']} (line {proc['lineNumber']})")
                if procedures_info:
                    user_message += f"\n\nMODULE PROCEDURES:\n" + "\n".join(procedures_info)
            
            if metadata.get("variables"):
                vars_info = []
                for var in metadata["variables"][:5]:  # Limit to first 5 variables
                    vars_info.append(f"- {var['name']} As {var['dataType']} ({var['scope']})")
                if vars_info:
                    user_message += f"\n\nMODULE VARIABLES:\n" + "\n".join(vars_info)
        
        # Add module context
        if context.get("currentSubroutine"):
            user_message += f"\n\nCURRENT SUBROUTINE: {context['currentSubroutine']}"
        
        if context.get("activeModule"):
            user_message += f"\n\nACTIVE MODULE: {context['activeModule']}"
        
        # Add available modules context
        if context.get("availableModules"):
            modules_info = []
            for module in context["availableModules"]:
                status = "ACTIVE" if module["isActive"] else "available"
                content_info = f"({module['contentLength']} chars)" if module["hasContent"] else "(empty)"
                modules_info.append(f"- {module['name']} {status} {content_info}")
            if modules_info:
                user_message += f"\n\nAVAILABLE MODULES:\n" + "\n".join(modules_info)
        
        # Add sheet context
        sheet_ctx = context.get("sheetContext")
        if sheet_ctx and not sheet_ctx.get("error"):
            user_message += self._format_sheet_context(sheet_ctx)
        
        # Add recent conversation context for VBA generation/modification
        if conversation_history:
            recent_vba_context = self._extract_vba_context_from_history(conversation_history)
            if recent_vba_context:
                user_message += f"\n\nRECENT CONTEXT: {recent_vba_context}"
            
            # For modifications, add specific discussion context
            modification_context = self._extract_modification_context(conversation_history)
            if modification_context:
                user_message += f"\n\nMODIFICATION DISCUSSION: {modification_context}"
        
        return user_message
    
    def _format_sheet_context(self, sheet_ctx: Dict[str, Any]) -> str:
        """Format sheet context information for the prompt"""
        context_text = "\n\n=== CURRENT EXCEL SHEET CONTEXT ==="
        context_text += f"\nWorkbook: {sheet_ctx.get('workbook_name', 'Unknown')}"
        context_text += f"\nActive Sheet: {sheet_ctx.get('sheet_name', 'Unknown')}"
        context_text += f"\nSheet Type: {sheet_ctx.get('sheet_type', 'Unknown')}"
        
        # Add used range information
        if sheet_ctx.get("used_range"):
            used_range = sheet_ctx["used_range"]
            context_text += f"\nUsed Range: {used_range['address']} ({used_range['rows']} rows × {used_range['columns']} columns)"
        
        # Add column headers
        if sheet_ctx.get("column_headers"):
            headers = sheet_ctx["column_headers"][:10]  # Limit to first 10
            context_text += f"\nColumn Headers: {', '.join(headers)}"
            if len(sheet_ctx["column_headers"]) > 10:
                context_text += f" (and {len(sheet_ctx['column_headers']) - 10} more...)"
        
        # Add data structure information
        if sheet_ctx.get("data_structure"):
            context_text += "\n\nData Structure:"
            for col_name, col_info in list(sheet_ctx["data_structure"].items())[:5]:  # Limit to 5 columns
                context_text += f"\n  - {col_name}: {col_info['type']}"
                if col_info.get("sample_values"):
                    sample_str = ", ".join(str(v) for v in col_info["sample_values"][:2])
                    context_text += f" (samples: {sample_str})"
        
        # Add named ranges
        if sheet_ctx.get("named_ranges"):
            ranges = [nr["name"] for nr in sheet_ctx["named_ranges"][:5]]
            if ranges:
                context_text += f"\nNamed Ranges: {', '.join(ranges)}"
        
        # Add chart objects
        if sheet_ctx.get("chart_objects"):
            charts = [c["name"] for c in sheet_ctx["chart_objects"][:3]]
            if charts:
                context_text += f"\nChart Objects: {', '.join(charts)}"
        
        # Add sample data (first few rows)
        if sheet_ctx.get("data_sample") and len(sheet_ctx["data_sample"]) > 1:
            context_text += "\n\nSample Data (first few rows):"
            headers = sheet_ctx.get("column_headers", [])
            sample_rows = sheet_ctx["data_sample"][:4]  # Header + 3 data rows
            
            for i, row in enumerate(sample_rows):
                row_data = row[:5]  # First 5 columns only
                if i == 0 and headers:
                    context_text += f"\n  Headers: {' | '.join(row_data)}"
                else:
                    context_text += f"\n  Row {i}: {' | '.join(row_data)}"
        
        context_text += "\n=== END SHEET CONTEXT ==="
        return context_text
    
    def _extract_vba_context_from_history(self, conversation_history: List[Dict]) -> str:
        """Extract relevant VBA context from recent conversation"""
        # Look for recent VBA-related discussions
        recent_messages = conversation_history[-3:] if len(conversation_history) > 3 else conversation_history
        vba_context = []
        
        for msg in recent_messages:
            content = msg.get('content', '')
            if msg.get('type') == 'user' and any(word in content.lower() for word in ['macro', 'vba', 'function', 'sub', 'code']):
                vba_context.append(f"User mentioned: {content[:100]}...")
        
        return " | ".join(vba_context) if vba_context else ""
    
    def _extract_modification_context(self, conversation_history: List[Dict]) -> str:
        """Extract specific modification requests from conversation"""
        # Look for recent modification discussions
        recent_messages = conversation_history[-5:] if len(conversation_history) > 5 else conversation_history
        modification_context = []
        
        for msg in recent_messages:
            content = msg.get('content', '').lower()
            msg_type = msg.get('type', '')
            
            # Look for specific modification requests
            if 'turn gridlines on' in content or 'make it turn gridlines on' in content:
                modification_context.append("User wants to turn gridlines ON (change False to True)")
            elif 'turn gridlines off' in content:
                modification_context.append("User wants to turn gridlines OFF (change True to False)")
            elif 'customizesheet' in content and msg_type == 'assistant':
                modification_context.append("Discussion about CustomizeSheet subroutine")
            elif any(phrase in content for phrase in ['change', 'modify', 'update', 'fix']):
                # Extract the change request
                change_snippet = content[:100] + "..." if len(content) > 100 else content
                modification_context.append(f"Change requested: {change_snippet}")
        
        return " | ".join(modification_context) if modification_context else ""
    
    def _separate_vba_and_text(self, ai_response: str) -> Dict[str, Any]:
        """
        Separate VBA code from explanatory text in AI response
        (Improved version of existing function)
        """
        lines = ai_response.split('\n')
        vba_lines = []
        text_lines = []
        in_vba_block = False
        vba_found = False
        
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            
            # Check for VBA code block start
            if re.match(r'^(Private\s+|Public\s+)?(Sub|Function)\s+\w+', line, re.IGNORECASE):
                in_vba_block = True
                vba_found = True
                vba_lines.append(lines[i])
                
            elif in_vba_block:
                vba_lines.append(lines[i])
                # Check for VBA code block end
                if re.match(r'^End\s+(Sub|Function)$', line, re.IGNORECASE):
                    in_vba_block = False
                    
            else:
                # This is explanatory text
                if text_lines or line:  # Skip empty lines at the beginning
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
    
    def process(self, prompt: str, context: Optional[Dict[str, Any]] = None, 
                conversation_history: Optional[List[Dict]] = None, 
                intent: str = 'vba_generation') -> Dict[str, Any]:
        """
        Process a VBA generation or modification request
        
        Args:
            prompt: User input text
            context: Editor/sheet context
            conversation_history: Previous conversation messages
            intent: Specific intent (vba_generation, vba_modification)
        
        Returns:
            Dictionary with VBA code and explanation
        """
        # Build system prompt based on context and intent
        system_prompt = self._build_vba_system_prompt(context, intent)
        
        # For modifications, add specific targeting instructions
        if intent == 'vba_modification' and context and context.get('fullCode'):
            system_prompt += self._build_modification_instructions(context, conversation_history)
        
        # Build enhanced user message with context
        user_message = self._build_vba_user_message(prompt, context, conversation_history)
        
        # Prepare messages for OpenAI API
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message}
        ]
        
        # Generate VBA response
        ai_response = self._call_openai(messages, temperature=0)
        
        # Separate VBA code from explanation
        separated = self._separate_vba_and_text(ai_response)
        
        # Build response with additional context
        response_data = {
            'type': 'vba',
            'has_vba': separated["has_vba"],
            'vba_code': separated["vba_code"],
            'explanation': separated["explanation"],
            'context': {
                'hasSelection': context.get("hasSelection", False) if context else False,
                'generationType': intent,
                'sheetContextUsed': bool(context and context.get("sheetContext") and not context["sheetContext"].get("error"))
            }
        }
        
        return response_data