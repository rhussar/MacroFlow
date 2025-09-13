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
        base_prompt = "You are an expert VBA assistant for Excel. Return ONLY the VBA code without any explanations, markdown formatting, or additional text."
        
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
        
        # Add code context
        if context.get("hasSelection") and context.get("selectedText"):
            user_message += f"\n\nSELECTED CODE TO WORK WITH:\n{context['selectedText']}"
            
            if context.get("surroundingCode"):
                user_message += f"\n\nSURROUNDING CONTEXT (for reference):\n{context['surroundingCode']}"
        
        elif context.get("surroundingCode"):
            user_message += f"\n\nEXISTING CODE CONTEXT:\n{context['surroundingCode']}"
        
        # Add module context
        if context.get("currentSubroutine"):
            user_message += f"\n\nCURRENT SUBROUTINE: {context['currentSubroutine']}"
        
        if context.get("activeModule"):
            user_message += f"\n\nACTIVE MODULE: {context['activeModule']}"
        
        # Add sheet context
        sheet_ctx = context.get("sheetContext")
        if sheet_ctx and not sheet_ctx.get("error"):
            user_message += self._format_sheet_context(sheet_ctx)
        
        # Add recent conversation context for VBA generation
        if conversation_history:
            recent_vba_context = self._extract_vba_context_from_history(conversation_history)
            if recent_vba_context:
                user_message += f"\n\nRECENT CONTEXT: {recent_vba_context}"
        
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