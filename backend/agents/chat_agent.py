"""
Chat Agent
Handles conversational responses, questions, and help requests
"""

from typing import Dict, Any, List, Optional
from .base_agent import BaseAgent


class ChatAgent(BaseAgent):
    """Handles conversational interactions and help requests"""
    
    def __init__(self):
        super().__init__()
        
        # Define well-formatted response templates for common scenarios
        self.response_templates = {
            'greeting': [
                "Hello! I'm your **MacroFlow AI assistant**. I'm here to help you create and manage VBA macros for Excel.\n\n• Generate custom VBA code\n• Answer Excel automation questions\n• Guide you through MacroFlow features\n\nWhat can I help you with today?",
                "Hi there! Ready to build some powerful Excel macros together?\n\n**I can help you:**\n• Create automated workflows\n• Solve Excel productivity challenges\n• Generate clean, efficient VBA code\n\nWhat's your goal?",
                "Hey! Welcome to MacroFlow.\n\n**Quick overview:**\n• Ask me to create any VBA macro\n• Get help with Excel automation\n• Learn about advanced Excel features\n\nLet's get started!"
            ],
            
            'thanks': [
                "You're welcome! **Happy to help** with your VBA development.\n\nFeel free to ask for more macros or Excel guidance anytime.",
                "**Glad I could assist!** Let me know if you need:\n\n• More VBA code\n• Excel automation help\n• MacroFlow guidance",
                "My pleasure! **Always here to help** with:\n\n• VBA development\n• Excel productivity\n• Automation solutions"
            ],
            
            'help_general': [
                "I can help you with:\n\n• **Create VBA Macros** - Automate any Excel task\n• **Excel Questions** - Formulas, features, best practices\n• **Debug Code** - Fix and improve existing macros\n\n**What would you like to work on?**",
                "**MacroFlow can:**\n\n• Generate custom VBA code\n• Answer Excel automation questions  \n• Help troubleshoot existing macros\n\n**What's your goal today?**"
            ]
        }
    
    def _get_base_system_prompt(self) -> str:
        """Get the base system prompt for conversational responses"""
        return """You are a helpful AI assistant for MacroFlow, a VBA code generation tool for Excel. 

Your role is to:
- Answer questions about Excel, VBA, and automation
- Explain and discuss existing VBA code modules
- Provide guidance on using MacroFlow features
- Help users understand VBA concepts and their existing code
- Offer friendly, conversational support

CODE DISCUSSION CAPABILITIES:
- You can see and analyze the full content of VBA modules
- Explain what existing code does and how it works
- Identify potential issues or improvements in code
- Discuss functions, subroutines, and variables in modules
- Help users understand code structure and logic

FORMATTING GUIDELINES - ALWAYS FOLLOW THESE:
- Use **bold** for emphasis and important terms
- Use bullet points (•) for lists, not dashes (-)
- Add blank lines between sections for better readability
- Keep paragraphs short (2-3 sentences maximum)
- Use numbered lists (1. 2. 3.) for sequential steps
- Add section headers when explaining complex topics
- End with a clear question or call-to-action when appropriate

RESPONSE STRUCTURE:
- Start with a brief, direct answer
- Add supporting details with proper formatting
- Use bullet points for multiple items
- Include examples when helpful
- End with engagement (question, suggestion, or call-to-action)

CONTENT GUIDELINES:
- Be friendly, professional, and helpful
- Keep responses SHORT and concise (3-4 bullet points maximum)
- When discussing VBA, focus on practical Excel applications
- **NEVER include VBA code snippets or code examples in your responses**
- **NEVER show code changes or modifications - only explain what needs to be done**
- If a user wants code changes, tell them to use "do it for me" instead of showing code
- Don't generate actual VBA code (that's handled by the VBA Agent)
- Avoid lengthy explanations - prefer clear, actionable bullet points
- When code is provided, analyze and explain it clearly WITHOUT showing modified code

You're having a conversation with someone who wants to improve their Excel productivity through automation."""
    
    def _detect_response_type(self, prompt: str) -> str:
        """
        Detect what type of conversational response is needed
        
        Args:
            prompt: User input text
        
        Returns:
            Response type: 'greeting', 'thanks', 'help_general', 'question', or 'general'
        """
        prompt_lower = prompt.lower().strip()
        
        # Greeting detection
        greeting_words = ['hi', 'hello', 'hey', 'good morning', 'good afternoon', 'good evening']
        if any(word in prompt_lower for word in greeting_words):
            return 'greeting'
        
        # Thanks detection
        thanks_words = ['thank', 'thanks', 'appreciate', 'grateful']
        if any(word in prompt_lower for word in thanks_words):
            return 'thanks'
        
        # Help requests
        help_words = ['help', 'guide', 'tutorial', 'how to use', 'getting started', 'what can you do']
        if any(word in prompt_lower for word in help_words):
            return 'help_general'
        
        # Question detection
        question_words = ['what', 'how', 'why', 'when', 'where', 'can', 'will', 'does', 'is', 'are']
        if any(prompt_lower.startswith(word) for word in question_words) or '?' in prompt:
            return 'question'
        
        return 'general'
    
    def _generate_contextual_response(self, prompt: str, response_type: str, 
                                    context: Optional[Dict[str, Any]] = None,
                                    conversation_history: Optional[List[Dict]] = None) -> str:
        """
        Generate a contextual response using AI
        
        Args:
            prompt: User input
            response_type: Type of response needed
            context: Editor/sheet context
            conversation_history: Previous conversation
        
        Returns:
            AI-generated response
        """
        system_prompt = self._get_base_system_prompt()
        
        # Add context-specific guidance
        if response_type == 'question':
            system_prompt += "\n\nThe user is asking a question. Provide a helpful, informative answer using proper formatting:\n- Start with a direct answer\n- Use bullet points for multiple aspects\n- Include examples if helpful\n- End with a follow-up question or suggestion"
        elif response_type == 'help_general':
            system_prompt += "\n\nThe user needs general help. Explain MacroFlow capabilities using clear structure:\n- Use **bold** headers for main sections\n- List capabilities with bullet points\n- Include brief descriptions for each feature\n- End with 'What would you like to work on?'"
        
        # Build messages for API call
        messages = [{"role": "system", "content": system_prompt}]
        
        # Add conversation history for context
        if conversation_history:
            history_messages = self._build_conversation_context(conversation_history, limit=3)
            messages.extend(history_messages)
        
        # Add current user message
        user_message = prompt
        
        # Add context information if available - now with full code visibility
        if context:
            context_info = []
            
            # Basic context
            if context.get('activeModule'):
                context_info.append(f"User is currently working in module: {context['activeModule']}")
            if context.get('hasSelection'):
                context_info.append("User has code selected in the editor")
            
            # Add full code context for chat discussions about existing modules (with size management)
            if context.get("fullCode") and context["fullCode"].strip():
                full_code = context["fullCode"]
                # Size management for chat - more aggressive truncation since focus is on discussion
                if len(full_code) > 4000:  # About 1000 tokens for chat
                    lines = full_code.split('\n')
                    # Keep first 50 lines for chat discussion
                    truncated_code = '\n'.join(lines[:50])
                    user_message += f"\n\nCURRENT MODULE CODE (excerpt):\n{truncated_code}\n\n[Note: Showing first 50 lines of {len(lines)} total lines]"
                else:
                    user_message += f"\n\nCURRENT MODULE CODE:\n{full_code}"
            
            # Add module metadata for better understanding
            if context.get("moduleMetadata"):
                metadata = context["moduleMetadata"]
                if metadata.get("procedures"):
                    procedures_list = [f"{proc['name']} ({proc['type']})" for proc in metadata["procedures"]]
                    if procedures_list:
                        user_message += f"\n\nModule contains: {', '.join(procedures_list)}"
            
            # Add available modules context
            if context.get("availableModules"):
                modules_with_content = [m["name"] for m in context["availableModules"] if m["hasContent"]]
                if modules_with_content:
                    context_info.append(f"Available modules: {', '.join(modules_with_content)}")
            
            # Sheet context
            if context.get('sheetContext') and not context['sheetContext'].get('error'):
                sheet_ctx = context['sheetContext']
                context_info.append(f"Active Excel sheet: {sheet_ctx.get('sheet_name', 'Unknown')}")
                if sheet_ctx.get('used_range'):
                    used_range = sheet_ctx['used_range']
                    context_info.append(f"Sheet has data in {used_range['rows']} rows × {used_range['columns']} columns")
            
            if context_info:
                user_message += f"\n\nContext: {' | '.join(context_info)}"
        
        messages.append({"role": "user", "content": user_message})
        
        ai_response = self._call_openai(messages, temperature=0.3)  # Slightly more creative for conversation
        
        # Filter out any VBA code that might have slipped through
        return self._filter_vba_code(ai_response)
    
    def _filter_vba_code(self, response: str) -> str:
        """
        Filter out VBA code snippets from chat responses
        
        Args:
            response: AI response text
            
        Returns:
            Filtered response with VBA code removed/replaced
        """
        import re
        
        # Patterns to detect VBA code
        vba_patterns = [
            r'```vba.*?```',  # Code blocks
            r'`[^`]*(?:Sub|Function|End Sub|End Function|Dim|ActiveSheet)[^`]*`',  # Inline VBA
            r'Sub\s+\w+\(\).*?End Sub',  # Sub blocks
            r'Function\s+\w+\(.*?\).*?End Function',  # Function blocks
            r'ActiveSheet\.[A-Za-z]+\s*=',  # ActiveSheet assignments
            r'ActiveWindow\.[A-Za-z]+\s*=',  # ActiveWindow assignments
        ]
        
        filtered_response = response
        for pattern in vba_patterns:
            filtered_response = re.sub(pattern, '[VBA code example removed - use "do it for me" to apply changes]', 
                                     filtered_response, flags=re.DOTALL | re.IGNORECASE)
        
        # Clean up multiple replacements
        filtered_response = re.sub(r'\[VBA code example removed[^\]]*\]\s*\[VBA code example removed[^\]]*\]', 
                                 '[VBA code examples removed - use "do it for me" to apply changes]', 
                                 filtered_response)
        
        return filtered_response
    
    def process(self, prompt: str, context: Optional[Dict[str, Any]] = None, 
                conversation_history: Optional[List[Dict]] = None) -> Dict[str, Any]:
        """
        Process a conversational request
        
        Args:
            prompt: User input text
            context: Editor/sheet context
            conversation_history: Previous conversation messages
        
        Returns:
            Dictionary with conversational response
        """
        response_type = self._detect_response_type(prompt)
        
        # For simple responses, use templates
        if response_type in self.response_templates and not context:
            import random
            response_content = random.choice(self.response_templates[response_type])
        else:
            # Generate contextual response with AI
            response_content = self._generate_contextual_response(
                prompt, response_type, context, conversation_history
            )
        
        return {
            'type': 'conversation',
            'content': response_content,
            'response_type': response_type,
            'has_vba': False,
            'vba_code': '',
            'explanation': response_content
        }