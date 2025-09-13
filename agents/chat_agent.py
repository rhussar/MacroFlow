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
        
        # Define response templates for common scenarios
        self.response_templates = {
            'greeting': [
                "Hello! I'm your MacroFlow AI assistant. I'm here to help you create and manage VBA macros for Excel.",
                "Hi there! Ready to build some powerful Excel macros together?",
                "Hey! I can help you generate VBA code, answer questions about Excel automation, and guide you through MacroFlow features."
            ],
            
            'thanks': [
                "You're welcome! Happy to help with your VBA development.",
                "Glad I could assist! Let me know if you need anything else.",
                "My pleasure! Feel free to ask if you have more questions."
            ],
            
            'help_general': [
                "I can help you with:\n• Creating VBA macros for Excel automation\n• Explaining how Excel features work\n• Troubleshooting VBA code\n• Answering questions about MacroFlow\n\nWhat would you like to work on?",
                "Here's what I can do for you:\n• Generate custom VBA macros\n• Help with Excel automation tasks\n• Debug and improve existing code\n• Provide guidance on MacroFlow features\n\nWhat's your goal today?"
            ]
        }
    
    def _get_base_system_prompt(self) -> str:
        """Get the base system prompt for conversational responses"""
        return """You are a helpful AI assistant for MacroFlow, a VBA code generation tool for Excel. 

Your role is to:
- Answer questions about Excel, VBA, and automation
- Provide guidance on using MacroFlow features
- Help users understand VBA concepts
- Offer friendly, conversational support

Guidelines:
- Be friendly, professional, and helpful
- Keep responses concise but informative
- When discussing VBA, focus on practical Excel applications
- If a user seems to want code generation, gently suggest they ask for specific macro creation
- Provide examples when helpful
- Don't generate actual VBA code (that's handled by the VBA Agent)

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
            system_prompt += "\n\nThe user is asking a question. Provide a helpful, informative answer."
        elif response_type == 'help_general':
            system_prompt += "\n\nThe user needs general help. Explain MacroFlow capabilities and how to get started."
        
        # Build messages for API call
        messages = [{"role": "system", "content": system_prompt}]
        
        # Add conversation history for context
        if conversation_history:
            history_messages = self._build_conversation_context(conversation_history, limit=3)
            messages.extend(history_messages)
        
        # Add current user message
        user_message = prompt
        
        # Add context information if available
        if context:
            context_info = []
            if context.get('activeModule'):
                context_info.append(f"User is currently working in module: {context['activeModule']}")
            if context.get('hasSelection'):
                context_info.append("User has code selected in the editor")
            if context.get('sheetContext') and not context['sheetContext'].get('error'):
                sheet_ctx = context['sheetContext']
                context_info.append(f"Active Excel sheet: {sheet_ctx.get('sheet_name', 'Unknown')}")
                if sheet_ctx.get('used_range'):
                    used_range = sheet_ctx['used_range']
                    context_info.append(f"Sheet has data in {used_range['rows']} rows × {used_range['columns']} columns")
            
            if context_info:
                user_message += f"\n\nContext: {' | '.join(context_info)}"
        
        messages.append({"role": "user", "content": user_message})
        
        return self._call_openai(messages, temperature=0.3)  # Slightly more creative for conversation
    
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