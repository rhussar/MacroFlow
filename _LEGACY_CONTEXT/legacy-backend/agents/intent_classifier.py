"""
Intent Classifier Agent
Determines whether user input is requesting code generation or conversation
"""

from typing import Dict, Any, List, Optional
from .base_agent import BaseAgent
import re


class IntentClassifier(BaseAgent):
    """Classifies user intent to route to appropriate agent"""
    
    def __init__(self):
        super().__init__()
        
        # Define intent patterns for keyword-based classification
        self.intent_patterns = {
            'vba_generation': [
                # Direct code requests
                r'\b(create|write|generate|make|build)\b.*\b(macro|function|sub|code|vba)\b',
                r'\b(macro|function|sub)\b.*\b(for|to|that)\b',
                r'\b(write|create)\b.*\b(that|to)\b',
                
                # Action-oriented requests
                r'\b(automate|calculate|format|sort|filter|find|replace|copy|paste)\b',
                r'\b(loop|iterate|repeat)\b.*\b(through|over|for)\b',
                r'\b(if|when|then)\b.*\b(do|run|execute)\b',
                
                # Excel-specific terms
                r'\b(cell|range|worksheet|workbook|column|row)\b',
                r'\b(pivot|chart|table|data|formula)\b',
            ],
            
            'vba_modification': [
                # Code improvement requests
                r'\b(fix|debug|improve|optimize|enhance|refactor)\b',
                r'\b(error|bug|issue|problem)\b.*\b(in|with)\b',
                r'\b(change|modify|update|edit)\b.*\b(code|macro|function)\b',
                r'\b(add|remove|delete)\b.*\b(to|from)\b.*\b(code|function)\b',
                
                # Follow-up modification requests (crucial for "do it for me" scenarios)
                r'\b(do it|make the change|apply|implement|update)\b.*\b(for me|please|now)\b',
                r'^(do it|make it|change it|update it|fix it)(\s+for me|\s+please|$)',
                r'\b(go ahead|please do|make that change)\b',
                r'\b(apply that|implement that|do that change)\b',
                
                # Direct modification commands
                r'\b(turn\s+(on|off)|enable|disable|set\s+to)\b.*\b(gridlines|formatting|option)\b',
                r'\b(make\s+it|change\s+it\s+to|set\s+it\s+to)\b',
            ],
            
            'conversation': [
                # Greetings and social
                r'\b(hi|hello|hey|thanks|thank you|please|sorry)\b',
                r'\b(good|nice|great|awesome|excellent)\b',
                
                # Questions about functionality
                r'\b(what|how|why|when|where|can|will|does|is)\b',
                r'\b(explain|tell me|show me|help)\b',
                r'\b(understand|know|learn)\b',
                
                # Code discussion and explanation (when code exists)
                r'\b(what does.*do|explain.*code|describe.*function)\b',
                r'\b(this (module|code|function|sub|macro))\b',
                r'\b(analyze|review|look at|check)\b.*\b(code|module)\b',
                
                # MacroFlow-specific help
                r'\b(macroflow|this tool|this app|this program)\b',
                r'\b(feature|capability|function)\b.*\b(of|in)\b',
            ],
            
            'help': [
                r'\b(help|support|guide|tutorial|documentation)\b',
                r'\b(how to|how do I|how can I)\b',
                r'\b(getting started|start|begin)\b',
                r'\b(example|sample|demo)\b',
            ]
        }
    
    def _keyword_classify(self, prompt: str, context: Optional[Dict[str, Any]] = None, 
                        conversation_history: Optional[List[Dict]] = None) -> str:
        """
        Classify intent using keyword patterns
        
        Args:
            prompt: User input text
            context: Additional context
        
        Returns:
            Intent classification: 'vba_generation', 'vba_modification', 'conversation', or 'help'
        """
        prompt_lower = prompt.lower()
        
        # Check conversation history for recent code discussions
        discussed_code_recently = False
        if conversation_history:
            recent_messages = conversation_history[-3:]  # Last 3 messages
            for msg in recent_messages:
                content = msg.get('content', '').lower()
                if any(word in content for word in ['function', 'sub', 'macro', 'code', 'vba']):
                    discussed_code_recently = True
                    break
        
        # Enhanced context-aware classification
        has_code = False
        if context:
            # Priority 1: Check if cursor is inside existing subroutine (strongest modification signal)
            if context.get('isInsideSubroutine') or context.get('currentSubroutine'):
                # Cursor inside existing function strongly suggests modification intent
                # Check for modification patterns OR general requests that could modify existing code
                modification_indicators = [
                    r'\b(add|change|modify|update|fix|improve|enhance|optimize)\b',
                    r'\b(make it|make this|change it to|set it to|turn on|turn off)\b',
                    r'\b(include|exclude|remove|delete|insert)\b',
                    r'\b(handle|check|validate|ensure)\b',  # Common modification requests
                    r'\b(make.*where|change.*to|set.*to)\b',  # Broader "make" patterns
                    r'^(make|change|set|add|remove|delete|fix|update|modify)',  # Start of sentence patterns
                ]

                # If cursor is inside subroutine, be very aggressive about detecting modifications
                for pattern in modification_indicators:
                    if re.search(pattern, prompt_lower, re.IGNORECASE):
                        return 'vba_modification'

                # If cursor is inside subroutine, also check standard modification patterns
                for pattern in self.intent_patterns['vba_modification']:
                    if re.search(pattern, prompt_lower, re.IGNORECASE):
                        return 'vba_modification'

                # AGGRESSIVE: If cursor is inside subroutine and prompt is short (< 10 words)
                # and contains action words, assume modification
                word_count = len(prompt_lower.split())
                if word_count <= 10:
                    action_words = ['make', 'change', 'add', 'set', 'put', 'use', 'get', 'take', 'move', 'do']
                    if any(word in prompt_lower.split() for word in action_words):
                        return 'vba_modification'

            # Priority 2: Check if user has selected code (medium modification signal)
            if context.get('hasSelection') and context.get('selectedText'):
                has_code = True
                # If user has selected code, likely wants modification
                for pattern in self.intent_patterns['vba_modification']:
                    if re.search(pattern, prompt_lower, re.IGNORECASE):
                        return 'vba_modification'
                # Check for conversation patterns when code exists
                for pattern in self.intent_patterns['conversation']:
                    if re.search(pattern, prompt_lower, re.IGNORECASE):
                        return 'conversation'
                # Default to VBA generation if code selected but no clear intent
                return 'vba_generation'
            
            # Check if current module has code content
            elif context.get('fullCode') and context['fullCode'].strip():
                has_code = True
                
                # Check for modification requests when code exists
                for pattern in self.intent_patterns['vba_modification']:
                    if re.search(pattern, prompt_lower, re.IGNORECASE):
                        return 'vba_modification'
                
                # When discussing existing code, prioritize conversation
                for pattern in self.intent_patterns['conversation']:
                    if re.search(pattern, prompt_lower, re.IGNORECASE):
                        return 'conversation'
        
        # Special handling for "do it for me" type requests when code was discussed recently
        if discussed_code_recently:
            simple_action_patterns = [
                r'^(do it|make it|change it|update it|fix it)(\s+for me|\s+please|$)',
                r'^(go ahead|please do|apply that|implement that)$',
                r'^(yes|ok|okay)\s*(do it|please|for me)$'
            ]
            for pattern in simple_action_patterns:
                if re.search(pattern, prompt_lower, re.IGNORECASE):
                    return 'vba_modification'
        
        # Score each intent based on pattern matches
        intent_scores = {}
        for intent, patterns in self.intent_patterns.items():
            score = 0
            for pattern in patterns:
                matches = len(re.findall(pattern, prompt_lower, re.IGNORECASE))
                score += matches
            intent_scores[intent] = score
        
        # Return intent with highest score, default to conversation if tie
        if max(intent_scores.values()) == 0:
            # No patterns matched, use AI to classify
            return self._ai_classify(prompt, context)
        
        return max(intent_scores.items(), key=lambda x: x[1])[0]
    
    def _ai_classify(self, prompt: str, context: Optional[Dict[str, Any]] = None) -> str:
        """
        Use AI to classify ambiguous intents
        
        Args:
            prompt: User input text
            context: Additional context
        
        Returns:
            Intent classification
        """
        system_prompt = """You are an intent classifier for MacroFlow, a VBA code generation tool.

Classify the user's intent into one of these categories:
- vba_generation: User wants to create new VBA code/macros
- vba_modification: User wants to modify/fix existing code  
- conversation: User is asking questions, greeting, or having a discussion
- help: User needs help using MacroFlow or understanding features

Respond with ONLY the classification category, nothing else."""

        user_message = f"User input: {prompt}"
        
        if context:
            context_info = []
            if context.get('hasSelection'):
                context_info.append("User has selected code in the editor")
            if context.get('activeModule'):
                context_info.append(f"User is working in module '{context['activeModule']}'")
            if context.get('fullCode') and context['fullCode'].strip():
                context_info.append("Current module contains existing VBA code")
            if context.get('moduleMetadata') and context['moduleMetadata'].get('procedures'):
                proc_count = len(context['moduleMetadata']['procedures'])
                context_info.append(f"Module has {proc_count} procedures")
            
            if context_info:
                user_message += f"\nContext: {' | '.join(context_info)}"
        
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message}
        ]
        
        response = self._call_openai(messages, temperature=0)
        
        # Validate response
        valid_intents = ['vba_generation', 'vba_modification', 'conversation', 'help']
        if response.strip() in valid_intents:
            return response.strip()
        
        # Default to conversation if AI response is invalid
        return 'conversation'
    
    def process(self, prompt: str, context: Optional[Dict[str, Any]] = None, 
                conversation_history: Optional[List[Dict]] = None) -> Dict[str, Any]:
        """
        Classify user intent
        
        Args:
            prompt: User input text
            context: Editor/sheet context
            conversation_history: Previous messages (not used for classification)
        
        Returns:
            Dictionary with intent classification and confidence
        """
        # Primary classification using keywords with conversation history
        intent = self._keyword_classify(prompt, context, conversation_history)
        
        # Add confidence scoring
        confidence = 'high' if intent in ['vba_generation', 'vba_modification'] else 'medium'
        
        return {
            'intent': intent,
            'confidence': confidence,
            'classifier': 'keyword' if intent != self._ai_classify(prompt, context) else 'ai',
            'context_used': bool(context and (context.get('hasSelection') or context.get('activeModule')))
        }