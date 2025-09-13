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
            ],
            
            'conversation': [
                # Greetings and social
                r'\b(hi|hello|hey|thanks|thank you|please|sorry)\b',
                r'\b(good|nice|great|awesome|excellent)\b',
                
                # Questions about functionality
                r'\b(what|how|why|when|where|can|will|does|is)\b',
                r'\b(explain|tell me|show me|help)\b',
                r'\b(understand|know|learn)\b',
                
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
    
    def _keyword_classify(self, prompt: str, context: Optional[Dict[str, Any]] = None) -> str:
        """
        Classify intent using keyword patterns
        
        Args:
            prompt: User input text
            context: Additional context
        
        Returns:
            Intent classification: 'vba_generation', 'vba_modification', 'conversation', or 'help'
        """
        prompt_lower = prompt.lower()
        
        # Check for code modification context (user has selected code)
        if context and context.get('hasSelection') and context.get('selectedText'):
            # If user has selected code, likely wants modification
            for pattern in self.intent_patterns['vba_modification']:
                if re.search(pattern, prompt_lower, re.IGNORECASE):
                    return 'vba_modification'
            # Even without modification keywords, selected code suggests VBA work
            return 'vba_generation'
        
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
        system_prompt = """You are an intent classifier for MacroFlow AI, a VBA code generation tool.

Classify the user's intent into one of these categories:
- vba_generation: User wants to create new VBA code/macros
- vba_modification: User wants to modify/fix existing code  
- conversation: User is asking questions, greeting, or having a discussion
- help: User needs help using MacroFlow or understanding features

Respond with ONLY the classification category, nothing else."""

        user_message = f"User input: {prompt}"
        
        if context:
            if context.get('hasSelection'):
                user_message += "\nContext: User has selected code in the editor"
            if context.get('activeModule'):
                user_message += f"\nContext: User is working in module '{context['activeModule']}'"
        
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
        # Primary classification using keywords
        intent = self._keyword_classify(prompt, context)
        
        # Add confidence scoring
        confidence = 'high' if intent in ['vba_generation', 'vba_modification'] else 'medium'
        
        return {
            'intent': intent,
            'confidence': confidence,
            'classifier': 'keyword' if intent != self._ai_classify(prompt, context) else 'ai',
            'context_used': bool(context and (context.get('hasSelection') or context.get('activeModule')))
        }