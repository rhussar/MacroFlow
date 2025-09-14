"""
Base Agent Class for MacroFlow AI
Provides common functionality for all agents
"""

from abc import ABC, abstractmethod
from openai import OpenAI
import os
from typing import Dict, Any, List, Optional


class BaseAgent(ABC):
    """Base class for all MacroFlow AI agents"""
    
    def __init__(self):
        # Initialize OpenAI client (shared across all agents)
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            raise ValueError("OPENAI_API_KEY environment variable is required")
        self.client = OpenAI(api_key=api_key)
    
    def _call_openai(self, messages: List[Dict[str, str]], model: str = "gpt-4o", temperature: float = 0) -> str:
        """
        Common OpenAI API call method
        
        Args:
            messages: List of message dictionaries with 'role' and 'content'
            model: OpenAI model to use
            temperature: Creativity level (0-1)
        
        Returns:
            AI response content as string
        """
        try:
            response = self.client.chat.completions.create(
                model=model,
                messages=messages,
                temperature=temperature
            )
            return response.choices[0].message.content
        except Exception as e:
            return f"Error calling OpenAI API: {str(e)}"
    
    def _build_conversation_context(self, conversation_history: Optional[List[Dict]] = None, limit: int = 5) -> List[Dict[str, str]]:
        """
        Build conversation context from history
        
        Args:
            conversation_history: List of previous messages
            limit: Maximum number of previous messages to include
        
        Returns:
            List of formatted messages for OpenAI API
        """
        if not conversation_history:
            return []
        
        # Take only the last `limit` messages to avoid token limits
        recent_history = conversation_history[-limit:] if len(conversation_history) > limit else conversation_history
        
        # Format for OpenAI API
        formatted_messages = []
        for msg in recent_history:
            if msg.get('type') in ['user', 'assistant']:
                role = 'user' if msg['type'] == 'user' else 'assistant'
                formatted_messages.append({
                    'role': role,
                    'content': msg.get('content', '')
                })
        
        return formatted_messages
    
    @abstractmethod
    def process(self, prompt: str, context: Optional[Dict[str, Any]] = None, 
                conversation_history: Optional[List[Dict]] = None) -> Dict[str, Any]:
        """
        Process a user request
        
        Args:
            prompt: User input text
            context: Additional context (editor state, sheet data, etc.)
            conversation_history: Previous conversation messages
        
        Returns:
            Dictionary with response data
        """
        pass