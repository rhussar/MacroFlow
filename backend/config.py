"""
Configuration management for MacroFlow AI Backend
"""
import os
from typing import Dict, Any


class Config:
    """Base configuration class"""
    
    # OpenAI Configuration
    OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
    
    # Flask Configuration
    DEBUG = False
    HOST = "localhost"
    PORT = 5000
    
    # CORS Configuration
    CORS_ORIGINS = ["https://localhost:3000"]
    
    # Cache Configuration
    SHEET_CONTEXT_CACHE_DURATION = 30  # seconds
    
    @classmethod
    def get_config(cls) -> Dict[str, Any]:
        """Get configuration as dictionary"""
        return {
            'openai_api_key': cls.OPENAI_API_KEY,
            'debug': cls.DEBUG,
            'host': cls.HOST,
            'port': cls.PORT,
            'cors_origins': cls.CORS_ORIGINS,
            'cache_duration': cls.SHEET_CONTEXT_CACHE_DURATION
        }


class DevelopmentConfig(Config):
    """Development configuration"""
    DEBUG = True


class ProductionConfig(Config):
    """Production configuration"""
    DEBUG = False
    HOST = "0.0.0.0"  # Allow external connections in production


# Configuration mapping
config = {
    'development': DevelopmentConfig,
    'production': ProductionConfig,
    'default': DevelopmentConfig
}


def get_config(environment: str = None) -> Config:
    """Get configuration for specified environment"""
    if environment is None:
        environment = os.getenv('FLASK_ENV', 'default')
    
    return config.get(environment, config['default'])