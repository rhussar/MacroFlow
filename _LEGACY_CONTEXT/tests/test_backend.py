"""
Test suite for MacroFlow backend
"""

import unittest
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'backend'))

from app import app

class TestBackendAPI(unittest.TestCase):
    
    def setUp(self):
        """Set up test client"""
        self.app = app.test_client()
        self.app.testing = True
    
    def test_generate_endpoint_exists(self):
        """Test that the generate endpoint exists"""
        response = self.app.post('/generate', 
                                json={"prompt": "Test prompt"})
        # Should not return 404 (endpoint exists)
        self.assertNotEqual(response.status_code, 404)
    
    def test_list_modules_endpoint(self):
        """Test that the list-modules endpoint exists"""
        response = self.app.get('/list-modules')
        # Should not return 404 (endpoint exists)  
        self.assertNotEqual(response.status_code, 404)

if __name__ == '__main__':
    unittest.main()