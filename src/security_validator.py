import re
import logging
from typing import Any, Optional

class SecurityValidator:
    """Security validation utilities for Excel operations"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def validate_file_path(self, file_path: str) -> bool:
        """Validate file path security"""
        # Prevent directory traversal attacks
        if '..' in file_path or file_path.startswith('/'):
            self.logger.warning(f"Potential directory traversal: {file_path}")
            return False
        return True
    
    def validate_sql_query(self, query: str) -> bool:
        """Basic SQL injection prevention"""
        # Look for potentially dangerous SQL patterns
        dangerous_patterns = [
            r'(?:DROP|DELETE|UPDATE|INSERT)\s+',
            r'(?:EXEC|EXECUTE)\s+',
            r'(?:UNION\s+SELECT)',
            r'(?:--|\/\*|\*\/)'
        ]
        
        for pattern in dangerous_patterns:
            if re.search(pattern, query, re.IGNORECASE):
                self.logger.warning(f"Potentially dangerous SQL pattern: {pattern}")
                return False
        return True
    
    def sanitize_input(self, value: Any) -> Any:
        """Basic input sanitization"""
        if isinstance(value, str):
            # Remove potentially dangerous characters
            value = re.sub(r'[<>"\'&]', '', value)
        return value
