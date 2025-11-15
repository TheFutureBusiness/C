"""
Moduł integracji z zewnętrznymi API
"""
from .openai_integration import generate_ai_summary
from .pagespeed import check_pagespeed

__all__ = ['generate_ai_summary', 'check_pagespeed']
