"""
Analizator sygnałów NAP (Name, Address, Phone) dla Local SEO
"""
import re
import json
from typing import Dict, Any
from bs4 import BeautifulSoup


def extract_nap_signals(soup: BeautifulSoup, text: str) -> Dict[str, Any]:
    """
    Wydobywa i analizuje sygnały NAP (Name, Address, Phone) dla Local SEO.

    Args:
        soup: Obiekt BeautifulSoup z HTML
        text: Oczyszczony tekst strony

    Returns:
        Słownik z wynikami analizy NAP
    """
    # Wzorce numerów telefonów
    phone_patterns = [
        r'\+?48\s?[\d\s\-]{9,}',
        r'\(\d{3}\)\s?\d{3}[\s\-]?\d{4}',
        r'\d{3}[\s\-]?\d{3}[\s\-]?\d{4}',
    ]

    phones = []
    for pattern in phone_patterns:
        phones.extend(re.findall(pattern, text))

    # Wskaźniki adresu
    address_indicators = ['ul.', 'ulica', 'al.', 'aleja', 'street', 'avenue', 'road']
    has_address_indicators = any(ind in text.lower() for ind in address_indicators)

    # Sprawdzenie Schema.org dla Local Business
    schema_scripts = soup.find_all('script', type='application/ld+json')
    has_local_schema = False

    for script in schema_scripts:
        try:
            data = json.loads(script.string)
            if isinstance(data, dict):
                schema_type = data.get('@type', '')
                if isinstance(schema_type, str):
                    if any(t in schema_type for t in ['LocalBusiness', 'Organization', 'Store', 'Restaurant']):
                        has_local_schema = True
                        break
        except:
            pass

    return {
        "phone_numbers_found": len(phones),
        "has_address_indicators": has_address_indicators,
        "has_local_business_schema": has_local_schema,
        "nap_score": sum([
            len(phones) > 0,
            has_address_indicators,
            has_local_schema
        ]),
    }
