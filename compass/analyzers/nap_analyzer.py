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
    Sprawdza zarówno tekst strony, footer, jak i dane strukturalne Schema.org.

    Args:
        soup: Obiekt BeautifulSoup z HTML
        text: Oczyszczony tekst strony

    Returns:
        Słownik z wynikami analizy NAP
    """
    # Wzorce numerów telefonów (rozszerzone o formaty niemieckie i międzynarodowe)
    phone_patterns = [
        r'\+?48\s?[\d\s\-]{9,}',           # Polski format
        r'\+?49\s?[\d\s\-/]{9,}',          # Niemiecki format
        r'\+?\d{1,3}\s?[\d\s\-/]{8,}',     # Międzynarodowy format
        r'\(\d{3,5}\)\s?[\d\s\-/]{5,}',    # Format z kodem obszaru w nawiasach
        r'\d{3,5}[\s\-/]?\d{3,5}[\s\-/]?\d{2,5}',  # Ogólny format
    ]

    phones = []
    for pattern in phone_patterns:
        phones.extend(re.findall(pattern, text))

    # Sprawdzenie footer'a osobno (bo tam często są dane kontaktowe)
    footer = soup.find('footer')
    footer_text = footer.get_text() if footer else ""

    # Sprawdzenie sekcji kontaktowych
    contact_sections = soup.find_all(
        ['div', 'section', 'aside'],
        class_=lambda x: x and any(c in x.lower() for c in ['contact', 'kontakt', 'footer', 'address'])
    )
    contact_text = " ".join(s.get_text() for s in contact_sections)

    combined_text = f"{text} {footer_text} {contact_text}".lower()

    # Rozszerzone wskaźniki adresu (niemieckie, polskie, angielskie)
    address_indicators = [
        'ul.', 'ulica', 'al.', 'aleja',  # polskie
        'str.', 'straße', 'strasse', 'weg', 'platz', 'allee',  # niemieckie
        'street', 'avenue', 'road', 'lane', 'drive', 'blvd',  # angielskie
    ]
    has_address_indicators = any(ind in combined_text for ind in address_indicators)

    # Sprawdzenie email w stopce/kontakcie
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    has_email = bool(re.search(email_pattern, combined_text))

    # Sprawdzenie Schema.org dla Organization/LocalBusiness
    schema_scripts = soup.find_all('script', type='application/ld+json')
    has_org_schema = False
    has_local_schema = False
    schema_has_phone = False
    schema_has_address = False
    schema_has_email = False

    for script in schema_scripts:
        try:
            script_text = script.string
            if not script_text:
                continue

            data = json.loads(script_text)

            # Obsługa @graph (wiele schematów w jednym skrypcie)
            items = [data]
            if isinstance(data, dict) and '@graph' in data:
                items = data.get('@graph', [])
            elif isinstance(data, list):
                items = data

            for item in items:
                if not isinstance(item, dict):
                    continue

                schema_type = item.get('@type', '')
                if isinstance(schema_type, list):
                    schema_type = ' '.join(schema_type)

                # Sprawdzenie typu schematu
                if any(t in str(schema_type) for t in ['LocalBusiness', 'Store', 'Restaurant', 'Hotel', 'Dentist', 'LegalService']):
                    has_local_schema = True
                if any(t in str(schema_type) for t in ['Organization', 'Corporation', 'Company']):
                    has_org_schema = True

                # Sprawdzenie danych kontaktowych w schemacie
                if item.get('telephone') or item.get('phone'):
                    schema_has_phone = True
                if item.get('address') or item.get('location'):
                    schema_has_address = True
                if item.get('email'):
                    schema_has_email = True

        except (json.JSONDecodeError, TypeError, AttributeError):
            pass

    # Rozszerzone obliczenie NAP score
    # Maksymalny wynik: 5 punktów
    nap_score = sum([
        len(phones) > 0 or schema_has_phone,       # Telefon (1 punkt)
        has_address_indicators or schema_has_address,  # Adres (1 punkt)
        has_email or schema_has_email,              # Email (1 punkt)
        has_local_schema or has_org_schema,         # Schema biznesowy (1 punkt)
        bool(footer_text.strip()),                   # Ma footer z treścią (1 punkt)
    ])

    # Normalizujemy do skali 0-3 dla kompatybilności wstecznej
    # (stary kod oczekuje nap_score < 2 jako "poor")
    normalized_score = min(3, nap_score) if nap_score >= 3 else (2 if nap_score >= 2 else nap_score)

    return {
        "phone_numbers_found": len(phones),
        "has_phone": len(phones) > 0 or schema_has_phone,
        "has_address_indicators": has_address_indicators or schema_has_address,
        "has_email": has_email or schema_has_email,
        "has_local_business_schema": has_local_schema,
        "has_organization_schema": has_org_schema,
        "has_footer_content": bool(footer_text.strip()),
        "nap_score": normalized_score,
        "nap_details_score": nap_score,  # Szczegółowy wynik (0-5)
    }
