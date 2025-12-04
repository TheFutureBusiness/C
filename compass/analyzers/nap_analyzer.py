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

    WAŻNE: NAP jest oceniany na poziomie CAŁEJ STRONY, nie tylko głównej treści.
    Footer zawiera zwykle dane kontaktowe wspólne dla całej witryny.

    Args:
        soup: Obiekt BeautifulSoup z HTML
        text: Oczyszczony tekst strony

    Returns:
        Słownik z wynikami analizy NAP
    """
    # Wzorce numerów telefonów (rozszerzone o formaty niemieckie, polskie i międzynarodowe)
    phone_patterns = [
        r'\+?48\s?[\d\s\-]{9,}',           # Polski format
        r'\+?49\s?[\d\s\-/]{9,}',          # Niemiecki format
        r'\+?\d{1,3}\s?[\d\s\-/]{8,}',     # Międzynarodowy format
        r'\(\d{3,5}\)\s?[\d\s\-/]{5,}',    # Format z kodem obszaru w nawiasach
        r'\d{3,5}[\s\-/]?\d{3,5}[\s\-/]?\d{2,5}',  # Ogólny format
        r'tel[:\s]*[\d\s\-\+/\(\)]+',       # Format z prefiksem "tel:"
        r'telefon[:\s]*[\d\s\-\+/\(\)]+',   # Format z prefiksem "telefon:"
        r'phone[:\s]*[\d\s\-\+/\(\)]+',     # Format z prefiksem "phone:"
    ]

    # Sprawdzenie footer'a NAJPIERW (bo tam zwykle są dane kontaktowe)
    footer = soup.find('footer')
    footer_text = footer.get_text() if footer else ""

    # Sprawdzenie header'a (czasem logo z nazwą firmy)
    header = soup.find('header')
    header_text = header.get_text() if header else ""

    # Sprawdzenie sekcji kontaktowych
    contact_sections = soup.find_all(
        ['div', 'section', 'aside', 'address'],
        class_=lambda x: x and any(c in str(x).lower() for c in ['contact', 'kontakt', 'footer', 'address', 'info', 'company'])
    )
    contact_text = " ".join(s.get_text() for s in contact_sections)

    # Sprawdzenie elementów z id sugerującym kontakt
    contact_ids = soup.find_all(id=re.compile(r'contact|kontakt|footer|address', re.I))
    contact_id_text = " ".join(s.get_text() for s in contact_ids)

    # Łączymy WSZYSTKIE źródła tekstu
    combined_text = f"{text} {footer_text} {header_text} {contact_text} {contact_id_text}".lower()

    # Szukamy telefonów we wszystkich źródłach
    phones = []
    for pattern in phone_patterns:
        phones.extend(re.findall(pattern, combined_text, re.I))

    # Rozszerzone wskaźniki adresu (niemieckie, polskie, angielskie)
    address_indicators = [
        'ul.', 'ulica', 'al.', 'aleja',  # polskie
        'str.', 'straße', 'strasse', 'weg', 'platz', 'allee',  # niemieckie
        'street', 'avenue', 'road', 'lane', 'drive', 'blvd',  # angielskie
        # Kody pocztowe
        r'\d{2}-\d{3}',  # Polski format (00-000)
        r'\d{5}',        # Niemiecki format (12345)
    ]
    has_address_indicators = any(
        (re.search(ind, combined_text) if ind.startswith('\\') else ind in combined_text)
        for ind in address_indicators
    )

    # Sprawdzenie email w stopce/kontakcie
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    has_email = bool(re.search(email_pattern, combined_text))

    # Sprawdzenie nazwy firmy (Business Name)
    # Szukamy wskaźników nazwy firmy w tekście
    business_name_indicators = [
        'gmbh', 'sp. z o.o.', 'sp.z.o.o', 's.a.', 'ag', 'e.k.', 'ohg', 'kg',
        'ltd', 'inc', 'corp', 'llc', 'co.', 'company',
        '©',  # Symbol copyright często przy nazwie firmy
    ]
    has_business_name = any(ind in combined_text for ind in business_name_indicators)

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
    # Maksymalny wynik: 6 punktów
    has_phone_final = len(phones) > 0 or schema_has_phone
    has_address_final = has_address_indicators or schema_has_address
    has_email_final = has_email or schema_has_email
    has_schema_final = has_local_schema or has_org_schema

    nap_score = sum([
        has_phone_final,                             # Telefon (1 punkt)
        has_address_final,                           # Adres (1 punkt)
        has_email_final,                             # Email (1 punkt)
        has_schema_final,                            # Schema biznesowy (1 punkt)
        bool(footer_text.strip()),                   # Ma footer z treścią (1 punkt)
        has_business_name,                           # Nazwa firmy (1 punkt)
    ])

    # Normalizujemy do skali 0-3 dla kompatybilności wstecznej
    # Jeśli ma telefon/email + adres + schema = 3 punkty = dobry wynik
    # (stary kod oczekuje nap_score < 2 jako "poor")
    if nap_score >= 4:
        normalized_score = 3
    elif nap_score >= 3:
        normalized_score = 2
    elif nap_score >= 2:
        normalized_score = 2
    else:
        normalized_score = nap_score

    return {
        "phone_numbers_found": len(phones),
        "has_phone": has_phone_final,
        "has_address_indicators": has_address_final,
        "has_email": has_email_final,
        "has_business_name": has_business_name,
        "has_local_business_schema": has_local_schema,
        "has_organization_schema": has_org_schema,
        "has_footer_content": bool(footer_text.strip()),
        "nap_score": normalized_score,
        "nap_details_score": nap_score,  # Szczegółowy wynik (0-6)
    }
