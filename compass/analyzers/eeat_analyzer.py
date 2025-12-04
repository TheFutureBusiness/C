"""
Analizator sygnałów E-E-A-T (Experience, Expertise, Authoritativeness, Trustworthiness)
"""
import re
from typing import Dict, Any
from bs4 import BeautifulSoup


def detect_page_type(url: str, soup: BeautifulSoup) -> str:
    """
    Wykrywa typ strony na podstawie URL i struktury.

    Returns:
        Typ strony: 'blog', 'service', 'contact', 'about', 'home', 'legal', 'other'
    """
    url_lower = url.lower()
    path = url_lower.split('/')[-2] if url_lower.endswith('/') else url_lower.split('/')[-1]

    # Strona główna
    if url_lower.rstrip('/').endswith(('.de', '.com', '.pl', '.net', '.org')) or path == '':
        return 'home'

    # Blog/artykuły
    blog_indicators = ['blog', 'news', 'artikel', 'article', 'post', 'beitrag', 'wpis', 'aktualnosci']
    if any(ind in url_lower for ind in blog_indicators):
        return 'blog'

    # Strony prawne
    legal_indicators = ['impressum', 'datenschutz', 'privacy', 'agb', 'terms', 'legal', 'regulamin', 'polityka']
    if any(ind in url_lower for ind in legal_indicators):
        return 'legal'

    # Strona kontaktowa
    if 'kontakt' in url_lower or 'contact' in url_lower:
        return 'contact'

    # O nas/O firmie
    about_indicators = ['about', 'ueber-uns', 'uber-uns', 'o-nas', 'team', 'firma', 'unternehmen']
    if any(ind in url_lower for ind in about_indicators):
        return 'about'

    # Usługi/Produkty
    service_indicators = ['service', 'leistung', 'produkt', 'product', 'angebot', 'uslugi']
    if any(ind in url_lower for ind in service_indicators):
        return 'service'

    return 'other'


def analyze_eeat_signals(soup: BeautifulSoup, text: str, url: str) -> Dict[str, Any]:
    """
    Analizuje sygnały E-E-A-T (Experience, Expertise, Authoritativeness, Trustworthiness).

    WAŻNE: E-E-A-T powinno być oceniane kontekstowo w zależności od typu strony:
    - Blog/artykuły: wymagają autora i daty publikacji
    - Strony usług: nie wymagają autora ani daty
    - Strona O nas: powinna mieć certyfikaty, doświadczenie
    - Wszystkie strony: HTTPS, dane kontaktowe (w stopce)

    Args:
        soup: Obiekt BeautifulSoup z HTML
        text: Oczyszczony tekst strony
        url: URL strony

    Returns:
        Słownik z wynikami analizy E-E-A-T
    """
    page_type = detect_page_type(url, soup)

    # Sprawdzenie informacji o autorze
    author_indicators = ['author', 'autor', 'written by', 'by', 'redaktor', 'verfasser']
    has_author = False

    for ind in author_indicators:
        if soup.find(attrs={'class': re.compile(ind, re.I)}) or \
                soup.find(attrs={'id': re.compile(ind, re.I)}) or \
                soup.find(attrs={'itemprop': ind}):
            has_author = True
            break

    # Sprawdzenie daty publikacji
    date_indicators = ['published', 'pubdate', 'datePublished', 'article:published_time', 'dateModified']
    has_date = False

    for ind in date_indicators:
        if soup.find(attrs={'itemprop': ind}) or soup.find('time') or soup.find('meta', property=ind):
            has_date = True
            break

    # Sygnały ekspertyzy (rozszerzone o niemieckie)
    expertise_keywords = [
        'certyfikat', 'certificate', 'zertifikat', 'zertifiziert',
        'licencja', 'license', 'lizenz',
        'dyplom', 'diploma', 'diplom',
        'doświadczenie', 'experience', 'erfahrung',
        'lat doświadczenia', 'years of experience', 'jahre erfahrung',
        'nagroda', 'award', 'auszeichnung', 'preis'
    ]
    has_expertise_signals = any(keyword in text.lower() for keyword in expertise_keywords)

    # Linki do wiarygodnych źródeł
    external_links = soup.find_all('a', href=True)
    external_quality_domains = ['.gov', '.edu', '.org', 'wikipedia.org', '.gov.de', '.gov.pl']
    has_quality_sources = any(
        any(domain in link.get('href', '') for domain in external_quality_domains)
        for link in external_links
    )

    # Informacje kontaktowe (sprawdzamy też footer)
    footer = soup.find('footer')
    footer_text = footer.get_text().lower() if footer else ""
    combined_text = f"{text.lower()} {footer_text}"

    contact_indicators = ['kontakt', 'contact', 'email', 'telefon', 'phone', 'adres', 'address', 'tel:', 'e-mail:']
    has_contact_info = any(ind in combined_text for ind in contact_indicators)

    # SSL/HTTPS
    has_ssl = url.startswith('https://')

    # Sygnały recenzji (rozszerzone o niemieckie)
    review_indicators = [
        'recenzja', 'review', 'rezension', 'bewertung',
        'opinia', 'opinion', 'meinung',
        'rating', 'ocena', 'note',
        'testimonial', 'kundenstimmen', 'referenzen'
    ]
    has_reviews = any(ind in text.lower() for ind in review_indicators)

    # Obliczenie wyniku E-E-A-T z uwzględnieniem typu strony
    # Różne wagi dla różnych typów stron
    if page_type == 'blog':
        # Blog wymaga autora i daty
        eeat_score = sum([
            has_author * 2.0,           # Autor ważny dla bloga
            has_date * 1.5,             # Data ważna dla bloga
            has_expertise_signals * 1.0,
            has_quality_sources * 2.0,
            has_contact_info * 0.5,     # Mniej ważne dla artykułu
            has_ssl * 1.0,
            has_reviews * 1.0,
        ])
        max_score = 9.0
    elif page_type in ('service', 'home', 'other'):
        # Strony usług nie wymagają autora ani daty
        eeat_score = sum([
            # Autor i data nie są wymagane - dajemy punkty domyślnie
            1.5,  # Brak wymogu autora
            1.0,  # Brak wymogu daty
            has_expertise_signals * 1.5,
            has_quality_sources * 1.5,
            has_contact_info * 1.5,     # Kontakt ważniejszy dla usług
            has_ssl * 1.0,
            has_reviews * 1.0,
        ])
        max_score = 9.0
    elif page_type == 'about':
        # Strona O nas - ekspertyza i certyfikaty ważne
        eeat_score = sum([
            has_author * 1.0,           # Mile widziany, ale nie wymagany
            1.0,                         # Data nie wymagana
            has_expertise_signals * 2.5, # Bardzo ważne dla About
            has_quality_sources * 1.0,
            has_contact_info * 1.5,
            has_ssl * 1.0,
            has_reviews * 1.0,
        ])
        max_score = 9.0
    elif page_type in ('legal', 'contact'):
        # Strony prawne/kontakt - głównie SSL i dane kontaktowe
        eeat_score = sum([
            1.5,  # Brak wymogu autora
            1.0,  # Brak wymogu daty
            1.0,  # Brak wymogu ekspertyzy
            1.0,  # Brak wymogu źródeł
            has_contact_info * 2.5,     # Bardzo ważne dla kontaktu
            has_ssl * 1.0,
            1.0,  # Brak wymogu recenzji
        ])
        max_score = 9.0
    else:
        # Domyślna logika
        eeat_score = sum([
            has_author * 1.5,
            has_date * 1.0,
            has_expertise_signals * 1.5,
            has_quality_sources * 2.0,
            has_contact_info * 1.0,
            has_ssl * 1.0,
            has_reviews * 1.0,
        ])
        max_score = 9.0

    return {
        "page_type": page_type,
        "has_author": has_author,
        "has_date": has_date,
        "has_expertise_signals": has_expertise_signals,
        "has_quality_external_links": has_quality_sources,
        "has_contact_info": has_contact_info,
        "has_ssl": has_ssl,
        "has_reviews": has_reviews,
        "eeat_score": round(eeat_score, 1),
        "eeat_max_score": max_score,
        "eeat_percentage": round((eeat_score / max_score) * 100, 1),
    }
