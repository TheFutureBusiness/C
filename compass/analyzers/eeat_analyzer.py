"""
Analizator sygnałów E-E-A-T (Experience, Expertise, Authoritativeness, Trustworthiness)
"""
import re
from typing import Dict, Any
from bs4 import BeautifulSoup


def analyze_eeat_signals(soup: BeautifulSoup, text: str, url: str) -> Dict[str, Any]:
    """
    Analizuje sygnały E-E-A-T (Experience, Expertise, Authoritativeness, Trustworthiness).

    Args:
        soup: Obiekt BeautifulSoup z HTML
        text: Oczyszczony tekst strony
        url: URL strony

    Returns:
        Słownik z wynikami analizy E-E-A-T
    """
    # Sprawdzenie informacji o autorze
    author_indicators = ['author', 'autor', 'written by', 'by', 'redaktor']
    has_author = False

    for ind in author_indicators:
        if soup.find(attrs={'class': re.compile(ind, re.I)}) or \
                soup.find(attrs={'id': re.compile(ind, re.I)}) or \
                soup.find(attrs={'itemprop': ind}):
            has_author = True
            break

    # Sprawdzenie daty publikacji
    date_indicators = ['published', 'pubdate', 'datePublished', 'article:published_time']
    has_date = False

    for ind in date_indicators:
        if soup.find(attrs={'itemprop': ind}) or soup.find('time') or soup.find('meta', property=ind):
            has_date = True
            break

    # Sygnały ekspertyzy
    expertise_keywords = [
        'certyfikat', 'certificate', 'licencja', 'license', 'dyplom', 'diploma',
        'doświadczenie', 'experience', 'lat doświadczenia', 'years of experience'
    ]
    has_expertise_signals = any(keyword in text.lower() for keyword in expertise_keywords)

    # Linki do wiarygodnych źródeł
    external_links = soup.find_all('a', href=True)
    external_quality_domains = ['.gov', '.edu', '.org', 'wikipedia.org']
    has_quality_sources = any(
        any(domain in link.get('href', '') for domain in external_quality_domains)
        for link in external_links
    )

    # Informacje kontaktowe
    contact_indicators = ['kontakt', 'contact', 'email', 'telefon', 'phone', 'adres', 'address']
    has_contact_info = any(ind in text.lower() for ind in contact_indicators)

    # SSL/HTTPS
    has_ssl = url.startswith('https://')

    # Sygnały recenzji
    review_indicators = ['recenzja', 'review', 'opinia', 'opinion', 'rating', 'ocena']
    has_reviews = any(ind in text.lower() for ind in review_indicators)

    # Obliczenie wyniku E-E-A-T
    eeat_score = sum([
        has_author * 1.5,
        has_date * 1.0,
        has_expertise_signals * 1.5,
        has_quality_sources * 2.0,
        has_contact_info * 1.0,
        has_ssl * 1.0,
        has_reviews * 1.0,
    ])

    return {
        "has_author": has_author,
        "has_date": has_date,
        "has_expertise_signals": has_expertise_signals,
        "has_quality_external_links": has_quality_sources,
        "has_contact_info": has_contact_info,
        "has_ssl": has_ssl,
        "has_reviews": has_reviews,
        "eeat_score": round(eeat_score, 1),
        "eeat_max_score": 9.0,
        "eeat_percentage": round((eeat_score / 9.0) * 100, 1),
    }
