"""
Analizator wyników crawlingu - znajdowanie duplikatów i problemów
"""
from collections import defaultdict
from typing import Dict, Any, List, Tuple
from datetime import datetime


def is_noindex_page(data: Dict[str, Any]) -> bool:
    """
    Sprawdza czy strona ma ustawiony noindex.

    Args:
        data: Dane strony

    Returns:
        True jeśli strona ma noindex
    """
    robots_meta = data.get('robots_meta', '').lower()
    return 'noindex' in robots_meta


def find_duplicates(all_pages: Dict[str, Any]) -> Dict[str, List]:
    """
    Znajduje duplikaty title i description w przeanalizowanych stronach.
    Pomija strony wykluczone i noindex.

    Args:
        all_pages: Słownik wszystkich przeanalizowanych stron

    Returns:
        Słownik z duplikatami title i description
    """
    title_map = defaultdict(list)
    desc_map = defaultdict(list)

    for url, data in all_pages.items():
        # Pomijamy strony wykluczone i noindex
        if data.get('is_excluded'):
            continue
        if is_noindex_page(data):
            continue

        title = data.get('title', '').strip()
        desc = data.get('meta_description', '').strip()

        if title:
            title_map[title].append(url)
        if desc:
            desc_map[desc].append(url)

    duplicates = {
        "title": {k: v for k, v in title_map.items() if len(v) > 1},
        "description": {k: v for k, v in desc_map.items() if len(v) > 1},
    }

    return duplicates


def analyze_issues(all_pages: Dict[str, Any]) -> Dict[str, Any]:
    """
    Analizuje wszystkie strony i znajduje problemy SEO/AEO/GEO/Security.
    Pomija strony wykluczone i noindex.

    Args:
        all_pages: Słownik wszystkich przeanalizowanych stron

    Returns:
        Słownik ze znalezionymi problemami w kategoriach
    """
    issues = {
        "critical_errors": [],
        "missing_title": [],
        "missing_description": [],
        "missing_canonical": [],
        "missing_h1": [],
        "multiple_h1": [],
        "images_no_alt": [],
        "no_viewport": [],
        "no_og_tags": [],
        "no_twitter_cards": [],
        "missing_schema": [],
        "weak_eeat": [],
        "poor_local_seo": [],
        "thin_content": [],
        "title_issues": [],
        "description_issues": [],
        "no_ssl": [],
        "missing_security_headers": [],
        "poor_security": [],
        "mixed_content": [],
        "info_disclosure": [],
    }

    for url, data in all_pages.items():
        # Pomijamy strony wykluczone i noindex
        if data.get('is_excluded'):
            continue
        if is_noindex_page(data):
            continue

        ct = data.get('content_type', '') or ''
        status = data.get('status')

        # Błędy krytyczne
        if data.get('error') or (status and 400 <= status < 500):
            if 'text/html' in ct or not ct:
                issues['critical_errors'].append({
                    'url': url,
                    'status': status,
                    'error': data.get('error', '')
                })

        if not status or status >= 400:
            continue

        # Brak title
        if not data.get('title'):
            issues['missing_title'].append(url)
        else:
            meta_scores = data.get('meta_scores', {})
            if meta_scores.get('title_too_short') or meta_scores.get('title_too_long'):
                issues['title_issues'].append({
                    'url': url,
                    'title': data.get('title'),
                    'length': meta_scores.get('title_length'),
                    'too_short': meta_scores.get('title_too_short'),
                    'too_long': meta_scores.get('title_too_long'),
                })

        # Brak description
        if not data.get('meta_description'):
            issues['missing_description'].append(url)
        else:
            meta_scores = data.get('meta_scores', {})
            if meta_scores.get('desc_too_short') or meta_scores.get('desc_too_long'):
                issues['description_issues'].append({
                    'url': url,
                    'description': data.get('meta_description')[:100],
                    'length': meta_scores.get('desc_length'),
                    'too_short': meta_scores.get('desc_too_short'),
                    'too_long': meta_scores.get('desc_too_long'),
                })

        # Brak canonical
        if not data.get('canonical'):
            issues['missing_canonical'].append(url)

        # Problemy z H1
        h1_count = data.get('h1_count', 0)
        if h1_count == 0:
            issues['missing_h1'].append(url)
        elif h1_count > 1:
            issues['multiple_h1'].append({
                'url': url,
                'h1_count': h1_count,
                'h1_list': data.get('h1', [])
            })

        # Obrazy bez ALT
        if data.get('img_without_alt', 0) > 0:
            issues['images_no_alt'].append({
                'url': url,
                'missing_alt': data.get('img_without_alt'),
                'total_images': data.get('img_total'),
                'alt_ratio': data.get('img_alt_ratio'),
            })

        # Brak viewport (mobile-friendly)
        if not data.get('is_mobile_friendly'):
            issues['no_viewport'].append(url)

        # Brak Open Graph
        if not data.get('has_og_image') or not data.get('has_og_title'):
            issues['no_og_tags'].append({
                'url': url,
                'has_og_image': data.get('has_og_image'),
                'has_og_title': data.get('has_og_title'),
                'has_og_description': data.get('has_og_description'),
            })

        # Brak Twitter Cards
        if not data.get('has_twitter_card'):
            issues['no_twitter_cards'].append(url)

        # Brak Schema.org
        if data.get('schema_count', 0) == 0:
            issues['missing_schema'].append(url)

        # Słabe sygnały E-E-A-T
        eeat = data.get('eeat_signals', {})
        if eeat.get('eeat_percentage', 100) < 50:
            issues['weak_eeat'].append({
                'url': url,
                'eeat_score': eeat.get('eeat_score'),
                'eeat_percentage': eeat.get('eeat_percentage'),
                'missing': [k for k, v in eeat.items() if k.startswith('has_') and not v]
            })

        # Słabe NAP (Local SEO)
        nap = data.get('nap_signals', {})
        if nap.get('nap_score', 0) < 2:
            issues['poor_local_seo'].append({
                'url': url,
                'nap_score': nap.get('nap_score'),
                'phone_numbers': nap.get('phone_numbers_found'),
                'has_address': nap.get('has_address_indicators'),
                'has_local_schema': nap.get('has_local_business_schema'),
            })

        # Thin content
        word_count = data.get('word_count', 0) or 0
        if word_count < 300 and word_count > 0:
            issues['thin_content'].append({
                'url': url,
                'word_count': word_count,
                'text_len': data.get('text_len', 0)
            })

        # Problemy z bezpieczeństwem
        security = data.get('security', {})

        # Brak SSL
        if not security.get('has_ssl'):
            issues['no_ssl'].append(url)

        # Słabe bezpieczeństwo
        sec_percentage = security.get('security_percentage', 100)
        if sec_percentage < 50:
            issues['poor_security'].append({
                'url': url,
                'security_percentage': sec_percentage,
                'security_level': security.get('security_level'),
                'missing_headers': security.get('missing_critical', []),
            })

        # Brakujące nagłówki bezpieczeństwa
        headers_count = security.get('headers_count', 0)
        if headers_count < 3:
            issues['missing_security_headers'].append({
                'url': url,
                'headers_count': headers_count,
                'missing_critical': security.get('missing_critical', []),
            })

        # Mixed content
        if security.get('has_mixed_content'):
            issues['mixed_content'].append(url)

        # Information disclosure
        if security.get('exposes_server_info') or security.get('exposes_tech_stack'):
            issues['info_disclosure'].append({
                'url': url,
                'server_header': security.get('server_header'),
                'powered_by': security.get('powered_by_header'),
            })

    return issues


def calculate_overall_score(summary: Dict[str, Any]) -> Tuple[int, str]:
    """
    Oblicza ogólny wynik audytu na podstawie różnych metryk.

    Args:
        summary: Słownik z podsumowaniem audytu

    Returns:
        Krotka: (wynik 0-100, ocena tekstowa)
    """
    pages = max(1, summary["pages_analyzed"])

    # Dostępność stron
    availability = summary["pages_ok"] / pages

    # Jakość meta tagów
    meta_ok_pages = pages - (summary["missing_title"] + summary["missing_description"])
    meta_quality = max(0.0, (meta_ok_pages - 0.25 * (
        summary["title_issues"] + summary["description_issues"]
    )) / pages)

    # Mobile-friendly
    mobile = summary["mobile_percentage"] / 100.0

    # Schema.org
    schema = (summary["pages_with_schema"] / pages) if pages else 0.0

    # E-E-A-T
    eeat = summary["avg_eeat_score"] / 100.0

    # Bezpieczeństwo
    security = summary["avg_security_score"] / 100.0
    if summary["pages_no_ssl"] > 0:
        security = max(0.0, security - 0.10)

    # Wagi dla różnych kategorii
    W = {
        "availability": 0.30,
        "meta": 0.15,
        "mobile": 0.15,
        "schema": 0.10,
        "eeat": 0.10,
        "security": 0.20
    }

    # Obliczenie wyniku
    score = (
        availability * W["availability"] +
        meta_quality * W["meta"] +
        mobile * W["mobile"] +
        schema * W["schema"] +
        eeat * W["eeat"] +
        security * W["security"]
    ) * 100.0

    score_int = max(0, min(100, int(round(score))))

    # Określenie oceny
    if score_int >= 85:
        grade = "Excellent"
    elif score_int >= 70:
        grade = "Good"
    elif score_int >= 50:
        grade = "Fair"
    else:
        grade = "Poor"

    return score_int, grade


def calculate_summary(all_pages: Dict[str, Any], issues: Dict[str, Any], duplicates: Dict) -> Dict[str, Any]:
    """
    Oblicza podsumowanie statystyk audytu.
    Pomija strony wykluczone i noindex.

    Args:
        all_pages: Słownik wszystkich przeanalizowanych stron
        issues: Słownik ze znalezionymi problemami
        duplicates: Słownik z duplikatami

    Returns:
        Słownik z podsumowaniem audytu
    """
    from compass.config import START_URL

    # Filtrujemy strony: wykluczamy is_excluded oraz noindex
    analyzed_pages = {
        url: data for url, data in all_pages.items()
        if not data.get('is_excluded') and not is_noindex_page(data)
    }
    excluded_count = len([p for p in all_pages.values() if p.get('is_excluded')])
    noindex_count = len([p for p in all_pages.values() if is_noindex_page(p)])

    pages_with_errors = len(issues['critical_errors'])
    pages_ok = len([p for p in analyzed_pages.values() if p.get('status') == 200])

    # Mobile-friendly
    mobile_friendly = sum(1 for p in analyzed_pages.values() if p.get('is_mobile_friendly'))
    mobile_percentage = round(mobile_friendly / max(1, len(analyzed_pages)) * 100, 1)

    # Schema.org
    pages_with_schema = sum(1 for p in analyzed_pages.values() if p.get('schema_count', 0) > 0)
    avg_schema_types = sum(p.get('schema_count', 0) for p in analyzed_pages.values()) / max(1, len(analyzed_pages))

    # E-E-A-T
    avg_eeat = sum(
        p.get('eeat_signals', {}).get('eeat_percentage', 0)
        for p in analyzed_pages.values()
    ) / max(1, len(analyzed_pages))

    # Local SEO (NAP)
    local_optimized = sum(
        1 for p in analyzed_pages.values()
        if p.get('nap_signals', {}).get('nap_score', 0) >= 2
    )

    # Bezpieczeństwo
    avg_security = sum(
        p.get('security', {}).get('security_percentage', 0)
        for p in analyzed_pages.values()
    ) / max(1, len(analyzed_pages))

    pages_with_ssl = sum(1 for p in analyzed_pages.values() if p.get('security', {}).get('has_ssl'))
    ssl_percentage = round(pages_with_ssl / max(1, len(analyzed_pages)) * 100, 1)

    result = {
        "start_url": START_URL,
        "pages_crawled": len(all_pages),
        "pages_analyzed": len(analyzed_pages),
        "pages_excluded": excluded_count,
        "pages_noindex": noindex_count,
        "pages_ok": pages_ok,
        "pages_with_errors": pages_with_errors,
        "missing_title": len(issues['missing_title']),
        "missing_description": len(issues['missing_description']),
        "title_issues": len(issues['title_issues']),
        "description_issues": len(issues['description_issues']),
        "duplicate_titles": len(duplicates['title']),
        "duplicate_descriptions": len(duplicates['description']),
        "missing_canonical": len(issues['missing_canonical']),
        "missing_h1": len(issues['missing_h1']),
        "multiple_h1": len(issues['multiple_h1']),
        "pages_with_alt_issues": len(issues['images_no_alt']),
        "total_images_without_alt": sum(i['missing_alt'] for i in issues['images_no_alt']),
        "mobile_friendly_pages": mobile_friendly,
        "mobile_percentage": mobile_percentage,
        "pages_without_viewport": len(issues['no_viewport']),
        "pages_without_og": len(issues['no_og_tags']),
        "pages_without_twitter": len(issues['no_twitter_cards']),
        "pages_with_schema": pages_with_schema,
        "pages_without_schema": len(issues['missing_schema']),
        "avg_schema_types": round(avg_schema_types, 1),
        "avg_eeat_score": round(avg_eeat, 1),
        "pages_weak_eeat": len(issues['weak_eeat']),
        "local_optimized_pages": local_optimized,
        "pages_poor_local_seo": len(issues['poor_local_seo']),
        "thin_content_pages": len(issues['thin_content']),
        "avg_security_score": round(avg_security, 1),
        "pages_with_ssl": pages_with_ssl,
        "ssl_percentage": ssl_percentage,
        "pages_no_ssl": len(issues['no_ssl']),
        "pages_poor_security": len(issues['poor_security']),
        "pages_missing_security_headers": len(issues['missing_security_headers']),
        "pages_with_mixed_content": len(issues['mixed_content']),
        "pages_with_info_disclosure": len(issues['info_disclosure']),
        "generated_at": datetime.now().strftime("%Y-%m-%d"),
    }

    # Obliczenie ogólnego wyniku
    overall_score, overall_grade = calculate_overall_score(result)
    result["overall_score"] = overall_score
    result["overall_grade"] = overall_grade

    return result
