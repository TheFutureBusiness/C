"""
Analizator bezpieczeństwa (nagłówki HTTP, SSL, mixed content)
"""
import re
from typing import Dict, Any


def analyze_security_headers(headers: Dict[str, str], url: str, html: str = "") -> Dict[str, Any]:
    """
    Analizuje nagłówki bezpieczeństwa HTTP i inne aspekty bezpieczeństwa strony.

    Args:
        headers: Słownik nagłówków HTTP
        url: URL strony
        html: Kod HTML strony (opcjonalnie, do sprawdzenia mixed content)

    Returns:
        Słownik z wynikami analizy bezpieczeństwa
    """
    headers_lower = {k.lower(): v for k, v in headers.items()}

    # Definicja sprawdzanych nagłówków bezpieczeństwa
    security_checks = {
        "hsts": {
            "name": "HTTP Strict Transport Security (HSTS)",
            "header": "strict-transport-security",
            "present": False,
            "value": "",
            "score": 0,
            "max_score": 15,
            "severity": "high",
            "description": ""
        },
        "x_frame_options": {
            "name": "X-Frame-Options",
            "header": "x-frame-options",
            "present": False,
            "value": "",
            "score": 0,
            "max_score": 10,
            "severity": "high",
            "description": ""
        },
        "x_content_type_options": {
            "name": "X-Content-Type-Options",
            "header": "x-content-type-options",
            "present": False,
            "value": "",
            "score": 0,
            "max_score": 10,
            "severity": "medium",
            "description": ""
        },
        "content_security_policy": {
            "name": "Content-Security-Policy (CSP)",
            "header": "content-security-policy",
            "present": False,
            "value": "",
            "score": 0,
            "max_score": 20,
            "severity": "high",
            "description": ""
        },
        "x_xss_protection": {
            "name": "X-XSS-Protection",
            "header": "x-xss-protection",
            "present": False,
            "value": "",
            "score": 0,
            "max_score": 5,
            "severity": "low",
            "description": ""
        },
        "referrer_policy": {
            "name": "Referrer-Policy",
            "header": "referrer-policy",
            "present": False,
            "value": "",
            "score": 0,
            "max_score": 10,
            "severity": "medium",
            "description": ""
        },
        "permissions_policy": {
            "name": "Permissions-Policy",
            "header": "permissions-policy",
            "present": False,
            "value": "",
            "score": 0,
            "max_score": 10,
            "severity": "low",
            "description": ""
        },
    }

    # Sprawdzenie obecności nagłówków
    total_score = 0
    max_possible_score = sum(check["max_score"] for check in security_checks.values())

    for key, check in security_checks.items():
        header_name = check["header"]
        if header_name in headers_lower:
            check["present"] = True
            check["value"] = headers_lower[header_name]
            check["score"] = check["max_score"]
        total_score += check["score"]

    # Sprawdzenie SSL/HTTPS
    has_ssl = url.startswith('https://')
    ssl_score = 20 if has_ssl else 0
    max_possible_score += 20
    total_score += ssl_score

    # Sprawdzenie mixed content (zasoby HTTP na stronach HTTPS)
    has_mixed_content = False
    if has_ssl and html:
        http_resources = re.findall(r'src=["\']http://[^"\']+["\']', html, re.I)
        http_resources += re.findall(r'href=["\']http://[^"\']+\.(?:css|js)["\']', html, re.I)
        has_mixed_content = len(http_resources) > 0

    # Sprawdzenie, czy serwer ujawnia informacje techniczne
    server_header = headers_lower.get('server', '')
    exposes_server_info = bool(server_header and server_header.lower() not in ['', 'cloudflare'])

    powered_by = headers_lower.get('x-powered-by', '')
    exposes_tech_stack = bool(powered_by)

    # Obliczenie procentowego wyniku bezpieczeństwa
    security_percentage = round((total_score / max_possible_score) * 100, 1)

    # Określenie poziomu bezpieczeństwa
    if security_percentage >= 90:
        security_level = "Excellent"
        security_emoji = "🟢"
    elif security_percentage >= 70:
        security_level = "Good"
        security_emoji = "🟡"
    elif security_percentage >= 50:
        security_level = "Fair"
        security_emoji = "🟠"
    else:
        security_level = "Poor"
        security_emoji = "🔴"

    # Lista brakujących krytycznych nagłówków
    missing_critical = [
        check["name"] for check in security_checks.values()
        if not check["present"] and check["severity"] in ["high", "medium"]
    ]

    # Określenie rzeczywistego ryzyka (nie tylko brakujące nagłówki)
    # WAŻNE: Brakujące nagłówki to "hardening" - nie oznaczają aktywnej podatności
    has_critical_issues = not has_ssl or has_mixed_content  # To są faktyczne problemy
    has_hardening_issues = len(missing_critical) > 0  # To jest brak "hardening"

    # Generowanie spójnego opisu bezpieczeństwa
    if has_critical_issues:
        if not has_ssl:
            security_description = "Brak HTTPS - dane nie są szyfrowane"
        else:
            security_description = "Mixed content - niektóre zasoby ładowane przez HTTP"
        security_risk = "high"
    elif security_percentage < 50:
        security_description = "Brak zalecanych nagłówków bezpieczeństwa (hardening)"
        security_risk = "medium"
    elif security_percentage < 70:
        security_description = "Częściowy hardening - brakuje niektórych nagłówków"
        security_risk = "low"
    else:
        security_description = "Dobre zabezpieczenia"
        security_risk = "none"

    return {
        "security_checks": security_checks,
        "has_ssl": has_ssl,
        "ssl_score": ssl_score,
        "has_mixed_content": has_mixed_content,
        "exposes_server_info": exposes_server_info,
        "server_header": server_header,
        "exposes_tech_stack": exposes_tech_stack,
        "powered_by_header": powered_by,
        "total_score": total_score,
        "max_score": max_possible_score,
        "security_percentage": security_percentage,
        "security_level": security_level,
        "security_emoji": security_emoji,
        "missing_critical": missing_critical,
        "headers_count": len([c for c in security_checks.values() if c["present"]]),
        "total_headers": len(security_checks),
        # Nowe pola dla spójnego raportowania
        "has_critical_issues": has_critical_issues,
        "has_hardening_issues": has_hardening_issues,
        "security_description": security_description,
        "security_risk": security_risk,  # "high", "medium", "low", "none"
    }
