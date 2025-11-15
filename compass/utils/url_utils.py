"""
Narzędzia do przetwarzania URL-i
"""
import re
import urllib.parse
import tldextract
from url_normalize import url_normalize
from compass.config import DOMAIN_SCOPE, EXCLUDED_PATTERNS


def same_site(u1: str, u2: str) -> bool:
    """
    Sprawdza, czy dwa URL-e należą do tej samej witryny.

    Args:
        u1: Pierwszy URL
        u2: Drugi URL

    Returns:
        True jeśli URL-e należą do tej samej witryny
    """
    a = urllib.parse.urlparse(u1)
    b = urllib.parse.urlparse(u2)

    if DOMAIN_SCOPE == "sub":
        ea = tldextract.extract(a.netloc)
        eb = tldextract.extract(b.netloc)
        return (a.scheme in ("http", "https")) and (ea.registered_domain == eb.registered_domain)
    else:
        return (a.scheme, a.netloc) == (b.scheme, b.netloc)


def absolutize(base: str, link: str) -> str:
    """
    Konwertuje względny URL na absolutny.

    Args:
        base: Bazowy URL
        link: Link do konwersji

    Returns:
        Znormalizowany absolutny URL
    """
    return url_normalize(urllib.parse.urljoin(base, link))


def is_excluded_url(url: str) -> bool:
    """
    Sprawdza, czy URL pasuje do wzorców wykluczonych.

    Args:
        url: URL do sprawdzenia

    Returns:
        True jeśli URL powinien być wykluczony
    """
    parsed = urllib.parse.urlparse(url)
    path = parsed.path.lower()

    for pattern in EXCLUDED_PATTERNS:
        if re.search(pattern, path):
            return True

    return False
