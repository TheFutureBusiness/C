"""
Moduł obsługi robots.txt i sitemap.xml
"""
import re
import urllib.parse
import urllib.robotparser as rps
import aiohttp
from typing import List

from .fetcher import fetch


async def build_robots(session: aiohttp.ClientSession, root: str) -> rps.RobotFileParser:
    """
    Pobiera i parsuje plik robots.txt.

    Args:
        session: Sesja aiohttp
        root: URL główny witryny

    Returns:
        Parser robots.txt
    """
    rp = rps.RobotFileParser()
    p = urllib.parse.urlparse(root)
    robots_url = f"{p.scheme}://{p.netloc}/robots.txt"
    status, final, ct, text, headers = await fetch(session, robots_url)
    rp.set_url(robots_url)

    if status and status < 400 and isinstance(text, str) and not text.startswith("__ERROR__"):
        rp.parse(text.splitlines())
    else:
        rp.parse([])

    return rp


async def discover_sitemaps(session: aiohttp.ClientSession, root: str) -> List[str]:
    """
    Wyszukuje pliki sitemap.xml na stronie.

    Args:
        session: Sesja aiohttp
        root: URL główny witryny

    Returns:
        Lista URL-i znalezionych plików sitemap
    """
    p = urllib.parse.urlparse(root)
    candidates = [
        f"{p.scheme}://{p.netloc}/sitemap.xml",
        f"{p.scheme}://{p.netloc}/sitemap_index.xml",
    ]

    found = []
    for u in candidates:
        status, final, ct, text, headers = await fetch(session, u)
        if status and status < 400 and "xml" in ct.lower():
            found.append(final)

    return found


def parse_sitemap_xml(xml_text: str) -> List[str]:
    """
    Parsuje XML sitemap i wydobywa URL-e.

    Args:
        xml_text: Zawartość pliku sitemap XML

    Returns:
        Lista URL-i ze sitemap
    """
    urls = re.findall(r"<loc>(.*?)</loc>", xml_text, flags=re.I | re.S)
    return [u.strip() for u in urls if u.strip()]


async def fetch_and_parse_sitemaps(session: aiohttp.ClientSession, sitemap_urls: List[str]) -> List[str]:
    """
    Pobiera i parsuje wszystkie pliki sitemap.

    Args:
        session: Sesja aiohttp
        sitemap_urls: Lista URL-i sitemap do pobrania

    Returns:
        Lista wszystkich URL-i znalezionych w sitemap
    """
    urls: List[str] = []
    for sm in sitemap_urls:
        status, final, ct, text, headers = await fetch(session, sm)
        if status and status < 400 and isinstance(text, str) and not text.startswith("__ERROR__"):
            urls += parse_sitemap_xml(text)

    return list(dict.fromkeys(urls))
