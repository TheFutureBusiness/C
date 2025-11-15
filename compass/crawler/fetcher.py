"""
Moduł pobierania i parsowania stron HTML
"""
import re
import aiohttp
from typing import Dict, Any, Tuple, Optional
from bs4 import BeautifulSoup
import extruct
from w3lib.html import get_base_url

from compass.config import TIMEOUT, USER_AGENT
from compass.utils import clean_text, absolutize, same_site, is_excluded_url
from compass.analyzers import (
    calculate_meta_score,
    extract_nap_signals,
    analyze_eeat_signals,
    analyze_security_headers,
)


async def fetch(session: aiohttp.ClientSession, url: str) -> Tuple[Optional[int], str, str, str, Dict[str, str]]:
    """
    Pobiera zawartość strony przy użyciu aiohttp.

    Args:
        session: Sesja aiohttp
        url: URL do pobrania

    Returns:
        Krotka: (status_code, final_url, content_type, html, headers)
    """
    try:
        async with session.get(
            url,
            allow_redirects=True,
            timeout=TIMEOUT,
            headers={"User-Agent": USER_AGENT}
        ) as r:
            ct = r.headers.get("Content-Type", "")
            txt = await r.text(errors="ignore")
            return r.status, str(r.url), ct, txt, {k: v for k, v in r.headers.items()}
    except Exception as e:
        return None, url, "", f"__ERROR__:{e}", {}


def parse_page(html: str, url: str) -> Dict[str, Any]:
    """
    Parsuje stronę HTML i wydobywa wszystkie istotne dane SEO/AEO/GEO.

    Args:
        html: Kod HTML strony
        url: URL strony

    Returns:
        Słownik z wynikami analizy strony
    """
    soup = BeautifulSoup(html, "lxml")

    # Podstawowe meta tagi
    title = (soup.title.string.strip() if soup.title and soup.title.string else "")
    desc = ""
    mr = soup.find("meta", attrs={"name": "description"})
    if mr and mr.get("content"):
        desc = mr["content"].strip()

    # Viewport i mobile-friendliness
    viewport = soup.find("meta", attrs={"name": "viewport"})
    has_viewport = bool(viewport)
    viewport_content = viewport.get("content", "") if viewport else ""

    # Robots meta
    robots_meta = ""
    mrobots = soup.find("meta", attrs={"name": "robots"})
    if mrobots and mrobots.get("content"):
        robots_meta = mrobots["content"].lower()

    # Canonical
    canonical = ""
    link_canon = soup.find("link", rel=lambda v: v and "canonical" in v)
    if link_canon and link_canon.get("href"):
        canonical = absolutize(url, link_canon["href"])

    # Nagłówki H1, H2, H3
    h1 = [h.get_text(strip=True) for h in soup.find_all("h1")]
    h2 = [h.get_text(strip=True) for h in soup.find_all("h2")]
    h3 = [h.get_text(strip=True) for h in soup.find_all("h3")]

    # Obrazy i ALT
    imgs = soup.find_all("img")
    img_without_alt = sum(1 for i in imgs if not i.get("alt"))
    img_total = len(imgs)

    # Linki
    a_tags = soup.find_all("a", href=True)
    links = [
        absolutize(url, a["href"])
        for a in a_tags
        if not a["href"].startswith("javascript:")
    ]

    # Open Graph
    og_data = {}
    og_tags = soup.find_all("meta", property=re.compile(r"^og:"))
    for tag in og_tags:
        prop = tag.get("property", "")
        content = tag.get("content", "")
        if prop and content:
            og_data[prop] = content

    # Twitter Cards
    twitter_data = {}
    twitter_tags = soup.find_all("meta", attrs={"name": re.compile(r"^twitter:")})
    for tag in twitter_tags:
        name = tag.get("name", "")
        content = tag.get("content", "")
        if name and content:
            twitter_data[name] = content

    # Structured Data (Schema.org)
    try:
        structured = extruct.extract(
            html,
            base_url=get_base_url(html, url),
            syntaxes=["json-ld", "microdata", "opengraph"],
            uniform=True
        )
    except Exception:
        structured = {"json-ld": [], "microdata": [], "opengraph": []}

    jsonld_types = []
    for node in structured.get("json-ld", []):
        t = node.get("@type")
        if isinstance(t, list):
            jsonld_types += t
        elif t:
            jsonld_types.append(t)
    jsonld_types = list(dict.fromkeys(jsonld_types))

    # Tekst i analiza treści
    text = clean_text(soup)
    text_len = len(text)

    # Sygnały NAP, E-E-A-T
    nap_signals = extract_nap_signals(soup, text)
    eeat_signals = analyze_eeat_signals(soup, text, url)
    meta_scores = calculate_meta_score(title, desc)

    # Sygnały GEO (Schema.org types)
    geo_signals = {
        "has_faq_schema": "FAQPage" in jsonld_types,
        "has_article_schema": any(t in jsonld_types for t in ("Article", "NewsArticle", "BlogPosting")),
        "has_org_schema": any(t in jsonld_types for t in ("Organization", "LocalBusiness")),
        "has_breadcrumbs": "BreadcrumbList" in jsonld_types,
        "has_review_schema": any(t in jsonld_types for t in ("Review", "AggregateRating")),
        "has_product_schema": "Product" in jsonld_types,
        "clear_hierarchy": len(h1) == 1 and len(h2) > 0,
        "sufficient_text": text_len >= 1200,
        "has_navigation_schema": any('navigation' in str(s).lower() for s in soup.find_all(['nav'])),
    }

    return {
        "title": title,
        "meta_description": desc,
        "robots_meta": robots_meta,
        "canonical": canonical,
        "h1": h1,
        "h2": h2,
        "h3": h3,
        "h1_count": len(h1),
        "h2_count": len(h2),
        "h3_count": len(h3),
        "img_total": img_total,
        "img_without_alt": img_without_alt,
        "img_alt_ratio": round((img_total - img_without_alt) / max(1, img_total) * 100, 1),
        "has_viewport": has_viewport,
        "viewport_content": viewport_content,
        "is_mobile_friendly": has_viewport and 'width=device-width' in viewport_content.lower(),
        "opengraph": og_data,
        "twitter_cards": twitter_data,
        "has_og_image": "og:image" in og_data,
        "has_og_title": "og:title" in og_data,
        "has_og_description": "og:description" in og_data,
        "has_twitter_card": "twitter:card" in twitter_data,
        "jsonld_types": jsonld_types,
        "schema_count": len(jsonld_types),
        "text_len": text_len,
        "word_count": len(text.split()),
        "links": links,
        "internal_links": len([l for l in links if same_site(url, l)]),
        "external_links": len([l for l in links if not same_site(url, l)]),
        "nap_signals": nap_signals,
        "eeat_signals": eeat_signals,
        "geo_signals": geo_signals,
        "meta_scores": meta_scores,
        "is_excluded": is_excluded_url(url),
    }
