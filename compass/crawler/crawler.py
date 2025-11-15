"""
Główny moduł crawlera - asynchroniczne przeszukiwanie witryny
"""
import asyncio
import aiohttp
from collections import deque
from typing import Dict, Any, Set

try:
    from tqdm import tqdm
    HAS_TQDM = True
except:
    HAS_TQDM = False

from compass.config import (
    MAX_PAGES,
    MAX_DEPTH,
    CONCURRENCY,
    RESPECT_ROBOTS,
    USER_AGENT,
    USE_PAGESPEED,
)
from compass.utils import same_site, is_excluded_url
from compass.analyzers import analyze_security_headers
from compass.integrations import check_pagespeed
from .fetcher import fetch, parse_page
from .robots import build_robots, discover_sitemaps, fetch_and_parse_sitemaps


async def crawl(start_url: str) -> Dict[str, Any]:
    """
    Przeszukuje witrynę internetową asynchronicznie, analizując wszystkie strony.

    Args:
        start_url: URL startowy do przeszukania

    Returns:
        Słownik z wynikami crawlingu (URL -> dane strony)
    """
    # Kolejka URL-i do odwiedzenia: (url, depth)
    q = deque([(start_url, 0)])
    seen: Set[str] = {start_url}
    results: Dict[str, Any] = {}

    # Pasek postępu (jeśli dostępny tqdm)
    pbar = tqdm(total=MAX_PAGES, desc="Crawling", unit="page") if HAS_TQDM else None

    async with aiohttp.ClientSession() as session:
        # Budowanie parsera robots.txt
        rp = await build_robots(session, start_url) if RESPECT_ROBOTS else None

        # Semaphore do kontroli współbieżności
        sem = asyncio.Semaphore(CONCURRENCY)

        # Odkrywanie i parsowanie sitemap.xml
        '''
        try:
            sitemaps = await discover_sitemaps(session, start_url)
            if sitemaps:
                urls_from_sm = await fetch_and_parse_sitemaps(session, sitemaps)
                for u in urls_from_sm[:max(0, MAX_PAGES - len(q))]:
                    if u not in seen and same_site(start_url, u) and not is_excluded_url(u):
                        seen.add(u)
                        q.append((u, 1))
        except Exception as e:
            print(f"⚠️  Błąd przy pobieraniu sitemap: {e}")
        '''
        async def worker():
            """Worker do przetwarzania URL-i z kolejki."""
            while q and len(results) < MAX_PAGES:
                url, depth = q.popleft()

                # Sprawdzenie robots.txt
                if RESPECT_ROBOTS and rp and not rp.can_fetch(USER_AGENT, url):
                    results[url] = {"url": url, "error": "blocked_by_robots"}
                    if pbar:
                        pbar.update(1)
                    continue

                async with sem:
                    # Pobranie strony
                    status, final, ct, html, headers = await fetch(session, url)

                    item: Dict[str, Any] = {
                        "url": url,
                        "final_url": final,
                        "status": status,
                        "content_type": ct,
                        "headers": headers
                    }

                    # Obsługa błędów
                    if not status or (isinstance(html, str) and html.startswith("__ERROR__")):
                        item["error"] = html if isinstance(html, str) else "fetch_error"
                        results[url] = item
                        if pbar:
                            pbar.update(1)
                        continue

                    # Parsowanie HTML
                    if ct and "text/html" in ct:
                        parsed = parse_page(html, final)
                        item.update(parsed)

                        # Analiza bezpieczeństwa
                        security_analysis = analyze_security_headers(headers, final, html)
                        item["security"] = security_analysis

                        # PageSpeed Insights (tylko dla pierwszych kilku stron)
                        if USE_PAGESPEED and len(results) < 5:
                            item["pagespeed"] = await check_pagespeed(final)

                        # Dodawanie nowych linków do kolejki
                        for link in item.get("links", []):
                            if not same_site(start_url, link):
                                continue
                            if is_excluded_url(link):
                                continue
                            if link not in seen and depth + 1 <= MAX_DEPTH:
                                seen.add(link)
                                q.append((link, depth + 1))
                    else:
                        item["note"] = "Pominięto (non-HTML)"

                    results[url] = item
                    if pbar:
                        pbar.update(1)

        # Uruchomienie workerów
        tasks = [asyncio.create_task(worker()) for _ in range(CONCURRENCY)]
        await asyncio.gather(*tasks)

    if pbar:
        pbar.close()

    return results
