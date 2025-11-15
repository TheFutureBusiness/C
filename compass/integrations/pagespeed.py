"""
Integracja z Google PageSpeed Insights API
"""
import aiohttp
from typing import Dict, Any
from compass.config import USE_PAGESPEED, PAGESPEED_API_KEY


async def check_pagespeed(url: str) -> Dict[str, Any]:
    """
    Sprawdza metryki PageSpeed Insights dla danego URL.

    Args:
        url: URL strony do sprawdzenia

    Returns:
        Słownik z wynikami PageSpeed (performance, accessibility, seo, core web vitals)
    """
    if not USE_PAGESPEED or not PAGESPEED_API_KEY:
        return {}

    try:
        api_url = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed"
        params = {
            "url": url,
            "key": PAGESPEED_API_KEY,
            "strategy": "mobile",
            "category": ["performance", "accessibility", "best-practices", "seo"]
        }

        async with aiohttp.ClientSession() as session:
            async with session.get(api_url, params=params, timeout=60) as resp:
                if resp.status == 200:
                    data = await resp.json()
                    lighthouse = data.get("lighthouseResult", {})
                    categories = lighthouse.get("categories", {})
                    audits = lighthouse.get("audits", {})

                    cwv = {
                        "lcp": audits.get("largest-contentful-paint", {}).get("numericValue"),
                        "cls": audits.get("cumulative-layout-shift", {}).get("numericValue"),
                        "fid": audits.get("max-potential-fid", {}).get("numericValue"),
                    }

                    return {
                        "performance_score": categories.get("performance", {}).get("score", 0) * 100,
                        "accessibility_score": categories.get("accessibility", {}).get("score", 0) * 100,
                        "seo_score": categories.get("seo", {}).get("score", 0) * 100,
                        "core_web_vitals": cwv,
                    }

    except Exception as e:
        print(f"⚠️  PageSpeed API error for {url}: {e}")

    return {}
