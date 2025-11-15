"""
Integracja z OpenAI API dla generowania AI Summary
"""
import json
from typing import Dict, Any
from compass.config import USE_AI_SUMMARY, OPENAI_API_KEY, OPENAI_MODEL


def generate_ai_summary(summary: Dict[str, Any], issues: Dict[str, Any]) -> str:
    """
    Generuje AI-powered podsumowanie audytu przy użyciu OpenAI API.

    Args:
        summary: Słownik z podsumowaniem audytu
        issues: Słownik ze znalezionymi problemami

    Returns:
        Tekst podsumowania wygenerowany przez AI (markdown)
    """
    if not USE_AI_SUMMARY or not OPENAI_API_KEY:
        return ""

    try:
        try:
            from openai import OpenAI
        except ImportError:
            print("⚠️  Brak biblioteki openai. Zainstaluj: pip install openai")
            return ""

        client = OpenAI(api_key=OPENAI_API_KEY)

        audit_data = {
            "url": summary["start_url"],
            "pages_analyzed": summary["pages_analyzed"],
            "seo": {
                "errors": summary["pages_with_errors"],
                "missing_title": summary["missing_title"],
                "missing_description": summary["missing_description"],
                "duplicate_titles": summary["duplicate_titles"],
                "duplicate_descriptions": summary["duplicate_descriptions"],
                "missing_canonical": summary["missing_canonical"],
                "mobile_percentage": summary["mobile_percentage"],
                "avg_eeat_score": summary["avg_eeat_score"],
            },
            "security": {
                "avg_score": summary["avg_security_score"],
                "ssl_percentage": summary["ssl_percentage"],
                "pages_poor_security": summary["pages_poor_security"],
            },
            "geo": {
                "pages_with_schema": summary["pages_with_schema"],
                "pages_without_schema": summary["pages_without_schema"],
            }
        }

        prompt = f"""Jesteś ekspertem SEO/GEO/Security. Przeanalizuj wyniki audytu i wygeneruj krótkie podsumowanie.

DANE AUDYTU:
{json.dumps(audit_data, indent=2, ensure_ascii=False)}

Wygeneruj podsumowanie (MAX 300 słów) zawierające:
1. OGÓLNA OCENA (1-2 zdania)
2. TOP 3 PRIORYTETY (lista) – konkretne action items, ale BEZ instrukcji wdrożeniowych i konfiguracji serwera
3. MOCNE STRONY (2-3 punkty)
4. OSTRZEŻENIA (jeśli są krytyczne problemy)
5. REKOMENDACJA BIZNESOWA (1 zdanie)

Styl: zwięzły, profesjonalny, emoji (🔴🟠🟡✅).
Język: polski
Format: Markdown, nagłówki ###
Nie podawaj konfiguracji serwerów (Apache/Nginx/Cloudflare) ani komend."""

        response = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[
                {
                    "role": "system",
                    "content": "Jesteś ekspertem SEO, GEO i cyberbezpieczeństwa. Tworzysz zwięzłe, actionable podsumowania audytów bez instrukcji wdrożeniowych."
                },
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=800
        )

        ai_summary = response.choices[0].message.content.strip()
        return ai_summary

    except Exception as e:
        print(f"⚠️  Błąd generowania AI Summary: {e}")
        return ""
