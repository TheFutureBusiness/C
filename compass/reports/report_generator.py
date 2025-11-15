"""
Generator raportów w różnych formatach (JSON, CSV, Word)
"""
import os
import json
from typing import Dict, Any

try:
    import pandas as pd
except Exception:
    pd = None

from compass.config import USE_AI_SUMMARY, OPENAI_API_KEY
from compass.integrations import generate_ai_summary
from .analyzer import find_duplicates, analyze_issues, calculate_summary
from .word_report import create_word_report


def save_reports(all_pages: Dict[str, Any], start_url: str, output_dir: str):
    """
    Generuje i zapisuje wszystkie raporty (JSON, CSV, Word).

    Args:
        all_pages: Słownik wszystkich przeanalizowanych stron
        start_url: URL startowy audytu
        output_dir: Katalog wyjściowy dla raportów
    """
    # Analiza duplikatów i problemów
    duplicates = find_duplicates(all_pages)
    issues = analyze_issues(all_pages)
    summary = calculate_summary(all_pages, issues, duplicates)

    # Generowanie AI Summary (jeśli włączone)
    ai_summary_text = ""
    if USE_AI_SUMMARY and OPENAI_API_KEY:
        print("🤖 Generuję AI Summary...")
        ai_summary_text = generate_ai_summary(summary, issues)
        if ai_summary_text:
            print("✅ AI Summary wygenerowane")

    # Zapis raportu JSON
    json_path = os.path.join(output_dir, "raport_szczegolowy.json")
    payload = {
        "summary": summary,
        "ai_summary": ai_summary_text,
        "pages": all_pages,
        "issues": issues,
        "duplicates": duplicates
    }

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    print(f"✅ JSON zapisany: {json_path}")

    # Zapis raportu CSV (jeśli pandas dostępne)
    if pd is not None:
        csv_path = os.path.join(output_dir, "raport_tabela.csv")
        rows = []

        for u, v in all_pages.items():
            rows.append({
                "url": u,
                "excluded": v.get("is_excluded", False),
                "status": v.get("status"),
                "title": v.get("title"),
                "title_length": v.get("meta_scores", {}).get("title_length"),
                "meta_description": v.get("meta_description"),
                "desc_length": v.get("meta_scores", {}).get("desc_length"),
                "canonical": v.get("canonical"),
                "h1_count": v.get("h1_count"),
                "h2_count": v.get("h2_count"),
                "mobile_friendly": v.get("is_mobile_friendly"),
                "has_og_image": v.get("has_og_image"),
                "has_twitter_card": v.get("has_twitter_card"),
                "schema_types": ",".join(v.get("jsonld_types", [])),
                "schema_count": v.get("schema_count"),
                "eeat_score": v.get("eeat_signals", {}).get("eeat_percentage"),
                "nap_score": v.get("nap_signals", {}).get("nap_score"),
                "word_count": v.get("word_count"),
                "img_total": v.get("img_total"),
                "img_without_alt": v.get("img_without_alt"),
                "error": v.get("error"),
                "has_ssl": v.get("security", {}).get("has_ssl"),
                "security_score": v.get("security", {}).get("security_percentage"),
                "security_level": v.get("security", {}).get("security_level"),
                "security_headers_count": v.get("security", {}).get("headers_count"),
                "has_hsts": v.get("security", {}).get("security_checks", {}).get("hsts", {}).get("present"),
                "has_csp": v.get("security", {}).get("security_checks", {}).get("content_security_policy", {}).get("present"),
                "has_mixed_content": v.get("security", {}).get("has_mixed_content"),
            })

        df = pd.DataFrame(rows)
        df.to_csv(csv_path, index=False, encoding="utf-8")
        print(f"✅ CSV zapisany: {csv_path}")

    # Zapis raportu Word
    word_path = os.path.join(output_dir, "raport_dla_klienta.docx")
    create_word_report(all_pages, summary, issues, duplicates, word_path)
