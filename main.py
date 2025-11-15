#!/usr/bin/env python3
"""
Compass - Audytor SEO/AEO/GEO Enhanced Edition

Główny plik uruchomieniowy audytora.
"""
import asyncio
import time

from compass.config import (
    START_URL,
    MAX_PAGES,
    MAX_DEPTH,
    CONCURRENCY,
    USE_PAGESPEED,
    USE_AI_SUMMARY,
    OPENAI_API_KEY,
    OPENAI_MODEL,
    get_output_dir,
)
from compass.crawler import crawl
from compass.reports import save_reports


def print_header(output_dir: str):
    """Wyświetla nagłówek z informacjami o audycie."""
    print("=" * 80)
    print("🚀 AUDYTOR SEO/AEO/GEO - ENHANCED EDITION".center(80))
    print("=" * 80)
    print(f"\n📁 Katalog wyników: {output_dir}")
    print(f"🌐 Audytowana strona: {START_URL}")
    print(f"⚙️  MAX_PAGES={MAX_PAGES}, MAX_DEPTH={MAX_DEPTH}, CONCURRENCY={CONCURRENCY}")
    print(f"🚫 Wykluczono m.in. /cdn-cgi/*")

    if USE_PAGESPEED:
        print("📊 PageSpeed Insights: WŁĄCZONY")
    else:
        print("📊 PageSpeed Insights: WYŁĄCZONY")

    if USE_AI_SUMMARY:
        if OPENAI_API_KEY:
            print(f"🤖 AI Summary: WŁĄCZONY (model: {OPENAI_MODEL})")
        else:
            print("🤖 AI Summary: WYŁĄCZONY (brak OPENAI_API_KEY)")
    else:
        print("🤖 AI Summary: WYŁĄCZONY")

    print("\n" + "=" * 80)
    print()


def print_footer(output_dir: str, elapsed_time: float):
    """Wyświetla stopkę z podsumowaniem."""
    print("\n" + "=" * 80)
    print("📊 AUDYT ZAKOŃCZONY".center(80))
    print("=" * 80)
    print(f"\n⏱️  Czas wykonania: {elapsed_time:.1f}s")
    print(f"\n📁 Wyniki: {output_dir}/")
    print("   • raport_dla_klienta.docx")
    print("   • raport_szczegolowy.json")
    print("   • raport_tabela.csv")
    print("\n✅ Gotowe!")


def main():
    """Główna funkcja uruchamiająca audyt."""
    # Utworzenie katalogu wyjściowego
    output_dir = get_output_dir()

    # Wyświetlenie nagłówka
    print_header(output_dir)

    # Start pomiaru czasu
    t0 = time.time()

    # Uruchomienie crawlera
    data = asyncio.run(crawl(START_URL))

    # Generowanie raportów
    save_reports(data, START_URL, output_dir)

    # Obliczenie czasu wykonania
    elapsed_time = time.time() - t0

    # Wyświetlenie stopki
    print_footer(output_dir, elapsed_time)


if __name__ == "__main__":
    main()
