"""
Analizator meta tagów (title, description)
"""
from typing import Dict, Any


def calculate_meta_score(title: str, description: str) -> Dict[str, Any]:
    """
    Analizuje jakość meta tagów (title i description).

    Args:
        title: Tytuł strony
        description: Meta description

    Returns:
        Słownik z wynikami analizy meta tagów
    """
    title_len = len(title)
    desc_len = len(description)

    title_optimal = 50 <= title_len <= 60
    title_acceptable = 30 <= title_len <= 65
    desc_optimal = 150 <= desc_len <= 160
    desc_acceptable = 120 <= desc_len <= 165

    return {
        "title_length": title_len,
        "title_optimal": title_optimal,
        "title_acceptable": title_acceptable,
        "title_too_short": title_len < 30,
        "title_too_long": title_len > 65,
        "desc_length": desc_len,
        "desc_optimal": desc_optimal,
        "desc_acceptable": desc_acceptable,
        "desc_too_short": desc_len < 120,
        "desc_too_long": desc_len > 165,
        "has_title": bool(title),
        "has_description": bool(description),
    }
