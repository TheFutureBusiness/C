"""
Narzędzia do przetwarzania tekstu
"""
from bs4 import BeautifulSoup


def clean_text(soup: BeautifulSoup) -> str:
    """
    Czyści tekst ze strony HTML, usuwając skrypty, style i nadmiarowe białe znaki.

    Args:
        soup: Obiekt BeautifulSoup z HTML

    Returns:
        Oczyszczony tekst
    """
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()

    text = soup.get_text(separator=" ", strip=True)
    text = " ".join(text.split())
    return text
