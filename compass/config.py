"""
Konfiguracja audytora SEO/AEO/GEO
"""
import os
from datetime import datetime

# ===== PODSTAWOWA KONFIGURACJA CRAWLINGU =====
START_URL = "https://ekantor.pl/"
MAX_PAGES = 300
MAX_DEPTH = 3
TIMEOUT = 25
CONCURRENCY = 10
USER_AGENT = "SiteAuditorBot/1.0 (+https://twojadomena.example/audyt)"
RESPECT_ROBOTS = True

# ===== ZAKRES DOMENY =====
# "root" - tylko ta sama domena (http://example.com)
# "sub" - cała domena z subdomenami (*.example.com)
DOMAIN_SCOPE = "root"

# ===== INTEGRACJE =====
USE_PAGESPEED = False
PAGESPEED_API_KEY = os.getenv(
    "PAGESPEED_API_KEY",
    "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"
)

USE_AI_SUMMARY = True
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
OPENAI_MODEL = "gpt-4o-mini"

# ===== KATALOG WYJŚCIOWY =====
# Katalog bazowy dla raportów
REPORTS_BASE_DIR = "raporty"

# Funkcja generująca nazwę katalogu dla konkretnego audytu
def get_output_dir() -> str:
    """Generuje nazwę katalogu dla bieżącego audytu."""
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output_dir = os.path.join(REPORTS_BASE_DIR, f"audyt_{timestamp}")
    os.makedirs(output_dir, exist_ok=True)
    return output_dir

# ===== WYKLUCZENIA =====
EXCLUDED_PATTERNS = [
    r'/polityka[_-]prywatnosci',
    r'/privacy[_-]policy',
    r'/regulamin',
    r'/terms',
    r'/sitemap',
    r'/robots\.txt',
    r'/cookie[s]?[_-]policy',
    r'/disclaimer',
    r'/terms-of-service',
    r'/legal',
    r'^/cdn-cgi/',
    r'/cdn-cgi/l/email-protection',
]

# ===== STRONY SYSTEMOWE (nie wymagają pełnej analizy SEO) =====
# Te strony nie powinny być oceniane pod kątem E-E-A-T, NAP, Meta Description itp.
SYSTEM_PAGE_PATTERNS = [
    r'/konto[_-]?uzytkownika',
    r'/mein[_-]?konto',
    r'/my[_-]?account',
    r'/account',
    r'/cart',
    r'/koszyk',
    r'/warenkorb',
    r'/checkout',
    r'/zamowienie',
    r'/bestellung',
    r'/login',
    r'/logowanie',
    r'/anmelden',
    r'/register',
    r'/rejestracja',
    r'/registrieren',
    r'/wholesale[_-]?login',
    r'/wp-admin',
    r'/wp-login',
    r'/wp-content/uploads/.*\.(pdf|doc|docx|xls|xlsx)$',
    r'/feed/?$',
    r'/rss/?$',
    r'/search',
    r'/suche',
    r'/szukaj',
    r'/404',
]

# ===== OPCJE RAPORTOWANIA =====
SHOW_REMEDIATIONS = False

# ===== JĘZYK RAPORTU =====
# Dostępne: "pl" (polski), "de" (niemiecki), "en" (angielski)
REPORT_LANGUAGE = "pl"
