# Compass - Audytor SEO/AEO/GEO

Zaawansowany audytor SEO/AEO/GEO z modułową architekturą, wsparciem dla E-E-A-T, Local SEO (NAP), bezpieczeństwa i AI-powered podsumowań.

## 🚀 Funkcje

- **SEO**: Analiza meta tagów, nagłówków, canonical, duplicates
- **AEO/GEO**: Schema.org, strukturalne dane, sygnały E-E-A-T
- **Local SEO**: Analiza NAP (Name, Address, Phone)
- **Bezpieczeństwo**: Audyt nagłówków HTTP, SSL, mixed content
- **AI Summary**: Automatyczne podsumowania przy użyciu OpenAI
- **PageSpeed**: Integracja z Google PageSpeed Insights
- **Raporty**: JSON, CSV, Word (DOCX)

## 📁 Struktura projektu

```
Compass/
├── compass/                    # Główny pakiet
│   ├── __init__.py
│   ├── config.py              # Konfiguracja
│   ├── utils/                 # Narzędzia pomocnicze
│   │   ├── url_utils.py       # Operacje na URL
│   │   └── text_utils.py      # Przetwarzanie tekstu
│   ├── analyzers/             # Analizatory SEO/AEO/Security
│   │   ├── meta_analyzer.py
│   │   ├── nap_analyzer.py
│   │   ├── eeat_analyzer.py
│   │   └── security_analyzer.py
│   ├── crawler/               # Moduł crawlera
│   │   ├── fetcher.py         # Pobieranie stron
│   │   ├── robots.py          # robots.txt i sitemap
│   │   └── crawler.py         # Główny crawler
│   ├── integrations/          # Integracje zewnętrzne
│   │   ├── openai_integration.py
│   │   └── pagespeed.py
│   └── reports/               # Generowanie raportów
│       ├── analyzer.py        # Analiza wyników
│       ├── word_report.py     # Raport Word
│       └── report_generator.py
├── raporty/                   # Folder na wygenerowane raporty
├── main.py                    # Główny plik uruchomieniowy
├── requirements.txt           # Zależności Python
└── README.md                  # Ten plik

```

## 🔧 Instalacja

1. Sklonuj repozytorium:
```bash
git clone <repo-url>
cd Compass
```

2. Zainstaluj zależności:
```bash
pip install -r requirements.txt
```

3. (Opcjonalnie) Skonfiguruj zmienne środowiskowe:
```bash
export OPENAI_API_KEY="twój-klucz-api"
export PAGESPEED_API_KEY="twój-klucz-pagespeed"
```

## ⚙️ Konfiguracja

Edytuj plik `compass/config.py` aby dostosować parametry:

```python
START_URL = "https://example.com/"  # URL do audytu
MAX_PAGES = 300                      # Maksymalna liczba stron
MAX_DEPTH = 3                        # Maksymalna głębokość crawlingu
CONCURRENCY = 10                     # Liczba równoległych requestów

USE_PAGESPEED = False                # Włącz PageSpeed Insights
USE_AI_SUMMARY = True                # Włącz AI Summary

DOMAIN_SCOPE = "root"                # "root" lub "sub" (subdomeny)
```

## 🚀 Użycie

### Podstawowe użycie

```bash
python main.py
```

### Jako moduł Python

```python
import asyncio
from compass.crawler import crawl
from compass.reports import save_reports
from compass.config import START_URL, get_output_dir

# Uruchom crawling
data = asyncio.run(crawl(START_URL))

# Wygeneruj raporty
output_dir = get_output_dir()
save_reports(data, START_URL, output_dir)
```

## 📊 Generowane raporty

Wszystkie raporty są zapisywane w folderze `raporty/audyt_YYYY-MM-DD_HH-MM-SS/`:

1. **raport_dla_klienta.docx** - Profesjonalny raport Word dla klienta
2. **raport_szczegolowy.json** - Pełne dane w formacie JSON
3. **raport_tabela.csv** - Dane tabelaryczne do analizy

## 🔍 Moduły

### Utils
- `url_utils.py` - Normalizacja URL, sprawdzanie domeny, wykluczenia
- `text_utils.py` - Czyszczenie tekstu HTML

### Analyzers
- `meta_analyzer.py` - Analiza title i description
- `nap_analyzer.py` - Local SEO (Name, Address, Phone)
- `eeat_analyzer.py` - Sygnały E-E-A-T (Experience, Expertise, Authoritativeness, Trustworthiness)
- `security_analyzer.py` - Nagłówki HTTP, SSL, bezpieczeństwo

### Crawler
- `fetcher.py` - Asynchroniczne pobieranie i parsowanie stron
- `robots.py` - Obsługa robots.txt i sitemap.xml
- `crawler.py` - Główny silnik crawlera z BFS

### Integrations
- `openai_integration.py` - Generowanie AI Summary przez OpenAI
- `pagespeed.py` - Google PageSpeed Insights API

### Reports
- `analyzer.py` - Analiza duplikatów i problemów
- `word_report.py` - Generator raportu Word (DOCX)
- `report_generator.py` - Orkiestracja wszystkich raportów

## 📝 Przykładowa konfiguracja dla różnych scenariuszy

### Audyt małej strony
```python
MAX_PAGES = 50
MAX_DEPTH = 2
CONCURRENCY = 5
```

### Audyt dużej strony
```python
MAX_PAGES = 1000
MAX_DEPTH = 5
CONCURRENCY = 20
```

### Audyt tylko głównych stron
```python
MAX_PAGES = 100
MAX_DEPTH = 1
EXCLUDED_PATTERNS = [
    r'/blog/',
    r'/archiwum/',
    # ... więcej wzorców
]
```

## 🤝 Wkład w projekt

Pull requesty są mile widziane! Przed wysłaniem PR:

1. Upewnij się, że kod jest zgodny z PEP 8
2. Dodaj testy dla nowych funkcji
3. Zaktualizuj dokumentację

## 📄 Licencja

MIT License - zobacz plik LICENSE

## 🙏 Podziękowania

- BeautifulSoup4 - parsowanie HTML
- aiohttp - asynchroniczne HTTP
- extruct - ekstrakcja strukturalnych danych
- python-docx - generowanie raportów Word

## 📞 Kontakt

W razie pytań lub problemów, otwórz issue na GitHubie.

---

**Compass** - Twój przewodnik w świecie audytów SEO/AEO/GEO 🧭
