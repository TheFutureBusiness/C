"""
Tłumaczenia dla raportów SEO/AEO/GEO

Obsługiwane języki: pl (polski), de (niemiecki), en (angielski)
"""

TRANSLATIONS = {
    "pl": {
        # Nagłówki główne
        "report_title": "Audyt SEO / AEO / GEO",
        "audit_date": "Data audytu",
        "recipients": "Adresaci: Zarząd / Marketing / Zespół IT",
        "table_of_contents": "Spis treści",

        # KPI
        "main_result": "Główny wynik",
        "technical_stability": "Stabilność techniczna",
        "mobility": "Mobilność",
        "audit_result": "Wynik audytu",
        "pages_ok": "Stron OK (200)",
        "mobile_friendly": "Mobile-friendly",
        "of_analyzed": "z {0} analizowanych",
        "pages": "stron",

        # Sekcje
        "scoring_composition": "Skład oceny ogólnej",
        "category": "Kategoria",
        "value": "Wartość (x z y)",
        "description": "Opis",
        "important_issues": "Najważniejsze problemy",
        "executive_summary": "Podsumowanie wykonawcze",
        "priorities": "Priorytety (od krytycznych do niskich)",
        "meta_analysis": "Analiza meta tagów",
        "technical_seo": "Techniczna analiza SEO",
        "mobility_responsive": "Mobilność i Responsive Design",
        "open_graph_twitter": "Open Graph i Twitter Cards",
        "structured_data": "Dane strukturalne (Schema.org)",
        "eeat": "E-E-A-T",
        "local_seo": "Lokalne SEO (NAP)",
        "content_quality": "Jakość treści",
        "security": "Bezpieczeństwo (Security Headers)",
        "legend": "Legenda i objaśnienia",
        "ai_summary": "Podsumowanie AI",

        # Status
        "status_excellent": "Doskonały!",
        "status_needs_improvement": "Wymaga poprawy",
        "status_requires_attention": "Wymaga uwagi!",
        "seo_status": "Status SEO",
        "security_status": "Status bezpieczeństwa",
        "critical_issues": "Problemy krytyczne",
        "warnings": "Ostrzeżenia",

        # Metryki
        "scanned_pages": "Przeskanowane strony",
        "analyzed_pages": "Analizowane strony",
        "content_pages": "Strony treściowe",
        "system_pages": "Strony systemowe",
        "excluded_pages": "Wykluczone strony",
        "noindex_pages": "Strony noindex",
        "pages_with_errors": "Strony z błędami",
        "missing_title": "Brak tytułu",
        "missing_description": "Brak opisu meta",
        "title_issues": "Problemy z tytułem",
        "description_issues": "Problemy z opisem",
        "duplicate_titles": "Duplikaty tytułów",
        "duplicate_descriptions": "Duplikaty opisów",
        "missing_canonical": "Brak canonical",
        "missing_h1": "Brak H1",
        "multiple_h1": "Wielokrotne H1",
        "images_without_alt": "Obrazy bez ALT",
        "missing_viewport": "Brak meta viewport",
        "missing_og": "Brak Open Graph",
        "missing_twitter": "Brak Twitter Cards",
        "pages_with_schema": "Strony ze Schema.org",
        "missing_schema": "Brak Schema.org",
        "avg_schema_types": "Śr. typów Schema/stronę",
        "avg_eeat": "Śr. wynik E-E-A-T",
        "weak_eeat": "Słabe E-E-A-T",
        "local_nap_ok": "Lokalne NAP OK",
        "poor_local_seo": "Słabe lokalne SEO",
        "thin_content": "Cienka treść (<300 słów)",
        "avg_security": "Śr. bezpieczeństwo",
        "weak_security": "Słabe bezpieczeństwo",
        "missing_security_headers": "Brak nagłówków bezpieczeństwa",
        "mixed_content": "Mixed content",
        "ssl_percentage": "Strony z SSL",

        # Opisy problemów
        "http_errors_desc": "Strony z błędami HTTP (4xx, 5xx). Blokują indeksowanie i powodują utratę ruchu.",
        "missing_title_desc": "Każda strona wymaga unikalnego tagu <title> (50-60 znaków) dla lepszej widoczności w wynikach wyszukiwania.",
        "missing_desc_desc": "Meta Description (150-160 znaków) to pierwszy kontakt użytkownika z Twoją stroną w wynikach Google.",
        "missing_canonical_desc": "Tag canonical zapobiega problemom z duplikatami treści i pomaga Google wybrać właściwą wersję strony.",
        "missing_viewport_desc": "Brak Meta Viewport - wymaga ręcznej weryfikacji wyświetlania na urządzeniach mobilnych.",
        "missing_schema_desc": "Dane strukturalne (Schema.org) pomagają Google lepiej zrozumieć treść i wyświetlać Rich Snippets.",
        "poor_security_desc": "Słabe bezpieczeństwo (<50%). Brakujące nagłówki bezpieczeństwa zagrażają użytkownikom i zmniejszają zaufanie.",
        "thin_content_desc": "Strony z mniej niż 300 słowami. Google preferuje wartościowe, szczegółowe treści.",
        "weak_eeat_desc": "Słabe sygnały E-E-A-T (<50%). Dodaj autora, datę publikacji, certyfikaty i linki do wiarygodnych źródeł.",

        # Priorytety
        "top_priorities": "Top 3 priorytety na następny sprint",
        "business_benefit": "Korzyść biznesowa",
        "what_you_gain": "Co zyskasz wdrażając zalecenia:",
        "more_traffic": "więcej organicznego ruchu z Google",
        "higher_ctr": "wyższa CTR z wyników wyszukiwania i social media",
        "better_security": "lepsze bezpieczeństwo i zaufanie użytkowników",

        # Bezpieczeństwo
        "security_excellent": "Doskonałe zabezpieczenia",
        "security_good": "Dobre zabezpieczenia",
        "security_fair": "Przeciętne zabezpieczenia",
        "security_poor": "Słabe zabezpieczenia",
        "no_critical_issues": "Brak krytycznych problemów bezpieczeństwa (HTTPS aktywne, brak mixed content)",
        "hardening_recommended": "Zalecane wzmocnienie zabezpieczeń poprzez dodanie nagłówków bezpieczeństwa",

        # Scoring
        "availability": "Dostępność (HTTP 200)",
        "availability_desc": "Procent stron bez błędów HTTP (4xx, 5xx)",
        "meta_tags": "Meta tagi",
        "meta_tags_desc": "Obecność i jakość Title i Description",
        "mobile_desc": "Responsywny design i Meta Viewport",
        "schema_desc": "Dane strukturalne JSON-LD",
        "eeat_desc": "Ekspertyza, Autorytet, Zaufanie",
        "security_desc": "HTTPS, Security Headers, brak Mixed Content",
    },

    "de": {
        # Główne nagłówki
        "report_title": "SEO / AEO / GEO Audit",
        "audit_date": "Audit-Datum",
        "recipients": "Adressaten: Geschäftsführung / Marketing / IT-Team",
        "table_of_contents": "Inhaltsverzeichnis",

        # KPI
        "main_result": "Hauptergebnis",
        "technical_stability": "Technische Stabilität",
        "mobility": "Mobilität",
        "audit_result": "Audit-Ergebnis",
        "pages_ok": "Seiten OK (200)",
        "mobile_friendly": "Mobile-friendly",
        "of_analyzed": "von {0} analysierten",
        "pages": "Seiten",

        # Sekcje
        "scoring_composition": "Zusammensetzung der Gesamtbewertung",
        "category": "Kategorie",
        "value": "Wert (x von y)",
        "description": "Beschreibung",
        "important_issues": "Wichtigste Probleme",
        "executive_summary": "Executive Summary – Kernzahlen",
        "priorities": "Prioritäten (von kritisch bis gering)",
        "meta_analysis": "Meta-Tag-Analyse",
        "technical_seo": "Technische SEO-Analyse",
        "mobility_responsive": "Mobilität und Responsive Design",
        "open_graph_twitter": "Open Graph und Twitter Cards",
        "structured_data": "Strukturierte Daten (Schema.org)",
        "eeat": "E-E-A-T",
        "local_seo": "Local SEO (NAP)",
        "content_quality": "Inhaltsqualität",
        "security": "Sicherheit (Security Headers)",
        "legend": "Legende und Erläuterungen",
        "ai_summary": "KI-gestützte Zusammenfassung",

        # Status
        "status_excellent": "Ausgezeichnet!",
        "status_needs_improvement": "Verbesserungsbedarf",
        "status_requires_attention": "Erfordert Aufmerksamkeit!",
        "seo_status": "SEO-Status",
        "security_status": "Sicherheitsstatus",
        "critical_issues": "Kritische Probleme",
        "warnings": "Warnungen",

        # Metryki
        "scanned_pages": "Gescannte Seiten",
        "analyzed_pages": "Analysierte Seiten",
        "content_pages": "Inhaltsseiten",
        "system_pages": "Systemseiten",
        "excluded_pages": "Ausgeschlossene Seiten",
        "noindex_pages": "Noindex-Seiten",
        "pages_with_errors": "Seiten mit Fehlern",
        "missing_title": "Fehlender Title",
        "missing_description": "Fehlende Meta Description",
        "title_issues": "Title-Probleme",
        "description_issues": "Description-Probleme",
        "duplicate_titles": "Title-Duplikate",
        "duplicate_descriptions": "Description-Duplikate",
        "missing_canonical": "Fehlende Canonical",
        "missing_h1": "Fehlende H1",
        "multiple_h1": "Mehrere H1",
        "images_without_alt": "Bilder ohne ALT",
        "missing_viewport": "Fehlendes Meta Viewport",
        "missing_og": "Fehlendes Open Graph",
        "missing_twitter": "Fehlende Twitter Cards",
        "pages_with_schema": "Seiten mit Schema.org",
        "missing_schema": "Fehlendes Schema.org",
        "avg_schema_types": "Durchschn. Schema-Typen/Seite",
        "avg_eeat": "Durchschn. E-E-A-T",
        "weak_eeat": "Schwaches E-E-A-T",
        "local_nap_ok": "Local NAP OK",
        "poor_local_seo": "Schwaches Local SEO",
        "thin_content": "Dünner Inhalt (<300 Wörter)",
        "avg_security": "Durchschn. Sicherheit",
        "weak_security": "Schwache Sicherheit",
        "missing_security_headers": "Fehlende Security Headers",
        "mixed_content": "Mixed Content",
        "ssl_percentage": "Seiten mit SSL",

        # Opisy problemów
        "http_errors_desc": "Seiten mit HTTP-Fehlercodes (4xx, 5xx). Sie verhindern die Indexierung und führen zu Traffic-Verlust.",
        "missing_title_desc": "Jede Seite benötigt ein eindeutiges <title>-Tag (50–60 Zeichen) für bessere Sichtbarkeit in Suchergebnissen.",
        "missing_desc_desc": "Meta Description (150–160 Zeichen) ist der erste Kontakt des Nutzers mit Ihrer Seite in Google-Ergebnissen.",
        "missing_canonical_desc": "Das Canonical-Tag verhindert Probleme mit doppelten Inhalten und hilft Google, die richtige Seitenversion zu wählen.",
        "missing_viewport_desc": "Fehlendes Meta Viewport - erfordert manuelle Überprüfung der Darstellung auf Mobilgeräten.",
        "missing_schema_desc": "Strukturierte Daten (Schema.org) helfen Google, Inhalte besser zu verstehen und Rich Snippets anzuzeigen.",
        "poor_security_desc": "Schwache Sicherheit (<50%). Fehlende Security Headers gefährden Nutzer und reduzieren das Vertrauen.",
        "thin_content_desc": "Seiten mit weniger als 300 Wörtern. Google bevorzugt wertvolle, detaillierte Inhalte.",
        "weak_eeat_desc": "Schwache E-E-A-T-Signale (<50%). Fügen Sie Autor, Veröffentlichungsdatum, Zertifikate und Links zu vertrauenswürdigen Quellen hinzu.",

        # Priorytety
        "top_priorities": "Top 3 Prioritäten für den nächsten Sprint",
        "business_benefit": "Geschäftlicher Nutzen",
        "what_you_gain": "Was Sie durch Umsetzung der Empfehlungen gewinnen:",
        "more_traffic": "mehr organischer Traffic von Google",
        "higher_ctr": "höhere CTR aus Suchergebnissen und Social Media",
        "better_security": "bessere Sicherheit und Nutzervertrauen",

        # Bezpieczeństwo
        "security_excellent": "Ausgezeichnete Sicherheit",
        "security_good": "Gute Sicherheit",
        "security_fair": "Durchschnittliche Sicherheit",
        "security_poor": "Schwache Sicherheit",
        "no_critical_issues": "Keine kritischen Sicherheitsprobleme (HTTPS aktiv, kein Mixed Content)",
        "hardening_recommended": "Empfohlene Sicherheitshärtung durch Hinzufügen von Security Headers",

        # Scoring
        "availability": "Verfügbarkeit (HTTP 200)",
        "availability_desc": "Prozentsatz der Seiten ohne HTTP-Fehler (4xx, 5xx)",
        "meta_tags": "Meta-Tags",
        "meta_tags_desc": "Vorhandensein und Qualität von Title und Description",
        "mobile_desc": "Responsive Design und Meta Viewport",
        "schema_desc": "Strukturierte Daten JSON-LD",
        "eeat_desc": "Expertise, Autorität, Vertrauen",
        "security_desc": "HTTPS, Security Headers, kein Mixed Content",
    },

    "en": {
        # Main headers
        "report_title": "SEO / AEO / GEO Audit",
        "audit_date": "Audit Date",
        "recipients": "Recipients: Management / Marketing / IT Team",
        "table_of_contents": "Table of Contents",

        # KPI
        "main_result": "Main Result",
        "technical_stability": "Technical Stability",
        "mobility": "Mobility",
        "audit_result": "Audit Result",
        "pages_ok": "Pages OK (200)",
        "mobile_friendly": "Mobile-friendly",
        "of_analyzed": "of {0} analyzed",
        "pages": "pages",

        # Sections
        "scoring_composition": "Overall Score Composition",
        "category": "Category",
        "value": "Value (x of y)",
        "description": "Description",
        "important_issues": "Key Issues",
        "executive_summary": "Executive Summary",
        "priorities": "Priorities (from critical to low)",
        "meta_analysis": "Meta Tag Analysis",
        "technical_seo": "Technical SEO Analysis",
        "mobility_responsive": "Mobility and Responsive Design",
        "open_graph_twitter": "Open Graph and Twitter Cards",
        "structured_data": "Structured Data (Schema.org)",
        "eeat": "E-E-A-T",
        "local_seo": "Local SEO (NAP)",
        "content_quality": "Content Quality",
        "security": "Security (Security Headers)",
        "legend": "Legend and Explanations",
        "ai_summary": "AI Summary",

        # Status
        "status_excellent": "Excellent!",
        "status_needs_improvement": "Needs Improvement",
        "status_requires_attention": "Requires Attention!",
        "seo_status": "SEO Status",
        "security_status": "Security Status",
        "critical_issues": "Critical Issues",
        "warnings": "Warnings",

        # Metrics
        "scanned_pages": "Scanned Pages",
        "analyzed_pages": "Analyzed Pages",
        "content_pages": "Content Pages",
        "system_pages": "System Pages",
        "excluded_pages": "Excluded Pages",
        "noindex_pages": "Noindex Pages",
        "pages_with_errors": "Pages with Errors",
        "missing_title": "Missing Title",
        "missing_description": "Missing Meta Description",
        "title_issues": "Title Issues",
        "description_issues": "Description Issues",
        "duplicate_titles": "Duplicate Titles",
        "duplicate_descriptions": "Duplicate Descriptions",
        "missing_canonical": "Missing Canonical",
        "missing_h1": "Missing H1",
        "multiple_h1": "Multiple H1",
        "images_without_alt": "Images without ALT",
        "missing_viewport": "Missing Meta Viewport",
        "missing_og": "Missing Open Graph",
        "missing_twitter": "Missing Twitter Cards",
        "pages_with_schema": "Pages with Schema.org",
        "missing_schema": "Missing Schema.org",
        "avg_schema_types": "Avg. Schema Types/Page",
        "avg_eeat": "Avg. E-E-A-T",
        "weak_eeat": "Weak E-E-A-T",
        "local_nap_ok": "Local NAP OK",
        "poor_local_seo": "Poor Local SEO",
        "thin_content": "Thin Content (<300 words)",
        "avg_security": "Avg. Security",
        "weak_security": "Weak Security",
        "missing_security_headers": "Missing Security Headers",
        "mixed_content": "Mixed Content",
        "ssl_percentage": "Pages with SSL",

        # Problem descriptions
        "http_errors_desc": "Pages with HTTP error codes (4xx, 5xx). They prevent indexing and cause traffic loss.",
        "missing_title_desc": "Every page needs a unique <title> tag (50-60 characters) for better visibility in search results.",
        "missing_desc_desc": "Meta Description (150-160 characters) is the user's first contact with your page in Google results.",
        "missing_canonical_desc": "The canonical tag prevents duplicate content issues and helps Google choose the correct page version.",
        "missing_viewport_desc": "Missing Meta Viewport - requires manual verification of display on mobile devices.",
        "missing_schema_desc": "Structured data (Schema.org) helps Google better understand content and display Rich Snippets.",
        "poor_security_desc": "Weak security (<50%). Missing security headers endanger users and reduce trust.",
        "thin_content_desc": "Pages with less than 300 words. Google prefers valuable, detailed content.",
        "weak_eeat_desc": "Weak E-E-A-T signals (<50%). Add author, publication date, certificates, and links to trusted sources.",

        # Priorities
        "top_priorities": "Top 3 Priorities for Next Sprint",
        "business_benefit": "Business Benefit",
        "what_you_gain": "What you gain by implementing recommendations:",
        "more_traffic": "more organic traffic from Google",
        "higher_ctr": "higher CTR from search results and social media",
        "better_security": "better security and user trust",

        # Security
        "security_excellent": "Excellent Security",
        "security_good": "Good Security",
        "security_fair": "Fair Security",
        "security_poor": "Poor Security",
        "no_critical_issues": "No critical security issues (HTTPS active, no mixed content)",
        "hardening_recommended": "Recommended security hardening by adding security headers",

        # Scoring
        "availability": "Availability (HTTP 200)",
        "availability_desc": "Percentage of pages without HTTP errors (4xx, 5xx)",
        "meta_tags": "Meta Tags",
        "meta_tags_desc": "Presence and quality of Title and Description",
        "mobile_desc": "Responsive Design and Meta Viewport",
        "schema_desc": "Structured Data JSON-LD",
        "eeat_desc": "Expertise, Authority, Trust",
        "security_desc": "HTTPS, Security Headers, no Mixed Content",
    }
}


def get_translation(key: str, lang: str = "pl") -> str:
    """
    Pobiera tłumaczenie dla podanego klucza w wybranym języku.

    Args:
        key: Klucz tłumaczenia
        lang: Kod języka (pl, de, en)

    Returns:
        Przetłumaczony tekst lub klucz jeśli tłumaczenie nie istnieje
    """
    if lang not in TRANSLATIONS:
        lang = "pl"  # Domyślny język

    return TRANSLATIONS[lang].get(key, key)


def t(key: str, lang: str = "pl") -> str:
    """Skrócona wersja get_translation."""
    return get_translation(key, lang)
