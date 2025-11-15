"""
Generator raportu Word (DOCX)
"""
from typing import Dict, Any
import re

from compass.config import SHOW_REMEDIATIONS, USE_AI_SUMMARY, OPENAI_API_KEY

try:
    from docx import Document as WordDocument
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    HAS_DOCX = True
except Exception:
    HAS_DOCX = False
    print("⚠️  Brak biblioteki python-docx. Raport Word nie będzie wygenerowany.")


def add_hyperlink(paragraph, text, url):
    """
    Dodaje hyperlink do paragrafu w dokumencie Word.

    Args:
        paragraph: Paragraf w dokumencie Word
        text: Tekst linku
        url: URL linku
    """
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True
    )
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0563C1')
    rPr.append(color)
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def add_section_heading(doc, text, level=1, icon=None, color=RGBColor(31, 71, 136)):
    """
    Dodaje spójny nagłówek sekcji z opcjonalną ikonką emoji.
    """
    if icon:
        full_text = f"{icon} {text}"
    else:
        full_text = text
    heading = doc.add_heading(full_text, level=level)
    for run in heading.runs:
        run.font.color.rgb = color
    return heading


def add_status_line(doc, label, status_text, color, extra=None):
    """
    Dodaje wyróżnioną linię statusu (np. Status SEO, Status Security).
    """
    p = doc.add_paragraph()
    r1 = p.add_run(f"{label}: ")
    r1.bold = True
    r2 = p.add_run(status_text)
    r2.bold = True
    r2.font.color.rgb = color
    if extra:
        p.add_run(f" – {extra}")


def create_word_report(all_pages: Dict[str, Any], summary: Dict[str, Any],
                       issues: Dict[str, Any], duplicates: Dict, word_path: str):
    if not HAS_DOCX:
        print("⚠️  Pomijam generowanie raportu Word")
        return

    doc = WordDocument()

    # =========================
    # OKŁADKA RAPORTU + KPI
    # =========================
    title = doc.add_heading('Audyt SEO / AEO / GEO', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.runs[0]
    title_run.font.color.rgb = RGBColor(31, 71, 136)
    title_run.font.size = Pt(34)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{summary['start_url']}\n")
    run.font.size = Pt(18)
    run.bold = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"Data audytu: {summary['generated_at']}")
    run.font.size = Pt(10)
    run.italic = True
    run.font.color.rgb = RGBColor(120, 120, 120)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Adresaci: Zarząd / Marketing / Zespół IT")
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(120, 120, 120)

    doc.add_paragraph()

    # MAŁA TABELKA KPI NA OKŁADCE
    kpi_table = doc.add_table(rows=2, cols=3)
    kpi_table.style = 'Light Shading Accent 1'

    hdr = kpi_table.rows[0].cells
    hdr[0].text = "🎯 Kluczowy wynik"
    hdr[1].text = "✅ Stabilność techniczna"
    hdr[2].text = "📱 Mobilność"

    kpis = [
        ("Wynik audytu", f"{summary['overall_score']}/100", summary['overall_grade']),
        ("Strony OK (200)", summary['pages_ok'], f"z {summary['pages_analyzed']} analizowanych"),
        ("Mobile-friendly", f"{summary['mobile_percentage']}%", f"{summary['mobile_friendly_pages']} stron"),
    ]

    row = kpi_table.rows[1].cells
    for cell, (label, value, sub) in zip(row, kpis):
        p = cell.paragraphs[0]
        r1 = p.add_run(f"{label}\n")
        r1.bold = True
        r2 = p.add_run(f"{value}\n")
        r2.font.size = Pt(14)
        r2.bold = True
        r3 = p.add_run(str(sub))
        r3.font.size = Pt(9)
        r3.italic = True

    doc.add_paragraph()

    # WYNIK AUDYTU jako wyróżniony box
    box = doc.add_paragraph()
    run = box.add_run(f"WYNIK AUDYTU: {summary['overall_score']}/100  ({summary['overall_grade']})")
    run.font.size = Pt(26)
    run.bold = True
    run.font.color.rgb = RGBColor(31, 71, 136)

    doc.add_paragraph()

    # SKŁAD OCENY OGÓLNEJ
    add_section_heading(doc, 'Skład oceny ogólnej', level=3)
    scoring_table = doc.add_table(rows=1, cols=3)
    scoring_table.style = 'Light Grid Accent 1'
    hdr = scoring_table.rows[0].cells
    hdr[0].text = 'Kategoria'
    hdr[1].text = 'Wartość (x z y)'
    hdr[2].text = 'Opis'

    # Obliczenie wartości dla każdej kategorii
    pages = max(1, summary['pages_analyzed'])

    # Dostępność
    availability_pct = (summary['pages_ok'] / pages) * 100
    availability_value = round(availability_pct * 0.30)
    availability_text = f"{availability_value} z 30 ({availability_pct:.1f}%)"

    # Meta tagi
    meta_ok_pages = pages - (summary['missing_title'] + summary['missing_description'])
    meta_pct = max(0.0, (meta_ok_pages - 0.25 * (summary['title_issues'] + summary['description_issues'])) / pages * 100)
    meta_value = round(meta_pct * 0.15)
    meta_text = f"{meta_value} z 15 ({meta_pct:.1f}%)"

    # Mobile-friendly
    mobile_value = round(summary['mobile_percentage'] * 0.15)
    mobile_text = f"{mobile_value} z 15 ({summary['mobile_percentage']:.1f}%)"

    # Schema.org
    schema_pct = (summary['pages_with_schema'] / pages) * 100
    schema_value = round(schema_pct * 0.10)
    schema_text = f"{schema_value} z 10 ({schema_pct:.1f}%)"

    # E-E-A-T
    eeat_value = round(summary['avg_eeat_score'] * 0.10)
    eeat_text = f"{eeat_value} z 10 ({summary['avg_eeat_score']:.1f}%)"

    # Bezpieczeństwo
    security_pct = summary['avg_security_score']
    if summary['pages_no_ssl'] > 0:
        security_pct = max(0.0, security_pct - 10)
    security_value = round(security_pct * 0.20)
    security_text = f"{security_value} z 20 ({security_pct:.1f}%)"

    scoring_components = [
        ('Dostępność (HTTP 200)', availability_text, 'Procent stron bez błędów HTTP (4xx, 5xx)'),
        ('Meta tagi', meta_text, 'Obecność i jakość Title oraz Description'),
        ('Mobile-friendly', mobile_text, 'Responsywność i meta viewport'),
        ('Schema.org', schema_text, 'Dane strukturalne JSON-LD'),
        ('E-E-A-T', eeat_text, 'Eksperckość, Autorytet, Zaufanie'),
        ('Bezpieczeństwo', security_text, 'HTTPS, Security Headers, brak mixed content'),
    ]

    for category, weight, description in scoring_components:
        row = scoring_table.add_row().cells
        row[0].text = category
        row[1].text = weight
        row[2].text = description

    doc.add_paragraph()

    # Najważniejsze problemy z licznikami X/Y i opisami
    add_section_heading(doc, 'Najważniejsze problemy', level=2, icon='⚠️')

    # Słownik z opisami problemów
    problem_descriptions = {
        "pages_with_errors": "Strony z kodami błędów HTTP (4xx, 5xx). Uniemożliwiają indeksację i powodują utratę ruchu.",
        "missing_title": "Każda strona potrzebuje unikalnego tagu <title> (50–60 znaków) dla lepszej widoczności w wynikach wyszukiwania.",
        "missing_description": "Meta description (150–160 znaków) to pierwszy kontakt użytkownika z Twoją stroną w wynikach Google.",
        "missing_canonical": "Tag canonical zapobiega problemom z duplikacją treści i pomaga Google wybrać właściwą wersję strony.",
        "pages_without_viewport": "Brak meta viewport - wymaga ręcznej weryfikacji wyświetlania na urządzeniach mobilnych.",
        "pages_without_schema": "Dane strukturalne (Schema.org) pomagają Google lepiej zrozumieć zawartość i wyświetlać rich snippets.",
        "pages_poor_security": "Słabe zabezpieczenia (<50%). Brak security headers naraża użytkowników i obniża zaufanie do witryny.",
        "thin_content_pages": "Strony z mniej niż 300 słowami. Google preferuje wartościowe, szczegółowe treści.",
        "pages_weak_eeat": "Słabe sygnały E-E-A-T (<50%). Dodaj autora, datę publikacji, certyfikaty i linki do wiarygodnych źródeł.",
    }

    top_issues = [
        ("Błędy HTTP (4xx/5xx)", summary["pages_with_errors"], summary["pages_analyzed"], "pages_with_errors"),
        ("Brak Title", summary["missing_title"], summary["pages_analyzed"], "missing_title"),
        ("Brak Meta Description", summary["missing_description"], summary["pages_analyzed"], "missing_description"),
        ("Brak canonical", summary["missing_canonical"], summary["pages_analyzed"], "missing_canonical"),
        ("Brak meta viewport (mobile)", summary["pages_without_viewport"], summary["pages_analyzed"],
         "pages_without_viewport"),
        ("Brak Schema.org", summary["pages_without_schema"], summary["pages_analyzed"], "pages_without_schema"),
        ("Słabe bezpieczeństwo (<50%)", summary["pages_poor_security"], summary["pages_analyzed"],
         "pages_poor_security"),
        ("Cienka treść (<300 słów)", summary["thin_content_pages"], summary["pages_analyzed"], "thin_content_pages"),
        ("Słabe E-E-A-T (<50%)", summary["pages_weak_eeat"], summary["pages_analyzed"], "pages_weak_eeat"),
    ]

    for label, problem_count, total_pages, key in sorted(top_issues, key=lambda kv: kv[1], reverse=True)[:6]:
        if problem_count > 0:
            p = doc.add_paragraph(style='List Bullet')
            run = p.add_run(f"{label}: {problem_count}/{total_pages}\n")
            run.bold = True
            if key in problem_descriptions:
                run2 = p.add_run(problem_descriptions[key])
                run2.font.size = Pt(10)
                run2.italic = True
                run2.font.color.rgb = RGBColor(80, 80, 80)

    doc.add_page_break()

    # =========================
    # SPIS TREŚCI
    # =========================
    add_section_heading(doc, 'Spis treści', 1, icon='📋')
    toc_items = [
        "1. Executive Summary – Kluczowe liczby",
        "2. Priorytety (od krytycznych do lekkich)",
        "3. Analiza Meta Tagów",
        "4. Analiza Techniczna SEO",
        "5. Mobilność i Responsywność",
        "6. Open Graph i Twitter Cards",
        "7. Dane Strukturalne (Schema.org)",
        "8. E-E-A-T",
        "9. Local SEO (NAP)",
        "10. Jakość Treści",
        "11. Bezpieczeństwo (Security Headers)",
        "12. Legenda i objaśnienia",
        "13. AI-Powered Executive Summary",
    ]
    for item in toc_items:
        p = doc.add_paragraph(item, style='List Number')
    doc.add_page_break()

    # =========================
    # 1. EXECUTIVE SUMMARY
    # =========================
    add_section_heading(doc, '1. Executive Summary – Kluczowe liczby', 1, icon='📊')
    add_section_heading(doc, 'Szybkie podsumowanie SEO', 2, icon='📌')

    p = doc.add_paragraph()
    run = p.add_run('Stan strony w pigułce:\n')
    run.font.size = Pt(11)
    run.bold = True

    # Problemy KRYTYCZNE: 4xx, brak SSL, brak Title, brak Description
    seo_critical = summary['pages_with_errors'] + summary['pages_no_ssl'] + summary['missing_title'] + summary[
        'missing_description']

    # Ostrzeżenia: problemy z długością, canonical, schema, itp.
    seo_warnings = summary['title_issues'] + summary['description_issues'] + summary['missing_canonical'] + summary[
        'pages_without_schema']

    if seo_critical == 0 and seo_warnings < 10:
        seo_status = "✅ Świetnie!"
        seo_color = RGBColor(0, 150, 0)
        seo_text = "Drobne optymalizacje – fundamenty SEO są w dobrej kondycji."
    elif seo_critical < 5 and seo_warnings < 30:
        seo_status = "⚠️ Do poprawy"
        seo_color = RGBColor(200, 100, 0)
        seo_text = f"Problemy krytyczne: {seo_critical} | Ostrzeżenia: {seo_warnings}"
    else:
        seo_status = "🔴 Wymaga uwagi!"
        seo_color = RGBColor(200, 0, 0)
        seo_text = (
            f"Problemy krytyczne: {seo_critical} (błędy 4xx, brak SSL/Title/Description) | "
            f"Ostrzeżenia: {seo_warnings} (długość meta, canonical, schema)"
        )

    add_status_line(doc, "Status SEO", seo_status, seo_color, extra=seo_text)

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('🎯 Top 3 priorytety na najbliższy sprint')
    run.font.size = Pt(12)
    run.bold = True

    priorities = []
    total_pages = summary['pages_analyzed']

    if summary['pages_with_errors'] > 0:
        pct = round((summary['pages_with_errors'] / total_pages) * 100, 1)
        priorities.append({
            'text': f"Napraw {summary['pages_with_errors']}/{total_pages} stron z błędami HTTP 4xx ({pct}%)",
            'desc': "Strony niedostępne dla użytkowników i botów Google",
            'count': summary['pages_with_errors']
        })

    if summary['missing_title'] > 0:
        pct = round((summary['missing_title'] / total_pages) * 100, 1)
        priorities.append({
            'text': f"Dodaj Title do {summary['missing_title']}/{total_pages} stron ({pct}%)",
            'desc': "Brak tytułu = niewidoczność w wynikach Google",
            'count': summary['missing_title']
        })

    if summary['missing_description'] > 0 and len(priorities) < 3:
        pct = round((summary['missing_description'] / total_pages) * 100, 1)
        priorities.append({
            'text': f"Dodaj Meta Description do {summary['missing_description']}/{total_pages} stron ({pct}%)",
            'desc': "Wpływa na CTR (Click-Through Rate) z wyszukiwarki",
            'count': summary['missing_description']
        })

    if summary['pages_without_schema'] > 0 and len(priorities) < 3:
        pct = round((summary['pages_without_schema'] / total_pages) * 100, 1)
        priorities.append({
            'text': f"Dodaj Schema.org do {summary['pages_without_schema']}/{total_pages} stron ({pct}%)",
            'desc': "Brak rich snippets w Google (gwiazdki, FAQ, breadcrumbs)",
            'count': summary['pages_without_schema']
        })

    if summary['missing_canonical'] > 0 and len(priorities) < 3:
        pct = round((summary['missing_canonical'] / total_pages) * 100, 1)
        priorities.append({
            'text': f"Dodaj Canonical do {summary['missing_canonical']}/{total_pages} stron ({pct}%)",
            'desc': "Zapobiega problemom z duplikacją treści",
            'count': summary['missing_canonical']
        })

    if summary['pages_without_viewport'] > 0 and len(priorities) < 3:
        pct = round((summary['pages_without_viewport'] / total_pages) * 100, 1)
        priorities.append({
            'text': f"Sprawdź wyświetlanie mobilne {summary['pages_without_viewport']}/{total_pages} stron bez viewport ({pct}%)",
            'desc': "Wymaga ręcznej weryfikacji na urządzeniach mobilnych",
            'count': summary['pages_without_viewport']
        })

    # Sortujemy po count (malejąco) i bierzemy top 3
    priorities.sort(key=lambda x: x['count'], reverse=True)

    for i, pr in enumerate(priorities[:3], 1):
        p = doc.add_paragraph(style='List Number')
        run = p.add_run(f"{i}. {pr['text']}\n")
        run.bold = True
        run2 = p.add_run(f"   Efekt biznesowy: {pr['desc']}")
        run2.font.size = Pt(10)
        run2.italic = True
        run2.font.color.rgb = RGBColor(80, 80, 80)

    doc.add_paragraph()

    # Krótkie "co zyskasz"
    p = doc.add_paragraph()
    run = p.add_run("Co zyskasz po wdrożeniu rekomendacji:")
    run.bold = True
    doc.add_paragraph("• więcej ruchu organicznego z Google", style='List Bullet')
    doc.add_paragraph("• wyższy CTR z wyników wyszukiwania i social media", style='List Bullet')
    doc.add_paragraph("• lepsze bezpieczeństwo i zaufanie użytkowników", style='List Bullet')

    doc.add_paragraph()
    stats_table = doc.add_table(rows=1, cols=3)
    stats_table.style = 'Light Grid Accent 1'
    hdr = stats_table.rows[0].cells
    hdr[0].text = 'Metryka'
    hdr[1].text = 'Wartość'
    hdr[2].text = 'Opis'

    # Funkcja pomocnicza do dodawania wierszy z opisem
    def add_metric_row(metric_name, value, description="", emoji=None):
        row = stats_table.add_row().cells
        label = f"{emoji} {metric_name}" if emoji else metric_name
        row[0].text = label
        row[1].text = str(value)
        row[2].text = description

    # Metryki z opisami
    add_metric_row(
        'Przeskanowanych stron',
        f"{summary['pages_crawled']}",
        "Całkowita liczba odkrytych URL-i podczas crawlingu",
        emoji="🌐"
    )

    add_metric_row(
        'Przeanalizowanych stron',
        f"{summary['pages_analyzed']} (wykluczono {summary['pages_excluded']})",
        "Strony poddane analizie SEO (bez technicznych URL jak /cdn-cgi/*)",
        emoji="📄"
    )

    add_metric_row(
        'Strony OK (200)',
        f"✅ {summary['pages_ok']}",
        "Strony działające poprawnie, zwracające kod HTTP 200",
        emoji="✅"
    )

    add_metric_row(
        'Strony z błędami (4xx)',
        f"🔴 {summary['pages_with_errors']}",
        "Błędy klienta (404 Not Found, 403 Forbidden itp.) – blokują indeksację",
        emoji="🚫"
    )

    add_metric_row(
        'Brak Title',
        f"{summary['missing_title']}",
        "Strony bez tagu <title> – kluczowego dla pozycjonowania i CTR",
        emoji="🧾"
    )

    add_metric_row(
        'Brak Description',
        f"{summary['missing_description']}",
        "Strony bez <meta name=\"description\"> – wpływa na snippet w Google",
        emoji="📝"
    )

    add_metric_row(
        'Problemy Title (długość)',
        f"{summary['title_issues']}",
        "Title zbyt krótki (<30 znaków) lub za długi (>65) – Google może obciąć",
        emoji="📏"
    )

    add_metric_row(
        'Problemy Description (długość)',
        f"{summary['description_issues']}",
        "Description poza zakresem 120–165 znaków – może być obcięty lub zbyt krótki",
        emoji="📐"
    )

    add_metric_row(
        'Duplikaty Title',
        f"{summary['duplicate_titles']}",
        "Różne strony z identycznym tytułem – konfunduje Google i użytkowników",
        emoji="🔁"
    )

    add_metric_row(
        'Duplikaty Description',
        f"{summary['duplicate_descriptions']}",
        "Różne strony z tą samą meta description – obniża unikatowość",
        emoji="🔁"
    )

    add_metric_row(
        'Brak Canonical',
        f"{summary['missing_canonical']}",
        "Brak tagu <link rel=\"canonical\"> – prowadzi do problemów z duplikacją treści",
        emoji="🏷️"
    )

    add_metric_row(
        'Brak H1',
        f"{summary['missing_h1']}",
        "Strona bez nagłówka głównego <h1> – kluczowy element struktury i SEO",
        emoji="🔤"
    )

    add_metric_row(
        'Wiele H1',
        f"{summary['multiple_h1']}",
        "Strona z >1 nagłówkiem H1 – może wprowadzać wyszukiwarki w błąd",
        emoji="⚠️"
    )

    add_metric_row(
        'Obrazy bez ALT (łączna liczba)',
        f"{summary['total_images_without_alt']}",
        "Obrazki bez atrybutu alt – problem dla SEO, dostępności i obrazków Google",
        emoji="🖼️"
    )

    add_metric_row(
        'Mobile-friendly',
        f"{summary['mobile_friendly_pages']} ({summary['mobile_percentage']}%)",
        "Strony dostosowane do urządzeń mobilnych (responsive design)",
        emoji="📱"
    )

    add_metric_row(
        'Brak meta viewport',
        f"{summary['pages_without_viewport']}",
        "Strony bez <meta name=\"viewport\"> – nie skalują się na mobile",
        emoji="🔍"
    )

    add_metric_row(
        'Brak Open Graph',
        f"{summary['pages_without_og']}",
        "Strony bez meta tagów OG (Facebook, LinkedIn) – brzydki podgląd linków",
        emoji="📢"
    )

    add_metric_row(
        'Brak Twitter Cards',
        f"{summary['pages_without_twitter']}",
        "Strony bez Twitter Card – brak atrakcyjnego podglądu na X (Twitter)",
        emoji="🐦"
    )

    add_metric_row(
        'Strony z danymi strukturalnymi',
        f"{summary['pages_with_schema']}",
        "Liczba stron z Schema.org (JSON-LD) – umożliwia rich snippets w Google",
        emoji="🔗"
    )

    add_metric_row(
        'Brak Schema',
        f"{summary['pages_without_schema']}",
        "Strony bez Schema.org – tracisz gwiazdki, FAQ, breadcrumbs w wynikach",
        emoji="⚠️"
    )

    add_metric_row(
        'Śr. typów Schema/stronę',
        f"{summary['avg_schema_types']}",
        "Ile różnych typów Schema jest na stronie (Article, Product, FAQ, etc.)",
        emoji="📚"
    )

    add_metric_row(
        'Śr. E-E-A-T',
        f"{summary['avg_eeat_score']}%",
        "Experience, Expertise, Authoritativeness, Trustworthiness – sygnały jakości Google",
        emoji="🏆"
    )

    add_metric_row(
        'Słabe E-E-A-T',
        f"{summary['pages_weak_eeat']}",
        "Strony z niskim E-E-A-T (<50%): brak autora, dat, certyfikatów, źródeł",
        emoji="⚠️"
    )

    add_metric_row(
        'Local NAP OK',
        f"{summary['local_optimized_pages']}",
        "Strony z poprawnymi danymi NAP (Name, Address, Phone) – istotne dla firm lokalnych",
        emoji="📍"
    )

    add_metric_row(
        'Słaby Local SEO',
        f"{summary['pages_poor_local_seo']}",
        "Strony bez NAP, Schema LocalBusiness, linków do mapy – słabo dla SEO lokalnego",
        emoji="📉"
    )

    add_metric_row(
        'Thin content (<300 słów)',
        f"{summary['thin_content_pages']}",
        "Strony z bardzo krótką treścią – Google może uznać za low-quality",
        emoji="✂️"
    )

    add_metric_row(
        'Śr. Security',
        f"{summary['avg_security_score']}%",
        "Średni poziom zabezpieczeń (HTTPS + security headers). {:.1f}% to bardzo nisko".format(
            summary['avg_security_score']),
        emoji="🔒"
    )

    add_metric_row(
        'Słabe bezpieczeństwo',
        f"{summary['pages_poor_security']}",
        "Strony z oceną <50%: brak kluczowych nagłówków (HSTS, CSP, X-Frame-Options)",
        emoji="🛑"
    )

    add_metric_row(
        'Brakujące security headers',
        f"{summary['pages_missing_security_headers']}",
        "Strony z <3 nagłówkami security. Sprawdź: HSTS, CSP, X-Frame-Options, X-Content-Type-Options, Referrer-Policy",
        emoji="🧱"
    )

    add_metric_row(
        'Mixed content',
        f"{summary['pages_with_mixed_content']}",
        "Strony HTTPS z zasobami HTTP (obrazki, skrypty) – warning w przeglądarce",
        emoji="⚡"
    )

    add_metric_row(
        'WYNIK AUDYTU',
        f"{summary['overall_score']}/100 ({summary['overall_grade']})",
        "Ocena łączna: Availability (30%) + Meta (15%) + Mobile (15%) + Schema (10%) + E-E-A-T (10%) + Security (20%)",
        emoji="🏁"
    )

    doc.add_page_break()

    # =========================
    # 2. PRIORYTETY
    # =========================
    add_section_heading(doc, '2. Priorytety (od krytycznych do lekkich)', 1, icon='🎯')

    if issues['critical_errors']:
        add_section_heading(doc, 'Błędy HTTP (4xx)', 2, icon='🔴')
        p = doc.add_paragraph()
        total = summary['pages_analyzed']
        count = len(issues['critical_errors'])
        pct = round((count / total) * 100, 1) if total > 0 else 0
        p.add_run(f"Znaleziono {count}/{total} stron z błędami 4xx ({pct}%).").bold = True
        doc.add_paragraph(
            "Strony te są niedostępne dla użytkowników i robotów Google, co skutkuje:\n"
            "• Utratą ruchu organicznego\n"
            "• Negatywnym wpływem na UX\n"
            "• Problemami z indeksacją"
        )
        doc.add_paragraph()
        p = doc.add_paragraph()
        run = p.add_run("Problematyczne adresy:")
        run.bold = True
        for err in issues['critical_errors'][:20]:
            status_code = err.get('status', 'N/A')
            error_msg = err.get('error', '')
            if error_msg and error_msg != '':
                doc.add_paragraph(f"• {err['url']} – Status: {status_code} ({error_msg})", style='List Bullet')
            else:
                doc.add_paragraph(f"• {err['url']} – Status: {status_code}", style='List Bullet')
        if len(issues['critical_errors']) > 20:
            doc.add_paragraph(f"...oraz {len(issues['critical_errors']) - 20} innych")

    if issues['missing_title'] or issues['title_issues'] or issues['missing_description'] or issues['description_issues'] or issues['missing_canonical']:
        add_section_heading(doc, 'Meta & Canonical', 2, icon='🟠')

        # ===== META TITLE =====
        if issues['missing_title']:
            p = doc.add_paragraph()
            total = summary['pages_analyzed']
            count = len(issues['missing_title'])
            pct = round((count / total) * 100, 1) if total > 0 else 0
            run = p.add_run(f"Brakujące Meta Title: {count}/{total} stron ({pct}%)")
            run.bold = True
            doc.add_paragraph(
                "Tag <title> to pierwszy element, jaki użytkownik widzi w wynikach Google. Jego brak oznacza:"
            )
            doc.add_paragraph("• Brak kontroli nad tym, co Google wyświetli w SERP", style='List Bullet')
            doc.add_paragraph("• Niższy CTR (Click-Through Rate)", style='List Bullet')
            doc.add_paragraph("• Słabsze pozycjonowanie", style='List Bullet')
            doc.add_paragraph()
            p = doc.add_paragraph()
            run = p.add_run("Strony bez Title:")
            run.bold = True
            for url in issues['missing_title'][:15]:
                doc.add_paragraph(f"• {url}", style='List Bullet')
            if len(issues['missing_title']) > 15:
                doc.add_paragraph(f"...oraz {len(issues['missing_title']) - 15} innych")
            doc.add_paragraph()

        if issues['title_issues']:
            # Rozdzielenie na za krótkie i za długie
            too_short = [item for item in issues['title_issues'] if item.get('too_short')]
            too_long = [item for item in issues['title_issues'] if item.get('too_long')]

            if too_short:
                p = doc.add_paragraph()
                total = summary['pages_analyzed']
                count = len(too_short)
                pct = round((count / total) * 100, 1) if total > 0 else 0
                run = p.add_run(f"Meta Title za krótkie (<30 znaków): {count}/{total} stron ({pct}%)")
                run.bold = True
                doc.add_paragraph(
                    "Zbyt krótki tytuł nie w pełni wykorzystuje dostępne miejsce w wynikach Google (50-60 znaków). "
                    "Tracisz możliwość zawarcia kluczowych słów i przyciągnięcia uwagi użytkowników."
                )
                doc.add_paragraph()
                p = doc.add_paragraph()
                run = p.add_run("Strony z za krótkim Title:")
                run.bold = True
                for item in too_short[:15]:
                    url = item['url']
                    title = item.get('title', '')[:80]
                    length = item.get('length', 0)
                    doc.add_paragraph(f"• {url}\n  Title ({length} znaków): \"{title}\"", style='List Bullet')
                if len(too_short) > 15:
                    doc.add_paragraph(f"...oraz {len(too_short) - 15} innych")
                doc.add_paragraph()

            if too_long:
                p = doc.add_paragraph()
                total = summary['pages_analyzed']
                count = len(too_long)
                pct = round((count / total) * 100, 1) if total > 0 else 0
                run = p.add_run(f"Meta Title za długie (>65 znaków): {count}/{total} stron ({pct}%)")
                run.bold = True
                doc.add_paragraph(
                    "Zbyt długi tytuł zostanie obcięty w wynikach Google (wyświetlane jest ~50-60 znaków). "
                    "Ważne informacje mogą nie być widoczne dla użytkowników."
                )
                doc.add_paragraph()
                p = doc.add_paragraph()
                run = p.add_run("Strony z za długim Title:")
                run.bold = True
                for item in too_long[:15]:
                    url = item['url']
                    title = item.get('title', '')[:100]
                    length = item.get('length', 0)
                    doc.add_paragraph(f"• {url}\n  Title ({length} znaków): \"{title}...\"", style='List Bullet')
                if len(too_long) > 15:
                    doc.add_paragraph(f"...oraz {len(too_long) - 15} innych")
                doc.add_paragraph()

        # ===== META DESCRIPTION =====
        if issues['missing_description']:
            p = doc.add_paragraph()
            total = summary['pages_analyzed']
            count = len(issues['missing_description'])
            pct = round((count / total) * 100, 1) if total > 0 else 0
            run = p.add_run(f"Brakujące Meta Description: {count}/{total} stron ({pct}%)")
            run.bold = True
            doc.add_paragraph("Meta description to 'zachęta' do kliknięcia w wynikach wyszukiwania. Bez niej:")
            doc.add_paragraph("• Google sam generuje opis (często nietrafiający w sedno)", style='List Bullet')
            doc.add_paragraph("• Tracisz kontrolę nad przekazem marketingowym", style='List Bullet')
            doc.add_paragraph("• CTR może spaść nawet o 30–40%", style='List Bullet')
            doc.add_paragraph()
            p = doc.add_paragraph()
            run = p.add_run("Strony bez Meta Description:")
            run.bold = True
            for url in issues['missing_description'][:15]:
                doc.add_paragraph(f"• {url}", style='List Bullet')
            if len(issues['missing_description']) > 15:
                doc.add_paragraph(f"...oraz {len(issues['missing_description']) - 15} innych")
            doc.add_paragraph()

        if issues['description_issues']:
            # Rozdzielenie na za krótkie i za długie
            too_short_desc = [item for item in issues['description_issues'] if item.get('too_short')]
            too_long_desc = [item for item in issues['description_issues'] if item.get('too_long')]

            if too_short_desc:
                p = doc.add_paragraph()
                total = summary['pages_analyzed']
                count = len(too_short_desc)
                pct = round((count / total) * 100, 1) if total > 0 else 0
                run = p.add_run(f"Meta Description za krótkie (<120 znaków): {count}/{total} stron ({pct}%)")
                run.bold = True
                doc.add_paragraph(
                    "Zbyt krótki opis nie wykorzystuje dostępnego miejsca (120-165 znaków) i może nie przekonać użytkownika do kliknięcia."
                )
                doc.add_paragraph()
                p = doc.add_paragraph()
                run = p.add_run("Strony z za krótkim Description:")
                run.bold = True
                for item in too_short_desc[:15]:
                    url = item['url']
                    desc = item.get('description', '')
                    length = item.get('length', 0)
                    doc.add_paragraph(f"• {url}\n  Description ({length} znaków): \"{desc}\"", style='List Bullet')
                if len(too_short_desc) > 15:
                    doc.add_paragraph(f"...oraz {len(too_short_desc) - 15} innych")
                doc.add_paragraph()

            if too_long_desc:
                p = doc.add_paragraph()
                total = summary['pages_analyzed']
                count = len(too_long_desc)
                pct = round((count / total) * 100, 1) if total > 0 else 0
                run = p.add_run(f"Meta Description za długie (>165 znaków): {count}/{total} stron ({pct}%)")
                run.bold = True
                doc.add_paragraph(
                    "Zbyt długi opis zostanie obcięty w wynikach Google. Ważne informacje mogą być ukryte."
                )
                doc.add_paragraph()
                p = doc.add_paragraph()
                run = p.add_run("Strony z za długim Description:")
                run.bold = True
                for item in too_long_desc[:15]:
                    url = item['url']
                    desc = item.get('description', '')
                    length = item.get('length', 0)
                    doc.add_paragraph(f"• {url}\n  Description ({length} znaków): \"{desc}...\"", style='List Bullet')
                if len(too_long_desc) > 15:
                    doc.add_paragraph(f"...oraz {len(too_long_desc) - 15} innych")
                doc.add_paragraph()

        # ===== CANONICAL =====
        if issues['missing_canonical']:
            p = doc.add_paragraph()
            total = summary['pages_analyzed']
            count = len(issues['missing_canonical'])
            pct = round((count / total) * 100, 1) if total > 0 else 0
            run = p.add_run(f"Brak Canonical: {count}/{total} stron ({pct}%)")
            run.bold = True
            doc.add_paragraph(
                "Tag canonical wskazuje Google, która wersja strony jest 'główna'. Jego brak prowadzi do:"
            )
            doc.add_paragraph(
                "• Problemów z duplikacją treści (Google nie wie, którą wersję indeksować)",
                style='List Bullet'
            )
            doc.add_paragraph("• Rozproszenia 'mocy' linków między duplikatami", style='List Bullet')
            doc.add_paragraph("• Słabszego pozycjonowania wszystkich wersji", style='List Bullet')
            doc.add_paragraph()
            p = doc.add_paragraph()
            run = p.add_run("Strony bez Canonical:")
            run.bold = True
            for url in issues['missing_canonical'][:15]:
                doc.add_paragraph(f"• {url}", style='List Bullet')
            if len(issues['missing_canonical']) > 15:
                doc.add_paragraph(f"...oraz {len(issues['missing_canonical']) - 15} innych")
            doc.add_paragraph()

    if issues['poor_security'] or issues['missing_security_headers']:
        add_section_heading(doc, 'Bezpieczeństwo – niski poziom / brak nagłówków', 2, icon='🟠')

        if issues['poor_security']:
            p = doc.add_paragraph()
            total = summary['pages_analyzed']
            count = len(issues['poor_security'])
            pct = round((count / total) * 100, 1) if total > 0 else 0
            run = p.add_run(f"Słabe bezpieczeństwo: {count}/{total} stron ({pct}%)")
            run.bold = True
            doc.add_paragraph(
                "Strony z oceną bezpieczeństwa <50% mają braki w podstawowych nagłówkach zabezpieczających:"
            )
            doc.add_paragraph()
            p = doc.add_paragraph()
            run = p.add_run("Przykłady stron ze słabym security:")
            run.bold = True
            for item in issues['poor_security'][:10]:
                url = item['url']
                sec_pct = item.get('security_percentage', 0)
                missing = item.get('missing_headers', [])
                if missing:
                    missing_str = ", ".join(missing[:3])
                    doc.add_paragraph(f"• {url} ({sec_pct}%) – Brak: {missing_str}", style='List Bullet')
                else:
                    doc.add_paragraph(f"• {url} ({sec_pct}%)", style='List Bullet')
            if len(issues['poor_security']) > 10:
                doc.add_paragraph(f"...oraz {len(issues['poor_security']) - 10} innych")
            doc.add_paragraph()

        if issues['missing_security_headers']:
            p = doc.add_paragraph()
            total = summary['pages_analyzed']
            count = len(issues['missing_security_headers'])
            pct = round((count / total) * 100, 1) if total > 0 else 0
            run = p.add_run(
                f"Braki w security headers: {count}/{total} stron ({pct}%)"
            )
            run.bold = True
            doc.add_paragraph(
                "Strony z mniej niż 3 nagłówkami bezpieczeństwa są podatne na ataki. Brak odpowiednich headerów oznacza:"
            )
            doc.add_paragraph("• Łatwiejsze przeprowadzenie ataków XSS, clickjacking", style='List Bullet')
            doc.add_paragraph("• Brak wymuszenia HTTPS (możliwy man-in-the-middle)", style='List Bullet')
            doc.add_paragraph("• Obniżone zaufanie użytkowników i Google", style='List Bullet')
            doc.add_paragraph()
            p = doc.add_paragraph()
            run = p.add_run("Kluczowe brakujące nagłówki:")
            run.bold = True
            p = doc.add_paragraph()
            run = p.add_run("• HSTS: ")
            run.bold = True
            p.add_run("Wymusza HTTPS, chroni przed atakami man-in-the-middle")
            p = doc.add_paragraph()
            run = p.add_run("• CSP: ")
            run.bold = True
            p.add_run("Zapobiega atakom XSS (wstrzykiwanie złośliwego kodu)")
            p = doc.add_paragraph()
            run = p.add_run("• X-Frame-Options: ")
            run.bold = True
            p.add_run("Chroni przed clickjacking (osadzenie strony w iframe)")
            doc.add_paragraph()

    if issues['no_viewport'] or issues['no_og_tags'] or issues['no_twitter_cards'] or issues['missing_schema']:
        add_section_heading(doc, 'Mobile / Social / Schema', 2, icon='🟡')

        if issues['no_viewport']:
            p = doc.add_paragraph()
            total = summary['pages_analyzed']
            count = len(issues['no_viewport'])
            pct = round((count / total) * 100, 1) if total > 0 else 0
            run = p.add_run(f"Brak meta viewport: {count}/{total} stron ({pct}%)")
            run.bold = True
            doc.add_paragraph("Od 2018 Google stosuje mobile-first indexing. Brak meta viewport oznacza, że:")
            doc.add_paragraph("• Strona nie skaluje się poprawnie na smartfonach", style='List Bullet')
            doc.add_paragraph("• Google może obniżyć ranking (mobile-first!)", style='List Bullet')
            doc.add_paragraph("• Użytkownicy mobile widzą 'desktopową' wersję (zła UX)", style='List Bullet')
            doc.add_paragraph()

        if issues['no_og_tags']:
            p = doc.add_paragraph()
            total = summary['pages_analyzed']
            count = len(issues['no_og_tags'])
            pct = round((count / total) * 100, 1) if total > 0 else 0
            run = p.add_run(f"Braki w Open Graph: {count}/{total} stron ({pct}%)")
            run.bold = True
            doc.add_paragraph(
                "Open Graph to meta tagi używane przez Facebook, LinkedIn, WhatsApp do generowania podglądu linków."
            )
            doc.add_paragraph()

        if issues['no_twitter_cards']:
            p = doc.add_paragraph()
            total = summary['pages_analyzed']
            count = len(issues['no_twitter_cards'])
            pct = round((count / total) * 100, 1) if total > 0 else 0
            run = p.add_run(f"Brak Twitter Cards: {count}/{total} stron ({pct}%)")
            run.bold = True
            doc.add_paragraph("Twitter Cards to odpowiednik OG dla platformy X (dawniej Twitter).")
            doc.add_paragraph()

        if issues['missing_schema']:
            p = doc.add_paragraph()
            total = summary['pages_analyzed']
            count = len(issues['missing_schema'])
            pct = round((count / total) * 100, 1) if total > 0 else 0
            run = p.add_run(f"Brak Schema.org: {count}/{total} stron ({pct}%)")
            run.bold = True
            doc.add_paragraph(
                "Schema.org (JSON-LD) to 'język', którym mówisz do Google o zawartości strony. Bez niego:"
            )
            doc.add_paragraph("• Tracisz rich snippets (gwiazdki, FAQ, breadcrumbs)", style='List Bullet')
            doc.add_paragraph("• Trudniej o featured snippet (pozycja 0)", style='List Bullet')
            doc.add_paragraph("• Google słabiej rozumie kontekst treści", style='List Bullet')
            doc.add_paragraph()

    if issues['weak_eeat'] or issues['thin_content']:
        add_section_heading(doc, 'E-E-A-T & Treść', 2, icon='🟡')

        if issues['weak_eeat']:
            p = doc.add_paragraph()
            total = summary['pages_analyzed']
            count = len(issues['weak_eeat'])
            pct = round((count / total) * 100, 1) if total > 0 else 0
            run = p.add_run(f"Słabe E-E-A-T: {count}/{total} stron ({pct}%)")
            run.bold = True
            doc.add_paragraph(
                "E-E-A-T (Experience, Expertise, Authoritativeness, Trustworthiness) to zestaw sygnałów jakości dla Google."
            )
            doc.add_paragraph()

        if issues['thin_content']:
            p = doc.add_paragraph()
            total = summary['pages_analyzed']
            count = len(issues['thin_content'])
            pct = round((count / total) * 100, 1) if total > 0 else 0
            run = p.add_run(f"Thin content: {count}/{total} stron ({pct}%)")
            run.bold = True
            doc.add_paragraph(
                "Strony z mniej niż 300 słowami mogą być uznane przez Google za 'cienkie' (low-quality)."
            )
            doc.add_paragraph()

    doc.add_page_break()

    # =========================
    # 5. MOBILNOŚĆ
    # =========================
    add_section_heading(doc, '5. Mobilność i Responsywność', 1, icon='📱')
    p = doc.add_paragraph()
    run = p.add_run('Status mobilności: ')
    run.bold = True
    if summary['mobile_percentage'] >= 90:
        run = p.add_run(f"✅ {summary['mobile_percentage']}% stron mobile-friendly")
        run.font.color.rgb = RGBColor(0, 150, 0)
    elif summary['mobile_percentage'] >= 70:
        run = p.add_run(f"⚠️ {summary['mobile_percentage']}% stron mobile-friendly")
        run.font.color.rgb = RGBColor(200, 100, 0)
    else:
        run = p.add_run(f"🔴 {summary['mobile_percentage']}% stron mobile-friendly")
        run.font.color.rgb = RGBColor(200, 0, 0)

    doc.add_paragraph()
    doc.add_paragraph(
        "Google analizuje najpierw wersję mobilną strony (mobile-first indexing). Brak responsywności oznacza:"
    )
    doc.add_paragraph("• Spadek pozycji w wynikach mobilnych (60%+ ruchu to mobile)", style='List Bullet')
    doc.add_paragraph("• Gorsze doświadczenie użytkownika = wyższy bounce rate", style='List Bullet')
    doc.add_paragraph("• Utratę potencjalnych klientów mobilnych", style='List Bullet')

    if issues['no_viewport']:
        doc.add_paragraph()
        p = doc.add_paragraph()
        run = p.add_run(f"Strony bez meta viewport ({len(issues['no_viewport'])}):")
        run.bold = True
        doc.add_paragraph(
            "⚠️ UWAGA: Sam brak meta viewport nie gwarantuje, że strona wyświetla się niepoprawnie na urządzeniach mobilnych. "
            "Jest to jedynie wskaźnik techniczny. Zalecamy ręczną weryfikację wyświetlania na rzeczywistych urządzeniach mobilnych "
            "lub przy użyciu narzędzi takich jak Google PageSpeed Insights, Lighthouse czy Chrome DevTools."
        )
        doc.add_paragraph()

        # Filtrujemy pliki multimedialne (mp4, mp3, jpg, png, pdf, etc.)
        multimedia_extensions = ['.mp4', '.mp3', '.avi', '.mov', '.wmv', '.flv', '.webm', '.mkv',
                                 '.jpg', '.jpeg', '.png', '.gif', '.webp', '.svg', '.pdf', '.zip',
                                 '.rar', '.doc', '.docx', '.xls', '.xlsx']

        filtered_urls = []
        for url in issues['no_viewport']:
            url_lower = url.lower()
            is_multimedia = any(url_lower.endswith(ext) for ext in multimedia_extensions)
            if not is_multimedia:
                filtered_urls.append(url)

        if filtered_urls:
            p = doc.add_paragraph()
            run = p.add_run("Strony HTML wymagające poprawy:")
            run.bold = True
            run.font.size = Pt(11)
            for url in filtered_urls[:15]:
                doc.add_paragraph(f"• {url}", style='List Bullet')
            if len(filtered_urls) > 15:
                doc.add_paragraph(f"...oraz {len(filtered_urls) - 15} innych stron")

        # Jeśli są pliki multimedialne, informujemy o nich osobno
        multimedia_count = len(issues['no_viewport']) - len(filtered_urls)
        if multimedia_count > 0:
            doc.add_paragraph()
            p = doc.add_paragraph()
            run = p.add_run(f"ℹ️  Dodatkowo znaleziono {multimedia_count} plików multimedialnych bez viewport ")
            run.font.size = Pt(9)
            run.italic = True
            run.font.color.rgb = RGBColor(100, 100, 100)
            run2 = p.add_run("(pliki video/obrazy – to normalne, nie wymaga poprawy)")
            run2.font.size = Pt(9)
            run2.italic = True
            run2.font.color.rgb = RGBColor(100, 100, 100)

        if SHOW_REMEDIATIONS:
            doc.add_paragraph()
            p = doc.add_paragraph()
            run = p.add_run('Jak dodać meta viewport:\n')
            run.bold = True
            code = '<meta name="viewport" content="width=device-width, initial-scale=1.0">'
            q = doc.add_paragraph()
            r = q.add_run(code)
            r.font.name = 'Courier New'
            r.font.size = Pt(10)

    doc.add_page_break()

    # =========================
    # 6. OPEN GRAPH / TWITTER
    # =========================
    add_section_heading(doc, '6. Open Graph i Twitter Cards', 1, icon='📢')

    doc.add_paragraph(
        "Social media to potężne źródło ruchu. Gdy ktoś udostępnia link do Twojej strony na Facebooku, LinkedIn czy X (Twitter), te platformy generują 'podgląd' – miniaturkę z obrazkiem, tytułem i opisem. "
        "To pierwsze wrażenie decyduje, czy użytkownik kliknie."
    )
    doc.add_paragraph()

    if issues['no_og_tags']:
        p = doc.add_paragraph()
        total = summary['pages_analyzed']
        count = len(issues['no_og_tags'])
        pct = round((count / total) * 100, 1) if total > 0 else 0
        run = p.add_run(f"Braki w Open Graph: {count}/{total} stron ({pct}%)")
        run.bold = True
        doc.add_paragraph("Open Graph to meta tagi używane przez Facebook, LinkedIn, WhatsApp, Messenger. Bez nich:")
        doc.add_paragraph("• Platformy same wybierają obrazek (często nietrafiający w sedno)", style='List Bullet')
        doc.add_paragraph("• Opis może być losowy fragment tekstu", style='List Bullet')
        doc.add_paragraph("• Tracisz kontrolę nad przekazem marketingowym", style='List Bullet')
        doc.add_paragraph("• Niższy CTR z social media (nawet o 50%!)", style='List Bullet')
        doc.add_paragraph()

        p = doc.add_paragraph()
        run = p.add_run("Strony wymagające dodania Open Graph:")
        run.bold = True
        run.font.size = Pt(11)
        for item in issues['no_og_tags'][:15]:
            url = item['url']
            has_image = item.get('has_og_image', False)
            has_title = item.get('has_og_title', False)
            has_desc = item.get('has_og_description', False)

            missing_parts = []
            if not has_image:
                missing_parts.append("obrazek")
            if not has_title:
                missing_parts.append("tytuł")
            if not has_desc:
                missing_parts.append("opis")

            if missing_parts:
                doc.add_paragraph(f"• {url}\n  Brak: {', '.join(missing_parts)}", style='List Bullet')
            else:
                doc.add_paragraph(f"• {url}", style='List Bullet')

        if len(issues['no_og_tags']) > 15:
            doc.add_paragraph(f"...oraz {len(issues['no_og_tags']) - 15} innych stron")
        doc.add_paragraph()

    if issues['no_twitter_cards']:
        p = doc.add_paragraph()
        total = summary['pages_analyzed']
        count = len(issues['no_twitter_cards'])
        pct = round((count / total) * 100, 1) if total > 0 else 0
        run = p.add_run(f"Brak Twitter Cards: {count}/{total} stron ({pct}%)")
        run.bold = True
        doc.add_paragraph(
            "Twitter Cards to odpowiednik Open Graph dla platformy X (dawniej Twitter). Działają analogicznie – "
            "kontrolują, jak Twój link wygląda po udostępnieniu."
        )
        doc.add_paragraph()

        p = doc.add_paragraph()
        run = p.add_run("Strony wymagające dodania Twitter Cards:")
        run.bold = True
        run.font.size = Pt(11)
        for url in issues['no_twitter_cards'][:15]:
            doc.add_paragraph(f"• {url}", style='List Bullet')
        if len(issues['no_twitter_cards']) > 15:
            doc.add_paragraph(f"...oraz {len(issues['no_twitter_cards']) - 15} innych stron")
        doc.add_paragraph()

    if SHOW_REMEDIATIONS:
        example_code = '''<!-- Open Graph -->
<meta property="og:title" content="Tytuł" />
<meta property="og:description" content="Opis" />
<meta property="og:image" content="https://twoja-domena.pl/obraz.jpg" />
<meta property="og:url" content="https://twoja-domena.pl/strona" />
<meta property="og:type" content="website" />
<!-- Twitter Cards -->
<meta name="twitter:card" content="summary_large_image" />
<meta name="twitter:title" content="Tytuł" />
<meta name="twitter:description" content="Opis" />
<meta name="twitter:image" content="https://twoja-domena.pl/obraz.jpg" />'''
        p = doc.add_paragraph()
        run = p.add_run('Przykładowe tagi:')
        run.bold = True
        q = doc.add_paragraph()
        r = q.add_run(example_code)
        r.font.name = 'Courier New'
        r.font.size = Pt(9)

    doc.add_page_break()

    # =========================
    # 7. SCHEMA
    # =========================
    add_section_heading(doc, '7. Dane Strukturalne (Schema.org)', 1, icon='🔗')

    doc.add_paragraph(
        "Schema.org (JSON-LD) to 'język techniczny', którym mówisz Google o zawartości swojej strony. "
        "Dzięki niemu Google może wyświetlić Twoją stronę w bardziej atrakcyjny sposób w wynikach wyszukiwania – tzw. rich snippets."
    )
    doc.add_paragraph()

    schema_percentage = (summary['pages_with_schema'] / max(1, summary['pages_analyzed'])) * 100

    if schema_percentage >= 70:
        p = doc.add_paragraph()
        run = p.add_run(f"✅ {summary['pages_with_schema']} stron ma dane strukturalne ({schema_percentage:.1f}%)")
        run.font.color.rgb = RGBColor(0, 150, 0)
        run.bold = True
    else:
        p = doc.add_paragraph()
        run = p.add_run(
            f"⚠️ Tylko {summary['pages_with_schema']} stron ma dane strukturalne ({schema_percentage:.1f}%)"
        )
        run.font.color.rgb = RGBColor(200, 100, 0)
        run.bold = True

    doc.add_paragraph()
    doc.add_paragraph("Co tracisz bez Schema.org:")
    doc.add_paragraph("• Rich snippets: gwiazdki ocen, ceny produktów, FAQ rozwijane w SERP", style='List Bullet')
    doc.add_paragraph("• Breadcrumbs (ścieżka nawigacji) w wynikach Google", style='List Bullet')
    doc.add_paragraph("• Featured snippet (pozycja 0) – trudniej bez struktury", style='List Bullet')
    doc.add_paragraph("• Karuzele produktów/artykułów w mobilnych wynikach", style='List Bullet')
    doc.add_paragraph("• Lepsze zrozumienie kontekstu przez Google (ważne dla AI)", style='List Bullet')
    doc.add_paragraph()

    if issues['missing_schema']:
        p = doc.add_paragraph()
        total = summary['pages_analyzed']
        count = len(issues['missing_schema'])
        pct = round((count / total) * 100, 1) if total > 0 else 0
        run = p.add_run(f"Strony bez Schema.org: {count}/{total} ({pct}%)")
        run.bold = True
        doc.add_paragraph()

        p = doc.add_paragraph()
        run = p.add_run("Przykłady stron wymagających Schema:")
        run.bold = True
        run.font.size = Pt(11)
        for url in issues['missing_schema'][:20]:
            doc.add_paragraph(f"• {url}", style='List Bullet')
        if len(issues['missing_schema']) > 20:
            doc.add_paragraph(f"...oraz {len(issues['missing_schema']) - 20} innych stron")

        doc.add_paragraph()
        p = doc.add_paragraph()
        run = p.add_run("💡 Rekomendacja:")
        run.bold = True
        run.font.color.rgb = RGBColor(50, 100, 200)
        doc.add_paragraph("Priorytetowo dodaj Schema do:")
        doc.add_paragraph("1. Strony głównej (Organization/LocalBusiness)", style='List Number')
        doc.add_paragraph("2. Stron produktów (Product z ceną i dostępnością)", style='List Number')
        doc.add_paragraph("3. Artykułów blogowych (Article/BlogPosting)", style='List Number')
        doc.add_paragraph("4. FAQ/Pytania (FAQPage)", style='List Number')
        doc.add_paragraph("5. Opinii klientów (Review/AggregateRating)", style='List Number')

    doc.add_page_break()

    # =========================
    # 8. E-E-A-T
    # =========================
    add_section_heading(doc, '8. E-E-A-T', 1, icon='🏆')

    doc.add_paragraph(
        "E-E-A-T to akronim od Experience, Expertise, Authoritativeness, Trustworthiness – po polsku: "
        "Doświadczenie, Eksperckość, Autorytet, Zaufanie."
    )
    doc.add_paragraph()

    doc.add_paragraph("Co sprawdza Google oceniając E-E-A-T:")
    doc.add_paragraph("• Czy autor jest ekspertem w temacie? (widoczne imię, nazwisko, bio)", style='List Bullet')
    doc.add_paragraph("• Czy treść jest aktualna? (data publikacji, data aktualizacji)", style='List Bullet')
    doc.add_paragraph("• Czy strona jest autorytetem? (linki z wiarygodnych źródeł .edu/.gov)", style='List Bullet')
    doc.add_paragraph("• Czy można zaufać? (HTTPS, dane kontaktowe, polityka prywatności)", style='List Bullet')
    doc.add_paragraph("• Czy są opinie/recenzje? (social proof)", style='List Bullet')
    doc.add_paragraph()

    p = doc.add_paragraph()
    run = p.add_run('Średnia ocena E-E-A-T: ')
    run.bold = True
    if summary['avg_eeat_score'] >= 70:
        run = p.add_run(f"✅ {summary['avg_eeat_score']}%")
        run.font.color.rgb = RGBColor(0, 150, 0)
    elif summary['avg_eeat_score'] >= 50:
        run = p.add_run(f"⚠️ {summary['avg_eeat_score']}%")
        run.font.color.rgb = RGBColor(200, 100, 0)
    else:
        run = p.add_run(f"🔴 {summary['avg_eeat_score']}%")
        run.font.color.rgb = RGBColor(200, 0, 0)

    doc.add_paragraph()

    if issues['weak_eeat']:
        p = doc.add_paragraph()
        total = summary['pages_analyzed']
        count = len(issues['weak_eeat'])
        pct = round((count / total) * 100, 1) if total > 0 else 0
        run = p.add_run(f"Strony ze słabym E-E-A-T (<50%): {count}/{total} ({pct}%)")
        run.bold = True
        doc.add_paragraph(
            "Te strony mają niską ocenę zaufania w oczach Google. To nie znaczy, że są 'złe' – po prostu brakuje im sygnałów jakości."
        )
        doc.add_paragraph()

        p = doc.add_paragraph()
        run = p.add_run("Przykłady stron ze słabym E-E-A-T:")
        run.bold = True
        run.font.size = Pt(11)
        for item in issues['weak_eeat'][:15]:
            url = item['url']
            eeat_pct = item.get('eeat_percentage', 0)
            missing = item.get('missing', [])

            missing_readable = []
            for key in missing:
                if key == 'has_author':
                    missing_readable.append("autor")
                elif key == 'has_date':
                    missing_readable.append("data publikacji")
                elif key == 'has_expertise_signals':
                    missing_readable.append("sygnały eksperckości")
                elif key == 'has_quality_external_links':
                    missing_readable.append("linki do źródeł")
                elif key == 'has_contact_info':
                    missing_readable.append("dane kontaktowe")
                elif key == 'has_reviews':
                    missing_readable.append("opinie")

            if missing_readable:
                doc.add_paragraph(
                    f"• {url} ({eeat_pct:.1f}%)\n  Brak: {', '.join(missing_readable[:3])}",
                    style='List Bullet'
                )
            else:
                doc.add_paragraph(f"• {url} ({eeat_pct:.1f}%)", style='List Bullet')

        if len(issues['weak_eeat']) > 15:
            doc.add_paragraph(f"...oraz {len(issues['weak_eeat']) - 15} innych stron")

        doc.add_paragraph()
        p = doc.add_paragraph()
        run = p.add_run("💡 Jak poprawić E-E-A-T:")
        run.bold = True
        run.font.color.rgb = RGBColor(50, 100, 200)
        doc.add_paragraph("1. Dodaj ramkę 'O autorze' z imieniem, nazwiskiem, zdjęciem, bio", style='List Number')
        doc.add_paragraph("2. Wstaw datę publikacji i 'Ostatnia aktualizacja: [data]'", style='List Number')
        doc.add_paragraph("3. Linkuj do wiarygodnych źródeł (.edu, .gov, badania naukowe)", style='List Number')
        doc.add_paragraph("4. Dodaj certyfikaty, nagrody, doświadczenie firmy", style='List Number')
        doc.add_paragraph("5. Umieść widoczne dane kontaktowe (telefon, email, adres)", style='List Number')

    doc.add_page_break()

    # =========================
    # 9. LOCAL SEO (NAP)
    # =========================
    add_section_heading(doc, '9. Local SEO (NAP)', 1, icon='📍')

    doc.add_paragraph(
        "NAP to skrót od Name, Address, Phone – czyli nazwa firmy, adres i telefon. To podstawa lokalnego SEO."
    )
    doc.add_paragraph("Google sprawdza, czy dane NAP są:",)
    doc.add_paragraph("• Spójne (takie same wszędzie: strona, Google Maps, Facebook, wizytówki)", style='List Bullet')
    doc.add_paragraph("• Widoczne (łatwo znaleźć na stronie)", style='List Bullet')
    doc.add_paragraph("• Ustrukturyzowane (Schema.org LocalBusiness)", style='List Bullet')
    doc.add_paragraph()

    local_percentage = (summary['local_optimized_pages'] / max(1, summary['pages_analyzed'])) * 100

    if local_percentage >= 50:
        p = doc.add_paragraph()
        run = p.add_run(
            f"✅ {summary['local_optimized_pages']} stron zoptymalizowanych pod NAP ({local_percentage:.1f}%)"
        )
        run.font.color.rgb = RGBColor(0, 150, 0)
        run.bold = True
    else:
        p = doc.add_paragraph()
        run = p.add_run(
            f"⚠️ Tylko {summary['local_optimized_pages']} stron posiada NAP ({local_percentage:.1f}%)"
        )
        run.font.color.rgb = RGBColor(200, 100, 0)
        run.bold = True

    doc.add_paragraph()
    doc.add_paragraph("Dlaczego NAP jest ważny:")
    doc.add_paragraph("• Google Local Pack (3 wyniki na mapie) wymaga spójnych danych", style='List Bullet')
    doc.add_paragraph("• Użytkownicy szukający 'firma + miasto' trafiają na lokalne wyniki", style='List Bullet')
    doc.add_paragraph("• Zaufanie: widoczny telefon i adres = większa konwersja", style='List Bullet')
    doc.add_paragraph("• Voice search ('Hey Google, znajdź X w pobliżu') preferuje NAP", style='List Bullet')
    doc.add_paragraph()

    if issues['poor_local_seo']:
        p = doc.add_paragraph()
        total = summary['pages_analyzed']
        count = len(issues['poor_local_seo'])
        pct = round((count / total) * 100, 1) if total > 0 else 0
        run = p.add_run(f"Strony bez poprawnego NAP: {count}/{total} ({pct}%)")
        run.bold = True
        doc.add_paragraph()

        p = doc.add_paragraph()
        run = p.add_run("Przykłady stron wymagających poprawy Local SEO:")
        run.bold = True
        run.font.size = Pt(11)
        for item in issues['poor_local_seo'][:15]:
            url = item['url']
            nap_score = item.get('nap_score', 0)
            phones = item.get('phone_numbers', 0)
            has_address = item.get('has_address', False)
            has_schema = item.get('has_local_schema', False)

            issues_list = []
            if phones == 0:
                issues_list.append("brak telefonu")
            if not has_address:
                issues_list.append("brak adresu")
            if not has_schema:
                issues_list.append("brak Schema LocalBusiness")

            if issues_list:
                doc.add_paragraph(
                    f"• {url} (NAP: {nap_score}/3)\n  Problem: {', '.join(issues_list)}",
                    style='List Bullet'
                )
            else:
                doc.add_paragraph(f"• {url} (NAP: {nap_score}/3)", style='List Bullet')

        if len(issues['poor_local_seo']) > 15:
            doc.add_paragraph(f"...oraz {len(issues['poor_local_seo']) - 15} innych stron")

        doc.add_paragraph()
        p = doc.add_paragraph()
        run = p.add_run("💡 Jak poprawić Local SEO:")
        run.bold = True
        run.font.color.rgb = RGBColor(50, 100, 200)
        doc.add_paragraph("1. Dodaj widoczną stopkę z: nazwa firmy, adres, telefon, email", style='List Number')
        doc.add_paragraph("2. Wdróż Schema.org LocalBusiness (JSON-LD) z pełnymi danymi NAP", style='List Number')
        doc.add_paragraph("3. Sprawdź spójność: te same dane na Google Maps, Facebook, stronie", style='List Number')
        doc.add_paragraph("4. Dodaj mapę Google (embed) ze wskazaniem lokalizacji", style='List Number')
        doc.add_paragraph("5. Stwórz dedykowaną podstronę 'Kontakt' z pełnymi danymi", style='List Number')

    doc.add_page_break()

    # =========================
    # 10. JAKOŚĆ TREŚCI
    # =========================
    add_section_heading(doc, '10. Jakość Treści', 1, icon='📝')

    doc.add_paragraph(
        "Google nie lubi 'cienkich' stron – czyli takich, które mają bardzo mało tekstu (poniżej 300 słów). "
        "Im więcej merytorycznej treści, tym lepiej."
    )
    doc.add_paragraph()

    doc.add_paragraph("Co ryzykujesz mając thin content:")
    doc.add_paragraph("• Google może uznać stronę za low-quality i obniżyć jej ranking", style='List Bullet')
    doc.add_paragraph("• Trudniej o featured snippet (pozycja 0) – potrzeba więcej kontekstu", style='List Bullet')
    doc.add_paragraph("• Użytkownicy szybko opuszczają stronę (wysoki bounce rate)", style='List Bullet')
    doc.add_paragraph("• Mniejsza szansa na linki zewnętrzne (nikt nie linkuje 100-słownego tekstu)", style='List Bullet')
    doc.add_paragraph()

    if issues['thin_content']:
        p = doc.add_paragraph()
        total = summary['pages_analyzed']
        count = len(issues['thin_content'])
        pct = round((count / total) * 100, 1) if total > 0 else 0
        run = p.add_run(f"Strony z thin content (<300 słów): {count}/{total} ({pct}%)")
        run.bold = True
        doc.add_paragraph()

        p = doc.add_paragraph()
        run = p.add_run("Strony wymagające rozbudowania treści:")
        run.bold = True
        run.font.size = Pt(11)
        for item in issues['thin_content']:
            url = item['url']
            word_count = item.get('word_count', 0)
            text_len = item.get('text_len', 0)

            doc.add_paragraph(f"• {url}\n  Słów: {word_count}, Znaków: {text_len}", style='List Bullet')

        doc.add_paragraph()
        p = doc.add_paragraph()
        run = p.add_run("💡 Jak poprawić:")
        run.bold = True
        run.font.color.rgb = RGBColor(50, 100, 200)
        doc.add_paragraph("1. Rozbuduj treść do minimum 600–800 słów (artykuły: 1500+ słów)", style='List Number')
        doc.add_paragraph("2. Dodaj wartość: praktyczne porady, case studies, przykłady", style='List Number')
        doc.add_paragraph("3. Strukturyzuj: nagłówki H2/H3, listy punktowane, wyróżnienia", style='List Number')
        doc.add_paragraph(
            "4. Multimedialność: obrazy, infografiki, video (zaliczają się do 'treści')",
            style='List Number'
        )
        doc.add_paragraph(
            "5. FAQ: dodaj sekcję pytań i odpowiedzi (boost dla SEO i użytkownika)",
            style='List Number'
        )
    else:
        p = doc.add_paragraph()
        run = p.add_run("✅ Brak problemów z thin content")
        run.font.color.rgb = RGBColor(0, 150, 0)
        run.bold = True
        doc.add_paragraph("Wszystkie strony mają odpowiednią ilość treści (>300 słów).")

    doc.add_page_break()

    # =========================
    # 11. BEZPIECZEŃSTWO
    # =========================
    add_section_heading(doc, '11. Bezpieczeństwo (Security Headers)', 1, icon='🔒')

    doc.add_paragraph(
        "Security headers to specjalne nagłówki HTTP, które serwer wysyła do przeglądarki, informując ją "
        "'jak ma się zachować' z punktu widzenia bezpieczeństwa."
    )
    doc.add_paragraph()

    if summary['avg_security_score'] >= 80:
        sec_status = "✅ Dobry"
        sec_color = RGBColor(0, 150, 0)
    elif summary['avg_security_score'] >= 60:
        sec_status = "⚠️ Średni"
        sec_color = RGBColor(200, 100, 0)
    elif summary['avg_security_score'] >= 40:
        sec_status = "🟠 Słaby"
        sec_color = RGBColor(200, 50, 0)
    else:
        sec_status = "🔴 Krytyczny"
        sec_color = RGBColor(200, 0, 0)

    add_status_line(
        doc,
        "Status Security",
        sec_status,
        sec_color,
        extra=f"Śr. ocena: {summary['avg_security_score']}%"
    )

    doc.add_paragraph()
    doc.add_paragraph("Co ryzykujesz bez security headers:")
    doc.add_paragraph("• Ataki XSS (Cross-Site Scripting) – wstrzykiwanie złośliwego kodu", style='List Bullet')
    doc.add_paragraph("• Clickjacking – nakładanie niewidocznych przycisków na Twoją stronę", style='List Bullet')
    doc.add_paragraph("• Man-in-the-middle – przechwytywanie danych użytkowników", style='List Bullet')
    doc.add_paragraph("• Obniżone zaufanie użytkowników (przeglądarki pokazują warningi)", style='List Bullet')
    doc.add_paragraph("• Gorsze pozycjonowanie (Google preferuje bezpieczne strony)", style='List Bullet')
    doc.add_paragraph()

    has_security_issues = (
        summary['pages_poor_security'] > 0 or
        summary['pages_missing_security_headers'] > 0 or
        summary['pages_with_mixed_content'] > 0
    )

    if not has_security_issues:
        p = doc.add_paragraph()
        run = p.add_run("✅ Brak poważnych problemów bezpieczeństwa.")
        run.font.color.rgb = RGBColor(0, 150, 0)
        run.bold = True
    else:
        if issues['poor_security']:
            p = doc.add_paragraph()
            total = summary['pages_analyzed']
            count = len(issues['poor_security'])
            pct = round((count / total) * 100, 1) if total > 0 else 0
            run = p.add_run(f"🟠 Słabe bezpieczeństwo: {count}/{total} stron ({pct}%)")
            run.bold = True
            doc.add_paragraph(
                "Strony z oceną security <50% mają krytyczne braki w podstawowych nagłówkach zabezpieczających. "
                "Poniżej endpointy z największymi problemami:"
            )
            doc.add_paragraph()

            for item in issues['poor_security'][:10]:
                url = item['url']
                sec_pct = item.get('security_percentage', 0)
                missing = item.get('missing_headers', [])

                p = doc.add_paragraph(style='List Bullet')
                run = p.add_run(f"{url} ")
                run.font.size = Pt(9)
                run2 = p.add_run(f"({sec_pct:.1f}%)\n")
                run2.font.color.rgb = RGBColor(200, 0, 0)
                run2.font.size = Pt(9)

                if missing:
                    run3 = p.add_run(f"  Brakuje: {', '.join(missing[:4])}")
                    run3.font.size = Pt(9)
                    run3.italic = True

            if len(issues['poor_security']) > 10:
                doc.add_paragraph(f"...oraz {len(issues['poor_security']) - 10} innych stron")
            doc.add_paragraph()

        if issues['missing_security_headers']:
            p = doc.add_paragraph()
            total = summary['pages_analyzed']
            count = len(issues['missing_security_headers'])
            pct = round((count / total) * 100, 1) if total > 0 else 0
            run = p.add_run(f"🟡 Braki w security headers: {count}/{total} stron ({pct}%)")
            run.bold = True
            doc.add_paragraph("Strony z mniej niż 3 nagłówkami bezpieczeństwa (z 7 możliwych).")
            doc.add_paragraph()

            doc.add_paragraph()
            p = doc.add_paragraph()
            run = p.add_run("Najważniejsze security headers (czego brakuje):")
            run.bold = True
            run.font.size = Pt(11)

            security_headers_info = [
                ("HSTS (Strict-Transport-Security)",
                 "Wymusza połączenia HTTPS przez określony czas. Bez niego: możliwy atak man-in-the-middle, gdzie "
                 "hacker przechwytuje dane przesyłane przez HTTP. Google od 2014 preferuje strony HTTPS w rankingu."),
                ("CSP (Content-Security-Policy)",
                 "Określa, z jakich źródeł można ładować zasoby (skrypty, obrazy, CSS). Bez niego: łatwe ataki XSS "
                 "(Cross-Site Scripting), gdzie hacker wstrzykuje złośliwy kod JavaScript na Twoją stronę."),
                ("X-Frame-Options",
                 "Zapobiega osadzeniu Twojej strony w iframe na innej witrynie. Bez niego: atak clickjacking – "
                 "użytkownik myśli, że klika w jeden przycisk, a w rzeczywistości w inny."),
                ("X-Content-Type-Options",
                 "Blokuje 'MIME sniffing' przeglądarek – zgadywanie typu pliku. Bez niego: przeglądarka może potraktować "
                 "plik tekstowy jako wykonywalny kod i uruchomić go (atak)."),
                ("Referrer-Policy",
                 "Kontroluje, ile informacji o źródle ruchu jest przekazywane innym stronom. Bez niego: pełny URL "
                 "(z parametrami, tokenami) może wyciec do zewnętrznych serwisów przez header Referer."),
                ("Permissions-Policy",
                 "Ogranicza dostęp do API przeglądarki (kamera, mikrofon, GPS, etc.). Bez niego: strony osadzone w "
                 "iframe mogą prosić o dostęp do wrażliwych zasobów użytkownika."),
            ]

            for header_name, description in security_headers_info:
                p = doc.add_paragraph(style='List Bullet')
                run = p.add_run(f"{header_name}\n")
                run.bold = True
                run.font.size = Pt(10)
                run2 = p.add_run(f"  {description}")
                run2.font.size = Pt(9)
                run2.italic = True

            doc.add_paragraph()
            p = doc.add_paragraph()
            run = p.add_run("Przykładowe endpointy z brakującymi headerami:")
            run.bold = True
            run.font.size = Pt(11)

            for item in issues['missing_security_headers'][:10]:
                url = item['url']
                headers_count = item['headers_count']
                missing = item.get('missing_critical', [])

                p = doc.add_paragraph(style='List Bullet')
                run = p.add_run(f"{url} ")
                run.font.size = Pt(9)
                run2 = p.add_run(f"({headers_count}/7 headerów)\n")
                run2.font.color.rgb = RGBColor(200, 100, 0)
                run2.font.size = Pt(9)

                if missing:
                    run3 = p.add_run(f"  Brak: {', '.join(missing[:3])}")
                    run3.font.size = Pt(9)
                    run3.italic = True

            if len(issues['missing_security_headers']) > 10:
                doc.add_paragraph(f"...oraz {len(issues['missing_security_headers']) - 10} innych stron")
            doc.add_paragraph()

        if issues['mixed_content']:
            p = doc.add_paragraph()
            total = summary['pages_analyzed']
            count = len(issues['mixed_content'])
            pct = round((count / total) * 100, 1) if total > 0 else 0
            run = p.add_run(f"🟡 Mixed content: {count}/{total} stron ({pct}%)")
            run.bold = True
            doc.add_paragraph(
                "Strony HTTPS zawierają zasoby ładowane przez HTTP (obrazki, CSS, JS). Przeglądarka pokazuje ostrzeżenie "
                "'Not Secure' mimo certyfikatu SSL. Użytkownik traci zaufanie, Google obniża ranking."
            )
            doc.add_paragraph()

            p = doc.add_paragraph()
            run = p.add_run("Strony z mixed content:")
            run.bold = True
            run.font.size = Pt(11)
            for url in issues['mixed_content'][:10]:
                doc.add_paragraph(f"• {url}", style='List Bullet')
            if len(issues['mixed_content']) > 10:
                doc.add_paragraph(f"...oraz {len(issues['mixed_content']) - 10} innych")

    if SHOW_REMEDIATIONS:
        doc.add_page_break()
        add_section_heading(doc, 'Jak wdrożyć security headers?', 2, icon='🔧')
        apache_code = '''Header always set Strict-Transport-Security "max-age=31536000; includeSubDomains"
Header always set X-Frame-Options "DENY"
Header always set X-Content-Type-Options "nosniff"
Header always set Referrer-Policy "strict-origin-when-cross-origin"
Header always set Permissions-Policy "geolocation=(), microphone=(), camera=()"'''
        q = doc.add_paragraph()
        r = q.add_run(apache_code)
        r.font.name = 'Courier New'
        r.font.size = Pt(9)
        doc.add_paragraph()
        nginx_code = '''add_header Strict-Transport-Security "max-age=31536000; includeSubDomains" always;
add_header X-Frame-Options "DENY" always;
add_header X-Content-Type-Options "nosniff" always;
add_header Referrer-Policy "strict-origin-when-cross-origin" always;
add_header Permissions-Policy "geolocation=(), microphone=(), camera=()" always;'''
        q = doc.add_paragraph()
        r = q.add_run(nginx_code)
        r.font.name = 'Courier New'
        r.font.size = Pt(9)

    doc.add_page_break()

    # =========================
    # 12. LEGENDA / SŁOWNIK
    # =========================
    add_section_heading(doc, '12. Legenda i objaśnienia – słownik dla biznesu', 1, icon='📚')

    doc.add_paragraph(
        "Poniżej znajdziesz wyjaśnienia najważniejszych terminów z raportu – językiem biznesowym, bez zbędnego żargonu IT."
    )
    doc.add_paragraph()

    # SEO
    p = doc.add_paragraph()
    run = p.add_run("SEO (Search Engine Optimization)")
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(31, 71, 136)
    doc.add_paragraph(
        "To fundamenty techniczne, które decydują, czy Google w ogóle 'widzi' Twoją stronę i rozumie, o czym ona jest. "
        "SEO to m.in. tytuły stron (tag <title>), meta opisy, nagłówki H1/H2/H3, poprawne statusy HTTP, responsywność. "
        "Bez tego Google nie wie, co indeksować i gdzie Cię pokazać."
    )
    doc.add_paragraph()

    # AEO
    p = doc.add_paragraph()
    run = p.add_run("AEO (Answer Engine Optimization)")
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(31, 71, 136)
    doc.add_paragraph(
        "To optymalizacja treści pod odpowiedzi na konkretne pytania użytkowników. "
        "Google i inne systemy (np. AI) starają się odpowiadać bezpośrednio w wynikach – bez klikania w link. "
        "AEO to m.in. sekcje FAQ, listy punktowane, tabelki, jasne definicje i konkrety."
    )
    doc.add_paragraph()

    # GEO
    p = doc.add_paragraph()
    run = p.add_run("GEO (Generative Engine Optimization)")
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(31, 71, 136)
    doc.add_paragraph(
        "To optymalizacja pod silniki generatywne (AI), takie jak ChatGPT, Gemini, Bing Chat czy SGE Google. "
        "Te systemy nie tylko wyszukują, ale też generują odpowiedzi, opierając się na wiarygodnych źródłach. "
        "GEO to m.in. E-E-A-T, dane strukturalne Schema.org, linki do źródeł, daty publikacji i autorzy."
    )
    doc.add_paragraph()

    # E-E-A-T
    p = doc.add_paragraph()
    run = p.add_run("E-E-A-T (Experience, Expertise, Authoritativeness, Trustworthiness)")
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(31, 71, 136)
    doc.add_paragraph(
        "Zestaw sygnałów jakości: doświadczenie autora, eksperckość, autorytet i zaufanie. "
        "Widoczni autorzy, daty artykułów, certyfikaty, linki do wiarygodnych źródeł, HTTPS i dane kontaktowe "
        "budują E-E-A-T i pomagają w pozycjonowaniu – szczególnie w tematach YMYL (finanse, zdrowie, prawo)."
    )
    doc.add_paragraph()

    # Schema.org
    p = doc.add_paragraph()
    run = p.add_run("Schema.org (Dane strukturalne JSON-LD)")
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(31, 71, 136)
    doc.add_paragraph(
        "Ustrukturyzowany sposób opisywania zawartości strony dla Google: artykuły, produkty, FAQ, wydarzenia itp. "
        "Pozwala uzyskać rich snippets (gwiazdki, FAQ, breadcrumbs), które zwiększają CTR i widoczność."
    )
    doc.add_paragraph()

    # NAP (Local SEO)
    p = doc.add_paragraph()
    run = p.add_run("NAP (Name, Address, Phone) – Local SEO")
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(31, 71, 136)
    doc.add_paragraph(
        "Spójne dane firmy (nazwa, adres, telefon) na stronie, w Google Maps i innych serwisach. "
        "Kluczowe dla widoczności w wynikach lokalnych i Google Local Pack."
    )
    doc.add_paragraph()

    # Mobile-first indexing
    p = doc.add_paragraph()
    run = p.add_run("Mobile-first indexing")
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(31, 71, 136)
    doc.add_paragraph(
        "Google indeksuje i ocenia w pierwszej kolejności wersję mobilną strony. "
        "Jeśli wersja mobilna jest słaba, ucierpi też pozycja wersji desktopowej."
    )
    doc.add_paragraph()

    # Security Headers
    p = doc.add_paragraph()
    run = p.add_run("Security Headers (nagłówki bezpieczeństwa HTTP)")
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(31, 71, 136)
    doc.add_paragraph(
        "Dodatkowe nagłówki HTTP (np. HSTS, CSP, X-Frame-Options), które chronią przed atakami i "
        "wymuszają bezpieczne zachowanie przeglądarki. Ich brak to realne ryzyko dla danych użytkowników."
    )
    doc.add_paragraph()

    # Canonical
    p = doc.add_paragraph()
    run = p.add_run("Canonical (tag kanoniczny)")
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(31, 71, 136)
    doc.add_paragraph(
        "Tag <link rel='canonical'> informuje Google, która wersja URL jest główna. "
        "Chroni przed duplikacją treści i rozproszeniem 'mocy' linków między podobnymi adresami."
    )
    doc.add_paragraph()

    # Thin content
    p = doc.add_paragraph()
    run = p.add_run("Thin content (cienka treść)")
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(31, 71, 136)
    doc.add_paragraph(
        "Strony z bardzo krótką, mało wartościową treścią. "
        "Google ocenia je jako low-quality, co skutkuje niższymi pozycjami i mniejszym ruchem."
    )
    doc.add_paragraph()

    # Open Graph & Twitter Cards
    p = doc.add_paragraph()
    run = p.add_run("Open Graph & Twitter Cards (meta tagi social media)")
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(31, 71, 136)
    doc.add_paragraph(
        "Tagi kontrolujące wygląd linku po udostępnieniu w mediach społecznościowych (obrazek, tytuł, opis). "
        "Bez nich tracisz kontrolę nad tym, jak Twoja marka wygląda w social media."
    )
    doc.add_paragraph()

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run("💡 Podsumowanie:")
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(50, 100, 200)
    doc.add_paragraph(
        "Każdy z opisanych elementów ma bezpośrednie przełożenie na ruch, konwersję i sprzedaż. "
        "Priorytetem są błędy krytyczne (4xx, brak Title/Description, brak viewport), następnie ostrzeżenia "
        "(canonical, schema), a na końcu optymalizacje jakościowe (E-E-A-T, content)."
    )

    doc.add_page_break()

    # =========================
    # 13. AI SUMMARY (opcjonalnie)
    # =========================
    if USE_AI_SUMMARY and OPENAI_API_KEY:
        # Zakładamy, że funkcja generate_ai_summary istnieje w innym module
        ai_summary_text = generate_ai_summary(summary, issues)  # noqa: F821
        if ai_summary_text:
            add_section_heading(doc, '13. AI-Powered Executive Summary', 1, icon='🤖')
            lines = ai_summary_text.split('\n')
            current_paragraph = None
            for line in lines:
                line = line.strip()
                if not line:
                    current_paragraph = None
                    continue
                if line.startswith('###'):
                    heading_text = line.replace('###', '').strip()
                    doc.add_heading(heading_text, level=3)
                    current_paragraph = None
                elif line.startswith('- ') or line.startswith('* '):
                    list_text = line[2:].strip()
                    doc.add_paragraph(list_text, style='List Bullet')
                    current_paragraph = None
                elif re.match(r'^\d+\.', line):
                    list_text = re.sub(r'^\d+\.\s*', '', line)
                    doc.add_paragraph(list_text, style='List Number')
                    current_paragraph = None
                else:
                    if current_paragraph is None:
                        current_paragraph = doc.add_paragraph()
                    parts = re.split(r'(\*\*.*?\*\*)', line)
                    for part in parts:
                        if part.startswith('**') and part.endswith('**'):
                            run = current_paragraph.add_run(part[2:-2])
                            run.bold = True
                        else:
                            current_paragraph.add_run(part)

    # STOPKA RAPORTU
    doc.add_paragraph('_' * 100)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('Raport wygenerowany przez Audytor SEO/AEO/GEO Enhanced Edition\n')
    run.font.size = Pt(9)
    run.italic = True
    run.font.color.rgb = RGBColor(120, 120, 120)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f'Data: {summary["generated_at"]}')
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(150, 150, 150)

    doc.save(word_path)
    print(f"✅ Raport Word zapisany: {word_path}")
