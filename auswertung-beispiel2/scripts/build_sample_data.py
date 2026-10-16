# -*- coding: utf-8 -*-
"""Erzeugt 4 Excel-Beispieldateien für Gewährleistungs-/Reklamations-Auswertung (Babyartikel)."""
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import date, timedelta
import random

random.seed(42)

base = Path(__file__).resolve().parents[1]
data_dir = base / "data"
data_dir.mkdir(parents=True, exist_ok=True)

PURPLE = "2E1477"
header_font = Font(bold=True, color="FFFFFF")
header_fill = PatternFill("solid", fgColor=PURPLE)
thin = Border(
    left=Side(style="thin", color="DDDDDD"),
    right=Side(style="thin", color="DDDDDD"),
    top=Side(style="thin", color="DDDDDD"),
    bottom=Side(style="thin", color="DDDDDD"),
)


def style_header(ws, headers):
    for col, h in enumerate(headers, 1):
        cell = ws.cell(1, col, h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = thin
    for col in range(1, len(headers) + 1):
        ws.column_dimensions[chr(64 + col) if col <= 26 else "A"].width = 18
    # better widths
    from openpyxl.utils import get_column_letter

    for col in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = max(14, min(28, len(headers[col - 1]) + 4))


def write_rows(ws, headers, rows):
    style_header(ws, headers)
    for r_i, row in enumerate(rows, 2):
        for c_i, h in enumerate(headers, 1):
            cell = ws.cell(r_i, c_i, row[h])
            cell.border = thin


# --- Artikelstamm (ERP) ---
# Lieferant A: viele Spielzeug-/Kuscheltier-Probleme (Demo-Story)
artikel = [
    {"Artikelnummer": "KW-110", "Artikelname": "Kinderwagen City Cruiser Pro", "Einkaufspreis_EUR": 189.0, "Warengruppe": "Kinderwagen", "Lieferant": "NordBaby GmbH"},
    {"Artikelnummer": "KW-220", "Artikelname": "Buggy Traunstein Light", "Einkaufspreis_EUR": 95.0, "Warengruppe": "Kinderwagen", "Lieferant": "AlpenKids KG"},
    {"Artikelnummer": "AS-310", "Artikelname": "Autositz BabySafe i-Size", "Einkaufspreis_EUR": 142.0, "Warengruppe": "Autositze", "Lieferant": "SüdImport AG"},
    {"Artikelnummer": "AS-320", "Artikelname": "Autositz DuoFix Reboard", "Einkaufspreis_EUR": 168.0, "Warengruppe": "Autositze", "Lieferant": "SüdImport AG"},
    {"Artikelnummer": "KT-410", "Artikelname": "Kuscheltier Hase Flecki", "Einkaufspreis_EUR": 8.5, "Warengruppe": "Spielzeug", "Lieferant": "NordBaby GmbH"},
    {"Artikelnummer": "KT-420", "Artikelname": "Kuscheltier Bär Alpen", "Einkaufspreis_EUR": 11.0, "Warengruppe": "Spielzeug", "Lieferant": "NordBaby GmbH"},
    {"Artikelnummer": "KT-430", "Artikelname": "Greifling Holzring Baby", "Einkaufspreis_EUR": 6.2, "Warengruppe": "Spielzeug", "Lieferant": "NordBaby GmbH"},
    {"Artikelnummer": "BK-510", "Artikelname": "Babybody Set 3er Bio", "Einkaufspreis_EUR": 12.4, "Warengruppe": "Babykleidung", "Lieferant": "AlpenKids KG"},
    {"Artikelnummer": "BK-520", "Artikelname": "Strampler Wolkenweich", "Einkaufspreis_EUR": 15.8, "Warengruppe": "Babykleidung", "Lieferant": "AlpenKids KG"},
    {"Artikelnummer": "BK-530", "Artikelname": "Mütze & Schal Set", "Einkaufspreis_EUR": 9.9, "Warengruppe": "Babykleidung", "Lieferant": "SüdImport AG"},
]

kunden = [
    ("Müller Anna", "DE"),
    ("Schmidt Tom", "DE"),
    ("Wagner Lisa", "AT"),
    ("Bauer Klaus", "DE"),
    ("Huber Maria", "AT"),
    ("Fischer Jan", "DE"),
    ("Weber Sophie", "CH"),
    ("Meyer Paul", "DE"),
    ("Koch Elena", "DE"),
    ("Richter Noah", "AT"),
    ("Lehmann Mia", "DE"),
    ("König Felix", "DE"),
]

# Freitext wie in Kundenmails — Klassifikation erfolgt später auf der Webseite
defekte = [
    "Hallo, das Radlager blockiert seit gestern komplett, wir kommen kaum noch vorwärts.",
    "Gurtschloss klickt nicht mehr richtig ein, Kind war schon fast raus.",
    "Am Kuscheltier ist die Naht am Ohr aufgetrennt und die Füllung kommt raus.",
    "Die Druckknöpfe am Body lösen sich nach 2× Waschen.",
    "Schiebegriff wackelt stark und knackt bei jeder Bordsteinkante.",
    "Bezug am Sitz reißt an der Naht, Stoff ist dünn geworden.",
    "Isofix-Fuß rastet nicht ein, Sitz wackelt im Auto.",
    "Farbe blutet stark aus beim Waschen, alles rosa verfärbt.",
    "Stoff am Sitzbezug scheuert durch, schon ein Loch.",
    "Auge vom Kuscheltier abgefallen – Verschluckungsgefahr!!!",
    "Vorderrad lässt sich nicht mehr schwenken, klemmt ständig.",
    "Luftreifen hat zum dritten Mal einen Platten.",
    "Hosenträgergurt lässt sich nicht mehr verkürzen.",
    "Magnethalter am Verdeck ist abgebrochen.",
    "Top-Tether Haken am Autositz klemmt.",
    "Nach dem Waschen ist der Strampler eingelaufen und hart.",
    "Plastikteil am Greifling ist angesprungen, scharfe Kante.",
    "Bremshebel am Kinderwagen greift nicht mehr zuverlässig.",
    "Kopfstütze am Autositz arretiert nicht in der Höhe.",
    "Reißverschluss am Fußsack klemmt und hat Zähne verloren.",
]

# Verkaufspreise (beantragte Erstattung) grob EK * Marge
vk = {a["Artikelnummer"]: round(a["Einkaufspreis_EUR"] * random.uniform(1.7, 2.4), 2) for a in artikel}

# --- Anträge über 2024-2026 ---
antraege = []
antrags_nr = 1000
# Bias: Spielzeug von NordBaby häufiger
weights = []
for a in artikel:
    w = 4 if a["Lieferant"] == "NordBaby GmbH" and a["Warengruppe"] == "Spielzeug" else 1.2
    if a["Warengruppe"] == "Kinderwagen":
        w = 1.5
    weights.append(w)

start = date(2024, 1, 5)
end = date(2026, 3, 28)
d = start
while d <= end:
    # ~2-5 Anträge pro Woche
    if d.weekday() < 5 and random.random() < 0.55:
        n = random.randint(1, 3)
        for _ in range(n):
            art = random.choices(artikel, weights=weights, k=1)[0]
            kunde = random.choice(kunden)
            kauf = d - timedelta(days=random.randint(20, 400))
            antrags_nr += 1
            antraege.append(
                {
                    "Kunde_Name": kunde[0],
                    "Kunde_Land": kunde[1],
                    "Antragsdatum": d.isoformat(),
                    "Antragsnummer": f"GW-{antrags_nr}",
                    "Artikelnummer": art["Artikelnummer"],
                    "Artikelname": art["Artikelname"],
                    "Kaufdatum": kauf.isoformat(),
                    "Beantragte_Kostenerstattung_EUR": vk[art["Artikelnummer"]],
                    "Defekt_Freitext": random.choice(defekte),
                }
            )
    d += timedelta(days=1)

# --- Gutschriften / Ersatz (nicht jeder Antrag anerkannt) ---
# Spielzeug NordBaby: höhere Anerkennungsquote + viele Fälle
gutschriften = []
gs_nr = 5000
for ant in antraege:
    art = next(x for x in artikel if x["Artikelnummer"] == ant["Artikelnummer"])
    p_ok = 0.55
    if art["Warengruppe"] == "Spielzeug" and art["Lieferant"] == "NordBaby GmbH":
        p_ok = 0.82
    elif art["Warengruppe"] == "Babykleidung":
        p_ok = 0.40
    if random.random() > p_ok:
        continue  # nicht anerkannt
    gs_nr += 1
    antrags_dt = date.fromisoformat(ant["Antragsdatum"])
    gs_dt = antrags_dt + timedelta(days=random.randint(3, 25))
    if gs_dt > end:
        gs_dt = end
    # Ersatz 1:1 vs Geld-Erstattung
    if art["Warengruppe"] == "Spielzeug" and random.random() < 0.45:
        art_typ = "Ersatz_1zu1"
        erstattet = 0.0
        menge = 1
        bemerkung = "Retoure + Neuausgabe gleicher Artikel, Bon Summe 0"
    else:
        art_typ = "Kostenerstattung"
        # oft EK oder Anteil VK
        erstattet = round(art["Einkaufspreis_EUR"] * random.uniform(0.9, 1.15), 2)
        menge = 1
        bemerkung = "Gutschrift an Kunden"
    gutschriften.append(
        {
            "Antragsnummer": ant["Antragsnummer"],
            "Gutschrift_Rechnung_Nr": f"GS-{gs_nr}",
            "Datum_Gutschrift": gs_dt.isoformat(),
            "Artikelnummer": ant["Artikelnummer"],
            "Artikelname": ant["Artikelname"],
            "Art": art_typ,
            "Menge": menge,
            "Erstattete_Kosten_EUR": erstattet,
            "Bemerkung": bemerkung,
        }
    )

# --- Umsatz je Artikel / Lieferant / Monat ---
umsatz = []
y, m = 2024, 1
while (y, m) <= (2026, 3):
    for art in artikel:
        # Basisumsatz, Spielzeug NordBaby relativ viel Absatz
        base_qty = random.randint(8, 35)
        if art["Warengruppe"] == "Spielzeug":
            base_qty = random.randint(25, 80)
        if art["Warengruppe"] == "Kinderwagen":
            base_qty = random.randint(3, 12)
        umsatz_eur = round(base_qty * vk[art["Artikelnummer"]], 2)
        umsatz.append(
            {
                "Jahr": y,
                "Monat": m,
                "Artikelnummer": art["Artikelnummer"],
                "Artikelname": art["Artikelname"],
                "Lieferant": art["Lieferant"],
                "Warengruppe": art["Warengruppe"],
                "Menge_verkauft": base_qty,
                "Umsatz_EUR": umsatz_eur,
            }
        )
    m += 1
    if m > 12:
        m = 1
        y += 1


def save_xlsx(filename, sheet_name, headers, rows, hinweis):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    write_rows(ws, headers, rows)
    ws2 = wb.create_sheet("Hinweis")
    ws2["A1"] = "Maurer Consulting – Beispieldaten (fiktiv)"
    ws2["A1"].font = Font(bold=True, color=PURPLE, size=13)
    ws2["A3"] = hinweis
    ws2["A4"] = "Keine echten Kundendaten. Sortiment angelehnt an typischen Babymarkt (z. B. früher ProBaby Traunstein / Wettbewerb BabyWalz, BabyOne)."
    ws2.column_dimensions["A"].width = 100
    path = data_dir / filename
    wb.save(path)
    print("Wrote", path, "rows", len(rows))


save_xlsx(
    "01_Antraege.xlsx",
    "Antraege",
    [
        "Kunde_Name",
        "Kunde_Land",
        "Antragsdatum",
        "Antragsnummer",
        "Artikelnummer",
        "Artikelname",
        "Kaufdatum",
        "Beantragte_Kostenerstattung_EUR",
        "Defekt_Freitext",
    ],
    antraege,
    "Tabelle 1: Kunden-Gewährleistungsanträge (z. B. aus Reklamations-Mails erfasst).",
)

save_xlsx(
    "02_Artikelstamm.xlsx",
    "Artikelstamm",
    ["Artikelnummer", "Artikelname", "Einkaufspreis_EUR", "Warengruppe", "Lieferant"],
    artikel,
    "Tabelle 2: Artikelstamm aus ERP inkl. Einkaufspreis, Warengruppe und Lieferant.",
)

save_xlsx(
    "03_Gutschriften.xlsx",
    "Gutschriften",
    [
        "Antragsnummer",
        "Gutschrift_Rechnung_Nr",
        "Datum_Gutschrift",
        "Artikelnummer",
        "Artikelname",
        "Art",
        "Menge",
        "Erstattete_Kosten_EUR",
        "Bemerkung",
    ],
    gutschriften,
    "Tabelle 3: Anerkannte Fälle – Kostenerstattung oder Ersatz 1:1 (Retoure/Neuausgabe, Bon oft 0).",
)

save_xlsx(
    "04_Umsatz.xlsx",
    "Umsatz",
    [
        "Jahr",
        "Monat",
        "Artikelnummer",
        "Artikelname",
        "Lieferant",
        "Warengruppe",
        "Menge_verkauft",
        "Umsatz_EUR",
    ],
    umsatz,
    "Tabelle 4: Umsatz und Verkaufsmenge je Artikel/Lieferant und Monat.",
)

# Meta für die Webseite
meta = data_dir / "daten_meta.txt"
meta.write_text(
    "Dateinamen unverändert lassen:\n"
    "01_Antraege.xlsx\n02_Artikelstamm.xlsx\n03_Gutschriften.xlsx\n04_Umsatz.xlsx\n"
    "Nach neuem Monats-Export: Dateien ersetzen, dann auf der Webseite „Daten aktualisieren“ klicken.\n",
    encoding="utf-8",
)
print("Antraege", len(antraege), "Gutschriften", len(gutschriften), "Umsatzzeilen", len(umsatz))
