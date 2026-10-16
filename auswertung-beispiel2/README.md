# Auswertung Beispiel 2 – Gewährleistung / Reklamationen

## Excel-Dateien (`data/`)

| Datei | Inhalt |
|-------|--------|
| `01_Antraege.xlsx` | Kundenanträge aus Reklamationen |
| `02_Artikelstamm.xlsx` | ERP: EK, Warengruppe, Lieferant |
| `03_Gutschriften.xlsx` | Anerkannte Erstattungen / Ersatz 1:1 |
| `04_Umsatz.xlsx` | Umsatz & Menge je Artikel/Lieferant/Monat |

Neu erzeugen:

```powershell
python scripts/build_sample_data.py
```

Auf der Webseite: **Daten aktualisieren** lädt die vier Dateien neu.
