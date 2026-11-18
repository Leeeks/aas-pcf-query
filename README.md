# 🌱 WAGO AAS PCF Query Tool

> **Studentenprojekt** – *„AAS & Nachhaltigkeitsdaten als Service: Neue Geschäftsmodelle für WAGO"*  
> Corporate Sustainability · WAGO Kontakttechnik GmbH & Co. KG

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python)](https://www.python.org/)
[![AAS](https://img.shields.io/badge/Standard-AAS%20Part%202%20v3-orange)](https://www.plattform-i40.de/IP/Redaktion/EN/Standardartikel/specification-administrationshell.html)
[![PCF](https://img.shields.io/badge/Submodel-CarbonFootprint-green)](https://industrialdigitaltwin.org/content-hub/aasspecifications/idta-2023-0-4_submodel-template-carbon-footprint)

---

## Kurzbeschreibung

Dieses Tool fragt den **Product Carbon Footprint (PCF)** von WAGO-Produkten automatisiert über das **WAGO AAS-Portal** (`https://c1.api.wago.com/smartservice-aas`) oder optional über einen eingebetteten **Offline-Snapshot (Q4 2026)** ab.

Es bietet drei integrierte Arbeitsmodi in einer Benutzeroberfläche im **WAGO Corporate Design**:

1. **Single Query:** Sofortige Einzelabfrage einer WAGO-Artikelnummer (Kurz-, Lang- oder Variantenformat) inkl. PCF-Wert, Einheit (`kg CO2e`), Datenquelle und Kopierfunktion für die Zwischenablage.
2. **Batch Query:** Struktur-agnostischer Import von `.xlsx`- und `.xls`-Dateien (auch per Drag-and-Drop), automatische Artikelnummern-Erkennung und Export einer neuen `.xlsx`-Arbeitsmappe mit einem standardisierten ersten Tabellenblatt **`PCF Results`** bei vollständiger Erhaltung der Originalblätter.
3. **PCF Portfolio Coverage:** Interne Analyse des gesamten WAGO-AAS-Produktportfolios mit Cursor-Paginierung, Deduplizierung, PCF-Verifikation, Abgleich fehlender AAS-Werte gegen den eingebetteten Q4-2026-Snapshot und Export eines 7-teiligen Management-Reports.

---

## Technische Architektur

### Überblick

```
1. Single Query          2. Batch Query (.xlsx / .xls + Drag & Drop)          3. PCF Portfolio Coverage
       |                                    |                                            |
       +------------------+-----------------+                                            |
                          |                                                              |
                          v                                                              v
               [  aas_pcf_query.py  ]                                       [  portfolio_coverage.py  ]
    (Auto-Detection + Normalisierung + Fast/Fallback/Offline)          (Cursor-Paginierung /shells + /submodels
                          |                                             + Q4-2026-Abgleich + 7-Sheet-Report)
                          v
       Excel Output (.xlsx): Sheet 1 = "PCF Results"
       + alle Original-Worksheets unverändert erhalten
```

### Verwendete AAS-Endpoints (SmartService AAS API v3.2)

| # | Endpoint | Methode | Rolle |
|---|----------|---------|-------|
| 1 | `https://c1.api.wago.com/smartservice-aas/submodels/{base64(ID)}` | GET | Fast Path: direkter CarbonFootprint-Submodel-Fetch |
| 2 | `https://c1.api.wago.com/smartservice-aas/shell-descriptors?assetIds={enc}` | GET | Fallback: Shell-Lookup über `globalAssetId` |
| 3 | `https://c1.api.wago.com/smartservice-aas/shells?limit=1000&cursor={cur}` | GET | Portfolio Coverage: Paginierte Shell-Enumeration |
| 4 | `https://c1.api.wago.com/smartservice-aas/submodels?idShort=CarbonFootprint&limit=500` | GET | Portfolio Coverage: Bulk-Abruf der CarbonFootprint-Submodels |

---

## Excel-Exportstruktur (`PCF Results`)

Bei der Batch-Verarbeitung bleibt die ursprüngliche Arbeitsmappe (alle Blätter, Formeln, Spalten und Formatierungen) vollständig erhalten. Als **erstes Tabellenblatt** wird **`PCF Results`** eingefügt:

1. **Informationskopf:** Titel (`WAGO Product Carbon Footprint Data`), Exportdatum/-uhrzeit, Datenquelle (`WAGO SmartService AAS API` bzw. `Embedded Snapshot (Q4 2026)`), Kennzahlen (abgefragte / verfügbare / fehlende Produkte), Link zum [WAGO AAS Portal](https://aas.wago.com/) sowie Hinweis zum Abfragemodus.
2. **Ergebnistabelle (exakt 4 Spalten):**

| Article Number (Short) | Article Number (Long) | PCF | PCF Unit |
|------------------------|-----------------------|-----|----------|
| 221-412 | 0221-0412 | 0.011 | kg CO2e |
| 280-101 | 0280-0101 | 0.0411 | kg CO2e |
| 855-9910 | 0855-9910 | Not yet available | |

> **Hinweis zu `.xls`-Dateien:** Ältere `.xls`-Dateien werden über `xlrd` eingelesen und bei der Ausgabe als `.xlsx` gespeichert. Dabei werden Zellwerte, Zahlenformate und Blattnamen übernommen; proprietäre Binär-Features älterer Excel-97-2003-Dateien (z. B. eingebettete Makros/OLE-Objekte oder komplexe bedingte Formatierungen) werden technisch bedingt nicht in `.xlsx` übertragen.

---

## Schnellstart (How-To)

### 1. Abhängigkeiten installieren

```bash
pip install -r requirements.txt
```

| Paket | Zweck |
|-------|-------|
| `pandas` | Datenverarbeitung |
| `openpyxl` | `.xlsx`-Lesen, -Schreiben und Formatierung |
| `xlrd` | Lesen von Legacy-Excel-Dateien (`.xls`) |
| `requests` / `urllib3` | HTTP-Abfragen der AAS-API inkl. Retry & Rate-Limit-Handling |
| `pillow` | Optional für WAGO-Logo in der GUI |

### 2. Tool ausführen

**Mit grafischer Oberfläche (empfohlen):**
```bash
python aas_pcf_query.py
```

**Über die Kommandozeile (CLI):**
```bash
# Einzelabfrage
python aas_pcf_query.py --single 221-412
python aas_pcf_query.py --single 0221-0412 --mode offline

# Batch-Verarbeitung (.xlsx oder .xls)
python aas_pcf_query.py input.xlsx output_pcf.xlsx --mode online

# Portfolio-Coverage-Analyse (optional mit --max-shells für Testläufe)
python aas_pcf_query.py --portfolio portfolio_report.xlsx --max-shells 200
```

---

## Projektstruktur

```
aas-pcf-query/
+-- aas_pcf_query.py             # Hauptskript: API-Client, Normalisierung, .xlsx/.xls-Import, Export, Drag&Drop, GUI
+-- offline_pcf_data.py          # Isoliertes Modul für den temporären Offline-Modus (eingebetteter Q4 2026 Snapshot)
+-- portfolio_coverage.py        # Modul für die interne WAGO AAS PCF Portfolio Coverage Analyse & 7-Sheet-Export
+-- tests/
|   +-- test_aas_pcf_query.py    # Automatisierte Pytest-Testsuite
+-- WAGO_PCF_Query.spec          # PyInstaller-Build-Spezifikation
+-- requirements.txt             # Python-Abhängigkeiten
+-- AAS_Performance_Analysis.md  # Performance-Analyse & Bottleneck-Report
+-- WAGO_LOGO.webp               # Logo für GUI
+-- README.md                    # Diese Datei
```

---

## Autor

**Alexander Flekler**  
Corporate Sustainability · WAGO Kontakttechnik GmbH & Co. KG
