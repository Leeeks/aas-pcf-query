# 🌱 WAGO AAS PCF Query Tool

> **Studentenprojekt** – *„AAS & Nachhaltigkeitsdaten als Service: Neue Geschäftsmodelle für WAGO"*  
> Corporate Sustainability · WAGO Kontakttechnik GmbH & Co. KG

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python)](https://www.python.org/)
[![AAS](https://img.shields.io/badge/Standard-AAS%20Part%202%20v3-orange)](https://www.plattform-i40.de/IP/Redaktion/EN/Standardartikel/specification-administrationshell.html)
[![PCF](https://img.shields.io/badge/Submodel-CarbonFootprint-green)](https://industrialdigitaltwin.org/content-hub/aasspecifications/idta-2023-0-4_submodel-template-carbon-footprint)

---

## Kurzbeschreibung

Dieses Tool fragt den **Product Carbon Footprint (PCF)** von WAGO-Produkten automatisiert über das **WAGO AAS-Portal** ab. Es liest eine Excel-Datei mit Artikelnummern ein, ruft für jede Nummer den PCF-Wert (kg CO2e) über die standardisierte **Asset Administration Shell (AAS) API** ab und schreibt die Ergebnisse in eine neue Excel-Datei.

Das Projekt entstand im Rahmen des **Corporate Sustainability**-Bereichs bei WAGO und dient gleichzeitig als praxisnahe Grundlage für die Untersuchung folgender Kernfrage:

> **Wie können AAS-Daten und Nachhaltigkeitskennzahlen (PCF) als Service- oder Geschäftsmodell für WAGO gedacht werden?**

---

## Studentenprojekt-Kontext

### Hintergrund

Mit der zunehmenden Regulierung im Bereich Nachhaltigkeitsberichterstattung (CSRD, CBAM, Lieferkettensorgfaltspflichtengesetz) wird der **Product Carbon Footprint** zu einem kritischen Datenpunkt in industriellen Wertschöpfungsketten. WAGO hat bereits für viele seiner Produkte PCF-Werte in der AAS hinterlegt und eine öffentliche API bereitgestellt.

### Forschungsfragen für Studierende

| # | Fragestellung | Bereich |
|---|--------------|---------|
| 1 | Welche Kunden- oder Marktgruppen haben den größten Bedarf an PCF-Daten? | Marktanalyse |
| 2 | Welche Geschäftsmodelle sind denkbar (API-as-a-Service, Data Product, Embedded Analytics)? | Business Model Design |
| 3 | Was müsste technisch ausgebaut werden, um einen skalierbaren PCF-Datendienst zu betreiben? | Technologie & API-Design |
| 4 | Wie verhält sich WAGO im Vergleich zu Wettbewerbern beim digitalen Produktpass? | Benchmarking |
| 5 | Welche regulatorischen Treiber (ESPR, CBAM, Scope-3-Reporting) erhöhen den Wert von PCF-Daten? | Regulatory Intelligence |

---

## Technische Architektur

### Überblick

```
Excel Input (.xlsx) - Struktur-agnostisch (beliebige Spaltennamen/-positionen, Kurz-/Langformat)
        |
        v
[  aas_pcf_query.py  ]  <- Hauptskript (Auto-Detection + Normalisierung + ThreadPoolExecutor)
        |
        v
    FAST PATH (1 Request, ~66 ms)
    Artikelnummer (normalisiert auf Kurzformat) -> CarbonFootprint-Submodel-ID
    GET /smartservice-aas/submodels/{base64(SubmodelID)}
        |
        +-- HTTP 200 -> PCF-Wert extrahieren (fertig)
        |
        +-- HTTP 404 -> FALLBACK (1-2 Requests, ~70 ms)
                |
                +-- GET /smartservice-aas/shell-descriptors?assetIds={base64url(AssetIdObj)}
                        |
                        +-- Submodel-ID ermitteln -> PCF-Submodel fetchen
        |
        v
    Excel Output (.xlsx) <- Article Number (Short) + Article Number (Long) + PCF + PCF Status
```

### Verwendete AAS-Endpoints (SmartService AAS API v3.2)

| # | Endpoint | Methode | Rolle |
|---|----------|---------|-------|
| 1 | `https://c1.api.wago.com/smartservice-aas/submodels/{base64(ID)}` | GET | Fast Path: direkter CarbonFootprint-Submodel-Fetch |
| 2 | `https://c1.api.wago.com/smartservice-aas/shell-descriptors?assetIds={enc}` | GET | Fallback: Shell-Lookup über `globalAssetId` |
| 3 | `https://c1.api.wago.com/smartservice-aas/submodels/{base64(SM-ID)}` | GET | Fallback: PCF-Submodel via gefundene ID |

### AAS-Submodel-URL-Konstruktion (Fast Path)

```python
# Artikelnummer (normalisiert auf Kurzformat) -> Submodel-URL
submodel_id = f"https://www.wago.com/submodel/{article_short}/CarbonFootprint"
encoded     = base64.b64encode(submodel_id.encode()).decode()
url         = f"https://c1.api.wago.com/smartservice-aas/submodels/{encoded}"
```

### PCF-Submodel-Struktur (AAS v3)

```
CarbonFootprint Submodel
+-- ProductCarbonFootprints  (SubmodelElementList)
    +-- [0]  (SubmodelElementCollection)
        +-- PcfCO2eq         <- Der gesuchte Wert (kg CO2e)
        +-- PcfLifeCyclePhase
        +-- PcfCalculationMethod
        +-- ...
```

### Performance-Ergebnisse (Benchmark mit 60 Artikeln, Aug 2026)

| Pfad | Median | Artikel-Anteil |
|------|--------|----------------|
| Fast Path (direkter Submodel-Fetch) | **66 ms** | 58 % |
| Fallback (Shell Registry-Lookup) | **15.551 ms** | 42 % |

> **Hauptbottleneck:** Der Shell Registry-Endpoint lädt ohne serverseitige Filterung **37,7 MB mit 117.418 Einträgen** in 17 Sekunden. Ein DB-Index auf `globalAssetId` könnte die Fallback-Zeit von 15+ s auf < 500 ms reduzieren. Details: [`AAS_Performance_Analysis.md`](AAS_Performance_Analysis.md)

---

## Schnellstart (How-To)

### Voraussetzungen

- Python 3.10 oder neuer
- Zugang zum WAGO-Netzwerk (VPN oder intern) für die API

### 1. Repository klonen

```bash
git clone https://github.com/<your-username>/aas-pcf-query.git
cd aas-pcf-query
```

### 2. Abhängigkeiten installieren

```bash
pip install -r requirements.txt
```

| Paket | Zweck |
|-------|-------|
| `pandas` | Excel-Lesen und -Schreiben |
| `openpyxl` | Excel-Engine für pandas |
| `requests` | HTTP-Abfragen der AAS-API |
| `urllib3` | HTTP-Verbindungsmanagement & Retry-Logik |

Optional für die GUI:
```bash
pip install pillow  # Für WAGO-Logo in der GUI
```

### 3. Excel-Datei vorbereiten

Erstelle eine `.xlsx`-Datei mit einer Spalte **`Article Number`**:

| Article Number |
|----------------|
| 280-101        |
| 283-107        |
| 221-412        |

### 4. Tool ausführen

**Mit grafischer Oberfläche (empfohlen):**
```bash
python aas_pcf_query.py
```

**Programmgesteuert als Python-Bibliothek:**
```python
from aas_pcf_query import get_pcf, get_pcf_result, main

# Einzelner Artikel - PCF-Wert (float oder None)
pcf = get_pcf("280-101")
print(f"PCF: {pcf} kg CO2e")

# Mit Status-Information
result = get_pcf_result("280-101")
# -> {"pcf": 0.0156, "status": "PCF available"}

# Batch-Verarbeitung einer Excel-Datei
summary = main("input.xlsx", "output_pcf.xlsx")
```

### 5. Ergebnis

| Article Number | Article Number (Short) | Article Number (Long) | PCF | PCF Status |
|----------------|------------------------|-----------------------|-----|------------|
| 280-101 | 280-101 | 0280-0101 | 0.0411 | PCF available |
| 0221-0412 | 221-412 | 0221-0412 | 0.0116 | PCF available |
| 855-9910 | 855-9910 | 0855-9910 | Not yet available | Not yet available |
| 999-999 | 999-999 | 0999-0999 | Not yet available | Not yet available |

---

## Projektstruktur

```
aas-pcf-query/
+-- aas_pcf_query.py             # Hauptskript: API-Client + Normalisierung + Excel-Detection + GUI + Batch
+-- offline_pcf_data.py          # Isoliertes Modul für den temporären Offline-Modus (eingebetteter Q4 2026 Snapshot)
+-- tests/
|   +-- test_aas_pcf_query.py    # Automatisierte Pytest-Testsuite
+-- requirements.txt             # Python-Abhängigkeiten
+-- AAS_Performance_Analysis.md  # Performance-Analyse & Bottleneck-Report
+-- input.xlsx                   # Beispiel-Eingabedatei (Artikelnummern)
+-- WAGO_LOGO.webp               # Logo für GUI
+-- README.md                    # Diese Datei
```

---

## Konfiguration

```python
# In aas_pcf_query.py:
AAS_BASE_URL           = "https://c1.api.wago.com/smartservice-aas"
SUBMODEL_ID_TEMPLATE   = "https://www.wago.com/submodel/{article}/CarbonFootprint"
CERT_VERIFY            = False              # TLS (Self-Signed Cert)
REQUEST_TIMEOUT        = (5, 10)           # (connect, read) in Sekunden
WORKERS                = 5                 # Parallele Worker-Threads
SLOW_REQUEST_THRESHOLD = 3.0               # Logging-Schwelle in Sekunden
```

---

## Status-Codes (Kunden-Export)

| Status | Bedeutung |
|--------|-----------|
| `PCF available` | PCF-Wert erfolgreich abgerufen |
| `Not yet available` | Für dieses Produkt liegt aktuell noch kein PCF-Wert vor |
| `Query error` | Technischer Abfragefehler (Timeout, Authentifizierungs- oder Verbindungsfehler) |

---

## Business-Kontext: AAS als Service für WAGO

### Regulatorische Treiber (Demand-Side)

| Regulierung | Relevanz | Zeitlinie |
|------------|---------|-----------|
| **CSRD** | Scope-3-Emissionen erfordern PCF-Daten von Lieferanten | Ab 2024 (gestaffelt) |
| **CBAM** | CO2-Intensität von Produkten muss nachgewiesen werden | Ab 2026 |
| **ESPR / Digital Product Passport** | Maschinenlesbare Produktdaten inkl. PCF verpflichtend | Ab 2026 (gestaffelt) |
| **ISO 14067** | PCF-Berechnung und -Kommunikation | Best Practice |

### Mögliche Geschäftsmodell-Dimensionen

```
            AAS / PCF als Service
                    |
        +-----------+-----------+
        |           |           |
  Data Product   API-as-a-   Embedded
                 Service     Sustainability

  PCF-Daten als  Kunden nutzen  PCF direkt in
  lizenziertes   WAGO-API für   Kunden-ERP/SCM
  Datenprodukt   Scope-3-       (Plugin/
  (CSV, API)     Reporting      Integration)
```

---

## Weiterführende Dokumentation

- **[`AAS_Performance_Analysis.md`](AAS_Performance_Analysis.md)** – Detaillierter Performance-Benchmark-Report
- [IDTA Submodel Template – Carbon Footprint](https://industrialdigitaltwin.org/content-hub/aasspecifications/idta-2023-0-4_submodel-template-carbon-footprint)
- [AAS Spezifikation Part 2 – Plattform Industrie 4.0](https://www.plattform-i40.de/IP/Redaktion/EN/Standardartikel/specification-administrationshell.html)
- [ESPR / Digital Product Passport – EU Kommission](https://ec.europa.eu/environment/ecodesign/)

---

## Autor

**Alexander Flekler**  
Corporate Sustainability · WAGO Kontakttechnik GmbH & Co. KG

---

*Dieses Projekt steht für interne WAGO-Zwecke und Bildungszwecke zur Verfügung.*
