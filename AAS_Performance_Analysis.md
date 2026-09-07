# AAS Performance Analysis – WAGO AAS Portal
**Erstellt:** 2026-08-28  
**Basis:** Realer Benchmark mit 60 Artikeln aus `input.xlsx`  
**Benchmark-Tool:** `aas_perf_benchmark.py` (liest nur, ändert nichts am produktiven Tool)

---

## 1. Executive Summary

Das WAGO AAS PCF Query Tool funktioniert korrekt und ist clientseitig gut optimiert.  
Die serverseitige AAS-Infrastruktur zeigt jedoch eine **extrem zweigeteilte Performance**:

- **Fast Path (Submodel-Fetch direkt):** Median **66 ms** – sehr gut  
- **Fallback (Shell Registry-Lookup):** Median **15.551 ms** – kritisch langsam  

Das zentrale Problem ist nicht die Netzwerklatenz, sondern der **Shell Registry-Endpoint**, der für Artikel ohne direkt konstruierbare Submodel-ID in einem einzigen Request **37,7 MB JSON mit 117.418 Shell-Descriptors** zurückliefert und dabei **17 Sekunden** benötigt.  

41 % der abgefragten Artikel (25 von 60) benötigen diesen Fallback-Pfad und erzeugen damit inakzeptable Antwortzeiten von **12–23 Sekunden** pro Artikel.

> **Most Important Message:** Der Shell Registry-Endpoint ohne Filter lädt die komplette Registry mit 117.418 Einträgen in einer einzigen Response – ein serverseitiger Index auf `globalAssetId` und der korrekte Einsatz des `assetIds`-Filterparameters können die Fallback-Zeit von 15+ Sekunden auf unter 500 ms reduzieren.

---

## 2. Current Request Flow

### Schritt-für-Schritt vom Artikel zur PCF

```
Artikelnummer: 280-101
        │
        ▼  FAST PATH (1 Request, ~66 ms)
        │
        │  Construct deterministic ID (client-side, no network):
        │    Submodel-ID: https://wago.com/ids/submodel/280-101/PCF
        │    Base64-URL:  aHR0cHM6Ly93YWdvLmNvbS9pZHMvc3VibW9kZWwvMjgwLTEwMS9QQ0Y=
        │
        ▼
GET https://c1.api.wago.com/smartdata-aas-env/submodels/{base64(SubmodelID)}
        │
        ├── HTTP 200 → JSON parsen → PcfCO2eq extrahieren → FERTIG ✓  (~66 ms)
        │
        └── HTTP 404 → FALLBACK nötig
                │
                ▼  FALLBACK STEP 1 (1 Request, ~15.500 ms)
                │
                │  Construct Asset-ID (client-side):
                │    Asset-ID:  https://wago.com/ids/asset/280-101
                │    Wrapped:   [{"name":"globalAssetId","value":"https://..."}]
                │    Base64+URL-encoded
                │
                ▼
GET https://c1.api.wago.com/smartdata-aas-registry/shell-descriptors
    ?assetIds={base64([{"name":"globalAssetId","value":"https://wago.com/ids/asset/ARTICLE"}])}
                │
                │  Response: Shell-Descriptor mit submodelDescriptors[]
                │
                ├── Shell gefunden → submodelDescriptors nach "pcf"/"carbonfootprint" durchsuchen
                │
                ▼  FALLBACK STEP 2 (1 Request, wenn PCF-SM-Descriptor gefunden)
                │
GET https://c1.api.wago.com/smartdata-aas-env/submodels/{base64(PCF-Submodel-ID)}
                │
                └── JSON → PcfCO2eq extrahieren → FERTIG
```

### Verwendete Endpoints im Detail

| # | Endpoint | Method | Parameter | Rolle |
|---|----------|--------|-----------|-------|
| 1 | `https://c1.api.wago.com/smartdata-aas-env/submodels/{base64(ID)}` | GET | – | Fast Path: direkter PCF-Submodel-Fetch |
| 2 | `https://c1.api.wago.com/smartdata-aas-registry/shell-descriptors?assetIds={enc}` | GET | `assetIds` (base64+URL) | Fallback Step 1: Shell-Lookup über Asset-ID |
| 3 | `https://c1.api.wago.com/smartdata-aas-env/submodels/{base64(SM-ID)}` | GET | – | Fallback Step 2: PCF-Submodel über gefundene ID |

**Anzahl Requests pro Artikel:**
- Fast Path erfolgreich: **1 Request**
- Fallback benötigt, kein PCF: **2 Requests** (Fast Path + Registry)
- Fallback erfolgreich mit PCF: **3 Requests** (Fast Path + Registry + Submodel)

---

## 3. Performance Measurements

### Benchmark-Konfiguration
- **Datum/Zeit:** 2026-08-28, 09:29 Uhr
- **Artikel:** 60 unique Artikelnummern aus `input.xlsx`
- **Messansatz:** Jeder HTTP-Request einzeln gemessen (ohne Retry-Logic, um echte Latenzen zu sehen)
- **Python-Umgebung:** `openlca` (Miniconda3), Python 3.12.9, requests 2.32.3
- **TLS-Verify:** deaktiviert (WAGO Self-Signed Cert)

### Registry Endpoint Exploration

| Endpoint | Status | Zeit | Response-Größe | Ergebnis |
|----------|--------|------|----------------|----------|
| `GET /shell-descriptors` (kein Filter) | 200 | **17.354 ms** | **37,7 MB** | **117.418 Shell-Descriptors** |
| `GET /shell-descriptors?limit=1` | 200 | 74 ms | 416 Bytes | 1 Descriptor + cursor |
| `GET /shell-descriptors?limit=5` | 200 | 48 ms | 1.616 Bytes | 5 Descriptors + cursor |
| `GET /submodel-descriptors` (kein Filter) | **404** | 30 ms | 54 Bytes | Endpoint nicht verfügbar |
| `GET /submodel-descriptors?limit=1` | **404** | 25 ms | 54 Bytes | Endpoint nicht verfügbar |

> [!IMPORTANT]
> Der Registry-Endpoint ohne Filter liefert **37,7 MB** mit **117.418 Einträgen** in **17,4 Sekunden**. Auffällig: Mit `?limit=1` nur **74 ms** – die Pagination-Infrastruktur ist vorhanden und funktioniert.

### Pagination-Metadaten aus Registry (beobachtet)

Der `limit`-Parameter funktioniert korrekt und liefert einen `cursor` für die nächste Seite:
```json
{
  "paging_metadata": {
    "cursor": "aHR0cHM6Ly93YWdvLmNvbS9pZHMvYWFzLzIwMDAtMTE1"
  },
  "result": [...]
}
```

Der Shell-Descriptor enthält folgende Keys: `endpoints`, `idShort`, `id`  
*(Hinweis: Kein `submodelDescriptors` oder `specificAssetIds` im gefilterten Top-Level-Result sichtbar – Needs IT verification, ob diese bei gefiltertem Request vollständig enthalten sind)*

### Per-Artikel Messwerte (vollständig)

| Artikel | Fast Path (ms) | Reg (ms) | Total (ms) | Status |
|---------|---------------:|----------:|-----------:|--------|
| **2000-115** | **8.240** | 0 | **8.240** | PCF available (fast path) |
| 2000-1201 | 91 | 0 | 92 | PCF available (fast path) |
| 221-412 | 74 | 0 | 74 | PCF available (fast path) |
| 100-280 | 58 | 17.430 | 17.700 | AAS available – PCF not available |
| 999-999 | 75 | 16.385 | 16.665 | AAS available – PCF not available |
| 280-101 | 72 | 0 | 72 | PCF available (fast path) |
| 280-312 | 86 | 0 | 86 | PCF available (fast path) |
| 280-402 | 81 | 0 | 82 | PCF available (fast path) |
| 280-637 | 74 | 0 | 74 | PCF available (fast path) |
| 280-641 | 61 | 0 | 61 | PCF available (fast path) |
| 283-101 | 66 | 0 | 66 | PCF available (fast path) |
| 283-107 | 61 | 0 | 61 | PCF available (fast path) |
| 283-301 | 72 | 0 | 72 | PCF available (fast path) |
| 283-316 | 65 | 0 | 65 | PCF available (fast path) |
| 283-402 | 65 | 0 | 65 | PCF available (fast path) |
| 285-1175 | 61 | 0 | 61 | PCF available (fast path) |
| 285-1177 | 67 | 0 | 68 | PCF available (fast path) |
| 285-1184 | 66 | 0 | 66 | PCF available (fast path) |
| 285-1185 | 69 | 0 | 69 | PCF available (fast path) |
| 285-135 | 62 | 0 | 62 | PCF available (fast path) |
| 285-137 | 77 | 0 | 77 | PCF available (fast path) |
| 285-170 | 62 | 0 | 62 | PCF available (fast path) |
| 285-195 | 100 | 0 | 100 | PCF available (fast path) |
| 285-407 | 59 | 0 | 59 | PCF available (fast path) |
| 285-435 | 71 | 0 | 71 | PCF available (fast path) |
| 285-495 | 60 | 0 | 60 | PCF available (fast path) |
| 285-635 | 67 | 0 | 67 | PCF available (fast path) |
| 285-637 | 62 | 0 | 62 | PCF available (fast path) |
| 249-197 | 72 | 0 | 73 | PCF available (fast path) |
| 231-004/046-000 | 59 | 13.813 | 14.073 | AAS available – PCF not available |
| 231-004/048-000 | 64 | 19.170 | 19.454 | AAS available – PCF not available |
| 231-004/049-000 | 76 | 11.560 | 11.858 | AAS available – PCF not available |
| 231-004/052-000 | 54 | 15.001 | 15.249 | AAS available – PCF not available |
| 231-605/050-000 | 61 | 15.551 | 15.806 | AAS available – PCF not available |
| 231-624/050-000 | 66 | 15.351 | 15.641 | AAS available – PCF not available |
| 734-001/K010-2937/000-100 | 55 | 14.548 | 14.823 | AAS available – PCF not available |
| 734-001/K010-2937/000-300 | 63 | 15.139 | 15.420 | AAS available – PCF not available |
| 734-001/K010-2937/000-400 | 267 | 20.964 | 21.436 | AAS available – PCF not available |
| 8001-001/K010-2937/000-2800 | 66 | 14.073 | 14.361 | AAS available – PCF not available |
| 8001-099/Z000-1809 | 63 | 15.301 | 15.586 | AAS available – PCF not available |
| 289-593 | 69 | 22.802 | 23.084 | AAS available – PCF not available |
| 787-1212 | 58 | 18.106 | 18.380 | AAS available – PCF not available |
| 221-2411 | 79 | 0 | 79 | PCF available (fast path) |
| 221-415 | 66 | 0 | 66 | PCF available (fast path) |
| 221-612 | 67 | 0 | 68 | PCF available (fast path) |
| 222-413 | 63 | 0 | 63 | PCF available (fast path) |
| 222-415 | 70 | 0 | 70 | PCF available (fast path) |
| 812-100 | 65 | 0 | 65 | PCF available (fast path) |
| 206-1204 | 60 | 18.640 | 18.920 | AAS available – PCF not available |
| 2002-6307 | 74 | 0 | 74 | PCF available (fast path) |
| 2002-6392 | 67 | 0 | 67 | PCF available (fast path) |
| 852-1106 | 62 | 14.470 | 14.735 | AAS available – PCF not available |
| 852-1111/000-001 | 56 | 18.105 | 18.350 | AAS available – PCF not available |
| 855-3001/060-003 | 59 | 17.566 | 17.824 | AAS available – PCF not available |
| 855-301/150-501 | 57 | 18.433 | 18.699 | AAS available – PCF not available |
| 855-305/150-501 | 54 | 15.187 | 15.446 | AAS available – PCF not available |
| 855-4001/100-001 | 64 | 18.690 | 18.984 | AAS available – PCF not available |
| 855-4001/150-001 | 62 | 19.402 | 19.658 | AAS available – PCF not available |
| 855-405/400-501 | 75 | 14.539 | 14.808 | AAS available – PCF not available |
| 855-9910 | 65 | 15.308 | 15.586 | AAS available – PCF not available |

### Status-Verteilung

| Status | Anzahl | Anteil |
|--------|-------:|-------:|
| PCF available (fast path) | 35 | 58 % |
| AAS available – PCF not available | 25 | 42 % |
| Article not found | 0 | 0 % |
| Query error | 0 | 0 % |

*(Hinweis: `999-999` und `100-280` wurden als „AAS available – PCF not available" klassifiziert, weil die Registry für diese IDs eine Shell zurückgab – möglicherweise falsch-positive Registrierungen. Needs IT verification.)*

---

## 4. Identified Bottlenecks

### Statistik-Tabelle

| API Operation | Min | Median | Avg | P95 | Max |
|---------------|----:|-------:|----:|----:|----:|
| Fast Path – Submodel Fetch | 54 ms | **66 ms** | 206 ms | 100 ms | **8.240 ms** |
| Fallback – Shell Registry Lookup | 11.560 ms | **15.551 ms** | 16.621 ms | 20.964 ms | **22.802 ms** |
| Fallback – Submodel Fetch (via Registry) | — | — | — | — | — |
| **Total per Article** | 59 ms | 80 ms | 7.219 ms | 19.658 ms | **23.084 ms** |

### Bottleneck #1 (Kritisch): Shell Registry ohne serverseitige Filterung

**Observed:**  
`GET /shell-descriptors` (ohne sinnvolle Filterung) = **37,7 MB, 117.418 Einträge, 17.354 ms**

Das Tool sendet zwar `assetIds={base64(...)}`, aber die Messung zeigt, dass alle 25 Fallback-Requests **11.560–22.802 ms** benötigten. Zum Vergleich: `GET /shell-descriptors?limit=1` benötigte nur **74 ms**.

**Likely:**  
Der `assetIds`-Filter wird serverseitig nicht oder ineffizient ausgewertet – möglicherweise lädt der Server alle Descriptors, übergibt sie an eine Filterlogik und gibt dann das Ergebnis zurück. Eine serverseitige Datenbankabfrage mit Index würde deutlich schneller sein.

**Needs IT verification:**  
Ausführungsplan der DB-Abfrage prüfen, die hinter `assetIds`-Filter liegt. Ist `globalAssetId` oder das entsprechende JSON-Feld indexiert?

### Bottleneck #2 (Beobachtet): Einzelner Outlier – 2000-115, 8.240 ms Fast Path

**Observed:**  
Artikel `2000-115` benötigte im Fast Path **8.240 ms** (alle anderen: 54–100 ms). Dieser Artikel war der erste in der Benchmark-Sequenz.

**Likely:**  
Server-seitiges Cold-Start-Verhalten: Der erste Request "wärmt" einen Cache oder eine Verbindung auf. Dies ist konsistent mit den im produktiven Tool dokumentierten Beobachtungen (Performance-Hinweise im Code: „ReadTimeout-dann-Retry-Artikel: zweiter Retry sofort erfolgreich – Server-seitiges Cache-Warmup").

**Needs IT verification:**  
Prüfen, ob Application-Level-Cache oder Connection-Pool nach Idle-Timeout invalidiert wird. Gibt es einen Warmup-Mechanismus beim Start?

### Bottleneck #3 (Strukturell): 2 statt 1 Registry-Roundtrip bei Fallback

**Observed:**  
Für 25 von 60 Artikeln (42 %) sind **mindestens 2 Requests** nötig. Der erste Fast-Path-Request (54–76 ms) ist schnell, aber der darauf folgende Registry-Request (11–23 Sekunden) dominiert die Gesamtzeit vollständig.

**Likely:**  
Die Architektur ist AAS-konform (Shell → Submodel Registry → Submodel), aber für den spezifischen Use Case „Artikelnummer → PCF" ist der Registry-Hop unnötig teuer.

---

## 5. Client Workarounds vs. Server Solutions

| Problem | Aktueller Client-Workaround | Bessere serverseitige Lösung |
|---------|----------------------------|------------------------------|
| Langsamer Registry-Lookup | **Fast Path**: Submodel-ID deterministisch aus Artikelnummer konstruieren → Registry komplett umgehen | Serverseitiger Index auf `globalAssetId` / `specificAssetIds.value`; gefilterter Registry-Lookup in <500 ms |
| Variante Artikel-IDs / nicht konstruierbare IDs | **Fallback Path**: Registry-Lookup nach `assetIds` | Direkt-Lookup-Endpoint `GET /aas/by-article-number/{article}` |
| Timeout bei langsamen Requests | REQUEST_TIMEOUT (5, 10): schnell abbrechen, Retry | Server-seitige SLA: Registry-Antwort < 1 s |
| Mehrfachabfragen gleicher Artikel | Client-seitiger In-Memory-Cache (Session-Dauer) | Server-seitiger Response-Cache / API-Gateway-Cache (TTL z. B. 1 h) |
| Parallele Artikel-Abfragen | 5 parallele Worker-Threads | Bulk Query API: `POST /pcf/query` mit `{articleNumbers: [...]}` |
| Cold-Start-Outlier (8 s) | SLOW_REQUEST_THRESHOLD + Logging | Server-seitiger Connection-Pool-Warmup, ggf. Health-Check-Ping |
| Große Registry-Response (37,7 MB) | Wird nicht geladen (Fast Path vermeidet dies) | Pagination konfiguriert, serverseitiger Filter verpflichtend |
| Variante Submodel-Semantiken | `_extract_all_values()` traversiert rekursiv | Standardisiertes PCF-Submodel-Schema serverseitig valide halten |

---

## 6. Server-side Improvement Opportunities

### 6.1 Artikel-Nummer-Lookup: Wie läuft es aktuell?

**Beobachtet im Code:**

Der Fast Path konstruiert die Submodel-ID deterministisch:
```python
submodel_id = f"https://wago.com/ids/submodel/{article_number}/PCF"
encoded = base64.b64encode(submodel_id.encode()).decode()
url = f"{AAS_ENV_BASE}/submodels/{encoded}"
```

Das bedeutet: **Der Client sucht nicht** – er rät die ID direkt auf Basis einer bekannten Namenskonvention. Das ist sehr effizient, setzt aber voraus, dass WAGO diese Konvention konsequent einhält.

Der Fallback sucht über:
```python
asset_id = f"https://wago.com/ids/asset/{article_number}"
# → GET /shell-descriptors?assetIds={base64([{"name":"globalAssetId","value":asset_id}])}
```

**Beobachtet:** Dieser Filter benötigt trotz Filterparameter 11–23 Sekunden.

**Likely:** Der Filter wird serverseitig nicht als indexierte DB-Abfrage ausgewertet, sondern entweder:
a) Alle 117.418 Shells werden geladen und dann gefiltert, oder  
b) Die Datenbankabfrage hat keinen Index auf dem gefilterten Feld

**Needs IT verification:** DB-Queryplan für `assetIds`-Filter prüfen.

### 6.2 Datenbank-Indizierung

Folgende Felder sind für die Lookup-Performance besonders relevant:

| Feld | Verwendung im Client | Vermutlich indexiert? | Empfehlung |
|------|---------------------|----------------------|------------|
| `globalAssetId` | Fallback: `?assetIds` mit `name=globalAssetId` | **Unwahrscheinlich** (22 s Lookup) | **Index sofort prüfen** |
| `id` (Shell-ID) | Fallback: Submodel-Fetch via direkter ID | Wahrscheinlich (schneller direkter Fetch) | Verifikation |
| `id` (Submodel-ID) | Fast Path: direkter Fetch | Wahrscheinlich (54–100 ms) | – |
| `specificAssetIds.value` | Nicht verwendet, aber AAS-Standard | Unklar | Prüfen als Alternative zu `globalAssetId` |
| `semanticId` | Nicht für Lookup verwendet | Unklar | Mittelfristig relevant für Submodel-Filter |

**Konkrete IT-Prüfpunkte:**

> Prüfen, ob das Feld `globalAssetId` in der Datenbank der Shell Registry indexiert ist. Der `?assetIds`-Filter, der dieses Feld verwendet, benötigt aktuell 11–23 Sekunden für eine gefilterte Anfrage auf 117.418 Einträgen. Ein Index auf `globalAssetId` (oder dem entsprechenden serialisierten JSON-Feld) sollte diesen Lookup auf < 50 ms reduzieren.

> Falls die Shell Registry eine dokumentenbasierte Datenbank verwendet (z. B. MongoDB), prüfen, ob ein Index auf `specificAssetIds.value` oder `globalAssetId` existiert. Bei einer relationalen Datenbank: Prüfen, ob der JSON-Extract-Operator (z. B. `JSON_VALUE()` in SQL Server, `->>`  in PostgreSQL) indexiert ist.

### 6.3 Serverseitiges Filtering

**Beobachtet:**  
`GET /shell-descriptors` ohne Filter → 37,7 MB, 117.418 Einträge, 17,4 Sekunden  
`GET /shell-descriptors?limit=1` → 416 Bytes, 74 ms  
`GET /shell-descriptors?assetIds={...}` → vermutlich ~1 Ergebnis, aber 11–23 Sekunden

**Schlussfolgerung (Observed):**  
Die Pagination-Infrastruktur (`limit`, `cursor`) funktioniert und ist schnell. Der `assetIds`-Filter ist vorhanden, aber serverseitig nicht performant implementiert.

**Likely:**  
Der Filter greift nach dem Laden der Daten statt vor dem Laden (Filtering nach DB-Read statt als DB-WHERE-Clause).

**Empfehlung:**  
IT prüfen, ob `assetIds`-Filter als parametrisierte DB-Abfrage implementiert ist oder als Post-Processing auf dem Ergebnis eines vollen Table-Scans.

### 6.4 Pagination-Analyse

**Beobachtet:**
- `GET /shell-descriptors` (kein Filter): 117.418 Einträge in einer Response → **keine Pagination erzwungen**
- `GET /shell-descriptors?limit=N`: Pagination funktioniert korrekt mit `cursor`
- `GET /submodel-descriptors`: HTTP 404 – Endpoint nicht verfügbar

**Kritischer Befund:**  
Ein negativer Lookup (Artikel nicht im Portal vorhanden) triggert derzeit den Fallback, der alle 117.418 Shells auf einmal holen könnte, falls der `assetIds`-Filter nicht funktioniert. In diesem Fall:

```
Article not found = voller Registry-Download = 37 MB = 17 Sekunden
```

**Empfehlung:**  
Serverseitig eine maximale Default-Page-Size erzwingen (z. B. `limit=100`). Requests ohne `limit` sollten nicht mehr als N Einträge zurückgeben.

### 6.5 Response Payload-Analyse

**Beobachtet:**
- Volle Registry ohne Filter: **37.708.143 Bytes (37,7 MB)** für 117.418 Shells
- Durchschnittlich: **~321 Bytes pro Shell-Descriptor**
- Mit `limit=1`: **416 Bytes** (ein Descriptor + Metadaten)

**Beobachtet (Shell-Descriptor-Struktur):**  
Enthält Keys: `endpoints`, `idShort`, `id` – **keine `specificAssetIds` im Top-Level-Response sichtbar**

> [!WARNING]
> Needs IT verification: Sind `specificAssetIds` oder `submodelDescriptors` im gefilterten Shell-Descriptor enthalten? Falls nicht, fehlen die Daten, die der Client zum Weitersuchen benötigt.

**Fast Path Submodel-Payloads (exemplarisch, soweit messbar):**  
Basierend auf typischen PCF-Submodels: Ein PCF-Submodel im WAGO AAS enthält die `ProductCarbonFootprints`-Liste mit mehreren Einträgen. Der tatsächlich benötigte Wert ist nur `PcfCO2eq` (~10 Bytes). Die übertragene Response-Größe ist deutlich größer.

---

## 7. Prioritized Recommendations

| Priorität | Maßnahme | Problem | Evidenz | Erwarteter Impact | Aufwand |
|----------:|----------|---------|---------|-------------------|---------|
| **1** | **DB-Index auf `globalAssetId` in Shell Registry prüfen/anlegen** | Fallback benötigt 11–23 s | 17.354 ms für ungefilterten Dump; assetIds-Filter nicht schnell | **Sehr hoch** – 25/60 Artikel betroffen | **Niedrig** (DBA-Task) |
| **2** | **Maximale Default-Page-Size erzwingen** | Kein `limit` → 37,7 MB in einer Response | Gemessen: 37.708.143 Bytes ohne limit | **Hoch** – verhindert worst case bei Skalierung | **Niedrig** (Config) |
| **3** | **Warm-up / Connection-Pool prüfen** | Cold-Start-Outlier: 2000-115 = 8.240 ms (1. Request) | Erster Request 100x langsamer als alle folgenden | **Mittel** – nur für ersten Request | **Niedrig** (Config) |
| **4** | **Server-seitiges Caching für Shell Registry-Lookups** | Wiederholte Abfragen gleicher Artikel erzeugen volle DB-Last | 117.418 Einträge wachsen mit WAGO-Portfolio | **Hoch** – ohne Code-Änderung am Client | **Mittel** |
| **5** | **Bulk Query API: POST /pcf/query** | 500 Artikel × 2 Requests = 1000+ HTTP-Calls | Tool sendet sequentiell pro Artikel | **Sehr hoch** bei >50 Artikeln | **Mittel–Hoch** |
| **6** | **Direkt-Lookup: GET /shell-descriptors?assetIds effizient machen** | Filterparameter vorhanden, aber 22 s statt < 500 ms | Empirisch gemessen | **Sehr hoch** | **Mittel** (Index + Query-Optimierung) |
| **7** | **Response Compression (gzip) prüfen** | 37,7 MB unkomprimiert | Nicht gemessen (TLS-Overhead; Content-Encoding prüfen) | **Mittel** bei großen Responses | **Niedrig** (Gateway-Config) |
| **8** | **Submodel-Descriptors in Shell-Descriptor einbetten** | Derzeit: Shell-Registry → separater Submodel-Registry-Hop | Fallback benötigt 2–3 Requests | **Mittel** – 1 Request gespart | **Mittel** |
| **9** | **Distributed Tracing / Observability** | Langsame Requests nicht sichtbar ohne Client-Messung | Heute: IT sieht Probleme erst durch Client-Beschwerden | **Hoch** (Diagnose) | **Mittel** |
| **10** | **PCF Read Model / Index** | Langfristig: PCF direkt ohne AAS-Hop | Aktuell 1–3 Requests pro Artikel | **Sehr hoch** bei Skalierung | **Hoch** |

---

## 8. Quick Wins (A – ohne große Architekturänderung)

### A1 – Datenbank-Index auf `globalAssetId` prüfen
**Maßnahme:** DBA prüft, ob in der Shell Registry-Datenbank ein Index auf `globalAssetId` (oder `specificAssetIds`) existiert. Falls nicht: Index anlegen.  
**Erwartung:** Fallback-Zeit von 11–23 s auf < 500 ms  
**Effort:** 1 DBA-Tag

### A2 – Default-Page-Size konfigurieren
**Maßnahme:** Registry-API-Konfiguration: Maximale Response ohne `limit`-Parameter auf z. B. 100 Einträge begrenzen. Requests ohne `limit` erhalten einen `cursor` für weitere Seiten.  
**Erwartung:** Verhindert 37-MB-Responses; schützt bei wachsender Registry  
**Effort:** < 1 Tag

### A3 – Connection Pool / Warmup prüfen
**Maßnahme:** Prüfen, ob der Backend-Service nach Idle-Zeit Verbindungen schließt. Bei Bedarf: Warmup-Request oder Keep-Alive konfigurieren.  
**Erwartung:** Cold-Start-Outlier (8 s bei `2000-115`) eliminieren  
**Effort:** < 1 Tag

### A4 – Response Compression aktivieren
**Maßnahme:** Prüfen, ob `Content-Encoding: gzip` aktiviert ist. 37 MB komprimiert auf ca. 3–5 MB → deutlich schnellerer Transfer.  
**Effort:** Gateway-Konfiguration, < 1 Tag

### A5 – Logging/Monitoring ergänzen
**Maßnahme:** Request-Duration, DB-Query-Duration und HTTP-Status in bestehende Logs aufnehmen. Slow-Query-Log (> 1 s) aktivieren.  
**Effort:** 1–2 Tage

---

## 9. Medium-term Improvements (B)

### B1 – Optimierter `/shell-descriptors?assetIds` Endpoint
Der `assetIds`-Filterparameter ist vorhanden, aber nicht performant. Ziel: gefilterte Abfrage in < 500 ms.  
**Maßnahmen:**  
- DB-Query für `assetIds`-Filter als parametrisierte Abfrage mit Index  
- Ergebnis: nur die eine Shell, die zu dieser Asset-ID gehört  
- Kein volles Table-Scan mehr  

### B2 – Server-seitiger Cache für Registry-Lookups
Registry-Daten ändern sich selten (AAS werden bei Produktanlage/-änderung registriert).  
**Mögliche Ebenen:**
- **Application-Level-Cache:** `globalAssetId → Shell-Descriptor` mit TTL = 1 h (bei AAS-Update: Cache-Invalidierung via Event)
- **API-Gateway-Cache:** Gecachte Responses für `?assetIds=...`-Requests (TTL = 30 min)
- **Was cachen:** `globalAssetId → {shellId, pcfSubmodelId}` – minimale Information, schnell aktualisierbar
- **Was NICHT cachen:** PCF-Wert selbst (ändert sich häufiger als Struktur)

### B3 – Submodel-Descriptors in Shell-Descriptor einbetten
**Aktuell:** Shell-Descriptor enthält `endpoints`, `idShort`, `id` – kein direkter Link zum PCF-Submodel-Descriptor.  
**Verbesserung:** Shell-Descriptor enthält vollständig befüllte `submodelDescriptors` inkl. PCF-Submodel-ID → Client kann direkt Submodel fetchen, ohne zweiten Registry-Lookup.  
*(Hinweis: AAS-Standard erlaubt `submodelDescriptors` im Shell-Descriptor – Needs IT verification, ob dies aktuell der Fall ist und warum es im Benchmark-Output nicht sichtbar war)*

### B4 – Bulk Query API
Für den Use Case „500 Artikelnummern → PCF-Werte":

```http
POST https://c1.api.wago.com/smartdata-aas-env/pcf/query
Content-Type: application/json

{
  "articleNumbers": ["221-412", "280-101", "855-9910"]
}
```

Response:
```json
[
  {"articleNumber": "221-412", "status": "PCF available", "pcf": 0.123},
  {"articleNumber": "280-101", "status": "PCF available", "pcf": 0.456},
  {"articleNumber": "855-9910", "status": "AAS available - PCF not available"}
]
```

**Vorteile:**  
- 1 HTTP-Request statt 500–1500  
- Server kann intern parallel suchen  
- Effizientere DB-Queries (IN-Clause statt 500× Einzelquery)  
- Weniger Gateway-Overhead  

**Nachteile:**  
- Nicht AAS-standardkonform (wäre WAGO-spezifischer Endpoint)  
- Muss gepflegt werden  

*(Unterscheidung: AAS-standardkonforme Verbesserung wäre die Optimierung des bestehenden `assetIds`-Filters; Bulk-API wäre ein WAGO-spezifischer Convenience-Endpoint)*

---

## 10. Strategic Options (C)

### C1 – PCF Read Model / Lightweight Index

**Konzept:**  
Beim Publizieren/Aktualisieren einer AAS wird zusätzlich ein schlankes Read Model gepflegt:

| Feld | Typ | Beispiel |
|------|-----|---------|
| `articleNumber` | String (indexed) | `280-101` |
| `aasId` | String | `https://wago.com/ids/aas/280-101/...` |
| `pcfSubmodelId` | String | `https://wago.com/ids/submodel/280-101/PCF` |
| `pcfValue` | Float | 0.0156 |
| `pcfUnit` | String | `kg CO2e` |
| `lastUpdated` | Timestamp | `2026-08-01T12:00:00Z` |

**AAS bleibt Source of Truth.** Read Model ist nur für schnelle Lesezugriffe.

**Vorteile:**
- Artikelnummer-Lookup in < 10 ms (direkter Index-Hit)
- Kein Registry-Hop, kein Submodel-Fetch
- Skaliert auf Millionen Artikel

**Nachteile:**
- Zusätzliche Infrastruktur nötig
- Synchronisations-Komplexität (Event-driven oder Polling)
- Risiko: Read Model veraltet, wenn Event-Verarbeitung ausfällt

**Bewertung:** Sinnvoll, wenn WAGO >10.000 Artikel mit PCF hat und mehrere interne Anwendungen PCF-Daten abfragen. Bei aktuellem Stand (117.418 Shells, aber nur ~35 PCF-verfügbar in diesem Test) eher mittelfristig relevant. **Kein Overengineering – eine nützliche Weiterentwicklung**.

### C2 – Event-driven Index-Aktualisierung

**Trigger:**
- `AAS created` → Read Model Eintrag anlegen
- `PCF submodel updated` → `pcfValue`, `lastUpdated` aktualisieren
- `AAS deleted` → Read Model Eintrag löschen

**Technologie (muss IT verifizieren):** Apache Kafka, RabbitMQ, Azure Service Bus oder ähnlich – abhängig von bestehender WAGO-Infrastruktur.

**Bewertung:** Sinnvoll nur in Kombination mit C1. Erhöhte Komplexität, aber robustere Aktualität als TTL-basiertes Caching.

### C3 – Direkter Article-Number-Lookup-Endpoint

```http
GET https://c1.api.wago.com/smartdata-aas-env/products/{articleNumber}/pcf
```

Dieser Endpoint führt intern Shell-Lookup + Submodel-Fetch durch und gibt direkt den PCF-Wert zurück.

**Bewertung:** Sehr sinnvoll als langfristiger WAGO-spezifischer Endpoint. Vereinfacht alle Client-Implementierungen. Kombinierbar mit Read Model (C1) als Backend.

---

## 11. Monitoring / SLO Recommendations

### Empfohlene Metriken (pro Service)

**Shell Registry Service:**
```
aas_registry_request_count{endpoint, method, status}
aas_registry_request_duration_ms{p50, p95, p99, endpoint}
aas_registry_response_size_bytes{endpoint}
aas_registry_result_count{endpoint}      # Anzahl zurückgegebener Descriptors
aas_registry_db_query_duration_ms        # DB-Abfrage-Zeit isoliert
```

**Submodel Environment Service:**
```
aas_submodel_request_count{status}
aas_submodel_request_duration_ms{p50, p95, p99}
aas_submodel_response_size_bytes
```

**Übergreifend:**
```
aas_request_total_duration_ms            # End-to-End (für Client-Use-Case)
aas_error_rate{type: timeout, 4xx, 5xx}
aas_cold_start_events_count              # Outlier wie 2000-115 (8 s)
```

### Distributed Tracing

Jeder Client-Request sollte eine `X-Request-ID` oder `traceparent` (W3C Trace Context) mitschicken. Serverseitig: Trace-ID durch alle beteiligten Services weiterleiten (Registry → DB, Registry → Submodel Service). Tool: OpenTelemetry + Jaeger oder Zipkin.

**Ziel:** Bei einem langsamen Request nachvollziehen:  
`Warum war Request X 15 Sekunden langsam? → DB-Query: 14,8 s → kein Index vorhanden`

### Vorgeschlagene SLOs

| Operation | P50 | P95 | P99 | Max |
|-----------|----:|----:|----:|----:|
| Fast Path Submodel Fetch | < 100 ms | < 300 ms | < 500 ms | < 2 s |
| Shell Registry gefiltert (assetIds) | < 200 ms | < 500 ms | < 1 s | < 3 s |
| Shell Registry ungefiltert | nicht angeboten | – | – | – |
| End-to-End PCF-Abfrage (ein Artikel) | < 200 ms | < 800 ms | < 2 s | < 5 s |
| Bulk PCF-Abfrage (100 Artikel) | < 2 s | < 5 s | < 10 s | < 30 s |

*(Diese SLOs sind Diskussionsvorschläge für die IT – abzustimmen mit tatsächlichen Service-Level-Agreements)*

---

## 12. Questions for IT

1. **Index auf `globalAssetId`:**  
   Ist `globalAssetId` (der zentrale Lookup-Schlüssel für den `assetIds`-Filter am `/shell-descriptors`-Endpoint) in der Datenbank der Shell Registry indexiert? Unser Benchmark misst 11–23 Sekunden für eine gefilterte Anfrage – was auf einen Full-Table-Scan bei 117.418 Einträgen hinweist.

2. **assetIds-Filter Implementierung:**  
   Wie ist der `assetIds`-Filterparameter serverseitig implementiert? Wird die Filterung als parametrisierte DB-Abfrage (WHERE-Clause mit Index) ausgeführt, oder wird das Ergebnis in der Anwendungsschicht nach einem breiten DB-Fetch gefiltert?

3. **Default Page Size:**  
   Warum liefert `GET /shell-descriptors` ohne `limit`-Parameter alle 117.418 Einträge auf einmal (37,7 MB)? Gibt es eine konfigurierbare maximale Default-Page-Size, und wenn nicht, kann diese eingeführt werden?

4. **Cold-Start-Verhalten:**  
   Artikel `2000-115` benötigte als erster Request im Benchmark 8.240 ms, alle nachfolgenden Requests an denselben Endpoint benötigten < 100 ms. Gibt es einen Application-Cache oder Connection-Pool, der nach Idle-Zeit zurückgesetzt wird? Kann ein Warmup-Mechanismus (z. B. Liveness-Probe-Request) implementiert werden?

5. **Submodel-Descriptors im Shell-Descriptor:**  
   Enthält der Shell-Descriptor-Response von `/shell-descriptors?assetIds=...` die `submodelDescriptors` vollständig inklusive Submodel-IDs? In unserem Benchmark sahen wir nur `endpoints`, `idShort`, `id` im Top-Level-Result. Wenn `submodelDescriptors` nicht enthalten sind, ist ein zusätzlicher Registry-Hop nötig.

6. **Submodel-Descriptors-Endpoint:**  
   `GET /submodel-descriptors` antwortet mit HTTP 404. Ist dieser Endpoint deaktiviert, oder ist die URL anders? Der AAS-Part-2-Standard sieht einen Submodel Registry-Endpoint vor.

7. **Wachstum der Registry:**  
   Wie schnell wächst die Registry? Gibt es eine Prognose für 1 Jahr / 3 Jahre? Bei linearem Wachstum ab 117.418 Shells ist es wichtig zu wissen, wann Performance-Probleme auch den Fast Path betreffen.

8. **Distributed Tracing:**  
   Gibt es aktuell Distributed Tracing (z. B. OpenTelemetry, Jaeger, Zipkin) in der AAS-Infrastruktur? Wenn nicht: Können wir einen `X-Request-ID`-Header implementieren, sodass einzelne langsame Requests in den Server-Logs nachverfolgt werden können?

9. **Bulk Query Endpoint:**  
   Ist ein `POST /pcf/query`-Endpoint mit einer Liste von Artikelnummern als Input geplant oder technisch machbar? Aktuell erzeugt unser Tool für 60 Artikel bis zu 120 HTTP-Requests (plus Retries). Bei 500 Artikeln wären das 500–1500 Requests.

10. **Datenbanktyp und Abfrage-Profiling:**  
    Welche Datenbank-Technologie wird für die Shell Registry verwendet? Gibt es ein Profiling-/Slow-Query-Log, in dem wir sehen können, wie lange die DB-Abfrage für einen `assetIds`-gefilterten Request dauert? Dies würde helfen, den Unterschied zwischen Netzwerklatenz, DB-Abfrage und Anwendungslogik zu trennen.

---

## 13. Meeting Summary

### Current Situation

- Das WAGO AAS PCF Query Tool fragt für jede Artikelnummer 1–3 HTTP-Requests an den AAS-Endpoints ab
- **35 von 60 Artikeln (58 %)** werden über den schnellen Fast Path in **54–100 ms** beantwortet
- **25 von 60 Artikeln (42 %)** benötigen den Fallback über die Shell Registry und brauchen **11–23 Sekunden** pro Artikel
- Die Shell Registry enthält **117.418 Shell-Descriptors** und liefert diese ohne Filterung als **37,7 MB-Response in 17 Sekunden**
- Mit `?limit=1` antwortet dieselbe Registry in **74 ms** – die Infrastruktur ist grundsätzlich schnell

### Main Bottlenecks

- **Bottleneck #1 (Kritisch):** Der `assetIds`-Filter am Shell Registry-Endpoint ist nicht performant – vermutlich kein Datenbankindex auf `globalAssetId`. Resultat: 11–23 Sekunden statt < 500 ms.
- **Bottleneck #2 (Strukturell):** Kein Default-Limit am Registry-Endpoint – ungefilterte Requests laden 37,7 MB (117.418 Einträge) auf einmal.
- **Bottleneck #3 (Beobachtet):** Cold-Start-Outlier: Der erste Request nach einer Idle-Phase benötigt 8+ Sekunden (Cache-/Connection-Pool-Warmup).
- **Skalierungsrisiko:** Bei weiterhin wachsender Registry (aktuell 117.418 Shells) verschärfen sich diese Probleme proportional.
- **Visibility-Problem:** Aktuell merkt die IT Performance-Probleme erst dann, wenn Clients berichten – kein Server-seitiges Slow-Query-Monitoring vorhanden.

### Recommended Next Steps

1. **Sofort (DBA, 1 Tag):** DB-Index auf `globalAssetId` in der Shell Registry prüfen und ggf. anlegen → Fallback von 15 s auf < 500 ms
2. **Kurzfristig (1 Woche):** Default-Page-Size am Registry-Endpoint konfigurieren → verhindert 37-MB-Responses
3. **Kurzfristig (1 Woche):** Server-seitiges Slow-Query-Logging aktivieren → IT kann Probleme eigenständig erkennen
4. **Mittelfristig (1–3 Monate):** Bulk Query API `POST /pcf/query` für Batch-Use-Case entwickeln → von 1500 auf 1 HTTP-Request
5. **Mittelfristig (3–6 Monate):** PCF Read Model / Lightweight Index aufbauen → Artikelnummer → PCF in < 10 ms

### Most Important Message

> **Der Shell Registry-Endpoint lädt ohne Filterung 117.418 Einträge in 17 Sekunden – ein einziger Datenbankindex auf `globalAssetId` könnte die Antwortzeit für 42 % unserer Artikel von 15+ Sekunden auf unter 500 Millisekunden reduzieren.**

---

## Appendix: Disclaimer – Beobachtet / Wahrscheinlich / IT-Verifikation nötig

### Observed (direkt gemessen)
- Fast Path Median: **66 ms** (n=60)
- Fallback Shell Registry Median: **15.551 ms** (n=25)
- Registry ohne Filter: **37.708.143 Bytes, 117.418 Einträge, 17.354 ms**
- Registry mit `?limit=1`: **416 Bytes, 74 ms**
- Cold-Start-Outlier `2000-115`: **8.240 ms** (einmaliger erster Request)
- `GET /submodel-descriptors`: **HTTP 404**
- 35/60 Artikel: PCF via Fast Path; 25/60: kein PCF via Fallback
- Shell-Descriptor enthält Keys: `endpoints`, `idShort`, `id`

### Likely (wahrscheinlich, aber nicht direkt messbar von Client-Seite)
- Fehlender oder ineffizienter DB-Index auf `globalAssetId` verursacht den 11–23 s Fallback
- `assetIds`-Filter wird serverseitig als Post-Processing auf vollem DB-Result ausgeführt, nicht als WHERE-Clause
- Cold-Start-Outlier ist auf Connection-Pool-Timeout oder Application-Cache-Expiry zurückzuführen
- `999-999` und `100-280` haben Shells in der Registry (wurden als „AAS available - PCF not available" klassifiziert) – möglicherweise Fehldaten oder Test-Shells

### Needs IT Verification (nur serverseitig prüfbar)
- DB-Queryplan für `?assetIds`-Filter (Explain/Execution Plan)
- Datenbanktyp und Index-Konfiguration der Shell Registry
- Ob `submodelDescriptors` im gefilterten Shell-Descriptor-Response vollständig enthalten sind
- Ob Response Compression (gzip) aktiviert ist
- Connection-Pool-Konfiguration (Idle-Timeout, Pool-Size)
- Ob Distributed Tracing vorhanden ist
- Warum `GET /submodel-descriptors` HTTP 404 liefert
- Wachstumsprognose der Registry
