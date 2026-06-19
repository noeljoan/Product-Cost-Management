# Product Cost Management

Eine moderne, performante und mehrbenutzerfähige Web-Anwendung zur automatisierten Produktkosten-Analyse. Dieses Projekt ist die Python-Web-Alternative (Migration) eines ursprünglichen Microsoft Access-basierten Tools (`.accdb` / `.mdb`).

Das Tool ermöglicht das schnelle Durchsuchen, Filtern, Markieren und Analysieren von Material- und Kostendatensätzen, berechnet volumengewichtete Kostentreiber (Ups/Downs) und bietet automatisierte Excel-Exporte.

## 🚀 Features

- **Live-Controlling-Dashboard:** Die Startseite aggregiert und berechnet beim Aufruf in Echtzeit die wichtigsten Key Performance Indicators (KPIs) direkt aus der `data.sqlite`. Es ermittelt dynamisch die kumulierten Einsparungen (Downs) und Kostensteigerungen (Ups) sowie die Top 5 SAP-Materialnummern mit der größten Hebelwirkung (Delta × Bedarfsvolumen) im Unternehmen.
- **Zentrales Daten-Grid:** Durchsuchen, Sortieren und Filtern von tausenden Datensätzen (`tbl_alle_anzeigen`) direkt im Browser mit serverseitiger Paginierung.
- **Interaktives Markierungssystem:** Live-Berechnung von Kennzahlen (Durchschnitts-VERPR, gewichtete Gesamtkosten) direkt im Dashboard beim Auswählen von Zeilen.
- **Planpreis-Validierung:** Strukturierte Auswertungen (Variante 1 und 2) inklusive Preiserhöhungs-Warnungen (wenn der Planpreis den aktuellen Verrechnungspreis übersteigt).
- **Multi-Sheet Excel-Exporte:** Gen++erierung von professionellen Excel-Berichten (Bedarfe, Neuteile, gesperrte Teile) über performante Python-Bibliotheken.
- **Sicherer Zugriffsschutz:** Cookie-basierte Session-Authentifizierung zum Schutz sensibler Unternehmensdaten.

## 📊 SAP Datenbasis & Sharding-Architektur

Die Anwendung verarbeitet und verdichtet eine massive Datenbasis von insgesamt ca. **15 Millionen Rohdatensätzen**, die über verschiedene SAP-Standardtransaktionen und Custom-Tabellen exportiert werden:
- **`MARA`**: Allgemeine Materialdaten (Artikelstamm, Kurztexte, Einheiten)
- **`MARC`**: Werksdaten zum Material (Dispositionselemente, Werksstatus)
- **`MBEW`**: Materialbewertung (Verrechnungspreise `VERPR`, Standardpreise `STPRS`)
- **`MD04`**: Aktuelle Dispositions-Situation (Dynamische Bestände und Bedarfe für `demand()`)
- ** sowie mehreren Custom-Tabellen **

### Aushebelung der 2-GB-Grenze (Data Sharding)
Um die systemseitige 2-GB-Dateigrößenbeschränkung von Microsoft Access zu umgehen, nutzt das System eine clevere Sharding-Architektur:
1. Die 15 Millionen SAP-Datensätze werden auf mehrere separate Access-Datenbanken** aufgeteilt.
2. Diese Teil-Datenbanken werden logisch miteinander **verknüpft (Linked Tables)**, um eine nahezu unbegrenzte Skalierbarkeit im Access-Umfeld zu erreichen.
3. Das Migrationsskript (`migrate_from_access.py`) zieht diese verteilten Daten über die DAO-Schnittstelle vollautomatisch zusammen und konsolidiert sie in einer einzigen, performanten **SQLite-Datenbasis** (`data.sqlite`) für das Web-Dashboard.

## 📦 Installation und Setup

Befolgen Sie diese Schritte, um die Anwendung lokal auf Ihrem Computer oder einem Test-Server einzurichten:

### 1. Repository klonen oder herunterladen
```bash
git clone https://github.com
cd product-cost-management
```

### 2. Virtuelle Umgebung (Venv) erstellen und aktivieren
```bash
# Virtuelle Umgebung erstellen
python -m venv venv

# Aktivieren unter Windows (Eingabeaufforderung)
venv\Scripts\activate

# Aktivieren unter Linux/macOS
source venv/bin/activate
```

### 3. Abhängigkeiten installieren
```bash
pip install -r requirements.txt
```

### 3.5 Access-Daten in SQLite importieren
Bevor Sie die Anwendung das erste Mal starten (oder wenn sich die Daten in Access geändert haben), führen Sie das Migrationsskript aus. Dieses liest die `.accdb`-Datei ein, bereinigt Umlaute und erstellt die lokale `data.sqlite`:

```bash
python migrate_from_access.py
```
*(Dauer: ca. 10–30 Sekunden, bricht mit einer Erfolgsmeldung `Done.` ab).*

### 4. Anwendung starten
Starten Sie den Uvicorn-Entwicklungsserver mit automatischem Reload bei Code-Änderungen:
```bash
python -m uvicorn app.main:app --reload
```

Die Anwendung ist nun im Browser erreichbar unter:  
🔗 **[http://127.0.0.1:8000](http://127.0.0.1:8000)**

---

## 🔐 Standard-Zugangsdaten (Test-Modus)

Für den lokalen Testlauf und die Validierung auf GitHub ist ein administrativer Benutzer vorkonfiguriert:
- **Benutzername:** `admin`
- **Passwort:** `geheim123`

*Hinweis: Die Sitzung wird über ein sicheres `HTTP-Only`-Cookie verwaltet. Nach dem Schließen des Browsers oder durch Klick auf "Abmelden" wird die Session automatisch zerstört.*

## 📁 Projektstruktur

```text
product-cost-management/
├── app/                        # Python Anwendungslogik (Backend)
│   ├── __init__.py
│   ├── main.py                 # FastAPI Routen, Authentifizierung & Kern-Konfiguration
│   ├── db.py                   # SQLite-Datenbankanbindung & Basismethoden
│   ├── analysis.py             # Portierte Access-SQL-Abfragen (CTEs, Demand, Ups/Downs)
│   ├── excel.py                # Excel-Berichtsgenerierung mit Pandas/OpenPyXL
│   └── static/                 # Statische Assets (CSS, Bilder, JS)
├── templates/                  # Jinja2 HTML-Templates (Frontend UI)
│   ├── base.html               # Globales Rahmenlayout und Navigation
│   ├── home.html               # Dashboard-Kacheln und Echtzeit-KPIs
│   ├── data.html               # Haupt-Daten-Grid mit Filtern und Modal
│   ├── login.html              # Autarke, moderne Anmeldemaske
│   ├── cost.html               # Kostenentwicklung (Tab-Ansicht)
│   └── planpreis.html          # Planpreis-Auswertungen
├── data.sqlite                 # Migrierte Access-Datenbank (11 MB)
├── requirements.txt            # Python Paket-Abhängigkeiten
└── README.md                   # Projektdokumentation
```

## ⚠️ System-Hinweis (SAP-Schnittstellen)
Im Gegensatz zur alten Access-Anwendung sind lokale Desktop-Automatisierungen wie **SAP GUI Scripting** oder **HTA-Hilfen** bewusst nicht Teil dieser Web-App. Diese Funktionen setzen zwingend eine lokale Windows-Installation mit einem installierten SAP-Desktop-Client voraus und wurden durch serverseitige Datenanalysen und Standard-Excel-Berichte ersetzt.
