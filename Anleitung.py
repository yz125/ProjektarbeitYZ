import streamlit as st

def show_anleitung():
    st.set_page_config(page_title=" Anleitung: Excel Matching Umgebung", layout="wide")
    st.title(" Anleitung zur Excel-Matching-Umgebung")

    # Abschnitt: Setup
    st.markdown("## Setup")
    st.markdown("""
Damit alle Funktionen genutzt werden können, muss zunächst der folgenden Ordner angelegt werden:  
**`C:/Matching/`**

Darin müssen folgende Dateien liegen:

1.  **`Matching.xlsx`** – enthält die Matching-Regeln und Bedingungen (für die Seite *Matching*).
2.  **`Datenmodell.xlsx`** – enthält das zentrale Datenmodell (für mehrere Seiten verwendet).
3.  **`masken.json`** – JSON-Datei mit definierten Extraktionsmasken (für die Seite *Datenpunkte extrahieren* ).

> Die Pfade und Dateinamen sind im Code fest kodiert. Bitte benennen Sie nichts um.
""")

    st.markdown("###  Beispielstruktur")
    st.code("""
C:/
└── Matching/
    ├── Datenmodell.xlsx
    ├── Matching.xlsx
    └── masken.json
""", language="text")

    st.markdown("---")

    # Abschnitt: Seitenübersicht
    st.markdown("##  Seitenübersicht")

    st.markdown("### 1. `Datenpunkte extrahieren` –  Daten aus Dateien extrahieren")
    st.markdown("""
- Lädt externe Excel-Dateien (beliebige Struktur)
- Nutzt eine definierte Maske (aus `masken.json`) zur Filterung
- Exportiert gefilterte Daten als neue Excel-Datei
- Verwaltung der Masken über Tabs:
  - Neue Maske erstellen
  - Bestehende bearbeiten oder löschen
  - Import/Export von Masken im JSON-Format
""")

    st.markdown("### 2. `Matching` –  Excel Topic Matcher")
    st.markdown("""
- Verarbeitet eine hochgeladene Shortlist (Excel mit Sheet „Shortlist“)
- Verknüpft Themen aus der Shortlist mit der Datei `Matching.xlsx`
- Prüft bedingte Datenpunkte („Ja“ oder „Nein“)
- Ergebnis: Gefilterte Excel-Datei mit passenden Topics zum Download
""")

    st.markdown("### 3. `Einträge erstellen` –  Daten manuell erfassen")
    st.markdown("""
- Enthält mehrere Tabs zur Pflege des Datenmodells:
  - **Regelwerk erstellen**: Neue Regelwerke anlegen
  - **Stakeholder hinzufügen**: Organisationen einpflegen
  - **Datenpunkte hochladen**: Datenpunkte aus Excel importieren (mit Mapping)
  - **Standorte hinzufügen**: Adresse und Zuordnung zu Stakeholdern
  - **Weitere Standards verknüpfen**: Paragrafen aus anderen Regelwerken übernehmen
  - **Kennzahl zuordnen**: Werte zu Datenpunkten und Stakeholdern hinzufügen
""")

    st.markdown("### 4. `Datenpunkte anzeigen` –  Übersicht & Filter")
    st.markdown("""
- Zeigt alle vorhandenen Datenpunkte aus `Datenmodell.xlsx`
- Filterbar nach:
  - Regelwerk
  - Stakeholder
  - Gruppe
- Zeigt bei Auswahl die zugehörigen **Kennzahlen**
""")

    st.markdown("### 5. `Daten ändern` –  Entitäten direkt bearbeiten")
    st.markdown("""
- Auswahl einer Entität (z. B. Stakeholder, Datenpunkt, Standort)
- Bearbeitung direkt im Web-Grid
- Änderungen werden sofort in `Datenmodell.xlsx` gespeichert
""")

    st.markdown("### 6. `Daten löschen` –  Einträge entfernen")
    st.markdown("""
- Auswahl eines Eintrags aus einer Entität
- Löscht den Eintrag sowie verknüpfte Daten aus allen Tabellen
- Achtung: Vorgang ist **nicht rückgängig** machbar
""")

    st.markdown("---")

    # Abschnitt: Voraussetzungen
    st.markdown("##  Voraussetzungen")
    st.markdown("""
- Alle Dateien im `.xlsx`-Format
- JSON-Dateien 
- Excel-Dateien **dürfen beim Speichern nicht geöffnet sein**
""")

    # Abschnitt: Hilfe
    st.markdown("##  Hilfe & Fehlerbehebung")
    st.markdown("""
Wenn beim Start Fehlermeldungen auftreten wie z. B.:

>  Datei `Matching.xlsx` wurde nicht gefunden.

dann prüfen Sie bitte:

-  Existiert der Ordner `C:/Matching/`?
-  Liegen alle Dateien mit den richtigen Namen dort?
-  Sind die Excel-Dateien geschlossen?
""")

    st.success(" Wenn alles eingerichtet ist, kann mit der Nutzung begonnen werden.")

show_anleitung()