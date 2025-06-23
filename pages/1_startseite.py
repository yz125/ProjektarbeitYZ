import streamlit as st
import pandas as pd

# Hartcodierter Dateiname
EXCEL_DATEI = "Datenmodell.xlsx"

def analyse_datenmodell():
    """
    Liest 'Datenmodell.xlsm', filtert nach 'Regelwerk', zählt Einträge mit/ohne 'Daten',
    und zeigt die Ergebnisse mit Streamlit-eigenen Funktionen (ohne externe Chartbibliotheken).
    """
    st.header("Analyse des Datenmodells")

    @st.cache_data
    def load_data():
        try:
            df = pd.read_excel(EXCEL_DATEI, engine="openpyxl")
            return df
        except Exception as e:
            st.error(f"Fehler beim Laden der Datei: {e}")
            return None

    df = load_data()

    if df is None:
        return

    if "Regelwerk" not in df.columns or "Daten" not in df.columns:
        st.error("Die Datei muss die Spalten 'Regelwerk' und 'Daten' enthalten.")
        return

    # Filter für Regelwerk
    regelwerke = df["Regelwerk"].dropna().unique()
    ausgewählte_regelwerke = st.multiselect(
        "Filtere nach 'Regelwerk':",
        options=sorted(regelwerke),
        default=sorted(regelwerke)
    )

    # Gefilterte Daten
    df_filtered = df[df["Regelwerk"].isin(ausgewählte_regelwerke)]

    # Berechnungen
    gesamtanzahl = len(df_filtered)
    mit_daten = df_filtered["Daten"].notna().sum()
    ohne_daten = gesamtanzahl - mit_daten

    # Ergebnisse anzeigen
    st.subheader("Ergebnisse")
    st.metric("Gesamtanzahl", gesamtanzahl)
    st.metric("Mit 'Daten'-Eintrag", mit_daten)
    st.metric("Ohne 'Daten'-Eintrag", ohne_daten)

    # Daten für Balkendiagramm vorbereiten
    chart_data = pd.DataFrame({
        "Anzahl": [mit_daten, ohne_daten]
    }, index=["Mit Daten", "Ohne Daten"])

    st.subheader("Visualisierung")
    st.bar_chart(chart_data)

    # Optional: Tabelle anzeigen
    with st.expander("Gefilterte Daten anzeigen"):
        st.dataframe(df_filtered)
analyse_datenmodell()