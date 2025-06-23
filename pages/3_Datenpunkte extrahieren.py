import pandas as pd
import streamlit as st
import os
from io import BytesIO
import json

MASKEN_DATEI = "masken.json"

def lade_masken():
    if not os.path.exists(MASKEN_DATEI):
        return {}
    with open(MASKEN_DATEI, "r", encoding="utf-8") as f:
        return json.load(f)

def speichere_masken(masken):
    with open(MASKEN_DATEI, "w", encoding="utf-8") as f:
        json.dump(masken, f, indent=2, ensure_ascii=False)

def process_excel():
    st.title("Excel-Daten Extraktion & Maskenverwaltung")

    masken = lade_masken()
    tabs = st.tabs([
        "📄 Daten extrahieren",
        "➕ Neue Maske erstellen",
        "✏️ Maske bearbeiten",
        "🗑️ Maske löschen",
        "📤📥 Masken exportieren/importieren"
    ])

    # -------------------------------
    # TAB 1: DATEN EXTRAHIEREN
    # -------------------------------
    with tabs[0]:
        if not masken:
            st.warning("Keine Masken vorhanden. Bitte erstelle eine unter 'Neue Maske erstellen'.")
        else:
            maske_name = st.selectbox("Wähle eine Maske", options=list(masken.keys()))
            maske = masken[maske_name]

            sheet_keyword = st.text_input("Suchbegriff für Sheetnamen", value=maske.get("sheet_keyword", ""))
            default_columns = maske.get("columns", [])

            uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xls", "xlsx"])

            if uploaded_file:
                try:
                    xls = pd.ExcelFile(uploaded_file)
                    matching_sheets = [s for s in xls.sheet_names if sheet_keyword.lower() in s.lower()]
                    if not matching_sheets:
                        st.warning("Keine passenden Sheets gefunden.")
                        return

                    st.write(f"Gefundene Sheets: {', '.join(matching_sheets)}")
                    beispiel_df = pd.read_excel(xls, sheet_name=matching_sheets[0], skiprows=1)
                    all_columns = list(beispiel_df.columns)

                    selected_columns = st.multiselect(
                        "Wähle Spalten aus der Datei",
                        options=all_columns,
                        default=[col for col in default_columns if col in all_columns]
                    )

                    additional_columns_raw = st.text_input(
                        "Weitere Spaltennamen manuell hinzufügen (durch Komma getrennt)"
                    )
                    additional_columns = [col.strip() for col in additional_columns_raw.split(",") if col.strip()]
                    all_selected_columns = list(dict.fromkeys(selected_columns + additional_columns))

                    if st.button("Daten extrahieren"):
                        if not all_selected_columns:
                            st.warning("Bitte wähle oder gib mindestens eine Spalte an.")
                            return

                        extracted_data = []
                        for sheet in matching_sheets:
                            df = pd.read_excel(xls, sheet_name=sheet, skiprows=1)
                            df_filtered = df[[col for col in all_selected_columns if col in df.columns]]
                            extracted_data.append(df_filtered)

                        final_df = pd.concat(extracted_data, ignore_index=True)

                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            final_df.to_excel(writer, sheet_name='Gefilterte_Daten', index=False)
                        output.seek(0)

                        st.success("Daten erfolgreich extrahiert!")

                        st.download_button(
                            label="Download gefilterte Excel-Datei",
                            data=output,
                            file_name="gefilterte_daten.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except Exception as e:
                    st.error(f"Fehler beim Verarbeiten der Datei: {e}")

    # -------------------------------
    # TAB 2: NEUE MASKE ERSTELLEN / DUPLIZIEREN
    # -------------------------------
    with tabs[1]:
        st.subheader("Neue Maske erstellen")

        vorlage_maske = st.selectbox("Vorhandene Maske als Vorlage nutzen (optional)", options=[""] + list(masken.keys()))
        neue_maske_name = st.text_input("Name der neuen Maske (muss eindeutig sein)")

        if vorlage_maske and vorlage_maske in masken:
            vorlage = masken[vorlage_maske]
            sheet_keyword_default = vorlage.get("sheet_keyword", "")
            columns_default = ", ".join(vorlage.get("columns", []))
        else:
            sheet_keyword_default = ""
            columns_default = ""

        neue_sheet_keyword = st.text_input("Sheet-Suchbegriff", value=sheet_keyword_default)
        neue_columns_raw = st.text_input("Spaltennamen (durch Komma getrennt)", value=columns_default)

        if neue_maske_name in masken:
            st.warning(f"Maske '{neue_maske_name}' existiert bereits.")
        elif st.button("Maske speichern"):
            if not neue_maske_name or not neue_sheet_keyword or not neue_columns_raw:
                st.warning("Bitte fülle alle Felder aus.")
            else:
                neue_columns = [col.strip() for col in neue_columns_raw.split(",") if col.strip()]
                masken[neue_maske_name] = {
                    "sheet_keyword": neue_sheet_keyword,
                    "columns": neue_columns
                }
                speichere_masken(masken)
                st.success(f"Maske '{neue_maske_name}' wurde gespeichert!")

    # -------------------------------
    # TAB 3: MASKE BEARBEITEN
    # -------------------------------
    with tabs[2]:
        st.subheader("Bestehende Maske bearbeiten")

        if not masken:
            st.info("Keine Masken zum Bearbeiten vorhanden.")
        else:
            maske_auswahl = st.selectbox("Wähle eine Maske zum Bearbeiten", options=list(masken.keys()))
            aktuelle_maske = masken[maske_auswahl]

            edit_sheet_keyword = st.text_input("Sheet-Suchbegriff bearbeiten", value=aktuelle_maske.get("sheet_keyword", ""))
            edit_columns_raw = st.text_input("Spaltennamen bearbeiten (durch Komma getrennt)", value=", ".join(aktuelle_maske.get("columns", [])))

            if st.button("Maske aktualisieren"):
                neue_columns = [col.strip() for col in edit_columns_raw.split(",") if col.strip()]
                if not edit_sheet_keyword or not neue_columns:
                    st.warning("Alle Felder müssen ausgefüllt sein.")
                else:
                    masken[maske_auswahl] = {
                        "sheet_keyword": edit_sheet_keyword,
                        "columns": neue_columns
                    }
                    speichere_masken(masken)
                    st.success(f"Maske '{maske_auswahl}' wurde aktualisiert!")

    # -------------------------------
    # TAB 4: MASKE LÖSCHEN
    # -------------------------------
    with tabs[3]:
        st.subheader("Maske löschen")

        if not masken:
            st.info("Keine Masken zum Löschen vorhanden.")
        else:
            maske_zum_loeschen = st.selectbox("Wähle eine Maske zum Löschen", options=list(masken.keys()))
            if st.button("Maske löschen"):
                del masken[maske_zum_loeschen]
                speichere_masken(masken)
                st.success(f"Maske '{maske_zum_loeschen}' wurde gelöscht.")

    # -------------------------------
    # TAB 5: MASKEN EXPORT / IMPORT
    # -------------------------------
    with tabs[4]:
        st.subheader("📤 Exportiere oder 📥 importiere Masken")

        # Export
        st.markdown("### 📤 Masken exportieren")
        export_json = json.dumps(masken, indent=2, ensure_ascii=False).encode("utf-8")
        st.download_button(
            label="Download masken.json",
            data=export_json,
            file_name="masken.json",
            mime="application/json"
        )

        # Import
        st.markdown("### 📥 Masken importieren")
        hochgeladene_datei = st.file_uploader("Lade eine masken.json Datei hoch", type=["json"])
        if hochgeladene_datei is not None:
            try:
                neue_masken = json.load(hochgeladene_datei)
                if not isinstance(neue_masken, dict):
                    st.error("Ungültiges Format – JSON muss ein Dictionary mit Masken sein.")
                else:
                    if st.checkbox("Bestehende Masken vollständig durch neue ersetzen?"):
                        speichere_masken(neue_masken)
                        st.success("Masken wurden erfolgreich importiert.")
            except Exception as e:
                st.error(f"Fehler beim Importieren der JSON-Datei: {e}")
process_excel()