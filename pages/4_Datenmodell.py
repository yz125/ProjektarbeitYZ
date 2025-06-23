import pandas as pd
import streamlit as st
import io
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode

data = pd.read_excel("Datenmodell.xlsx")
data = data.dropna(how="all")

def online_changes(df: pd.DataFrame):
    st.header("Datenverwaltung")

    # Sicherstellen, dass die Spalte 'Id' existiert
    if "Id" not in df.columns:
        st.error("Die Spalte 'Id' ist im DataFrame nicht vorhanden.")
        st.stop()

    data = df.copy()
    spalten = data.columns.tolist()

    action = st.selectbox(
        "Bitte Aktion wählen",
        [
            "Neuen Datenpunkt erstellen",
            "Datenpunkt aktualisieren",
            "Datenpunkte anzeigen",
            "Datenpunkt löschen",
        ],
    )

    if action == "Neuen Datenpunkt erstellen":
        with st.form("eingabe_formular"):
            neue_werte = {}
            for spalte in spalten:
                neue_werte[spalte] = st.text_input(f"{spalte}*")

            submit_button = st.form_submit_button("Datenpunkt erstellen")

            if submit_button:
                if not neue_werte["Id"]:
                    st.warning("Pflichtfeld 'Id' muss ausgefüllt werden.")
                elif data["Id"].astype(str).str.contains(str(neue_werte["Id"])).any():
                    st.warning(f"Ein Eintrag mit Id = '{neue_werte['Id']}' existiert bereits.")
                else:
                    neue_daten = pd.DataFrame([neue_werte])
                    updated_df = pd.concat([data, neue_daten], ignore_index=True)
                    updated_df.to_excel("Datenmodell.xlsx", index=False)
                    st.success("Neuer Datenpunkt wurde hinzugefügt.")

    elif action == "Datenpunkt aktualisieren":
        ids = data["Id"].astype(str).tolist()
        auswahl_id = st.selectbox("Id zum Bearbeiten auswählen", options=ids)

        vorhandener_eintrag = data[data["Id"].astype(str) == auswahl_id].iloc[0]

        with st.form("update_formular"):
            aktualisierte_werte = {}
            for spalte in spalten:
                aktualisierte_werte[spalte] = st.text_input(f"{spalte}", value=str(vorhandener_eintrag[spalte]))

            update_button = st.form_submit_button("Datenpunkt aktualisieren")

            if update_button:
                for spalte in spalten:
                    data.loc[data["Id"].astype(str) == auswahl_id, spalte] = aktualisierte_werte[spalte]
                data.to_excel("Datenmodell.xlsx", index=False)
                st.success("Datenpunkt erfolgreich aktualisiert.")

    elif action == "Datenpunkte anzeigen":
        st.dataframe(data)

    elif action == "Datenpunkt löschen":
        ids = data["Id"].astype(str).tolist()
        loesch_id = st.selectbox("Id zum Löschen auswählen", options=ids)

        if st.button("Datenpunkt löschen"):
            data.drop(data[data["Id"].astype(str) == loesch_id].index, inplace=True)
            data.to_excel("Datenmodell.xlsx", index=False)
            st.success("Datenpunkt wurde gelöscht.")


def convert_df_to_excel(dataframe):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Daten")
    return output.getvalue()

# 📥 Download-Button anzeigen
def download_to_excel(df):
    st.download_button(
        label="📥 Excel-Datei herunterladen",
        data=convert_df_to_excel(df),
        file_name="geänderte_daten.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# 🧩 Datenanzeige und Bearbeitung mit AG Grid
def display_and_edit_excel_data(excel_file_path):
    try:
        df = pd.read_excel(excel_file_path, dtype=str)  # Alles als String einlesen
    except FileNotFoundError:
        st.error(f"Die Datei '{excel_file_path}' wurde nicht gefunden.")
        return pd.DataFrame()

    # Leere Felder in "Daten" und "Regelwerk" explizit mit leeren Strings befüllen
    df['Daten'] = df['Daten'].fillna("")
    df['Regelwerk'] = df['Regelwerk'].fillna("")

    regelwerke = sorted(set(
        sum([rw.split(', ') for rw in df['Regelwerk'] if isinstance(rw, str)], [])
    ))

    selected_regelwerk = st.selectbox("Regelwerk filtern:", ["Alle"] + regelwerke)

    df_filtered = df.copy()
    if selected_regelwerk != "Alle":
        df_filtered = df[df['Regelwerk'].str.contains(selected_regelwerk, na=False)]

    gb = GridOptionsBuilder.from_dataframe(df_filtered)
    gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=15)
    gb.configure_side_bar()
    gb.configure_grid_options(domLayout='autoHeight')
    gb.configure_default_column(flex=1, minWidth=100, resizable=True, autoHeight=True, editable=True)
    gb.configure_grid_options(editable=True)

    grid_response = AgGrid(
        df_filtered,
        gridOptions=gb.build(),
        height=500,
        fit_columns_on_grid_load=True,
        enable_enterprise_modules=True,
        data_return_mode=DataReturnMode.AS_INPUT,
    )

    updated_df = pd.DataFrame(grid_response['data'])

    if st.button("💾 Änderungen speichern"):
        try:
            df_original = pd.read_excel(excel_file_path, dtype=str).fillna("")

            for index, row in updated_df.iterrows():
                mask = df_original['Id'] == row['Id']
                if mask.any():
                    for col in df_original.columns:
                        df_original.loc[mask, col] = str(row.get(col, "") or "")

            with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='w') as writer:
                df_original.to_excel(writer, index=False, sheet_name='Daten')

            st.success("Geänderte Zeilen wurden erfolgreich gespeichert!")
        except PermissionError:
            st.error("❌ Fehler: Die Excel-Datei ist geöffnet. Bitte schließen Sie die Datei und versuchen Sie es erneut.")
        except Exception as e:
            st.error(f"❌ Fehler beim Speichern der Daten: {e}")

    download_to_excel(updated_df)
    return updated_df

display_and_edit_excel_data("Datenmodell.xlsx")
online_changes(data)