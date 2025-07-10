import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
import os

def edit_entity_and_save_to_model(excel_file_path):
    st.title(" Entitäten aus Datenmodell bearbeiten und speichern")

    try:
        if not os.path.exists(excel_file_path):
            st.error(" Die Datei wurde nicht gefunden.")
            return

        xls = pd.ExcelFile(excel_file_path)
        tabellen = xls.sheet_names
        entitaeten = [name for name in tabellen if "-" not in name]

        selected_table = st.selectbox(" Entität auswählen", entitaeten)

        df = xls.parse(selected_table).fillna("")
        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

        if df.empty:
            st.info(" Diese Entität enthält keine Daten.")
            return

        #  Sucheingabe
        search_term = st.text_input(" Suche in der Tabelle").strip().lower()
        if search_term:
            df = df[df.apply(lambda row: row.astype(str).str.lower().str.contains(search_term).any(), axis=1)]

        #  AgGrid-Konfiguration
        gb = GridOptionsBuilder.from_dataframe(df)
        gb.configure_pagination()
        gb.configure_default_column(editable=True, wrapText=True, autoHeight=True)
        gb.configure_grid_options(domLayout="autoHeight")

        #  Alle Spalten mit "ID" im Namen nicht bearbeitbar machen
        id_spalten = [col for col in df.columns if "id" in col.lower()]
        for id_col in id_spalten:
            gb.configure_column(id_col, editable=False)

        #  Spalte "Wert" explizit als Textspalte definieren (Text statt Zahl)
        if "Wert" in df.columns:
            gb.configure_column("Wert", type=["textColumn"], cellDataType="text", editable=True)

        grid_options = gb.build()

        #  Interaktive Tabelle siehe Projektarbeit Kap. 5
        grid_response = AgGrid(
            df,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.VALUE_CHANGED,
            fit_columns_on_grid_load=True,
            theme="streamlit"
        )

        edited_df = grid_response["data"]

        if st.button(" Änderungen ins Datenmodell speichern"):
            try:
                with pd.ExcelWriter(excel_file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    edited_df.to_excel(writer, sheet_name=selected_table, index=False)
                st.success(f" Änderungen in '{selected_table}' wurden erfolgreich gespeichert.")
            except Exception as e:
                st.error(f" Fehler beim Schreiben in die Datei: {e}")

    except Exception as e:
        st.error(f" Fehler beim Verarbeiten der Datei: {e}")

edit_entity_and_save_to_model("/Matching/Datenmodell.xlsx")