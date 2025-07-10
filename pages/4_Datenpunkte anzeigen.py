import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode

def display_data_model(excel_file_path):
    try:
        # Excel einlesen
        xls = pd.ExcelFile(excel_file_path)
        data = {sheet: xls.parse(sheet).fillna("") for sheet in xls.sheet_names}

        # Tabellen extrahieren
        dp_df = data.get("Datenpunkt", pd.DataFrame())
        kennzahlen_df = data.get("Kennzahl", pd.DataFrame())
        regelwerk_df = data.get("Regelwerk", pd.DataFrame())
        stakeholder_df = data.get("Stakeholder", pd.DataFrame())
        regelwerk_dp_df = data.get("Regelwerk-Datenpunkt", pd.DataFrame())

        if dp_df.empty:
            st.warning(" Keine Datenpunkte gefunden.")
            return

        st.title(" Datenmodell anzeigen")
        st.write("Klicke auf einen Datenpunkt, um verknüpfte Kennzahlen anzuzeigen.")
        st.write("Sortier die Datenpunkte mithilfe der drei Auswahlmöglichkeiten.")
        # Filter
        col1, col2, col3 = st.columns(3)
        with col1:
            regelwerk_options = ["Alle"] + regelwerk_df.get("Name", pd.Series()).dropna().unique().tolist()
            selected_regelwerk = st.selectbox(" Regelwerk", regelwerk_options)
        with col2:
            stakeholder_options = ["Alle"] + stakeholder_df.get("Name", pd.Series()).dropna().unique().tolist()
            selected_stakeholder = st.selectbox("Stakeholder", stakeholder_options)
        with col3:
            gruppen_options = ["Alle"] + sorted(dp_df.get("Gruppe", pd.Series()).dropna().unique().tolist())
            selected_gruppe = st.selectbox(" Gruppe", gruppen_options)

        # Filterlogik
        filtered_dp_df = dp_df.copy()

        if selected_stakeholder != "Alle" and not stakeholder_df.empty and not kennzahlen_df.empty:
            match = stakeholder_df[stakeholder_df["Name"] == selected_stakeholder]
            if not match.empty:
                stakeholder_id = match.iloc[0]["Stakeholder-ID"]
                relevant_dp_ids = kennzahlen_df[
                    kennzahlen_df["Stakeholder-ID"] == stakeholder_id
                ]["Datenpunkt-ID"].unique()
                filtered_dp_df = filtered_dp_df[filtered_dp_df["Datenpunkt-ID"].isin(relevant_dp_ids)]

        if selected_regelwerk != "Alle" and not regelwerk_df.empty and not regelwerk_dp_df.empty:
            match = regelwerk_df[regelwerk_df["Name"] == selected_regelwerk]
            if not match.empty:
                regelwerk_id = match.iloc[0]["Regelwerk-ID"]
                relevant_dp_ids = regelwerk_dp_df[
                    regelwerk_dp_df["Regelwerk-ID"] == regelwerk_id
                ]["Datenpunkt-ID"].unique()
                filtered_dp_df = filtered_dp_df[filtered_dp_df["Datenpunkt-ID"].isin(relevant_dp_ids)]

        if selected_gruppe != "Alle":
            filtered_dp_df = filtered_dp_df[filtered_dp_df["Gruppe"] == selected_gruppe]

        st.subheader(" Gefilterte Datenpunkte")

        # Haupttabelle
        gb = GridOptionsBuilder.from_dataframe(filtered_dp_df)
        gb.configure_selection("single", use_checkbox=False)
        grid_options = gb.build()

        grid_response = AgGrid(
            filtered_dp_df,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.SELECTION_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            theme="streamlit",
            fit_columns_on_grid_load=True,
            domLayout="autoHeight"
        )

        # Selektion verarbeiten
        raw_selection = grid_response.get("selected_rows", None)

        if raw_selection is None:
            selected = []
        elif isinstance(raw_selection, pd.DataFrame):
            selected = raw_selection.to_dict(orient="records")
        elif isinstance(raw_selection, list):
            selected = raw_selection
        else:
            selected = []

        if selected:
            selected_obj = selected[0]
            selected_dp_id = selected_obj.get("Datenpunkt-ID") if isinstance(selected_obj, dict) else None

            if isinstance(selected_dp_id, (int, float, str)) and not pd.isna(selected_dp_id):
                selected_dp_id_str = str(selected_dp_id)

                st.markdown(f"###  Zugehörige Kennzahlen zu Datenpunkt-ID `{selected_dp_id_str}`")

                matching_kz = kennzahlen_df[
                    kennzahlen_df["Datenpunkt-ID"].astype(str) == selected_dp_id_str
                ]

                if not matching_kz.empty:
                    AgGrid(
                        matching_kz,
                        theme="streamlit",
                        fit_columns_on_grid_load=True,
                        domLayout="autoHeight"
                    )
                else:
                    st.info(" Für diesen Datenpunkt sind keine Kennzahlen vorhanden.")
            else:
                st.warning(" Ungültige oder leere Datenpunkt-ID ausgewählt.")
        else:
            st.info("Bitte wähle einen Datenpunkt aus der Tabelle.")

    except Exception as e:
        st.error(f" Fehler beim Verarbeiten der Datei: {e}")
    
display_data_model("/Matching/Datenmodell.xlsx")