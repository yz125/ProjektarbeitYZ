import streamlit as st
import pandas as pd
from io import BytesIO

def run_excel_matcher():
    matching_datei = "/Matching/Matching.xlsx"
    st.set_page_config(page_title="Excel Topic Matcher", layout="wide")
    st.title(" Excel Topic Matcher")

    uploaded_file = st.file_uploader(" Ziehe eine Shortlist Excel-Datei hierher", type=["xlsx"])

    try:
        matching_df = pd.read_excel(matching_datei, sheet_name=0, dtype=str)
        matching_df = matching_df.loc[:, ~matching_df.columns.str.contains("^Unnamed")]
    except FileNotFoundError:
        st.error(" Die Datei 'Matching.xlsx' wurde nicht gefunden.")
        return

    if uploaded_file:
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            if "Shortlist" not in excel_file.sheet_names:
                st.error(" Die hochgeladene Excel-Datei enthält kein Blatt namens 'Shortlist'.")
                return
            shortlist_df = excel_file.parse("Shortlist", dtype=str)
            shortlist_df = shortlist_df.loc[:, ~shortlist_df.columns.str.contains("^Unnamed")]
        except Exception as e:
            st.error(f" Fehler beim Lesen der Excel-Datei: {e}")
            return

        required_columns = ["Unter-Unterthema", "Unterthema"]
        if not all(col in shortlist_df.columns for col in required_columns):
            st.error(" Die Shortlist-Datei muss die Spalten 'Unter-Unterthema' und 'Unterthema' enthalten.")
            return

        if "Topic" not in matching_df.columns:
            st.error(" Die Matching-Datei muss eine Spalte 'Topic' enthalten.")
            return

        if "Default" not in matching_df.columns:
            matching_df["Default"] = ""

        matched_rows = []

        unter_unterthemen = shortlist_df["Unter-Unterthema"].dropna().astype(str).unique()
        unterthemen = shortlist_df["Unterthema"].dropna().astype(str).unique()

        def topic_matches(topic_str):
            if str(topic_str).strip().lower() == "immer":
                return True
            topic_str = str(topic_str).lower()
            return any(t.lower() in topic_str for t in unter_unterthemen) or \
                   any(t.lower() in topic_str for t in unterthemen)

        st.subheader(" Bedingungen bestätigen")
        for idx, row in matching_df.iterrows():
            topic = str(row.get("Topic", "")).strip()
            condition = str(row.get("Bedingung", "")).strip()
            note = str(row.get("Bemerkung", "")).strip()

            if topic_matches(topic):
                if condition.lower() == "bedingt":
                    default_value = str(row.get("Default", "Ja")).strip().capitalize()
                    default_value = default_value if default_value in ("Ja", "Nein") else "Ja"
                    user_input = st.radio(
                        f" {note} (Zeile {idx+2})",
                        ("Ja", "Nein"),
                        index=0 if default_value == "Ja" else 1,
                        key=f"bedingung_{idx}"
                    )
                    matching_df.at[idx, "Default"] = user_input
                    if user_input == "Ja":
                        matched_rows.append(row)
                else:
                    matching_df.at[idx, "Default"] = ""
                    matched_rows.append(row)

        if not matched_rows:
            st.info("ℹ Keine passenden Zeilen gefunden oder alle Bedingungen wurden abgelehnt.")
            return

        result_df = pd.DataFrame(matched_rows).dropna(how="all")
        result_df = result_df.loc[:, ~result_df.columns.str.contains("^Unnamed")]

        # Spalten-Mapping
        spalten_mapping = {
            "Id": "Id",
            "DR": "Regelwerk",
            "Paragraph": "Paragraph",
            "Name": "Name",
            "Datentyp": "Datentyp",
            "Topic": "Gruppe"
        }

        mapped_df = result_df.rename(columns=spalten_mapping)
        mapped_df = mapped_df.loc[:, ~mapped_df.columns.duplicated()]
        zielspalten = list(spalten_mapping.values())
        mapped_df = mapped_df[[col for col in zielspalten if col in mapped_df.columns]]

        #  Bereitstellen als Download 
        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            mapped_df.to_excel(writer, index=False, sheet_name="Matched")

        st.download_button(
            label=" Ergebnis herunterladen",
            data=excel_buffer.getvalue(),
            file_name="Ergebnis_Matching.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
run_excel_matcher()