import streamlit as st
import pandas as pd
import os

def run_excel_matcher():
    st.set_page_config(page_title="Excel Topic Matcher", layout="wide")
    st.title("🔗 Excel Topic Matcher")

    uploaded_file = st.file_uploader("📥 Ziehe deine Longlist Excel-Datei hierher", type=["xlsx"])

    try:
        matching_df = pd.read_excel("Matching.xlsx", sheet_name=0, dtype=str)
        matching_df = matching_df.loc[:, ~matching_df.columns.str.contains("^Unnamed")]
    except FileNotFoundError:
        st.error("⚠️ Die Datei 'Matching.xlsx' wurde nicht gefunden.")
        return

    result_df = None  # für späteres Speichern

    if uploaded_file:
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            if "Longlist" not in excel_file.sheet_names:
                st.error("⚠️ Die hochgeladene Excel-Datei enthält kein Blatt namens 'Longlist'.")
                return
            longlist_df = excel_file.parse("Longlist", dtype=str)
            longlist_df = longlist_df.loc[:, ~longlist_df.columns.str.contains("^Unnamed")]
        except Exception as e:
            st.error(f"❌ Fehler beim Lesen der Excel-Datei: {e}")
            return

        required_columns = ["Unter-Unterthema", "Unterthema"]
        if not all(col in longlist_df.columns for col in required_columns):
            st.error("⚠️ Die Longlist-Datei muss die Spalten 'Unter-Unterthema' und 'Unterthema' enthalten.")
            return

        if "Topic" not in matching_df.columns:
            st.error("⚠️ Die Matching-Datei muss eine Spalte 'Topic' enthalten.")
            return

        if "Default" not in matching_df.columns:
            matching_df["Default"] = ""

        matched_rows = []

        unter_unterthemen = longlist_df["Unter-Unterthema"].dropna().astype(str).unique()
        unterthemen = longlist_df["Unterthema"].dropna().astype(str).unique()

        def topic_matches(topic_str):
            if str(topic_str).strip().lower() == "immer":
                return True
            topic_str = str(topic_str).lower()
            return any(t.lower() in topic_str for t in unter_unterthemen) or \
                   any(t.lower() in topic_str for t in unterthemen)

        st.subheader("🔍 Bedingungen bestätigen")
        for idx, row in matching_df.iterrows():
            topic = str(row.get("Topic", "")).strip()
            condition = str(row.get("Bedingung", "")).strip()
            note = str(row.get("Bemerkung", "")).strip()

            if topic_matches(topic):
                if condition.lower() == "bedingt":
                    default_value = str(row.get("Default", "Ja")).strip().capitalize()
                    default_value = default_value if default_value in ("Ja", "Nein") else "Ja"
                    user_input = st.radio(
                        f"💬 {note} (Zeile {idx+2})",
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

        # Speichere aktualisierte Matching.xlsx
        try:
            with pd.ExcelWriter("Matching.xlsx", engine="openpyxl", mode="w") as writer:
                matching_df.to_excel(writer, index=False)
        except Exception as e:
            st.error(f"❌ Fehler beim Speichern der Matching-Datei: {e}")
            return

        if not matched_rows:
            st.info("ℹ️ Keine passenden Zeilen gefunden oder alle Bedingungen wurden abgelehnt.")
            return

        result_df = pd.DataFrame(matched_rows).dropna(how="all")
        result_df = result_df.loc[:, ~result_df.columns.str.contains("^Unnamed")]

        # Spalten-Mapping
        spalten_mapping = {
            "Id": "Id",
            "DR": "Regelwerk",
            "Paragraph": "Paragraph",
            "Name": "Name",
            "Datentyp": "Datentyp"
        }

        mapped_df = result_df.rename(columns=spalten_mapping)
        mapped_df = mapped_df.loc[:, ~mapped_df.columns.duplicated()]
        zielspalten = list(spalten_mapping.values())
        mapped_df = mapped_df[[col for col in zielspalten if col in mapped_df.columns]]

        # Auswahl zur Id-Verarbeitung
        id_handling = st.selectbox(
            "🆔 Wie sollen Zeilen mit bereits vorhandener Id in 'Datenmodell.xlsx' behandelt werden?",
            ("Ersetzen", "Beibehalten")
        )

        # Button zum Speichern
        if st.button("💾 Speichern in Datenmodell.xlsx"):
            if os.path.exists("Datenmodell.xlsx"):
                try:
                    bestehende_df = pd.read_excel("Datenmodell.xlsx", dtype=str)
                    bestehende_df = bestehende_df.loc[:, ~bestehende_df.columns.str.contains("^Unnamed")]
                except Exception as e:
                    st.error(f"❌ Fehler beim Lesen von 'Datenmodell.xlsx': {e}")
                    return
            else:
                bestehende_df = pd.DataFrame(columns=zielspalten)

            if "Id" not in bestehende_df.columns:
                bestehende_df["Id"] = ""

            if id_handling == "Ersetzen":
                bestehende_df = bestehende_df[~bestehende_df["Id"].isin(mapped_df["Id"])]
                final_df = pd.concat([bestehende_df, mapped_df], ignore_index=True)
            else:  # Beibehalten
                mapped_df = mapped_df[~mapped_df["Id"].isin(bestehende_df["Id"])]
                final_df = pd.concat([bestehende_df, mapped_df], ignore_index=True)

            final_df.dropna(how="all", inplace=True)

            try:
                with pd.ExcelWriter("Datenmodell.xlsx", engine="openpyxl", mode="w") as writer:
                    final_df.to_excel(writer, index=False)
                st.success("✅ Daten erfolgreich in 'Datenmodell.xlsx' gespeichert.")
            except Exception as e:
                st.error(f"❌ Fehler beim Schreiben in 'Datenmodell.xlsx': {e}")
run_excel_matcher()