import streamlit as st
import pandas as pd
import os

def delete_entity_and_cascade(excel_file_path):
    st.title(" Entitätseintrag löschen (inkl. verknüpfte Daten)")

    if not os.path.exists(excel_file_path):
        st.error(" Datei wurde nicht gefunden.")
        return

    try:
        xls = pd.ExcelFile(excel_file_path)
        tabellen = xls.sheet_names
        entitaeten = [name for name in tabellen if "-" not in name]

        selected_entity = st.selectbox(" Entität auswählen", entitaeten)
        df_entity = xls.parse(selected_entity).fillna("")
        df_entity = df_entity.loc[:, ~df_entity.columns.str.contains("^Unnamed")]

        if df_entity.empty:
            st.info(" Keine Einträge in dieser Entität.")
            return

        # ID-Spalte identifizieren
        id_spalten = [col for col in df_entity.columns if "id" in col.lower()]
        id_col = next((c for c in id_spalten if c.lower().startswith(selected_entity.lower())), id_spalten[0])

        # Auswahl per Anzeige einer Zeile (nicht ID-basiert, sondern optisch)
        anzeige_spalte = next((c for c in df_entity.columns if c != id_col), df_entity.columns[0])
        auswahl = st.selectbox("🔍 Eintrag auswählen", df_entity[anzeige_spalte])
        zeile = df_entity[df_entity[anzeige_spalte] == auswahl]
        if zeile.empty:
            st.warning(" Kein gültiger Eintrag ausgewählt.")
            return

        ziel_id = zeile.iloc[0][id_col]

        if st.button(f" '{auswahl}' löschen und alle zugehörigen Einträge entfernen"):
            
            writer = pd.ExcelWriter(excel_file_path, engine="openpyxl", mode="a", if_sheet_exists="replace")

            for sheet in tabellen:
                try:
                    sheet_df = pd.read_excel(excel_file_path, sheet_name=sheet)

                    # Wenn die ID-Spalte vorhanden ist → filtern
                    if id_col in sheet_df.columns:
                        sheet_df = sheet_df[sheet_df[id_col] != ziel_id]
                    writer.book.remove(writer.book[sheet])  
                    sheet_df.to_excel(writer, sheet_name=sheet, index=False)
                except Exception as inner:
                    pass  

            writer.close()
            st.success(f" Eintrag '{auswahl}' und alle verknüpften Daten mit `{id_col} = {ziel_id}` wurden entfernt.")
    except PermissionError:
                st.error(" Zugriff verweigert: Die Datei ist derzeit geöffnet. Bitte schließe sie in Excel und versuche es erneut.")
    except Exception as e:
        st.error(f" Fehler: {e}")

delete_entity_and_cascade("/Matching/Datenmodell.xlsx")