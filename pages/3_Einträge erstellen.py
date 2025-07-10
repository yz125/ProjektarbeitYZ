import streamlit as st
import pandas as pd

st.set_page_config(
    layout="wide",           # Nutzt die volle Breite der Seite
    page_title="Einträge erstellen",  # Optional: Titel des Browser-Tabs
)
# Lade die Excel-Tabelle
DATA_FILE = "/Matching/Datenmodell.xlsx"

def load_data():
    xls = pd.ExcelFile(DATA_FILE)
    return {sheet: xls.parse(sheet) for sheet in xls.sheet_names}

def save_data(sheet_name, df):
    # Bestehende Daten aus Datei laden
    existing_data = load_data()

    # Sheet aktualisieren
    existing_data[sheet_name] = df

    # Alles neu schreiben (jedes Sheet nacheinander)
    with pd.ExcelWriter(DATA_FILE, engine='openpyxl', mode='w') as writer:
        for sheet, data in existing_data.items():
            data.to_excel(writer, sheet_name=sheet, index=False)



st.title(" Datenmodellverwaltung")

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["Neues Regelwerk", "Stakeholder hinzufügen", "Datenpunkte hochladen", "Standorte hinzufügen","Weitere Standards verknüpfen", "Kennzahl zuordnen"])
def eintraege():
    data = load_data()
    # TAB 1: Neues Regelwerk
    with tab1:
        st.header(" Regelwerk erstellen")

        # Lade bestehende Regelwerke aus Excel
        regelwerke_df = data.get("Regelwerk", pd.DataFrame())
        regelwerke_df.columns = regelwerke_df.columns.str.strip()  

        # Warnung, wenn falsche Spalten vorhanden sind
        expected_cols = {"Regelwerk-ID", "Name"}
        if not expected_cols.issubset(set(regelwerke_df.columns)):
            st.warning(f" Unerwartete Spalten in der Regelwerk-Tabelle: {regelwerke_df.columns.tolist()}")

        # Eingabe für neues Regelwerk
        name = st.text_input("Name des neuen Regelwerks")

        if st.button(" Regelwerk speichern"):
            if not name.strip():
                st.error(" Bitte gib einen Namen für das Regelwerk ein.")
            else:
                next_id = regelwerke_df["Regelwerk-ID"].max() + 1 if not regelwerke_df.empty else 1

                # Neue Zeile mit exakten Spaltennamen
                new_row = pd.DataFrame([[next_id, name.strip()]], columns=["Regelwerk-ID", "Name"])

                # Anhängen an bestehende Regelwerke
                regelwerke_df = pd.concat([regelwerke_df, new_row], ignore_index=True)

                # Speichern 
                save_data("Regelwerk", regelwerke_df)

                st.success(f" Regelwerk '{name}' wurde gespeichert.")
                #  Reload nach Speichern
                regelwerke_df = pd.read_excel("/Matching/Datenmodell.xlsx", sheet_name="Regelwerk")
        # Anzeige bestehender Regelwerke
        st.subheader(" Bestehende Regelwerke")
        if not regelwerke_df.empty:
            st.dataframe(regelwerke_df[["Regelwerk-ID", "Name"]])
        else:
            st.info("Es sind noch keine Regelwerke vorhanden.")

    # TAB 2: Stakeholder hinzufügen
    with tab2:
        st.header(" Stakeholder hinzufügen")

        # Stakeholder-Tabelle laden
        stakeholders = data.get("Stakeholder", pd.DataFrame())
        stakeholders.columns = stakeholders.columns.str.strip()  

        # Neue ID berechnen
        new_id = stakeholders["Stakeholder-ID"].max() + 1 if not stakeholders.empty else 1

        # Eingabefelder
        name = st.text_input("Name des Stakeholders")
        branche = st.text_input("Branche")

        if st.button(" Stakeholder speichern"):
            if not name.strip():
                st.error("Bitte gib einen Namen für den Stakeholder ein.")
            else:
                new_entry = pd.DataFrame([{
                    "Stakeholder-ID": new_id,
                    "Name": name.strip(),
                    "Branche": branche.strip()
                }])
                updated_df = pd.concat([stakeholders, new_entry], ignore_index=True)
                save_data("Stakeholder", updated_df)
                st.success(f" Stakeholder '{name}' gespeichert mit ID {new_id}.")
                #  Neu einlesen
                stakeholders = pd.read_excel("/Matching/Datenmodell.xlsx", sheet_name="Stakeholder")
        # Bestehende Stakeholder anzeigen + Suchfeld
        st.subheader(" Bestehende Stakeholder")

        if stakeholders.empty:
            st.info(" Noch keine Stakeholder vorhanden.")
        else:
            search = st.text_input(" Suche nach Name")

            filtered_df = stakeholders[
                stakeholders["Name"].astype(str).str.contains(search, case=False, na=False)
            ] if search else stakeholders

            st.dataframe(filtered_df[["Stakeholder-ID", "Name", "Branche"]])

    # TAB 3: Datenpunkte und Paragrafen aus Excel
    with tab3:
        st.header(" Datenpunkte und Paragrafen hochladen")

        # Daten aus dem Storage laden
        regelwerke = data["Regelwerk"]
        stakeholders = data["Stakeholder"]
        datenpunkte_df = data["Datenpunkt"]
        paragraf_df = data["Paragraf"]
        rw_dp_df = data["Regelwerk-Datenpunkt"]
        sh_dp_df = data["Stakeholder-Datenpunkt"]

        # Sicherstellen, dass Spalten konsistent benannt sind
        datenpunkte_df.columns = datenpunkte_df.columns.str.strip()
        if "Namen" in datenpunkte_df.columns:
            datenpunkte_df = datenpunkte_df.rename(columns={"Namen": "Name"})

        if regelwerke.empty or stakeholders.empty:
            st.warning(" Bitte stelle sicher, dass mindestens ein Regelwerk und ein Stakeholder existieren, bevor Datenpunkte hochgeladen werden.")
        else:
            regelwerk_name = st.selectbox("Regelwerk auswählen", regelwerke["Name"])
            regelwerk_id = regelwerke[regelwerke["Name"] == regelwerk_name]["Regelwerk-ID"].values[0]

            if stakeholders.empty:
                st.warning(" Es sind keine Stakeholder vorhanden. Bitte füge zuerst Stakeholder hinzu.")
                
            else:
                stakeholder_name = st.selectbox("Stakeholder auswählen", stakeholders["Name"], key="stakeholder_dp_upload")
                selected_stakeholder = stakeholders[stakeholders["Name"] == stakeholder_name]

            if selected_stakeholder.empty:
                st.error(" Der ausgewählte Stakeholder konnte nicht gefunden werden.")
                
            else:
                stakeholder_id = selected_stakeholder["Stakeholder-ID"].values[0]

            uploaded = st.file_uploader("Excel-Datei mit Datenpunkten hochladen", type=["xlsx"])
            if uploaded:
                try:
                    df = pd.read_excel(uploaded)

                    if df.empty or len(df.columns) < 5:
                        st.error(" Die hochgeladene Datei scheint leer zu sein oder enthält zu wenige Spalten.")
                    else:
                        st.markdown("###  Spaltenzuordnung")
                        st.write("Ordne die Spalten der hochgeladenen Datei den Attributen zu")
                        col_names = df.columns.tolist()
                        col_dp_name = st.selectbox("Spalte für Datenpunkt-Name", col_names)
                        col_datentyp = st.selectbox("Spalte für Datentyp", col_names)
                        col_gruppe = st.selectbox("Spalte für Gruppe", col_names)
                        col_standard = st.selectbox("Spalte für Standard", col_names)
                        col_paragraf = st.selectbox("Spalte für Paragraf", col_names)

                        all_mapped = all([col_dp_name, col_datentyp, col_gruppe, col_standard, col_paragraf])
                        if not all_mapped:
                            st.error(" Bitte alle Felder für die Zuordnung der Spalten korrekt auswählen.")
                        elif st.button("Datenpunkte und Paragrafen importieren"):
                            next_dp_id = datenpunkte_df["Datenpunkt-ID"].max() + 1 if not datenpunkte_df.empty else 1
                            next_pg_id = paragraf_df["Paragraf-ID"].max() + 1 if not paragraf_df.empty else 1

                            dp_list, pg_list, rw_dp_list, sh_dp_list = [], [], [], []

                            for _, row in df.iterrows():
                                dp_id = next_dp_id
                                next_dp_id += 1

                                dp_list.append({
                                    "Datenpunkt-ID": dp_id,
                                    "Name": row[col_dp_name],  # Wichtig: "Name", nicht "Namen"
                                    "Datentyp": row[col_datentyp],
                                    "Gruppe": row[col_gruppe]
                                })

                                pg_list.append({
                                    "Paragraf-ID": next_pg_id,
                                    "Regelwerk-ID": regelwerk_id,
                                    "Datenpunkt-ID": dp_id,
                                    "Standard": row[col_standard],
                                    "Paragraf": row[col_paragraf]
                                })
                                next_pg_id += 1

                                rw_dp_list.append({"Regelwerk-ID": regelwerk_id, "Datenpunkt-ID": dp_id})
                                sh_dp_list.append({"Stakeholder-ID": stakeholder_id, "Datenpunkt-ID": dp_id})

                            # Tabellen aktualisieren
                            updated_dp = pd.concat([datenpunkte_df, pd.DataFrame(dp_list)], ignore_index=True)
                            updated_pg = pd.concat([paragraf_df, pd.DataFrame(pg_list)], ignore_index=True)
                            # Neue DataFrames
                            new_rw_dp_df = pd.DataFrame(rw_dp_list)
                            new_sh_dp_df = pd.DataFrame(sh_dp_list)

                            # Nur neue Regelwerk-Datenpunkt-Kombinationen, die noch nicht existieren
                            merged_rw = new_rw_dp_df.merge(rw_dp_df, on=["Regelwerk-ID", "Datenpunkt-ID"], how="left", indicator=True)
                            new_rw_only = merged_rw[merged_rw["_merge"] == "left_only"].drop(columns=["_merge"])
                            updated_rw_dp = pd.concat([rw_dp_df, new_rw_only], ignore_index=True)

                            # Nur neue Stakeholder-Datenpunkt-Kombinationen, die noch nicht existieren
                            merged_sh = new_sh_dp_df.merge(sh_dp_df, on=["Stakeholder-ID", "Datenpunkt-ID"], how="left", indicator=True)
                            new_sh_only = merged_sh[merged_sh["_merge"] == "left_only"].drop(columns=["_merge"])
                            updated_sh_dp = pd.concat([sh_dp_df, new_sh_only], ignore_index=True)

                            save_data("Datenpunkt", updated_dp)
                            save_data("Paragraf", updated_pg)
                            save_data("Regelwerk-Datenpunkt", updated_rw_dp)
                            save_data("Stakeholder-Datenpunkt", updated_sh_dp)

                            st.success(" Datenpunkte und Paragrafen erfolgreich importiert.")
                except Exception as e:
                    st.error(f" Fehler beim Verarbeiten der Datei: {e}")

    # TAB 4: Standort hinzufügen
    with tab4:
        st.header(" Standort hinzufügen")

        # Aktuelle Daten laden
        current_data = load_data()
        stakeholders = current_data["Stakeholder"]
        standorte = current_data["Standort"]

        if stakeholders.empty or selected_stakeholder.empty:
            st.warning("Es sind keine Stakeholder vorhanden. Bitte füge zuerst Stakeholder hinzu.")
        else:
            if stakeholders.empty:
                st.warning(" Es sind keine Stakeholder vorhanden. Bitte füge zuerst Stakeholder hinzu.")
                return

            stakeholder_name = st.selectbox("Stakeholder auswählen", stakeholders["Name"], key="stakeholder_standort")
            selected_stakeholder = stakeholders[stakeholders["Name"] == stakeholder_name]

            if selected_stakeholder.empty:
                st.error(" Der ausgewählte Stakeholder konnte nicht gefunden werden.")
                return

            stakeholder_id = selected_stakeholder["Stakeholder-ID"].values[0]

            # Eingabefelder
            plz = st.text_input("Postleitzahl")
            hausnummer = st.text_input("Hausnummer")
            strasse = st.text_input("Straße")
            land = st.text_input("Land")

            if st.button("Standort speichern"):
                #  Daten neu laden, um aktuelle ID zu berechnen
                current_data = load_data()
                standorte = current_data["Standort"]

                new_id = standorte["Standort-ID"].max() + 1 if not standorte.empty else 1

                new_entry = pd.DataFrame([{
                    "Standort-ID": new_id,
                    "Postleitzahl": plz,
                    "Hausnummer": hausnummer,
                    "Straße": strasse,
                    "Land": land,
                    "Stakeholder-ID": stakeholder_id
                }])

                updated_df = pd.concat([standorte, new_entry], ignore_index=True)
                save_data("Standort", updated_df)

                st.success(f" Standort erfolgreich gespeichert mit ID {new_id}.")
                st.rerun()
                # Daten erneut laden
                current_data = load_data()
                standorte = current_data["Standort"]

            #  Standortübersicht mit Suche
            st.markdown("###  Aktuelle Standorte")

            search = st.text_input(" Standortsuche (PLZ, Straße, Land, etc.)")
            filtered = standorte.copy()

            if search:
                search_lower = search.lower()
                filtered = standorte[standorte.apply(lambda row: search_lower in str(row).lower(), axis=1)]

            st.dataframe(filtered, use_container_width=True)
    # TAB 5: Weitere Standards verknüpfen (flexibles Mapping)

    with tab5:
        st.header(" Paragrafen aus anderem Regelwerk übernehmen (flexibles Mapping)")

        data = load_data()
        regelwerke = data.get("Regelwerk", pd.DataFrame())
        paragraf_df = data.get("Paragraf", pd.DataFrame())
        rw_dp_df = data.get("Regelwerk-Datenpunkt", pd.DataFrame())

        if regelwerke.empty:
            st.warning(" Es sind keine Regelwerke vorhanden.")
        else:

            ziel_regelwerk_name = st.selectbox("Ziel-Regelwerk auswählen", regelwerke["Name"])
            ziel_regelwerk_id = regelwerke[regelwerke["Name"] == ziel_regelwerk_name]["Regelwerk-ID"].values[0]

            uploaded_file = st.file_uploader("Excel-Datei mit Paragrafen-Mapping hochladen", type=["xlsx"], key="mapping_upload")

            if uploaded_file:
                try:
                    df = pd.read_excel(uploaded_file)

                    if len(df.columns) < 3:
                        st.error(" Die Datei muss mindestens drei Spalten enthalten.")
                        return

                    st.subheader(" Spalten zuordnen")
                    spalten = list(df.columns)

                    mapping_basis = st.radio(" Mapping über:", ["Paragraf-ID", "Paragraf-Text"])
                    st.write("Wähle Mapping über ID nur, wenn die Paragrafen ID des Datenmodells verwendet wird!")
                    alte_pg_spalte = st.selectbox(f"Spalte mit dem alten Paragrafen ({mapping_basis})", spalten)
                    neue_standard_spalte = st.selectbox("Spalte mit dem **neuen Standardnamen**", spalten)
                    neuer_pg_text_spalte = st.selectbox("Spalte mit dem **neuen Paragraftext**", spalten)

                    if st.button(" Paragrafen übernehmen und verknüpfen"):
                        st.session_state["run_mapping"] = True

                    if st.session_state.get("run_mapping"):
                        next_pg_id = paragraf_df["Paragraf-ID"].max() + 1 if not paragraf_df.empty else 1
                        new_paragraphs = []
                        new_rw_dp_links = []

                        for _, row in df.iterrows():
                            referenzwert = row[alte_pg_spalte]

                            if mapping_basis == "Paragraf-ID":
                                try:
                                    referenzwert = int(str(referenzwert).strip())
                                    matching_pg = paragraf_df[paragraf_df["Paragraf-ID"] == referenzwert]
                                except:
                                    st.warning(f" Ungültiger Wert in '{alte_pg_spalte}': {referenzwert}")
                                    continue
                            else:  
                                referenzwert = str(referenzwert).strip()
                                matching_pg = paragraf_df[paragraf_df["Paragraf"].str.strip() == referenzwert]

                            if matching_pg.empty:
                                st.warning(f" Kein passender Paragraf gefunden für: {referenzwert}")
                                continue

                            datenpunkt_id = matching_pg.iloc[0]["Datenpunkt-ID"]

                            new_paragraphs.append({
                                "Paragraf-ID": next_pg_id,
                                "Regelwerk-ID": ziel_regelwerk_id,
                                "Datenpunkt-ID": datenpunkt_id,
                                "Standard": row[neue_standard_spalte],
                                "Paragraf": row[neuer_pg_text_spalte]
                            })

                            exists = rw_dp_df[
                                (rw_dp_df["Regelwerk-ID"] == ziel_regelwerk_id) &
                                (rw_dp_df["Datenpunkt-ID"] == datenpunkt_id)
                            ]
                            if exists.empty:
                                new_rw_dp_links.append({
                                    "Regelwerk-ID": ziel_regelwerk_id,
                                    "Datenpunkt-ID": datenpunkt_id
                                })

                            next_pg_id += 1

                        if new_paragraphs:
                            paragraf_df_updated = pd.concat([paragraf_df, pd.DataFrame(new_paragraphs)], ignore_index=True)
                            save_data("Paragraf", paragraf_df_updated)
                            st.success(f" {len(new_paragraphs)} neue Paragrafen hinzugefügt.")

                        if new_rw_dp_links:
                            rw_dp_df_updated = pd.concat([rw_dp_df, pd.DataFrame(new_rw_dp_links)], ignore_index=True)
                            save_data("Regelwerk-Datenpunkt", rw_dp_df_updated)
                            st.success(f" {len(new_rw_dp_links)} neue Regelwerk-Datenpunkt-Verknüpfungen gespeichert.")

                        if not new_paragraphs and not new_rw_dp_links:
                            st.info(" Es wurden keine neuen Einträge erkannt.")

                        st.session_state["run_mapping"] = False

                except Exception as e:
                    st.error(f" Fehler beim Verarbeiten der Datei: {e}")
    with tab6:
        st.header(" Kennzahl zu Datenpunkt hinzufügen")

        data = load_data()
        datenpunkt_df = data.get("Datenpunkt", pd.DataFrame())
        kennzahl_df = data.get("Kennzahl", pd.DataFrame())
        stakeholder_df = data.get("Stakeholder", pd.DataFrame())

        if datenpunkt_df.empty:
            st.warning(" Es sind keine Datenpunkte vorhanden.")
            return

        if stakeholder_df.empty:
            st.warning(" Es sind keine Stakeholder vorhanden.")
            return

        # DATENPUNKT SUCHEN
        st.subheader(" Datenpunkt auswählen")
        dp_suchbegriff = st.text_input("Datenpunktname suchen")
        gefundene_dp = datenpunkt_df[datenpunkt_df["Name"].astype(str).str.contains(dp_suchbegriff, case=False, na=False)] if dp_suchbegriff else datenpunkt_df

        if not gefundene_dp.empty:
            dp_name = st.selectbox("Gefundene Datenpunkte:", gefundene_dp["Name"])
            ausgewählter_dp = datenpunkt_df[datenpunkt_df["Name"] == dp_name].iloc[0]

            st.markdown("#### Ausgewählter Datenpunkt:")
            st.write(ausgewählter_dp)

            # KENNZAHL-EINGABE
            st.markdown("---")
            st.subheader(" Neue Kennzahl erfassen")

            wert = st.text_input("Wert der Kennzahl (Zahl oder Text)")
            jahr = st.number_input("Jahr", min_value=1900, max_value=2100, step=1, value=2024)
            quelle = st.text_input("Quelle der Kennzahl")

            st.markdown("#### Stakeholder auswählen")
            stakeholder_suchbegriff = st.text_input("Stakeholdername suchen")
            gefundene_stk = stakeholder_df[stakeholder_df["Name"].astype(str).str.contains(stakeholder_suchbegriff, case=False, na=False)] if stakeholder_suchbegriff else stakeholder_df

            if not gefundene_stk.empty:
                stk_name = st.selectbox("Gefundene Stakeholder:", gefundene_stk["Name"])
                stakeholder_id = stakeholder_df[stakeholder_df["Name"] == stk_name]["Stakeholder-ID"].values[0]

                if st.button(" Kennzahl hinzufügen"):
                    neue_id = kennzahl_df["Kennzahl-ID"].max() + 1 if not kennzahl_df.empty else 1

                    neue_kennzahl = {
                        "Kennzahl-ID": neue_id,
                        "Datenpunkt-ID": ausgewählter_dp["Datenpunkt-ID"],
                        "Wert": wert,
                        "Jahr": jahr,
                        "Stakeholder-ID": stakeholder_id,
                        "Quelle": quelle
                    }

                    try:
                        kennzahl_df_updated = pd.concat([kennzahl_df, pd.DataFrame([neue_kennzahl])], ignore_index=True)
                        save_data("Kennzahl", kennzahl_df_updated)
                        st.success(" Kennzahl erfolgreich hinzugefügt.")
                    except PermissionError:
                        st.error(" Fehler beim Speichern. Bitte schließen Sie die Excel-Datei und versuchen Sie es erneut.")
            else:
                st.info(" Bitte geben Sie einen gültigen Stakeholdernamen ein.")
        else:
            st.info(" Bitte geben Sie einen Datenpunktnamen ein.")
eintraege()