from google.oauth2 import service_account
import streamlit as st
import gspread
from gspread.exceptions import SpreadsheetNotFound, WorksheetNotFound
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from datetime import datetime
import builtins
import os
import pandas as pd
import traceback
import time


# ==== Seiteneinstellungen ====
st.set_page_config(page_title="Verkaufszahlen", layout="wide")
        
# ==== Globales Styling ====
st.markdown("""
    <style>
    .stTextInput > div > input {
        padding: 10px;
        font-size: 16px;
    }
    .stButton > button {
        padding: 10px 20px;
        font-size: 16px;
        border-radius: 6px;
    }
    .produkt-zeile:hover {
        background-color: #f5f5f5;
    }
    input[type="number"], input[type="text"] {
        height: 40px !important;
        font-size: 16px !important;
    }
    div[data-testid="stTextInput"] input {
        margin-top: -4px !important;
    }
    textarea {
        font-size: 16px !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==== Dateipfade ====
DATEN_ORDNER = "daten"
PRODUKTE_DATEI = os.path.join(DATEN_ORDNER, "App_Produkte_mit_EAN.xlsx")
HAENDLER_DATEI = os.path.join(DATEN_ORDNER, "App_Haendler_mit_Passwoertern.xlsx")

# ==== Google Drive Setup ====
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds = service_account.Credentials.from_service_account_info(
    st.secrets["google_service_account"],
    scopes=SCOPES
)

gclient = gspread.authorize(creds)
# Google Drive Zielordner und Freigabe-Einstellungen
FOLDER_ID = "1oh4nWRqYE2k4K-EUts93cU7pKjatbfB4"
SHARE_EMAIL = None  # z. B. "your.email@domain.com" oder None, falls keine automatische Freigabe gewünscht

# ==== Dateiüberwachung ====
def get_modified_time(path):
    try:
        return os.path.getmtime(path)
    except FileNotFoundError:
        return 0

# ==== Ladefunktionen ====
def lade_produkte():
    if not os.path.exists(PRODUKTE_DATEI):
        st.warning(f"⚠️ Produktdatei nicht gefunden: {PRODUKTE_DATEI}")
        return pd.DataFrame()
    modified_time = get_modified_time(PRODUKTE_DATEI)
    if "produkte_modified_time" not in st.session_state or st.session_state["produkte_modified_time"] != modified_time:
        df = pd.read_excel(PRODUKTE_DATEI)
        df.columns = df.columns.str.strip().str.lower()
        df = df[df["aktiv"].fillna("").str.lower() == "x"]
        st.session_state["produkte"] = df
        st.session_state["produkte_modified_time"] = modified_time
        st.session_state["produkte_ladezeit"] = datetime.now().strftime("%H:%M:%S")
    return st.session_state.get("produkte", pd.DataFrame())

def lade_haendler():
    if not os.path.exists(HAENDLER_DATEI):
        st.warning(f"⚠️ Händlerdatei nicht gefunden: {HAENDLER_DATEI}")
        return pd.DataFrame()
    modified_time = get_modified_time(HAENDLER_DATEI)
    if "haendler_modified_time" not in st.session_state or st.session_state["haendler_modified_time"] != modified_time:
        df = pd.read_excel(HAENDLER_DATEI)
        df.columns = df.columns.str.strip()
        st.session_state["haendler"] = df
        st.session_state["haendler_modified_time"] = modified_time
        st.session_state["haendler_ladezeit"] = datetime.now().strftime("%H:%M:%S")
    return st.session_state.get("haendler", pd.DataFrame())
#google drive upload
def google_drive_upload(modus, haendler_name, eintrag,
                        filename="verkaufsdaten.xlsx",
                        folder_id="1oh4nWRqYE2k4K-EUts93cU7pKjatbfB4"):
    import pandas as pd
    import io
    from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
    import time

    try:
        temp_dir = "temp_excel"
        os.makedirs(temp_dir, exist_ok=True)
        temp_path = os.path.join(temp_dir, filename)

        # Verbindung zur Drive API
        drive_service = build("drive", "v3", credentials=creds)

        # Suche nach Datei im Zielordner
        query = f"name='{filename}' and '{folder_id}' in parents and trashed=false"
        result = drive_service.files().list(q=query, fields="files(id, name)").execute()
        files = result.get("files", [])
        file_id = files[0]["id"] if files else None

        # Händlerdaten
        haendler_info = st.session_state.get("haendler_info", {})
        login_nr = haendler_info.get("Login-Nr.", haendler_info.get("Login", "unbekannt"))

        # Einheitliche Spaltenstruktur inkl. Cashback-Betrag
        from datetime import datetime

        daten = []
        zeitstempel = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for produkt in eintrag.get("eintraege", []):
            daten.append([
                zeitstempel,  # <--- NEU: Zeitstempel
                eintrag.get("datum", datetime.today().strftime("%Y-%m-%d")),
                eintrag.get("kw", ""),
                modus,
                haendler_name,
                login_nr,
                produkt.get("Produktname", ""),
                produkt.get("EAN", ""),
                produkt.get("Menge", ""),
                produkt.get("Preis", ""),
                produkt.get("Seriennummer", ""),
                eintrag.get("cashback_typ", ""),
                produkt.get("Cashback-Betrag", ""),
                produkt.get("Soundbar", ""),
                produkt.get("Seriennummer_SB", ""),
                eintrag.get("kommentar", "")
            ])


        new_df = pd.DataFrame(daten, columns=[
            "Zeitstempel","Datum", "KW", "Modus", "Händler", "Login-Nr.",
            "Produkt", "EAN", "Menge", "Preis", "Seriennummer",
            "Cashback-Typ", "Cashback-Betrag", "Soundbar", "Seriennummer Soundbar", "Kommentar"
        ])

        if file_id:
            # Datei herunterladen
            request = drive_service.files().get_media(fileId=file_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
            fh.seek(0)
            existing_df = pd.read_excel(fh)
            full_df = pd.concat([existing_df, new_df], ignore_index=True)
            st.info("📥 Bestehende Datei geladen und ergänzt")
        else:
            full_df = new_df
            st.info("📄 Neue zentrale Datei wird erstellt")

        # Excel lokal speichern
        with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
            full_df.to_excel(writer, index=False)

        time.sleep(0.5)  # Windows-Dateisperre umgehen

        # Upload zu Google Drive
        media = MediaFileUpload(temp_path, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if file_id:
            drive_service.files().update(fileId=file_id, media_body=media).execute()
            st.success("✅ Zentrale Datei auf Google Drive ergänzt")
        else:
            file_metadata = {"name": filename, "parents": [folder_id]}
            uploaded = drive_service.files().create(
                body=file_metadata,
                media_body=media,
                fields="id, webViewLink"
            ).execute()
            file_id = uploaded.get("id")
            st.success("✅ Neue zentrale Datei auf Google Drive erstellt")

        file_link = f"https://drive.google.com/file/d/{file_id}/view"
        st.markdown(f"🔗 [Zur Datei öffnen]({file_link})")

        # Lokale Datei löschen
        for i in range(5):
            try:
                os.remove(temp_path)
                st.info("🧹 Lokale Datei gelöscht")
                break
            except PermissionError:
                time.sleep(0.5)
                if i == 4:
                    st.warning("⚠️ Datei konnte nicht gelöscht werden. Bitte manuell entfernen.")

        return True

    except Exception as e:
        st.error("❌ Fehler beim Upload in zentrale Datei")
        traceback.print_exc()
        st.exception(e)
        return False



# ==== Produktsuche ====
def suche_produkte(df, suchbegriff):
    if not suchbegriff:
        return df
    suchbegriff = suchbegriff.lower()
    return df[df["produktname"].str.lower().str.contains(suchbegriff) | df["ean"].astype(str).str.contains(suchbegriff)]

# ==== Produktzeile anzeigen ====
def zeige_produktzeile(row, modus, index, reset):
    ean = str(row["ean"])
    produktname = row["produktname"]
    menge_key = f"menge_{ean}_{modus}_{index}"
    preis_key = f"preis_{ean}_{modus}_{index}"

    if reset:
        st.session_state[menge_key] = 0
        if modus in ["projekt", "bestellung"]:
            st.session_state[preis_key] = ""

    # Kompaktere Darstellung in einer Zeile mit 3–4 Spalten
    cols = st.columns([1, 1.5, 1.5, 1.5, 2]) if modus in ["verkauf"] else st.columns([1, 1.5, 1.5, 2.5, 2])

    # Produktinfo links (Name + EAN)
    with cols[0]:
        st.markdown(f"**{produktname}**  \nEAN: `{ean}`", unsafe_allow_html=True)

    # Menge
    with cols[1]:
        menge = st.number_input("", min_value=0, step=1, key=menge_key, label_visibility="collapsed")

    # Preis (nur bei projekt/bestellung)
    preis = None
    if modus in ["projekt", "bestellung"]:
        with cols[2]:
            preis = st.text_input(
                "", key=preis_key, label_visibility="collapsed",
                placeholder="Zielpreis exkl. MwSt" if modus == "projekt" else "Bestpreis exkl. MwSt"
            )

    return {"EAN": ean, "Produktname": produktname, "Menge": menge, "Preis": preis}

# ==== Formularansicht ====# ==== Formularansicht (mit Aufruf für Cashback) ====
def formular_ansicht(modus):
    # Wenn Cashback gewählt wurde: eigenes Formular und Sidebar-Verlauf anzeigen
    if modus == "cashback":
        zeige_cashback_formular()
        return

    # ==== Normale Ansicht für Verkaufszahlen, Projekt, Bestellung ====
    produkte_df = lade_produkte()
    if produkte_df.empty:
        return

    # Gruppierung vorbereiten
    gruppen = list(produkte_df["gruppe"].dropna().unique()) if "gruppe" in produkte_df.columns else []

    # Such-/Gruppenlogik
    def gruppe_geaendert():
        if st.session_state.get("suchbegriff"):
            st.session_state["suchbegriff"] = ""
        st.session_state["skip_suchlogik"] = True
        st.query_params.update({"refresh": "1"})

    if st.session_state.get("trigger_reset_filter"):
        st.session_state["suchbegriff"] = ""
        st.session_state["gruppe_filter"] = "Alle"
        st.session_state.pop("trigger_reset_filter")

    if "previous_gruppe_filter" not in st.session_state:
        st.session_state["previous_gruppe_filter"] = st.session_state.get("gruppe_filter", "Alle")
    if "previous_suchbegriff" not in st.session_state:
        st.session_state["previous_suchbegriff"] = st.session_state.get("suchbegriff", "")

    if (
        not st.session_state.get("skip_suchlogik")
        and st.session_state.get("suchbegriff") != st.session_state.get("previous_suchbegriff")
        and st.session_state.get("gruppe_filter") != "Alle"
    ):
        neuer_suchbegriff = st.session_state.get("suchbegriff", "")
        st.session_state["gruppe_filter"] = "Alle"
        st.session_state["suchbegriff"] = neuer_suchbegriff
        st.session_state["previous_suchbegriff"] = neuer_suchbegriff
        st.rerun()

    if st.session_state.get("skip_suchlogik"):
        st.session_state.pop("skip_suchlogik")

    col1, col2 = st.columns([2, 2])
    with col1:
        suchbegriff = st.text_input("🔍 Produkt suchen (Name oder EAN)", key="suchbegriff")
    with col2:
        gruppe = st.selectbox(
            "Gruppe/Kategorie filtern",
            options=["Alle"] + sorted(gruppen),
            key="gruppe_filter",
            on_change=gruppe_geaendert
        )

    st.session_state["previous_gruppe_filter"] = st.session_state["gruppe_filter"]
    st.session_state["previous_suchbegriff"] = st.session_state["suchbegriff"]

    col_reset, col_absenden = st.columns([3, 3])
    with col_reset:
        if st.button("🔁 Filter zurücksetzen", key="btn_filter_zurueck"):
            st.session_state["trigger_reset_filter"] = True
            st.rerun()
    with col_absenden:
        if st.button("Absenden", key="btn_absenden"):
            st.session_state["senden_geklickt"] = True

    produkte_df = suche_produkte(produkte_df, st.session_state["suchbegriff"])
    if "gruppe" in produkte_df.columns and st.session_state["gruppe_filter"] != "Alle":
        produkte_df = produkte_df[produkte_df["gruppe"] == st.session_state["gruppe_filter"]]

    reset = st.session_state.pop("reset_felder", False)
    erfolg_modus = st.session_state.pop("zeige_bestaetigung", None)
    if erfolg_modus == modus:
        meldung = {
            "verkauf": ("✅ Verkaufszahlen gespeichert", "📊"),
            "projekt": ("✅ Projektanfrage übermittelt", "📁"),
            "bestellung": ("✅ Bestellung übermittelt", "💯")
        }.get(modus, ("✅ Daten gespeichert", "✅"))
        st.markdown(f"<div style='background-color:#e6f4ea;border-left:5px solid #34a853;padding:12px 16px;border-radius:6px;margin-top:1rem;margin-bottom:1rem;'><span style='font-size:18px;'>{meldung[1]} <strong>{meldung[0]}</strong></span></div>", unsafe_allow_html=True)

    eintraege = []
    for i, row in produkte_df.iterrows():
        with st.container():
            st.markdown("<div style='margin-bottom: 0.1rem;'>", unsafe_allow_html=True)
            eintrag = zeige_produktzeile(row, modus, i, reset)
            if eintrag["Menge"] > 0 or (eintrag["Preis"] and str(eintrag["Preis"]).strip()):
                eintraege.append(eintrag)

    kommentar = ""
    kommentar_key = f"kommentar_{modus}"
    if modus in ["projekt", "bestellung"]:
        if reset:
            st.session_state[kommentar_key] = ""
        kommentar = st.text_area("📬 Was möchten Sie uns mitteilen?", placeholder="Z. B. Projektinformationen, Lieferwünsche ...", key=kommentar_key)

    if st.session_state.get("senden_geklickt"):
        st.session_state["senden_geklickt"] = False
        if not eintraege:
            st.warning("Bitte mindestens eine Menge oder einen Preis angeben.")
        else:
            neuer_eintrag = {
                "modus": modus,
                "uhrzeit": datetime.today().strftime("%H:%M:%S"),
                "datum": datetime.today().strftime("%Y-%m-%d"),
                "kw": datetime.today().isocalendar()[1],
                "eintraege": eintraege,
                "kommentar": kommentar,
            }
            haendler = st.session_state.get("haendler_info", {})
            haendlername = haendler.get("Firmenname", "unbekannt")
            success = google_drive_upload(modus, haendlername, neuer_eintrag)
            if success:
                historie_key = f"historie_{modus}"
                if historie_key not in st.session_state:
                    st.session_state[historie_key] = []
                st.session_state[historie_key].insert(0, neuer_eintrag)
                st.session_state["reset_felder"] = True
                st.session_state["zeige_bestaetigung"] = modus
                st.rerun()
            else:
                st.error("❌ Upload zu Google Drive fehlgeschlagen. Bitte erneut versuchen.")

    # ✅ Sidebar anzeigen
    zeige_sidebar_verlauf(modus)


    historie_key = f"historie_{modus}"
def zeige_sidebar_verlauf(modus):
    with st.sidebar:
        haendler = st.session_state.get("haendler_info", {})
        firmenname = haendler.get("Firmenname", "unbekannt")
        login_nr = haendler.get("Login-Nr.", "unbekannt")

        st.markdown(f"""
            <div style="padding: 0.5rem; border-radius: 0.5rem; background-color: #f5f5f5; margin-bottom: 1rem;">
                <span style="color: #333; font-size: 0.9rem;">👤 Eingeloggt als:</span><br>
                <strong style="font-size: 1rem;">{firmenname}</strong>
            </div>
        """, unsafe_allow_html=True)

        if st.button("Logout", use_container_width=True, key=f"logout_{modus}"):
            st.session_state.clear()
            st.rerun()

        titel_map = {
            "verkauf": "Sell Out",
            "projekt": "Projektanfrage",
            "bestellung": "Bestellung",
            "cashback": "Cashback-Anfrage"
        }
        titel = titel_map.get(modus, "Meldung")
        historie_key = f"historie_{modus}"
        st.markdown(f"### Verlauf: {titel}")

        sort_option = st.selectbox("Sortierung", ["Neueste zuerst", "Älteste zuerst"], key=f"sort_{modus}_sidebar")
        if historie_key in st.session_state and st.session_state[historie_key]:
            historie_eintraege = st.session_state[historie_key]
            if sort_option == "Älteste zuerst":
                historie_eintraege = list(reversed(historie_eintraege))

            for h in historie_eintraege:
                zeitinfo = h.get("uhrzeit", "??:??")
                st.markdown(f"**{titel} vom {h['datum']} (KW {h['kw']}) um {zeitinfo}**")
                for e in h["eintraege"]:
                    details = f"- {e['Produktname']}"
                    if e.get("Menge"):
                        details += f": {e['Menge']} Stk"
                    if e.get("Preis"):
                        details += f", {e['Preis']} CHF"
                    if e.get("Seriennummer"):
                        details += f", SN: {e['Seriennummer']}"
                    if e.get("cashback_typ"):
                        details += f", {e['cashback_typ']} Cashback"
                    if e.get("Cashback-Betrag"):
                        details += f" ({e['Cashback-Betrag']} CHF)"
                    st.write(details)
                if h.get("kommentar"):
                    st.info(f"💬 {h['kommentar']}")
                st.markdown("---")

        try:
            from googleapiclient.http import MediaIoBaseDownload
            import io
            import tempfile

            drive_service = build("drive", "v3", credentials=creds)
            query = f"name='verkaufsdaten.xlsx' and '{FOLDER_ID}' in parents and trashed=false"
            result = drive_service.files().list(q=query, fields="files(id)").execute()
            files = result.get("files", [])
            if files:
                file_id = files[0]["id"]
                request = drive_service.files().get_media(fileId=file_id)
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                fh.seek(0)

                df = pd.read_excel(fh)
                df_eigen = df[df["Login-Nr."] == login_nr]

                if not df_eigen.empty:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        df_eigen.to_excel(tmp.name, index=False)
                        with open(tmp.name, "rb") as f:
                            st.download_button(
                                label="📥 Meine Historie herunterladen",
                                data=f,
                                file_name=f"historie_{login_nr}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
        except Exception as e:
            st.warning("📂 Historie konnte nicht geladen werden.")

def zeige_cashback_formular():
    import pandas as pd
    from datetime import datetime

    # 🔁 Reset bei Bedarf VOR der Widget-Erstellung
    if st.session_state.get("trigger_reset_cb"):
        if "suchbegriff_cb" in st.session_state:
            del st.session_state["suchbegriff_cb"]
        st.session_state["gruppe_cb"] = "Alle"
        st.session_state.pop("trigger_reset_cb")

    if st.session_state.get("trigger_reset_cb_suchbegriff"):
        if "suchbegriff_cb" in st.session_state:
            del st.session_state["suchbegriff_cb"]
        st.session_state.pop("trigger_reset_cb_suchbegriff")

    if st.session_state.get("reset_felder"):
        # Alles löschen, auch Single-Felder und ggf. Expander-Keys
        keys_to_delete = [k for k in st.session_state.keys()
                          if k.startswith("sn_")      # Seriennummern (Single & Double)
                          or k.startswith("cbtyp_")   # Auswahl Radio
                          or k.startswith("sb_")      # Soundbar Auswahl
                          or k.startswith("sbsn_")    # Soundbar Seriennummer
                          or k.startswith("exp_")]    # Expander-Zustand
        for k in keys_to_delete:
            del st.session_state[k]
        from uuid import uuid4
        st.session_state["reset_token"] = str(uuid4())  # Expander neu initialisieren
        st.session_state["reset_felder"] = False

    st.subheader("💳 Direct Cashback Promotion")

    df_produkte = pd.read_excel("daten/App_Produkte_mit_EAN.xlsx")

    cashback_produkte = df_produkte[
        (df_produkte["aktiv Cashback"].str.lower() == "x") &
        (
            df_produkte["Cashback single"].notna() |
            df_produkte["Cashback Double"].notna()
        )
    ]

    soundbars = df_produkte[
        (df_produkte["aktiv Cashback"].str.lower() == "x") &
        (df_produkte["Gruppe"].str.contains("BRAVIA THEATER", case=False, na=False))
    ]

    gruppen = sorted(cashback_produkte["Gruppe"].dropna().unique())

    # Callback
    def gruppe_cb_geaendert():
        st.session_state["trigger_reset_cb_suchbegriff"] = True

    # UI: Filter
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        suchbegriff = st.text_input("🔍 Produkt suchen (Name oder EAN)", key="suchbegriff_cb")
    with col2:
        gruppe = st.selectbox("Gruppe filtern", ["Alle"] + gruppen, key="gruppe_cb", on_change=gruppe_cb_geaendert)
    with col3:
        if st.button("🔁 Reset", key="reset_cb"):
            st.session_state["trigger_reset_cb"] = True
            st.rerun()

    # Filter anwenden
    filtered_df = cashback_produkte.copy()
    suchbegriff_val = st.session_state.get("suchbegriff_cb", "").strip().lower()
    gruppe_val = st.session_state.get("gruppe_cb", "Alle")

    if suchbegriff_val:
        filtered_df = filtered_df[
            filtered_df["Produktname"].str.lower().str.contains(suchbegriff_val) |
            filtered_df["EAN"].astype(str).str.contains(suchbegriff_val)
        ]
    elif gruppe_val != "Alle":
        filtered_df = filtered_df[filtered_df["Gruppe"] == gruppe_val]

    # Formular
    st.divider()
    senden = st.button("✅ Cashback-Einforderung absenden", use_container_width=True, key="submit_cb")

    eintraege = []
    for i, row in filtered_df.iterrows():
        produkt_key = f"cb_{i}"
        with st.expander(f"➕ {row['Produktname']} ({row['EAN']})", expanded=False):
            col1, col2 = st.columns([1, 1])
            seriennummer = col1.text_input("Seriennummer", key=f"sn_{i}")
            cashback_typ = col2.radio("Cashback-Typ", ["Single", "Double"], key=f"cbtyp_{i}", horizontal=True)
            cashback_betrag = row["Cashback single"] if cashback_typ == "Single" else row["Cashback Double"]
            st.info(f"💰 Cashback-Betrag: {cashback_betrag} CHF")

            soundbar = ""
            seriennummer_sb = ""

            if cashback_typ == "Double":
                col3, col4 = st.columns([1, 1])
                soundbar = col3.selectbox("Soundbar wählen", [""] + soundbars["Produktname"].tolist(), key=f"sb_{i}")
                seriennummer_sb = col4.text_input("Seriennummer der Soundbar", key=f"sbsn_{i}")

            if cashback_typ == "Single" and seriennummer.strip():
                eintraege.append({
                    "Produktname": row["Produktname"],
                    "EAN": row["EAN"],
                    "Seriennummer": seriennummer,
                    "cashback_typ": cashback_typ,
                    "Soundbar": "",
                    "Seriennummer_SB": "",
                    "Cashback-Betrag": cashback_betrag
                })
            elif cashback_typ == "Double" and seriennummer.strip() and seriennummer_sb.strip() and soundbar:
                eintraege.append({
                    "Produktname": row["Produktname"],
                    "EAN": row["EAN"],
                    "Seriennummer": seriennummer,
                    "cashback_typ": cashback_typ,
                    "Soundbar": soundbar,
                    "Seriennummer_SB": seriennummer_sb,
                    "Cashback-Betrag": cashback_betrag
                })

    kommentar = st.text_area("💬 Kommentar (optional)", key="kommentar_cb")

    if senden:
        if not eintraege:
            st.warning("Bitte mindestens eine Seriennummer eingeben.")
        else:
            eintrag = {
                "datum": datetime.today().strftime("%Y-%m-%d"),
                "uhrzeit": datetime.today().strftime("%H:%M"),
                "kw": datetime.today().isocalendar()[1],
                "eintraege": eintraege,
                "cashback_typ": "Double" if any(e["cashback_typ"] == "Double" for e in eintraege) else "Single",
                "kommentar": kommentar
            }
            haendler = st.session_state.get("haendler_info", {})
            haendlername = haendler.get("Firmenname", "unbekannt")
            erfolg = google_drive_upload("Cashback", haendlername, eintrag)

            if erfolg:
                historie_key = "historie_cashback"
                if historie_key not in st.session_state:
                    st.session_state[historie_key] = []
                st.session_state[historie_key].insert(0, eintrag)
                st.success("🎉 Cashback-Anfrage erfolgreich eingereicht!")
                st.session_state["reset_felder"] = True
                st.session_state["trigger_reset_cb"] = True  # <- wie beim echten Reset-Button!
                st.rerun()

            else:
                st.error("❌ Fehler beim Hochladen. Bitte später erneut versuchen.")

    zeige_sidebar_verlauf("cashback")





# ==== Login Funktion ====
def login():
    haendler_df = lade_haendler()
    login_nr = st.text_input("Login-Nr.", placeholder="Ihre Login-Nummer", label_visibility="visible")
    password = st.text_input("Passwort", type="password", placeholder="Ihr Passwort", label_visibility="visible")
    if st.button("Einloggen") and not haendler_df.empty:
        try:
            login_col = [c for c in haendler_df.columns if "login" in c.lower()][0]
            pw_col = [c for c in haendler_df.columns if "passwort" in c.lower()][0]
            user = haendler_df[
                (haendler_df[login_col].astype(str).str.strip() == login_nr.strip()) &
                (haendler_df[pw_col].astype(str).str.strip() == password.strip())
            ]
            if not user.empty:
                st.session_state["login_success"] = True
                st.session_state["haendler_info"] = user.iloc[0].to_dict()
                if not user.empty:
                    st.session_state["login_success"] = True
                    st.session_state["haendler_info"] = user.iloc[0].to_dict()
                    st.session_state["haendler_info"]["Login-Nr."] = login_nr.strip()  # ✅ HIER ergänzen
                    st.rerun()
                st.rerun()
            else:
                st.error("Login fehlgeschlagen – bitte Zugangsdaten prüfen.")
        except Exception as e:
            st.error(f"Fehler beim Login: {str(e)}")

# ==== Hauptprogramm ====
if "login_success" not in st.session_state:
    st.session_state["login_success"] = False
if st.session_state["login_success"] and "haendler_info" in st.session_state:
    haendler = st.session_state["haendler_info"]
    # NEU: Header mit KW oben rechts
    from datetime import datetime
    heute = datetime.today()
    kw = heute.isocalendar()[1]
    datum = heute.strftime('%d.%m.%Y')

    st.markdown(f"""
        <style>
            .main > div:first-child {{
                padding-top: 0rem !important;
            }}
            .header-flex {{
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 1rem;
                margin-top: -1.5rem;
                padding-left: 0rem;
                padding-right: 1rem;
            }}
            .header-title h2 {{
                margin: 0;
                font-size: 38px;
            }}
            .header-kw {{
                font-size: 18px;
                text-align: right;
            }}
        </style>
        <div class='header-flex'>
            <div class='header-title'>
                <h2>SONY Partnerprogramm <span style='color: #000;'>P</span><span style='color: #d1008f;'>5</span></h2>
            </div>
            <div class='header-kw'>
                Kalenderwoche **{kw}**, Datum: {datum}
            </div>
        </div>
    """, unsafe_allow_html=True)
    modus = st.radio("Was möchten Sie tun?", [
    "Verkaufszahlen melden", 
    "Projektanfrage", 
    "Bestellung zum Bestpreis", 
    "Cashback Promotion"  # ✅ NEU
    ], horizontal=True)

    moduskürzel = {
        "Verkaufszahlen melden": "verkauf",
        "Projektanfrage": "projekt",
        "Bestellung zum Bestpreis": "bestellung",
        "Cashback Promotion": "cashback"  # ✅ NEU
    }

    formular_ansicht(moduskürzel[modus])
else:
    st.markdown(""" 
        <div class='login-wrapper'>
            <h1>SONY Partnerprogramm <span style='color: #000000;'>P</span><span style='color: #d1008f;'>5</span></h1>
            <p>Zugang zum Händler-Portal für Verkaufszahlen, Projekte & Bestpreis-Bestellungen</p>
        </div>
    """, unsafe_allow_html=True)
    login()

st.markdown("---")
#st.header("🚀 Upload-Test (manuell)")
#
#if st.button("🔼 Test-Upload starten"):
#    print("✅ Upload-Block wurde manuell ausgelöst!")
#    test_eintrag = {
#        "datum": "2025-05-05",
#        "kw": "19",
#        "kommentar": "Testeintrag manuell",
#        "eintraege": [{
#            "Produktname": "Testprodukt",
#            "EAN": "1234567890123",
#            "Menge": 1,
#            "Preis": "99.00"
#        }]
#    }
#    google_drive_upload("testmodus", "Demo-Händler", test_eintrag)
