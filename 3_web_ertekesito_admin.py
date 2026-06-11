import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

# --- KONFIGURÁCIÓ ---
ADMIN_JELSZO = "admin123"
st.set_page_config(page_title="Balettcipő Raktár", layout="wide")

# --- FIREBASE ---
@st.cache_resource
def get_db():
    if not firebase_admin._apps:
        secrets = st.secrets["firestore"]
        cred_dict = dict(secrets)
        cred_dict["private_key"] = cred_dict["private_key"].replace("\\n", "\n")
        cred = credentials.Certificate(cred_dict)
        firebase_admin.initialize_app(cred)
    return firestore.client()

db = get_db()

# --- SEGÉDFÜGGVÉNYEK ---
def get_firebase_data():
    try:
        adatok = {}
        docs = db.collection("keszlet").stream()
        for doc in docs:
            adatok[doc.id] = int(doc.to_dict().get("mennyiseg", 0))
        return adatok
    except: return {}

def iras_blokkba(ws, adat_szotar, kezdo_sor, cim):
    # Fejlécek
    ws.merge_cells(start_row=kezdo_sor, start_column=1, end_row=kezdo_sor, end_column=15)
    ws.cell(row=kezdo_sor, column=1, value=cim).font = Font(bold=True, size=14)
    napok = ["Hétfő", "Kedd", "Szerda", "Csütörtök", "Péntek"]
    for i, nap in enumerate(napok):
        c = i * 3 + 1
        ws.merge_cells(kezdo_sor+1, c, kezdo_sor+1, c+2)
        ws.cell(kezdo_sor+1, c, nap).font = Font(bold=True)
        ws.cell(kezdo_sor+2, c, "MÉRET"); ws.cell(kezdo_sor+2, c+1, "KEM."); ws.cell(kezdo_sor+2, c+2, "DB")
    
    # Adatok
    termekek = sorted({(k[1], k[2]) for k in adat_szotar.keys() if isinstance(k, tuple)})
    for i, (msz, kem) in enumerate(termekek):
        row = kezdo_sor + 3 + i
        for nap_idx in range(5):
            val = adat_szotar.get((nap_idx, msz, kem), 0)
            ws.cell(row=row, column=nap_idx*3+1, value=msz.upper())
            ws.cell(row=row, column=nap_idx*3+2, value=kem.upper())
            ws.cell(row=row, column=nap_idx*3+3, value=val if val > 0 else "")
            # Keretezés
            for c in range(1, 16):
                ws.cell(row=row, column=c).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    return kezdo_sor + 3 + len(termekek) + 2

# --- FŐMENÜ ---
funkcio = st.sidebar.radio("Válassz felületet:", ["📱 Raktári Kiszedés", "📊 Értékesítő", "🔐 Admin"])

if funkcio == "📱 Raktári Kiszedés":
    st.title("📱 Raktári Mozgás")
    # (Ide illeszd vissza a saját rádiógombos kiszedő logikádat)

elif funkcio == "📊 Értékesítő":
    st.title("📊 Értékesítői Nézet")
    adatok = get_firebase_data()
    widths = ["M", "W", "XW", "XXW"]
    for w in widths:
        st.subheader(f"📦 {w} szélesség")
        # (Ide illeszd vissza a saját get_matrix hívásodat)

elif funkcio == "🔐 Admin":
    st.title("🔐 Adminisztráció")
    if st.sidebar.text_input("Jelszó:", type="password") == ADMIN_JELSZO:
        if st.button("📥 Heti riport generálása"):
            wb = openpyxl.Workbook()
            ws = wb.active
            docs = list(db.collection("naplo").stream())
            kiszedesek, visszarakasok = {}, {}
            for doc in docs:
                adat = doc.to_dict()
                ni = pd.to_datetime(adat['datum']).dayofweek
                parts = adat['sku'].split('_')
                if 0 <= ni <= 4 and len(parts) >= 3:
                    k = (ni, parts[0]+parts[1], parts[2])
                    if adat['tipus'] == 'kiszedes': kiszedesek[k] = kiszedesek.get(k, 0) + int(adat.get('darabszam',0))
                    else: visszarakasok[k] = visszarakasok.get(k, 0) + int(adat.get('darabszam',0))
            
            row = iras_blokkba(ws, kiszedesek, 1, "Heti Kiszedések")
            iras_blokkba(ws, visszarakasok, row, "Heti Visszarakások")
            
            output = BytesIO()
            wb.save(output)
            st.download_button("📥 Letöltés: Heti Riport", data=output.getvalue(), file_name="heti_riport.xlsx")
