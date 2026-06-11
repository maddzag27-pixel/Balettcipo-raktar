import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

# --- JELSZÓ BEÁLLÍTÁSA ---
ADMIN_JELSZO = "admin123"

st.set_page_config(page_title="Balettcipő Raktár", layout="wide")

# --- FIREBASE INDÍTÁSA ---
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

# --- FIX ADATOK ---
widths = ["M", "W", "XW", "XXW"]
sizes = [str(i) for i in range(5, 15)] 
hardnesses = ["LGH", "SFT", "FLX", "SUP", "REG", "FRM", "STR", "XFR", "XST"]
kemenyseg_szinek = {"LGH": "#FFD1DC", "SFT": "#FFFFFF", "FLX": "#FF91A4", "SUP": "#E0E0E0", "REG": "#FFC000", "FRM": "#CD7F32", "STR": "#4682B4", "XFR": "#A6A6A6", "XST": "#CC0000"}

@st.cache_data(ttl=0)
def get_firebase_data():
    try:
        adatok = {}
        docs = db.collection("keszlet").stream()
        for doc in docs:
            adatok[doc.id] = int(doc.to_dict().get("mennyiseg", 0))
        return adatok
    except: return {}

def get_matrix(adatok, w):
    matrix = pd.DataFrame(0, index=hardnesses, columns=sizes)
    for m in sizes:
        for k in hardnesses:
            matrix.at[k, m] = adatok.get(f"{m}_{w}_{k}", 0)
    matrix = matrix.replace(0, "")
    final_df = matrix.copy()
    final_df.insert(0, "Keménység", hardnesses)
    return final_df

# --- ADMIN FUNKCIÓ JAVÍTVA ---
def iras_blokkba(ws, adat_szotar, kezdo_sor, cim):
    if not isinstance(adat_szotar, dict): adat_szotar = {}
    
    ws.merge_cells(start_row=kezdo_sor, start_column=1, end_row=kezdo_sor, end_column=15)
    ws.cell(row=kezdo_sor, column=1, value=cim).font = Font(bold=True, size=14)
    
    napok = ["Hétfő", "Kedd", "Szerda", "Csütörtök", "Péntek"]
    for i, nap in enumerate(napok):
        c = i * 3 + 1
        ws.merge_cells(kezdo_sor+1, c, kezdo_sor+1, c+2)
        ws.cell(kezdo_sor+1, c, nap).font = Font(bold=True)
        ws.cell(kezdo_sor+2, c, "MÉRET"); ws.cell(kezdo_sor+2, c+1, "KEM."); ws.cell(kezdo_sor+2, c+2, "DB")

    termek_halmaz = { (k[1], k[2]) for k in adat_szotar.keys() if isinstance(k, tuple) }
    osszes_termek = sorted(list(termek_halmaz))
    
    for i, (msz, kem) in enumerate(osszes_termek):
        row = kezdo_sor + 3 + i
        for nap_idx in range(5):
            col = nap_idx * 3 + 1
            val = adat_szotar.get((nap_idx, msz, kem), 0)
            ws.cell(row=row, column=col, value=msz.upper())
            ws.cell(row=row, column=col+1, value=kem.upper())
            ws.cell(row=row, column=col+2, value=val if val > 0 else "")
            
    osszes_sor = kezdo_sor + 3 + max(1, len(osszes_termek))
    for nap_idx in range(5):
        col = nap_idx * 3 + 3
        start_cell = ws.cell(row=kezdo_sor+3, column=col).coordinate
        end_cell = ws.cell(row=osszes_sor-1, column=col).coordinate
        ws.cell(row=osszes_sor, column=col, value=f"=SUM({start_cell}:{end_cell})").font = Font(bold=True)
    
    return osszes_sor + 2

# --- APP LOGIKA ---
funkcio = st.sidebar.radio("Válassz felületet:", ["📱 Raktári Kiszedés", "📊 Értékesítő", "🔐 Admin"])

if funkcio == "📱 Raktári Kiszedés":
    st.title("📱 Raktári Mozgás")
    adatok = get_firebase_data()
    c1, c2, c3 = st.columns(3)
    meret = c1.radio("Méret:", sizes)
    szelesseg = c2.radio("Szélesség:", widths)
    kemenyseg = c3.radio("Keménység:", hardnesses)
    sku = f"{meret}_{szelesseg}_{kemenyseg}"
    db_val = adatok.get(sku, 0)
    st.info(f"Kiválasztva: **{sku}** | Aktuális készlet: **{db_val}**")
    b1, b2 = st.columns(2)
    ma = datetime.now().strftime("%Y-%m-%d")
    if b1.button("❌ KISZEDÉS (-1)"):
        if db_val > 0:
            db.collection("keszlet").document(sku).set({"mennyiseg": db_val - 1}, merge=True)
            db.collection("naplo").add({"datum": ma, "sku": sku, "tipus": "kiszedes", "darabszam": 1})
            st.rerun()
    if b2.button("✅ VISSZARAKÁS (+1)"):
        db.collection("keszlet").document(sku).set({"mennyiseg": db_val + 1}, merge=True)
        db.collection("naplo").add({"datum": ma, "sku": sku, "tipus": "visszarakas", "darabszam": 1})
        st.rerun()

elif funkcio == "📊 Értékesítő":
    st.title("📊 Értékesítői Nézet")
    if st.button("🔄 Adatok frissítése"): st.rerun()
    adatok = get_firebase_data()
    for w in widths:
        st.subheader(f"📦 \"{w}\" Szélesség")
        df = get_matrix(adatok, w)
        st.dataframe(df.style.hide(axis="index"), use_container_width=True)

elif funkcio == "🔐 Admin":
    st.title("🔐 Adminisztráció")
    if st.sidebar.text_input("Jelszó:", type="password") == ADMIN_JELSZO:
        if st.button("📥 Heti sablon kitöltése"):
            wb = openpyxl.Workbook()
            ws = wb.active
            docs = list(db.collection("naplo").stream())
            kiszedesek, visszarakasok = {}, {}
            for doc in docs:
                adat = doc.to_dict()
                ni = pd.to_datetime(adat['datum']).dayofweek
                sku_str = adat['sku'] # Ez egy string (pl: 8_M_LGH)
                parts = sku_str.split('_')
                if 0 <= ni <= 4 and len(parts) >= 3:
                    k = (ni, parts[0]+parts[1], parts[2])
                    if adat['tipus'] == 'kiszedes': kiszedesek[k] = kiszedesek.get(k, 0) + int(adat.get('darabszam',0))
                    else: visszarakasok[k] = visszarakasok.get(k, 0) + int(adat.get('darabszam',0))
            
            hét_szám = datetime.now().strftime("%V. Hét")
            row = iras_blokkba(ws, kiszedesek, 1, f"{hét_szám} - KISZEDÉSEK")
            iras_blokkba(ws, visszarakasok, row, f"{hét_szám} - VISSZARAKÁSOK")
            
            output = BytesIO()
            wb.save(output)
            st.download_button("📥 Letöltés: Heti Riport", data=output.getvalue(), file_name="heti_riport.xlsx")
