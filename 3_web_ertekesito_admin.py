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

# --- SEGÉDFÜGGVÉNYEK ---
def get_firebase_data():
    try:
        docs = db.collection("keszlet").stream()
        return {doc.id: int(doc.to_dict().get("mennyiseg", 0)) for doc in docs}
    except: return {}

def get_matrix(adatok, w):
    sizes = [str(i) for i in range(5, 15)]
    hardnesses = ["LGH", "SFT", "FLX", "SUP", "REG", "FRM", "STR", "XFR", "XST"]
    matrix = pd.DataFrame(0, index=hardnesses, columns=sizes)
    for m in sizes:
        for k in hardnesses:
            matrix.at[k, m] = adatok.get(f"{m}_{w}_{k}", 0)
    matrix['ÖSSZESEN'] = matrix.sum(axis=1)
    matrix.loc['ÖSSZESEN'] = matrix.sum(axis=0)
    df = matrix.reset_index().rename(columns={"index": "Keménység"})
    return df

def szinezo(row):
    color = {"LGH": "#FFD1DC", "SFT": "#FFFFFF", "FLX": "#FF91A4", "SUP": "#E0E0E0", "REG": "#FFC000", "FRM": "#CD7F32", "STR": "#4682B4", "XFR": "#A6A6A6", "XST": "#CC0000"}.get(row["Keménység"], "#FFFFFF")
    if row["Keménység"] == "ÖSSZESEN": return ['background-color: #f0f0f0; font-weight: bold'] * len(row)
    return [f'background-color: {color}'] * len(row)

def iras_blokkba(ws, adat_szotar, kezdo_sor, cim):
    ws.merge_cells(start_row=kezdo_sor, start_column=1, end_row=kezdo_sor, end_column=15)
    ws.cell(row=kezdo_sor, column=1, value=cim).font = Font(bold=True, size=14)
    for i, nap in enumerate(["Hétfő", "Kedd", "Szerda", "Csütörtök", "Péntek"]):
        c = i * 3 + 1
        ws.merge_cells(kezdo_sor+1, c, kezdo_sor+1, c+2)
        ws.cell(kezdo_sor+1, c, nap).font = Font(bold=True)
    termekek = sorted({(k[1], k[2]) for k in adat_szotar.keys() if isinstance(k, tuple)})
    for i, (msz, kem) in enumerate(termekek):
        row = kezdo_sor + 3 + i
        for nap_idx in range(5):
            val = adat_szotar.get((nap_idx, msz, kem), 0)
            ws.cell(row=row, column=nap_idx*3+1, value=msz.upper())
            ws.cell(row=row, column=nap_idx*3+2, value=kem.upper())
            ws.cell(row=row, column=nap_idx*3+3, value=val if val > 0 else "")
    return kezdo_sor + 3 + len(termekek) + 2

# --- APP LOGIKA ---
funkcio = st.sidebar.radio("Válassz felületet:", ["📱 Raktári Kiszedés", "📊 Értékesítő", "🔐 Admin"], key="nav")

if funkcio == "📱 Raktári Kiszedés":
    st.title("📱 Raktári Mozgás")
    adatok = get_firebase_data()
    sizes = [str(i) for i in range(5, 15)]
    widths = ["M", "W", "XW", "XXW"]
    hardnesses = ["LGH", "SFT", "FLX", "SUP", "REG", "FRM", "STR", "XFR", "XST"]
    c1, c2, c3 = st.columns(3)
    meret = c1.radio("Méret:", sizes)
    szelesseg = c2.radio("Szélesség:", widths)
    kemenyseg = c3.radio("Keménység:", hardnesses)
    sku = f"{meret}_{szelesseg}_{kemenyseg}"
    st.info(f"Kiválasztva: {sku} | Készlet: {adatok.get(sku, 0)}")
    b1, b2 = st.columns(2)
    if b1.button("❌ KISZEDÉS (-1)"):
        db.collection("keszlet").document(sku).set({"mennyiseg": adatok.get(sku, 0) - 1}, merge=True)
        db.collection("naplo").add({"datum": datetime.now().strftime("%Y-%m-%d"), "sku": sku, "tipus": "kiszedes", "darabszam": 1})
        st.rerun()
    if b2.button("✅ VISSZARAKÁS (+1)"):
        db.collection("keszlet").document(sku).set({"mennyiseg": adatok.get(sku, 0) + 1}, merge=True)
        db.collection("naplo").add({"datum": datetime.now().strftime("%Y-%m-%d"), "sku": sku, "tipus": "visszarakas", "darabszam": 1})
        st.rerun()

elif funkcio == "📊 Értékesítő":
    st.title("📊 Értékesítői Nézet")
    adatok = get_firebase_data()
    if st.button("📥 Összes export egy lapra"):
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            row_offset = 0
            for w in ["M", "W", "XW", "XXW"]:
                df = get_matrix(adatok, w)
                df.to_excel(writer, sheet_name="Keszlet", startrow=row_offset, index=False)
                row_offset += len(df) + 2
        st.download_button("✅ Letöltés", buffer.getvalue(), "Keszlet_Osszes.xlsx")
    for w in ["M", "W", "XW", "XXW"]:
        st.subheader(f"📦 {w} szélesség")
        st.dataframe(get_matrix(adatok, w).style.apply(szinezo, axis=1), use_container_width=True)

elif funkcio == "🔐 Admin":
    st.title("🔐 Adminisztráció")
    if st.sidebar.text_input("Jelszó:", type="password") == ADMIN_JELSZO:
        if st.button("📥 Heti riport generálása"):
            wb = openpyxl.Workbook(); ws = wb.active
            docs = list(db.collection("naplo").stream())
            kiszedesek, visszarakasok = {}, {}
            for doc in docs:
                adat = doc.to_dict(); ni = pd.to_datetime(adat['datum']).dayofweek; parts = adat['sku'].split('_')
                if 0 <= ni <= 4 and len(parts) >= 3:
                    k = (ni, parts[0]+parts[1], parts[2])
                    if adat['tipus'] == 'kiszedes': kiszedesek[k] = kiszedesek.get(k, 0) + int(adat.get('darabszam',0))
                    else: visszarakasok[k] = visszarakasok.get(k, 0) + int(adat.get('darabszam',0))
            row = iras_blokkba(ws, kiszedesek, 1, "Heti Kiszedések")
            iras_blokkba(ws, visszarakasok, row, "Heti Visszarakások")
            out = BytesIO(); wb.save(out)
            st.download_button("📥 Letöltés riport", data=out.getvalue(), file_name="heti_riport.xlsx")
