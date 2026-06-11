import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

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
    except Exception as e:
        st.error(f"Hiba: {e}")
        return {}

def get_matrix(adatok, w):
    sizes = [str(i) for i in range(5, 15)]
    hardnesses = ["LGH", "SFT", "FLX", "SUP", "REG", "FRM", "STR", "XFR", "XST"]
    matrix = pd.DataFrame(0, index=hardnesses, columns=sizes)
    for m in sizes:
        for k in hardnesses:
            matrix.at[k, m] = adatok.get(f"{m}_{w}_{k}", 0)
    
    osszeg_sor = matrix.sum(axis=0)
    df = matrix.reset_index().rename(columns={"index": "Keménység"})
    df.loc[len(df)] = ["ÖSSZESEN"] + list(osszeg_sor)
    df["Keménység_Jobb"] = df["Keménység"]
    return df

def szinezo(row):
    szinek = {
        "LGH": "#FFD1DC", "SFT": "#FFFFFF", "FLX": "#FF91A4", 
        "SUP": "#E0E0E0", "REG": "#FFC000", "FRM": "#CD7F32", 
        "STR": "#4682B4", "XFR": "#A6A6A6", "XST": "#CC0000"
    }
    if row["Keménység"] == "ÖSSZESEN": 
        return ['background-color: #f0f0f0; font-weight: bold'] * len(row)
    color = szinek.get(row["Keménység"], "#FFFFFF")
    return [f'background-color: {color}'] * len(row)

def iras_blokkba(ws, adat_szotar, kezdo_sor, cim):
    ws.merge_cells(start_row=kezdo_sor, start_column=1, end_row=kezdo_sor, end_column=15)
    ws.cell(row=kezdo_sor, column=1, value=cim).font = Font(bold=True, size=14)
    # Formázott riport írása
    napok = ["Hétfő", "Kedd", "Szerda", "Csütörtök", "Péntek"]
    for i, nap in enumerate(napok):
        ws.cell(row=kezdo_sor+1, column=3+i, value=nap).font = Font(bold=True)
    
    termekek = sorted({(k[1], k[2]) for k in adat_szotar.keys()})
    for i, (msz, kem) in enumerate(termekek):
        row = kezdo_sor + 2 + i
        ws.cell(row=row, column=1, value=msz)
        ws.cell(row=row, column=2, value=kem)
        for nap in range(5):
            val = adat_szotar.get((nap, msz, kem), "")
            ws.cell(row=row, column=3+nap, value=val if val != 0 else "")
    return kezdo_sor + len(termekek) + 4

# --- APP LOGIKA ---
funkcio = st.sidebar.radio("Válassz felületet:", ["📱 Raktári Kiszedés", "📊 Értékesítő", "🔐 Admin"], key="nav")

if funkcio == "📱 Raktári Kiszedés":
    st.title("📱 Raktári Mozgás")
    adatok = get_firebase_data()
    c1, c2, c3 = st.columns(3)
    meret = c1.selectbox("Méret:", [str(i) for i in range(5, 15)])
    szelesseg = c2.selectbox("Szélesség:", ["M", "W", "XW", "XXW"])
    kemenyseg = c3.selectbox("Keménység:", ["LGH", "SFT", "FLX", "SUP", "REG", "FRM", "STR", "XFR", "XST"])
    sku = f"{meret}_{szelesseg}_{kemenyseg}"
    st.write(f"Jelenlegi készlet: **{adatok.get(sku, 0)}**")
    
    col1, col2 = st.columns(2)
    if col1.button("❌ Kiszedés"):
        db.collection("keszlet").document(sku).set({"mennyiseg": adatok.get(sku, 0) - 1}, merge=True)
        db.collection("naplo").add({"datum": datetime.now().strftime("%Y-%m-%d"), "sku": sku, "tipus": "kiszedes", "darabszam": 1})
        st.rerun()
    if col2.button("✅ Visszarakás"):
        db.collection("keszlet").document(sku).set({"mennyiseg": adatok.get(sku, 0) + 1}, merge=True)
        db.collection("naplo").add({"datum": datetime.now().strftime("%Y-%m-%d"), "sku": sku, "tipus": "visszarakas", "darabszam": 1})
        st.rerun()

elif funkcio == "📊 Értékesítő":
    st.title("📊 Értékesítői Nézet")
    adatok = get_firebase_data()
    if st.button("📥 Exportálás"):
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            row = 0
            for w in ["M", "W", "XW", "XXW"]:
                get_matrix(adatok, w).replace(0, "").to_excel(writer, sheet_name="Keszlet", startrow=row, index=False)
                row += 15
        st.download_button("✅ Letöltés (Excel)", buffer.getvalue(), "Keszlet_Osszes.xlsx")
    
    for w in ["M", "W", "XW", "XXW"]:
        st.subheader(f"📦 {w} szélesség")
        df = get_matrix(adatok, w).replace(0, "")
        st.dataframe(df.style.apply(szinezo, axis=1), use_container_width=True)

elif funkcio == "🔐 Admin":
    st.title("🔐 Adminisztráció")
    if st.sidebar.text_input("Jelszó:", type="password") == ADMIN_JELSZO:
        st.subheader("🛠 Készlet módosítás")
        sku = st.selectbox("SKU:", [f"{m}_{w}_{k}" for m in range(5,15) for w in ["M","W","XW","XXW"] for k in ["LGH","SFT","FLX","SUP","REG","FRM","STR","XFR","XST"]])
        uj = st.number_input("Új érték:", value=0)
        if st.button("Mentés"):
            db.collection("keszlet").document(sku).set({"mennyiseg": uj}, merge=True)
            st.rerun()
        
        st.divider()
        st.subheader("📥 Heti riport generálása")
        if st.button("Riport készítése"):
            wb = openpyxl.Workbook(); ws = wb.active
            docs = list(db.collection("naplo").stream())
            kiszedesek, visszarakasok = {}, {}
            for doc in docs:
                adat = doc.to_dict()
                dt = datetime.strptime(adat['datum'], "%Y-%m-%d")
                nap = dt.weekday()
                if nap < 5:
                    key = (nap, adat['sku'].split('_')[0]+adat['sku'].split('_')[1], adat['sku'].split('_')[2])
                    if adat['tipus'] == 'kiszedes': kiszedesek[key] = kiszedesek.get(key, 0) + 1
                    else: visszarakasok[key] = visszarakasok.get(key, 0) + 1
            row = iras_blokkba(ws, kiszedesek, 1, "Heti Kiszedések")
            iras_blokkba(ws, visszarakasok, row, "Heti Visszarakások")
            out = BytesIO(); wb.save(out)
            st.download_button("📥 Letöltés", out.getvalue(), "heti_riport.xlsx")
    else:
        st.warning("Add meg a jelszót!")
