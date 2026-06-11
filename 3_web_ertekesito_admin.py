import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
from datetime import datetime, timedelta
from io import BytesIO
import openpyxl

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
    
    # 1. Oszloponkénti összeg (az ÖSSZESEN sorhoz)
    osszeg_sor = matrix.sum(axis=0)
    
    # 2. DataFrame összeállítása
    df = matrix.reset_index().rename(columns={"index": "Keménység"})
    
    # 3. ÖSSZESEN sor hozzáadása
    df.loc[len(df)] = ["ÖSSZESEN"] + list(osszeg_sor)
    
    # 4. Jobb szélső oszlop: Keménység másolata
    df["Keménység_Jobb"] = df["Keménység"]
    
    # 5. Végösszeg beírása a jobb alsó cellába (ÖSSZESEN sor, utolsó oszlop)
    df.at[len(df)-1, "Keménység_Jobb"] = int(osszeg_sor.sum())
    
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

# --- RIPORT GENERÁLÁS ---
def generate_weekly_report(year, week):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "FRD kiszedés"
    ws.cell(row=1, column=1, value=f"Riport: {year}. év, {week}. hét")
    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()

# --- NAVIGÁCIÓ ---
funkcio = st.sidebar.radio("Válassz felületet:", ["📱 Raktári Kiszedés", "📊 Értékesítő", "🔐 Admin"], key="nav")

# --- RAKTÁRI KISZEDÉS ---
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

    st.divider()
    st.subheader("📥 Heti riport export")
    ev, het = st.columns(2)
    ev_in = ev.number_input("Év", value=datetime.now().year)
    het_in = het.number_input("Hét", value=datetime.now().isocalendar()[1])
    if st.button("Riport készítése"):
        st.download_button("📥 Letöltés (Excel)", generate_weekly_report(ev_in, het_in), "heti_riport.xlsx")

# --- ÉRTÉKESÍTŐ ---
elif funkcio == "📊 Értékesítő":
    st.title("📊 Értékesítői Nézet")
    adatok = get_firebase_data()
    
    if st.button("📥 Összes leltár exportálása"):
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            row = 0
            for w in ["M", "W", "XW", "XXW"]:
                get_matrix(adatok, w).replace(0, "").to_excel(writer, sheet_name="Keszlet", startrow=row, index=False)
                row += 20
        st.download_button("✅ Letöltés (Excel)", buffer.getvalue(), "Leltar_Osszes.xlsx")
    
    for w in ["M", "W", "XW", "XXW"]:
        st.subheader(f"📦 {w} szélesség")
        df = get_matrix(adatok, w).replace(0, "")
        st.dataframe(df.style.apply(szinezo, axis=1), use_container_width=True)

# --- ADMIN ---
elif funkcio == "🔐 Admin":
    st.title("🔐 Adminisztráció")
    if st.sidebar.text_input("Jelszó:", type="password") == ADMIN_JELSZO:
        adatok = get_firebase_data()
        for w in ["M", "W", "XW", "XXW"]:
            with st.expander(f"📦 {w} szélesség"):
                st.dataframe(get_matrix(adatok, w).replace(0, ""), use_container_width=True)
        st.divider()
        sku = st.selectbox("SKU:", [f"{m}_{w}_{k}" for m in range(5,15) for w in ["M","W","XW","XXW"] for k in ["LGH","SFT","FLX","SUP","REG","FRM","STR","XFR","XST"]])
        uj = st.number_input("Új érték:", value=0)
        if st.button("Mentés"):
            db.collection("keszlet").document(sku).set({"mennyiseg": uj}, merge=True)
            st.rerun()
    else: st.warning("Add meg a jelszót!")
