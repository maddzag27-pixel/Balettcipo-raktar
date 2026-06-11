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

# Színek
kemenyseg_szinek = {"LGH": "#FFD1DC", "SFT": "#FFFFFF", "FLX": "#FF91A4", "SUP": "#E0E0E0", "REG": "#FFC000", "FRM": "#CD7F32", "STR": "#4682B4", "XFR": "#A6A6A6", "XST": "#CC0000"}

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
    
    # Összesítés
    matrix.loc["ÖSSZESEN"] = matrix.sum()
    df = matrix.reset_index().rename(columns={"index": "Keménység"})
    df["Keménység "] = df["Keménység"]
    return df

def szinezo(row):
    color = kemenyseg_szinek.get(row["Keménység"], "#FFFFFF")
    if row["Keménység"] == "ÖSSZESEN": return ['background-color: #f0f0f0; font-weight: bold'] * len(row)
    return [f'background-color: {color}'] * len(row)

# --- FŐMENÜ ---
funkcio = st.sidebar.radio("Válassz felületet:", ["📱 Raktári Kiszedés", "📊 Értékesítő", "🔐 Admin"])

if funkcio == "📱 Raktári Kiszedés":
    st.title("📱 Raktári Mozgás")
    adatok = get_firebase_data()
    # (Ide mehet a korábbi rádiógombos logikád)

elif funkcio == "📊 Értékesítő":
    st.title("📊 Értékesítői Nézet")
    adatok = get_firebase_data()
    widths = ["M", "W", "XW", "XXW"]
    
    # Exportálás egybe
    if st.button("📥 Összes export egybe"):
        buffer = BytesIO()
        with pd.ExcelWriter(buffer) as writer:
            for w in widths:
                get_matrix(adatok, w).to_excel(writer, sheet_name=w, index=False)
        st.download_button("✅ Letöltés (Összes)", buffer.getvalue(), "Osszes_Keszlet.xlsx")

    for w in widths:
        st.subheader(f"📦 {w} szélesség")
        df = get_matrix(adatok, w)
        st.dataframe(df.style.apply(szinezo, axis=1), use_container_width=True)

elif funkcio == "🔐 Admin":
    st.title("🔐 Adminisztráció")
    if st.sidebar.text_input("Jelszó:", type="password") == ADMIN_JELSZO:
        st.subheader("📊 Heti riport generálás")
        if st.button("📥 Heti sablon generálása"):
            # Itt a korábbi iras_blokkba logika
            st.success("Riport generálva!")
    else:
        st.warning("Add meg a jelszót!")
