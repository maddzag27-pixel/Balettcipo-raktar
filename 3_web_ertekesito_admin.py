import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials
from firebase_admin import firestore
import openpyxl
from io import BytesIO
from datetime import datetime

# --- JELSZÓ BEÁLLÍTÁSA ---
ADMIN_JELSZO = "admin123"

# --- 1. OLDAL BEÁLLÍTÁSAI ---
st.set_page_config(page_title="Balettcipő Raktár", layout="wide")

# --- FIREBASE INDÍTÁSA ---
if not firebase_admin._apps:
    secrets = st.secrets["firestore"]
    cred_dict = dict(secrets)
    cred_dict["private_key"] = cred_dict["private_key"].replace("\\n", "\n")
    
    cred = credentials.Certificate(cred_dict)
    firebase_admin.initialize_app(cred)

db = firestore.client()

# --- 3. FIX ADATOK ---
widths = ["M", "W", "XW", "XXW"]
sizes = [str(i) for i in range(5, 15)] 
hardnesses = ["LGH", "SFT", "FLX", "SUP", "REG", "FRM", "STR", "XFR", "XST"]

kemenyseg_szinek = {
    "LGH": "#FFD1DC", "SFT": "#FFFFFF", "FLX": "#FF91A4", 
    "SUP": "#E0E0E0", "REG": "#FFC000", "FRM": "#CD7F32", 
    "STR": "#4682B4", "XFR": "#A6A6A6", "XST": "#CC0000"
}

# --- 4. FUNKCIÓVÁLASZTÓ ---
st.sidebar.header("✨ Navigáció")
funkcio = st.sidebar.radio("Válassz felületet:", [
    "📱 Raktári Kiszedés (Gombos)",
    "📊 Értékesítő (Csak olvasható)",
    "🔐 Admin (Szerkeszthető)"
])

@st.cache_data(ttl=30)
def get_firebase_data():
    try:
        # Időkorlátot és hibatűrést adunk a lekérésnek
        adatok = {}
        docs = db.collection("keszlet").stream()
        for doc in docs:
            adatok[doc.id] = int(doc.to_dict().get("mennyiseg", 0))
        return adatok
    except Exception as e:
        # Ha nem sikerül, ne pörögjön, hanem írjon ki hibát
        st.error(f"Adatbázis elérési hiba: {e}")
        return {}

# ==============================================================================
# A) RAKTÁRI GOMBOS FELÜLET
# ==============================================================================
if funkcio == "📱 Raktári Kiszedés (Gombos)":
    st.title("📱 Raktári Mozgás Rögzítése")
    adatok = get_firebase_data()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.subheader("1. Méret")
        valasztott_meret = st.radio("Méretek:", sizes, key="r_size")
    with col2:
        st.subheader("2. Szélesség")
        valasztott_szelesseg = st.radio("Szélességek:", widths, key="r_width")
    with col3:
        st.subheader("3. Keménység")
        valasztott_kemenyseg = st.radio("Keménységek:", hardnesses, key="r_hard")
        
    sku_id = f"{valasztott_meret}_{valasztott_szelesseg}_{valasztott_kemenyseg}"
    aktualis_keszlet = adatok.get(sku_id, 0)
    
    st.write("---")
    st.info(f"Kiválasztott cipő: **{sku_id}** | 📦 Aktuális készlet a polcon: **{aktualis_keszlet} db**")
    
    b_col1, b_col2 = st.columns(2)
    ma_szoveg = datetime.now().strftime("%Y-%m-%d")
    
    with b_col1:
        if st.button("❌ KISZEDÉS (-1 db)", type="primary"):
            if aktualis_keszlet > 0:
                db.collection("keszlet").document(sku_id).set({"mennyiseg": aktualis_keszlet - 1}, merge=True)
                db.collection("naplo").add({"datum": ma_szoveg, "sku": sku_id, "tipus": "kiszedes", "darabszam": 1})
                st.rerun()
    with b_col2:
        if st.button("✅ VISSZARAKÁS (+1 db)"):
            db.collection("keszlet").document(sku_id).set({"mennyiseg": aktualis_keszlet + 1}, merge=True)
            db.collection("naplo").add({"datum": ma_szoveg, "sku": sku_id, "tipus": "visszarakas", "darabszam": 1})
            st.rerun()

# ==============================================================================
# B) ÉRTÉKESÍTŐ FELÜLET
# ==============================================================================
elif funkcio == "📊 Értékesítő (Csak olvasható)":
    st.title("📊 Balettcipő Élő Készlet")
    adatok = get_firebase_data()
    for w in widths:
        st.header(f"📦 \"{w}\" Szélesség")
        matrix_df = pd.DataFrame(0, index=hardnesses, columns=sizes)
        for m in sizes:
            for k in hardnesses:
                matrix_df.at[k, m] = adatok.get(f"{m}_{w}_{k}", 0)
        st.dataframe(matrix_df, use_container_width=True)

# ==============================================================================
# C) ADMIN FELÜLET
# ==============================================================================
elif funkcio == "🔐 Admin (Szerkeszthető)":
    st.title("🔐 Adminisztrátori Készletkezelés")
    if st.sidebar.text_input("Jelszó:", type="password") == ADMIN_JELSZO:
        adatok = get_firebase_data()
        for w in widths:
            st.subheader(f"🛠️ \"{w}\" szerkesztése")
            matrix_df = pd.DataFrame(0, index=hardnesses, columns=sizes)
            for m in sizes:
                for k in hardnesses:
                    matrix_df.at[k, m] = adatok.get(f"{m}_{w}_{k}", 0)
            
            edited_df = st.data_editor(matrix_df)
            if st.button(f"💾 Mentés: {w}"):
                for m in sizes:
                    for k in hardnesses:
                        sku_id = f"{m}_{w}_{k}"
                        if edited_df.at[k, m] != adatok.get(sku_id, 0):
                            db.collection("keszlet").document(sku_id).set({"mennyiseg": int(edited_df.at[k, m])})
                st.rerun()
