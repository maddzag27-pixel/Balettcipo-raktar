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

# --- FÜGGVÉNYEK ---
def get_firebase_data():
    docs = db.collection("keszlet").stream()
    return {doc.id: int(doc.to_dict().get("mennyiseg", 0)) for doc in docs}

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
    szinek = {"LGH": "#FFD1DC", "SFT": "#FFFFFF", "FLX": "#FF91A4", "SUP": "#E0E0E0", "REG": "#FFC000", "FRM": "#CD7F32", "STR": "#4682B4", "XFR": "#A6A6A6", "XST": "#CC0000"}
    if row["Keménység"] == "ÖSSZESEN": return ['background-color: #f0f0f0; font-weight: bold'] * len(row)
    color = szinek.get(row["Keménység"], "#FFFFFF")
    return [f'background-color: {color}'] * len(row)

# --- APP ---
funkcio = st.sidebar.radio("Válassz felületet:", ["📱 Raktári Kiszedés", "📊 Értékesítő", "🔐 Admin"], key="nav")

if funkcio == "📱 Raktári Kiszedés":
    st.title("📱 Raktári Mozgás")
    adatok = get_firebase_data()
    c1, c2, c3 = st.columns(3)
    meret = c1.selectbox("Méret:", [str(i) for i in range(5, 15)])
    szelesseg = c2.selectbox("Szélesség:", ["M", "W", "XW", "XXW"])
    kemenyseg = c3.selectbox("Keménység:", ["LGH", "SFT", "FLX", "SUP", "REG", "FRM", "STR", "XFR", "XST"])
    sku = f"{meret}_{szelesseg}_{kemenyseg}"
    st.write(f"Készlet: **{adatok.get(sku, 0)}**")
    if st.button("❌ Kiszedés"):
        db.collection("keszlet").document(sku).set({"mennyiseg": adatok.get(sku, 0) - 1}, merge=True)
        st.rerun()
    if st.button("✅ Visszarakás"):
        db.collection("keszlet").document(sku).set({"mennyiseg": adatok.get(sku, 0) + 1}, merge=True)
        st.rerun()

elif funkcio == "📊 Értékesítő":
    st.title("📊 Értékesítői Nézet")
    adatok = get_firebase_data()
    if st.button("📥 Összes export"):
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            row = 0
            for w in ["M", "W", "XW", "XXW"]:
                get_matrix(adatok, w).to_excel(writer, sheet_name="Keszlet", startrow=row, index=False)
                row += 15
        st.download_button("✅ Letöltés", buffer.getvalue(), "Keszlet.xlsx")
    for w in ["M", "W", "XW", "XXW"]:
        st.subheader(f"📦 {w} szélesség")
        st.dataframe(get_matrix(adatok, w).style.apply(szinezo, axis=1), use_container_width=True)

elif funkcio == "🔐 Admin":
    st.title("🔐 Adminisztráció")
    if st.sidebar.text_input("Jelszó:", type="password") == ADMIN_JELSZO:
        st.subheader("📊 Készlet áttekintése és módosítása")
        
        # Táblázatos nézet (színezés nélkül, nullák elrejtésével)
        adatok = get_firebase_data()
        for w in ["M", "W", "XW", "XXW"]:
            with st.expander(f"📦 {w} szélesség"):
                df_admin = get_matrix(adatok, w).replace(0, "") # Nullák elrejtése
                st.dataframe(df_admin, use_container_width=True)
        
        st.divider()
        st.subheader("🛠 Manuális készlet módosítása")
        skus = [f"{m}_{w}_{k}" for m in range(5, 15) for w in ["M", "W", "XW", "XXW"] for k in ["LGH", "SFT", "FLX", "SUP", "REG", "FRM", "STR", "XFR", "XST"]]
        valasztott_sku = st.selectbox("SKU kiválasztása:", skus)
        uj_ertek = st.number_input("Új készletérték beállítása:", value=0, min_value=0)
        
        if st.button("Mentés az adatbázisba"):
            db.collection("keszlet").document(valasztott_sku).set({"mennyiseg": uj_ertek}, merge=True)
            st.success(f"A {valasztott_sku} új értéke: {uj_ertek}")
            st.rerun() # Frissítjük az oldalt, hogy az új érték látszódjon
            
    else:
        st.warning("Add meg a jelszót a kezelőfelület eléréséhez!")
