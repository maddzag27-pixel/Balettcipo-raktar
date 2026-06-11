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

# --- ADATKEZELÉS ---
def get_firebase_data():
    docs = db.collection("keszlet").stream()
    return {doc.id: int(doc.to_dict().get("mennyiseg", 0)) for doc in docs}

def get_matrix(adatok, w):
    sizes = [str(i) for i in range(5, 15)]
    hardnesses = ["LGH", "SFT", "FLX", "SUP", "REG", "FRM", "STR", "XFR", "XST"]
    
    # 1. Alap mátrix
    matrix = pd.DataFrame(0, index=hardnesses, columns=sizes)
    for m in sizes:
        for k in hardnesses:
            matrix.at[k, m] = adatok.get(f"{m}_{w}_{k}", 0)
    
    # 2. DataFrame előkészítése
    df = matrix.reset_index().rename(columns={"index": "Keménység"})
    
    # 3. Oszloponkénti összegek számítása
    osszeg_sor = matrix.sum(axis=0)
    vegs_osszeg = osszeg_sor.sum()
    
    # 4. Utolsó oszlop (ismételt Keménység) hozzáadása
    df["Keménység.1"] = df["Keménység"] 
    # Megjegyzés: A pandas nem enged két azonos nevű oszlopot, 
    # de a megjelenítésnél átnevezhetjük, hogy vizuálisan "Keménység" legyen.
    
    # 5. Végső összesítő sor
    osszeg_sor_row = ["ÖSSZESEN"] + list(osszeg_sor) + [vegs_osszeg]
    df.loc[len(df)] = osszeg_sor_row
    
    # Oszlopok sorrendje: Keménység, 5-14, Keménység (ismétlés)
    cols = ["Keménység"] + sizes + ["Keménység.1"]
    df = df[cols]
    
    # Átnevezés a megjelenítéshez
    df = df.rename(columns={"Keménység.1": "Keménység"})
    
    return df
def szinezo(row):
    color = {"LGH": "#FFD1DC", "SFT": "#FFFFFF", "FLX": "#FF91A4", "SUP": "#E0E0E0", "REG": "#FFC000", "FRM": "#CD7F32", "STR": "#4682B4", "XFR": "#A6A6A6", "XST": "#CC0000"}.get(row["Keménység"], "#FFFFFF")
    if row["Keménység"] == "ÖSSZESEN": return ['background-color: #f0f0f0; font-weight: bold'] * len(row)
    return [f'background-color: {color}'] * len(row)

# --- NAVIGÁCIÓ ---
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
    if st.button("❌ Kiszedés"):
        db.collection("keszlet").document(sku).set({"mennyiseg": adatok.get(sku, 0) - 1}, merge=True)
        st.rerun()
    if st.button("✅ Visszarakás"):
        db.collection("keszlet").document(sku).set({"mennyiseg": adatok.get(sku, 0) + 1}, merge=True)
        st.rerun()

elif funkcio == "📊 Értékesítő":
    st.title("📊 Értékesítői Nézet")
    adatok = get_firebase_data()
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        start_row = 0
        for w in ["M", "W", "XW", "XXW"]:
            df = get_matrix(adatok, w)
            df.to_excel(writer, sheet_name="Keszlet", startrow=start_row, index=False)
            start_row += len(df) + 2
    st.download_button("📥 Letöltés: Összes egy lapon", buffer.getvalue(), "Keszlet_Osszes.xlsx")
    for w in ["M", "W", "XW", "XXW"]:
        st.subheader(f"📦 {w} szélesség")
        st.dataframe(get_matrix(adatok, w).style.apply(szinezo, axis=1), use_container_width=True)

elif funkcio == "🔐 Admin":
    st.title("🔐 Adminisztráció")
    if st.sidebar.text_input("Jelszó:", type="password") == ADMIN_JELSZO:
        st.subheader("🛠 Készlet módosítása")
        skus = [f"{m}_{w}_{k}" for m in range(5, 15) for w in ["M", "W", "XW", "XXW"] for k in ["LGH", "SFT", "FLX", "SUP", "REG", "FRM", "STR", "XFR", "XST"]]
        valasztott_sku = st.selectbox("SKU kiválasztása:", skus)
        uj_ertek = st.number_input("Új érték:", value=0)
        if st.button("Mentés"):
            db.collection("keszlet").document(valasztott_sku).set({"mennyiseg": uj_ertek}, merge=True)
            st.success("Mentve!")
        st.divider()
        st.subheader("📥 Heti Riport")
        if st.button("Riport generálása"):
            st.info("Logika a riport generáláshoz (lásd: iras_blokkba)...")
    else:
        st.warning("Add meg a jelszót!")
