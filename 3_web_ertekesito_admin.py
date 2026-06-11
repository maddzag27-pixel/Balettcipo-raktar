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
    osszeg_sor = matrix.sum(axis=0)
    df = matrix.reset_index().rename(columns={"index": "Keménység"})
    df.loc[len(df)] = ["ÖSSZESEN"] + list(osszeg_sor)
    df["Keménység_Jobb"] = df["Keménység"]
    return df

# --- FŐMENÜ ---
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
    if st.button("❌ Kiszedés"):
        db.collection("keszlet").document(sku).set({"mennyiseg": adatok.get(sku, 0) - 1}, merge=True)
        db.collection("naplo").add({"datum": datetime.now().strftime("%Y-%m-%d"), "sku": sku, "tipus": "kiszedes", "darabszam": 1})
        st.rerun()

# --- ÉRTÉKESÍTŐI NÉZET ---
elif funkcio == "📊 Értékesítő":
    st.title("📊 Értékesítői Nézet")
    adatok = get_firebase_data()
    for w in ["M", "W", "XW", "XXW"]:
        st.subheader(f"📦 {w} szélesség")
        df = get_matrix(adatok, w).replace(0, "")
        st.dataframe(df, use_container_width=True)

# --- ADMINISZTRÁCIÓ ---
elif funkcio == "🔐 Admin":
    st.title("🔐 Adminisztráció")
    if st.sidebar.text_input("Jelszó:", type="password") == ADMIN_JELSZO:
        st.subheader("📊 Készlet áttekintése")
        adatok = get_firebase_data()
        for w in ["M", "W", "XW", "XXW"]:
            with st.expander(f"📦 {w} szélesség"):
                st.dataframe(get_matrix(adatok, w).replace(0, ""), use_container_width=True)
        
        st.divider()
        st.subheader("🛠 Készlet módosítása")
        valasztott_sku = st.selectbox("SKU:", [f"{m}_{w}_{k}" for m in range(5, 15) for w in ["M", "W", "XW", "XXW"] for k in ["LGH", "SFT", "FLX", "SUP", "REG", "FRM", "STR", "XFR", "XST"]])
        uj_ertek = st.number_input("Új érték:", value=0)
        if st.button("Mentés"):
            db.collection("keszlet").document(valasztott_sku).set({"mennyiseg": uj_ertek}, merge=True)
            st.success("Mentve!")
            st.rerun()
            
        st.divider()
        st.subheader("📅 Napi kiszedések")
        naplo_docs = db.collection("naplo").where("datum", "==", datetime.now().strftime("%Y-%m-%d")).stream()
        naplo_adatok = [d.to_dict() for d in naplo_docs]
        if naplo_adatok:
            st.table(pd.DataFrame(naplo_adatok))
    else:
        st.warning("Add meg a jelszót!")
