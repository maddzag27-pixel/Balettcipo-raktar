import streamlit as st
import pandas as pd
from firebase_admin import credentials, firestore
import firebase_admin
import json
from datetime import datetime, timedelta
from io import BytesIO
import openpyxl
from collections import defaultdict

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

@st.cache_data(ttl=60)
def get_firebase_data():
    try:
        docs = db.collection("keszlet").stream()
        data = {}
        for doc in docs:
            d = doc.to_dict()
            data[doc.id] = {
                "mennyiseg": int(d.get("mennyiseg", 0)),
                "min_ertek": int(d.get("min_ertek", 0))
            }
        return data
    except: return {}

def get_matrix(adatok, w):
    sizes = [str(i) for i in range(5, 15)]
    hardnesses = ["LGH", "SFT", "FLX", "SUP", "REG", "FRM", "STR", "XFR", "XST"]
    matrix = pd.DataFrame(0, index=hardnesses, columns=sizes)
    
    for m in sizes:
        for k in hardnesses:
            termek_info = adatok.get(f"{m}_{w}_{k}", {"mennyiseg": 0})
            matrix.at[k, m] = termek_info.get("mennyiseg", 0)
    
    matrix.loc["ÖSSZESEN"] = matrix.sum(axis=0)
    df = matrix.reset_index()
    df.columns.values[0] = "Keménység"
    
    df = df.astype(object)
    df["Keménység "] = df["Keménység"] 
    
    teljes_osszeg = df.iloc[:-1, 1:-1].sum().sum()
    df.iloc[-1, -1] = teljes_osszeg
    
    return df

# --- APP LOGIKA ---
st.set_page_config(layout="wide")
funkcio = st.sidebar.radio("Válassz felületet:", ["📱 Raktári Kiszedés", "📊 Értékesítő", "🔐 Admin"])
adatok = get_firebase_data()

if funkcio == "📊 Értékesítő":
    st.title("📊 Értékesítői Nézet")
    for w in ["M", "W", "XW", "XXW"]:
        st.subheader(f"📦 {w} szélesség")
        df = get_matrix(adatok, w).replace(0, "")
        st.dataframe(df, use_container_width=True, hide_index=True)

elif funkcio == "🔐 Admin":
    st.title("🔐 Adminisztráció")
    if st.sidebar.text_input("Jelszó:", type="password") == "admin123":
        for w in ["M", "W", "XW", "XXW"]:
            with st.expander(f"📦 {w} szélesség"):
                df = get_matrix(adatok, w)
                edited_df = st.data_editor(df, hide_index=True, use_container_width=True)
                
                if st.button(f"Mentés: {w} szélesség"):
                    for _, row in edited_df.iterrows():
                        kem = row.iloc[0]
                        if kem == "ÖSSZESEN": continue
                        # Csak a középső (méret) oszlopokon megyünk végig
                        for col in edited_df.columns[1:-1]: 
                            val = row[col]
                            new_val = int(val) if str(val).isdigit() else 0
                            sku = f"{col}_{w}_{kem}"
                            db.collection("keszlet").document(sku).set({"mennyiseg": new_val}, merge=True)
                    st.success("Mentve!")
                    st.rerun()
    else: st.warning("Add meg a jelszót!")
