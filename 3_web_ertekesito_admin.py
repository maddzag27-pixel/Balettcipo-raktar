import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
from datetime import datetime

# --- JELSZÓ BEÁLLÍTÁSA ---
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

widths = ["M", "W", "XW", "XXW"]
sizes = [str(i) for i in range(5, 15)] 
hardnesses = ["LGH", "SFT", "FLX", "SUP", "REG", "FRM", "STR", "XFR", "XST"]

kemenyseg_szinek = {
    "LGH": "#FFD1DC", "SFT": "#FFFFFF", "FLX": "#FF91A4", 
    "SUP": "#E0E0E0", "REG": "#FFC000", "FRM": "#CD7F32", 
    "STR": "#4682B4", "XFR": "#A6A6A6", "XST": "#CC0000"
}

@st.cache_data(ttl=30)
def get_firebase_data():
    try:
        adatok = {}
        docs = db.collection("keszlet").stream()
        for doc in docs:
            adatok[doc.id] = int(doc.to_dict().get("mennyiseg", 0))
        return adatok
    except Exception as e:
        return {}

def get_matrix(adatok, w):
    matrix = pd.DataFrame(0, index=hardnesses, columns=sizes)
    for m in sizes:
        for k in hardnesses:
            matrix.at[k, m] = adatok.get(f"{m}_{w}_{k}", 0)
    
    # Teljes összesítés a sarokba
    total_sum = matrix.values.sum()
    
    final_df = matrix.copy()
    final_df.insert(0, "Keménység", hardnesses)
    final_df["Keménység"] = hardnesses
    
    # Összesítő sor
    total_row = matrix.sum(axis=0).to_dict()
    total_row["Keménység"] = "ÖSSZESEN"
    
    # A jobb oldali "Keménység" oszlop az utolsó elem:
    final_df = pd.concat([final_df, pd.DataFrame([total_row])], ignore_index=True)
    # A jobb alsó sarokba beírjuk a teljes összeget
    final_df.at[len(final_df)-1, "Keménység"] = str(total_sum) 
    
    return final_df

funkcio = st.sidebar.radio("Navigáció", ["📱 Raktár", "📊 Értékesítő", "🔐 Admin"])

if funkcio == "📊 Értékesítő":
    st.title("📊 Balettcipő Élő Készlet")
    adatok = get_firebase_data()
    for w in widths:
        st.subheader(f"📦 \"{w}\" Szélesség")
        df = get_matrix(adatok, w)
        
        def szinezo_df(row):
            kemenyseg = df.loc[row.name, "Keménység"]
            if kemenyseg == "ÖSSZESEN": return ['background-color: #f0f0f0'] * len(row)
            szin = kemenyseg_szinek.get(kemenyseg, "#FFFFFF")
            return [f'background-color: {szin}'] * len(row)

        styled_df = df.style.apply(szinezo_df, axis=1).hide(axis="index")
        st.dataframe(styled_df, use_container_width=True)

elif funkcio == "🔐 Admin":
    st.title("🔐 Adminisztráció")
    if st.sidebar.text_input("Jelszó:", type="password") == ADMIN_JELSZO:
        adatok = get_firebase_data()
        for w in widths:
            st.subheader(f"🛠️ \"{w}\" szerkesztése")
            matrix_df = pd.DataFrame(0, index=hardnesses, columns=sizes)
            for m in sizes:
                for k in hardnesses:
                    matrix_df.at[k, m] = adatok.get(f"{m}_{w}_{k}", 0)
            
            edited_df = st.data_editor(matrix_df, use_container_width=True, key=f"ed_{w}")
            if st.button(f"💾 Mentés: {w}", key=f"btn_{w}"):
                batch = db.batch()
                for m in sizes:
                    for k in hardnesses:
                        if int(edited_df.at[k, m]) != adatok.get(f"{m}_{w}_{k}", 0):
                            batch.set(db.collection("keszlet").document(f"{m}_{w}_{k}"), {"mennyiseg": int(edited_df.at[k, m])})
                batch.commit()
                st.rerun()
