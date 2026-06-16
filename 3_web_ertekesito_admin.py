import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
from datetime import datetime, timedelta
from io import BytesIO
import openpyxl
from collections import defaultdict

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
    
    matrix["ÖSSZESEN"] = matrix.sum(axis=1)
    matrix.loc["ÖSSZESEN"] = matrix.sum(axis=0)
    
    # A bal felső cella üres ("")
    df = matrix.reset_index().rename(columns={"index": ""})
    return df

def szinezo(row):
    szinek = {
        "LGH": "#FFD1DC", "SFT": "#FFFFFF", "FLX": "#FF91A4", 
        "SUP": "#E0E0E0", "REG": "#FFC000", "FRM": "#CD7F32", 
        "STR": "#ADD8E6", "XFR": "#A6A6A6", "XST": "#FFB6C1" 
    }
    # row.iloc[0] az első oszlop értéke, függetlenül a nevétől
    cell_value = row.iloc[0]
    
    # Alap stílus minden cellára
    style = ['font-weight: bold'] * len(row)
    
    if cell_value == "ÖSSZESEN": 
        return ['background-color: #f0f0f0; font-weight: bold'] * len(row)
    
    color = szinek.get(cell_value, "#FFFFFF")
    return [f'background-color: {color}; font-weight: bold'] * len(row)

# --- RIPORT GENERÁLÁS (AGGREGÁLT) ---
def generate_weekly_report(year, week):
    jan4 = datetime(year, 1, 4)
    start_date = jan4 + timedelta(days=(week - 1) * 7 - jan4.weekday())
    naplo_docs = db.collection("naplo").where("datum", ">=", start_date.strftime("%Y-%m-%d")).stream()
    
    osszesites = defaultdict(int)
    for d in naplo_docs:
        doc = d.to_dict()
        if doc.get("tipus") == "kiszedes":
            key = (doc.get("datum"), doc.get("sku"))
            osszesites[key] += doc.get("darabszam", 0)

    wb = openpyxl.load_workbook("template.xlsx")
    ws = wb.active
    ws['O1'] = week
    for (datum, sku), mennyiseg in osszesites.items():
        datum_obj = datetime.strptime(datum, "%Y-%m-%d")
        nap_index = datum_obj.weekday()
        col_offset = nap_index * 3 + 1
        for r in range(4, 34):
            if ws.cell(row=r, column=col_offset).value is None:
                sku_parts = sku.split("_")
                ws.cell(row=r, column=col_offset, value=f"{sku_parts[0]}{sku_parts[1]}")
                ws.cell(row=r, column=col_offset+1, value=sku_parts[2])
                ws.cell(row=r, column=col_offset+2, value=mennyiseg)
                break
    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()

# --- APP LOGIKA ---
funkcio = st.sidebar.radio("Válassz felületet:", ["📱 Raktári Kiszedés", "📊 Értékesítő", "🔐 Admin"], key="nav")

if funkcio == "📊 Értékesítő":
    st.title("📊 Értékesítői Nézet")
    adatok = get_firebase_data()
    for w in ["M", "W", "XW", "XXW"]:
        st.subheader(f"📦 {w} szélesség")
        df = get_matrix(adatok, w).replace(0, "")
        st.dataframe(df.style.apply(szinezo, axis=1), use_container_width=True, hide_index=True)

elif funkcio == "🔐 Admin":
    st.title("🔐 Adminisztráció")
    if st.sidebar.text_input("Jelszó:", type="password") == ADMIN_JELSZO:
        adatok = get_firebase_data()
        for w in ["M", "W", "XW", "XXW"]:
            with st.expander(f"📦 {w} szélesség"):
                df = get_matrix(adatok, w)
                adat_df = df[df.iloc[:, 0] != "ÖSSZESEN"]
                osszesen_df = df[df.iloc[:, 0] == "ÖSSZESEN"]
                
                edited_df = st.data_editor(adat_df, hide_index=True, use_container_width=True)
                st.dataframe(osszesen_df.style.set_properties(**{'font-weight': 'bold', 'background-color': '#f0f0f0'}), 
                             hide_index=True, use_container_width=True)
                
                if st.button(f"Mentés: {w} szélesség"):
                    for index, row in edited_df.iterrows():
                        for col in edited_df.columns[1:-1]:
                            new_val = int(row[col]) if str(row[col]).isdigit() else 0
                            sku = f"{col}_{w}_{row.iloc[0]}"
                            db.collection("keszlet").document(sku).set({"mennyiseg": new_val}, merge=True)
                    st.success(f"{w} szélesség frissítve!")
                    st.rerun()
    else: st.warning("Add meg a jelszót!")
