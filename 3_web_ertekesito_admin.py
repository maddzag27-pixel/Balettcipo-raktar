import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
from datetime import datetime
from io import BytesIO
import openpyxl

# --- JELSZÓ BEÁLLÍTÁSA ---
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

# --- FIX ADATOK ---
widths = ["M", "W", "XW", "XXW"]
sizes = [str(i) for i in range(5, 15)] 
hardnesses = ["LGH", "SFT", "FLX", "SUP", "REG", "FRM", "STR", "XFR", "XST"]
kemenyseg_szinek = {"LGH": "#FFD1DC", "SFT": "#FFFFFF", "FLX": "#FF91A4", "SUP": "#E0E0E0", "REG": "#FFC000", "FRM": "#CD7F32", "STR": "#4682B4", "XFR": "#A6A6A6", "XST": "#CC0000"}

@st.cache_data(ttl=30)
def get_firebase_data():
    try:
        adatok = {}
        docs = db.collection("keszlet").stream()
        for doc in docs:
            adatok[doc.id] = int(doc.to_dict().get("mennyiseg", 0))
        return adatok
    except: return {}

def get_matrix(adatok, w):
    matrix = pd.DataFrame(0, index=hardnesses, columns=sizes)
    for m in sizes:
        for k in hardnesses:
            matrix.at[k, m] = adatok.get(f"{m}_{w}_{k}", 0)
    
    total_sum = matrix.values.sum()
    final_df = matrix.copy()
    final_df.insert(0, "Keménység", hardnesses)
    final_df["Keménység "] = hardnesses 
    
    total_row = matrix.sum(axis=0).to_dict()
    total_row["Keménység"] = "ÖSSZESEN"
    total_row["Keménység "] = str(total_sum)
    
    final_df = pd.concat([final_df, pd.DataFrame([total_row])], ignore_index=True)
    return final_df

# --- NAVIGÁCIÓ ---
funkcio = st.sidebar.radio("Válassz felületet:", ["📱 Raktári Kiszedés", "📊 Értékesítő", "🔐 Admin"])

if funkcio == "📱 Raktári Kiszedés":
    st.title("📱 Raktári Mozgás")
    adatok = get_firebase_data()
    c1, c2, c3 = st.columns(3)
    meret = c1.radio("Méret:", sizes)
    szelesseg = c2.radio("Szélesség:", widths)
    kemenyseg = c3.radio("Keménység:", hardnesses)
    sku = f"{meret}_{szelesseg}_{kemenyseg}"
    db_val = adatok.get(sku, 0)
    
    st.info(f"Kiválasztva: **{sku}** | Aktuális készlet: **{db_val}**")
    
    b1, b2 = st.columns(2)
    ma = datetime.now().strftime("%Y-%m-%d")
    
    if b1.button("❌ KISZEDÉS (-1)"):
        if db_val > 0:
            db.collection("keszlet").document(sku).set({"mennyiseg": db_val - 1}, merge=True)
            db.collection("naplo").add({"datum": ma, "sku": sku, "tipus": "kiszedes", "darabszam": 1})
            st.rerun()
    if b2.button("✅ VISSZARAKÁS (+1)"):
        db.collection("keszlet").document(sku).set({"mennyiseg": db_val + 1}, merge=True)
        db.collection("naplo").add({"datum": ma, "sku": sku, "tipus": "visszarakas", "darabszam": 1})
        st.rerun()

elif funkcio == "📊 Értékesítő":
    st.title("📊 Értékesítői Nézet")
    adatok = get_firebase_data()
    for w in widths:
        st.subheader(f"📦 \"{w}\" Szélesség")
        df = get_matrix(adatok, w)
        def szinezo(row):
            k = df.loc[row.name, "Keménység"]
            if k == "ÖSSZESEN": return ['background-color: #f0f0f0'] * len(row)
            return [f'background-color: {kemenyseg_szinek.get(k, "#FFFFFF")}'] * len(row)
        st.dataframe(df.style.apply(szinezo, axis=1).hide(axis="index"), use_container_width=True)

elif funkcio == "🔐 Admin":
    st.title("🔐 Adminisztráció")
    if st.sidebar.text_input("Jelszó:", type="password") == ADMIN_JELSZO:
# --- EXPORTÁLÁS: Fix sablonba (kiszedes_sablon.xlsx) ---
        st.subheader("📊 Adatkezelés")
        if st.button("📥 Heti sablon kitöltése adatokkal"):
            try:
                fajlnev = "kiszedes_sablon.xlsx" 
                wb = openpyxl.load_workbook(fajlnev)
                ws = wb.active

# 2. Adatok lekérése a Firebase-ből
                docs = db.collection("naplo").stream()
                
                # 3. Napok oszlopainak megfeleltetése: Hétfő (C=3), Kedd (F=6), Szerda (I=9), Csütörtök (L=12), Péntek (O=15)
                nap_oszlop_darab = {0: 3, 1: 6, 2: 9, 3: 12, 4: 15}

# 3. Napok oszlopainak megfeleltetése:
                # A napokhoz tartozó [Méret, Keménység, Darab] blokkok kezdete
                # Hétfő=1, Kedd=4, Szerda=7, Csütörtök=10, Péntek=13
                nap_blokk_kezdete = {0: 1, 1: 4, 2: 7, 3: 10, 4: 13}

                for doc in docs:
                    adat = doc.to_dict()
                    datum = pd.to_datetime(adat['datum'])
                    nap_index = datum.dayofweek
                    
                    if 0 <= nap_index <= 4:
                        sku_reszek = adat['sku'].split('_') 
                        meret_db = str(sku_reszek[0]).strip()
                        kemenyseg_db = str(sku_reszek[2]).strip()
                        darab = int(adat.get('darabszam', 0))
                        
                        kezdo_oszlop = nap_blokk_kezdete[nap_index]
                        
                        # Sor keresése (A 4. sortól a 30-ig)
                        for row in range(4, 31): 
                            # Az adott nap blokkjának 1. és 2. oszlopát nézzük
                            cell_m = str(ws.cell(row=row, column=kezdo_oszlop).value or "").strip()
                            cell_k = str(ws.cell(row=row, column=kezdo_oszlop + 1).value or "").strip()
                            
                            # Ha egyezik, írjuk a 3. oszlopba (ami a kezdo_oszlop + 2)
                            if cell_m == meret_db and cell_k == kemenyseg_db:
                                cel_oszlop = kezdo_oszlop + 2
                                current_val = ws.cell(row=row, column=cel_oszlop).value or 0
                                ws.cell(row=row, column=cel_oszlop).value = int(current_val) + darab
                                break
                output = BytesIO()
                wb.save(output)
                st.download_button(
                    label="📥 Kitöltött sablon letöltése",
                    data=output.getvalue(),
                    file_name="Kitoltott_kiszedes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Hiba történt a kitöltés során: {e}")
        # --- KÉSZLET SZERKESZTÉSE ---
        st.subheader("📦 Készlet Szerkesztése")
        adatok = get_firebase_data()
        for w in widths:
            st.markdown(f"**\"{w}\" szélesség**")
            m_df = pd.DataFrame(0, index=hardnesses, columns=sizes)
            for m in sizes:
                for k in hardnesses:
                    m_df.at[k, m] = adatok.get(f"{m}_{w}_{k}", 0)
            edit = st.data_editor(m_df, use_container_width=True, key=f"ed_{w}")
            if st.button(f"💾 Mentés: {w}", key=f"btn_{w}"):
                batch = db.batch()
                for m in sizes:
                    for k in hardnesses:
                        if int(edit.at[k, m]) != adatok.get(f"{m}_{w}_{k}", 0):
                            batch.set(db.collection("keszlet").document(f"{m}_{w}_{k}"), {"mennyiseg": int(edit.at[k, m])})
                batch.commit()
                st.rerun()
