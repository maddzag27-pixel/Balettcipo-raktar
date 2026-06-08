import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Font

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
        st.subheader("📊 Adatkezelés")
        if st.button("📥 Heti sablon kitöltése adatokkal"):
            try:
                wb = openpyxl.Workbook()
                ws = wb.active
                
                # Fejlécek függvénye
                def iras_fejlec(ws, kezdo_sor, cim):
                    ws.cell(row=kezdo_sor, column=1, value=cim).font = Font(bold=True, size=14)
                    napok = ["Hétfő", "Kedd", "Szerda", "Csütörtök", "Péntek"]
                    for i, nap in enumerate(napok):
                        col = i*4 + 1
                        ws.cell(row=kezdo_sor+1, column=col, value=nap)
                        ws.cell(row=kezdo_sor+2, column=col, value="Méret")
                        ws.cell(row=kezdo_sor+2, column=col+1, value="Keménység")
                        ws.cell(row=kezdo_sor+2, column=col+2, value="Darab")

                # Adatok szétválogatása
                docs = list(db.collection("naplo").stream())
                kiszedesek, visszarakasok = {}, {}
                for doc in docs:
                    adat = doc.to_dict()
                    nap_index = pd.to_datetime(adat['datum']).dayofweek
                    if 0 <= nap_index <= 4:
                        sku_reszek = adat['sku'].split('_')
                        msz = (str(sku_reszek[0]) + str(sku_reszek[1])).lower()
                        kem = str(sku_reszek[2]).lower()
                        db_sz = int(adat.get('darabszam', 0))
                        kulcs = (nap_index, msz, kem)
                        if adat['tipus'] == 'kiszedes': kiszedesek[kulcs] = kiszedesek.get(kulcs, 0) + db_sz
                        else: visszarakasok[kulcs] = visszarakasok.get(kulcs, 0) + db_sz

                # Blokkok kiírása
                def iras_blokkba(adat_szotar, kezdo_sor, cim):
                    iras_fejlec(ws, kezdo_sor, cim)
                    osszes_termek = sorted(list(set((k[1], k[2]) for k in adat_szotar.keys())))
                    data_start = kezdo_sor + 3
                    
                    for i, (msz, kem) in enumerate(osszes_termek):
                        sor = data_start + i
                        for nap_index in range(5):
                            kezdo_oszlop = nap_index * 4 + 1
                            mennyiseg = adat_szotar.get((nap_index, msz, kem), 0)
                            if mennyiseg > 0:
                                ws.cell(row=sor, column=kezdo_oszlop, value=msz)
                                ws.cell(row=sor, column=kezdo_oszlop + 1, value=kem)
                                ws.cell(row=sor, column=kezdo_oszlop + 2, value=mennyiseg)
                    
                    utolso_adat_sor = data_start + len(osszes_termek) - 1
                    osszes_sor = data_start + len(osszes_termek)
                    ws.cell(row=osszes_sor, column=1, value="ÖSSZESEN").font = Font(bold=True)
                    for nap_index in range(5):
                        oszlop = nap_index * 4 + 3
                        range_str = f"{ws.cell(row=data_start, column=oszlop).coordinate}:{ws.cell(row=utolso_adat_sor, column=oszlop).coordinate}"
                        ws.cell(row=osszes_sor, column=oszlop, value=f"=SUM({range_str})")
                    return osszes_sor + 3

                kovetkezo_sor = iras_blokkba(kiszedesek, 1, "KISZEDÉSEK")
                iras_blokkba(visszarakasok, kovetkezo_sor, "VISSZARAKÁSOK")

                output = BytesIO()
                wb.save(output)
                st.download_button("📥 Letöltés: Heti Riport", data=output.getvalue(), file_name="Heti_Riport.xlsx")
                st.success("Sikeres generálás!")
            except Exception as e:
                st.error(f"Hiba történt: {e}")

        st.subheader("📦 Készlet Szerkesztése")
        # ... (készlet szerkesztő rész maradhat ahogy volt)
