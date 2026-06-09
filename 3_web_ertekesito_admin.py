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
            st.toast("✅ Kiszéve: 1 db", icon="📉")
            st.cache_data.clear() # Töröljük a régi készletadatokat
            st.rerun() # Újratöltés, ami már a friss adatot kéri le
        else:
            st.error("Hiba: Nincs elég készlet!")

    if b2.button("✅ VISSZARAKÁS (+1)"):
        db.collection("keszlet").document(sku).set({"mennyiseg": db_val + 1}, merge=True)
        db.collection("naplo").add({"datum": ma, "sku": sku, "tipus": "visszarakas", "darabszam": 1})
        st.toast("✅ Visszarakva: 1 db", icon="📈")
        st.cache_data.clear() # Töröljük a régi készletadatokat
        st.rerun() # Újratöltés, ami már a friss adatot kéri le

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
                
                # Stílusok
                from openpyxl.styles import Alignment, Border, Side, Font
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                                     top=Side(style='thin'), bottom=Side(style='thin'))
                center = Alignment(horizontal="center", vertical="center")

                def iras_blokkba(adat_szotar, kezdo_sor, cim):
                    # Cím egyesítése
                    ws.merge_cells(start_row=kezdo_sor, start_column=1, end_row=kezdo_sor, end_column=19)
                    ws.cell(row=kezdo_sor, column=1, value=cim).font = Font(bold=True, size=14)
                    ws.cell(row=kezdo_sor, column=1).alignment = center
                    
                    # Napok fejlécei
                    napok = ["Hétfő", "Kedd", "Szerda", "Csütörtök", "Péntek"]
                    for i, nap in enumerate(napok):
                        col = i*4 + 1
                        ws.merge_cells(start_row=kezdo_sor+1, start_column=col, end_row=kezdo_sor+1, end_column=col+2)
                        ws.cell(row=kezdo_sor+1, column=col, value=nap).font = Font(bold=True)
                        ws.cell(row=kezdo_sor+1, column=col).alignment = center
                        ws.cell(row=kezdo_sor+2, column=col, value="Méret")
                        ws.cell(row=kezdo_sor+2, column=col+1, value="Keménység")
                        ws.cell(row=kezdo_sor+2, column=col+2, value="Darab")

                    osszes_termek = sorted(list(set((k[1], k[2]) for k in adat_szotar.keys())))
                    data_start = kezdo_sor + 3
                    
                    osszes_sor = data_start
                    utolso_adat_sor = data_start - 1
                    
                    # Adatok kiírása
                    for i, (msz, kem) in enumerate(osszes_termek):
                        sor = data_start + i
                        utolso_adat_sor = sor 
                        for nap_index in range(5):
                            c = nap_index * 4 + 1
                            val = adat_szotar.get((nap_index, msz, kem), 0)
                            if val > 0:
                                ws.cell(row=sor, column=c, value=msz)
                                ws.cell(row=sor, column=c+1, value=kem)
                                ws.cell(row=sor, column=c+2, value=int(val))
                    
                    # Összesítő sor számítása
                    osszes_sor = data_start + len(osszes_termek)
                    
                    # Napi összesítők
                    for nap_index in range(5):
                        c = nap_index * 4 + 3
                        if utolso_adat_sor >= data_start:
                            r_str = f"{ws.cell(row=data_start, column=c).coordinate}:{ws.cell(row=utolso_adat_sor, column=c).coordinate}"
                            ws.cell(row=osszes_sor, column=c, value=f"=SUM({r_str})").font = Font(bold=True)
                        else:
                            ws.cell(row=osszes_sor, column=c, value=0).font = Font(bold=True)
                    
                    # Heti összesítő (T oszlopba, 20. oszlop)
                    ws.cell(row=osszes_sor, column=20, value=f"=SUM(C{osszes_sor},G{osszes_sor},K{osszes_sor},O{osszes_sor},S{osszes_sor})").font = Font(bold=True)
                    
                    # Szegélyek
                    for r in range(kezdo_sor, osszes_sor + 1):
                        for c in range(1, 21):
                            ws.cell(row=r, column=c).border = thin_border
                    return osszes_sor + 2

                # Adatok begyűjtése
                docs = list(db.collection("naplo").stream())
                kiszedesek, visszarakasok = {}, {}
                for doc in docs:
                    adat = doc.to_dict()
                    ni = pd.to_datetime(adat['datum']).dayofweek
                    if 0 <= ni <= 4:
                        sku = adat['sku'].split('_')
                        k = (ni, (sku[0]+sku[1]).lower(), sku[2].lower())
                        if adat['tipus'] == 'kiszedes': kiszedesek[k] = kiszedesek.get(k, 0) + int(adat.get('darabszam',0))
                        else: visszarakasok[k] = visszarakasok.get(k, 0) + int(adat.get('darabszam',0))

                # Generálás
                hét_szám = datetime.now().strftime("%V. Hét")
                kovetkezo_sor = iras_blokkba(kiszedesek, 1, f"{hét_szám} - KISZEDÉSEK")
                iras_blokkba(visszarakasok, kovetkezo_sor, f"{hét_szám} - VISSZARAKÁSOK")

                # Fájlnév és letöltés
                fajlnev = f"frd_kiszedes_{datetime.now().strftime('%Y_%V')}.xlsx"
                output = BytesIO()
                wb.save(output)
                
                st.download_button(
                    label="📥 Letöltés: Heti Riport", 
                    data=output.getvalue(), 
                    file_name=fajlnev, 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success("Sikeres generálás!")
            except Exception as e:
                st.error(f"Hiba: {e}")
        # --- KÉSZLET SZERKESZTÉSE ---
        st.subheader("📦 Készlet Szerkesztése")
        adatok = get_firebase_data()
        
        for w in widths:
            st.markdown(f"**\"{w}\" szélesség**")
            m_df = pd.DataFrame(0, index=hardnesses, columns=sizes)
            for m in sizes:
                for k in hardnesses:
                    m_df.at[k, m] = adatok.get(f"{m}_{w}_{k}", 0)
            
            # Táblázatos szerkesztő
            edit = st.data_editor(m_df, use_container_width=True, key=f"ed_{w}")
            
            # Mentés gomb
            if st.button(f"💾 Mentés: {w}", key=f"btn_{w}"):
                batch = db.batch()
                for m in sizes:
                    for k in hardnesses:
                        if int(edit.at[k, m]) != adatok.get(f"{m}_{w}_{k}", 0):
                            batch.set(db.collection("keszlet").document(f"{m}_{w}_{k}"), {"mennyiseg": int(edit.at[k, m])})
                batch.commit()
                st.success(f"A(z) {w} szélesség készlete frissítve!")
                st.rerun()
        # ... (készlet szerkesztő rész maradhat ahogy volt)
