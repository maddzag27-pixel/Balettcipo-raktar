import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials
from firebase_admin import firestore
import openpyxl
from io import BytesIO
from datetime import datetime

# --- JELSZÓ BEÁLLÍTÁSA (Ezt írd át amire szeretnéd!) ---
ADMIN_JELSZO = "admin123"

# --- 1. OLDAL BEÁLLÍTÁSAI ---
st.set_page_config(page_title="Balettcipő Raktár", layout="wide")

# --- 2. FIREBASE INDÍTÁSA ---
if not firebase_admin._apps:
    cred = credentials.Certificate("secrets.json")
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

# --- ADATOK LEKÉRÉSE ---
docs = db.collection("keszlet").stream()
firebase_adatok = {doc.id: doc.to_dict().get("mennyiseg", 0) for doc in docs}

# ==============================================================================
# A) RAKTÁRI GOMBOS FELÜLET (Mindenki eléri)
# ==============================================================================
if funkcio == "📱 Raktári Kiszedés (Gombos)":
    st.title("📱 Raktári Mozgás Rögzítése")
    
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
    aktualis_keszlet = firebase_adatok.get(sku_id, 0)
    
    st.write("---")
    st.info(f"Kiválasztott cipő: **{sku_id}** | 📦 Aktuális készlet a polcon: **{aktualis_keszlet} db**")
    
    b_col1, b_col2 = st.columns(2)
    ma_szoveg = datetime.now().strftime("%Y-%m-%d")
    
    with b_col1:
        if st.button("❌ KISZEDÉS (-1 db)", type="primary", use_container_width=True):
            if aktualis_keszlet > 0:
                uj_db = aktualis_keszlet - 1
                db.collection("keszlet").document(sku_id).update({"mennyiseg": uj_db})
                db.collection("naplo").add({"datum": ma_szoveg, "sku": sku_id, "tipus": "kiszedes", "darabszam": 1})
                st.success(f"Sikeresen kiszedve 1 db! Új készlet: {uj_db}")
                st.rerun()
            else:
                st.error("A készlet nem mehet 0 alá!")
    with b_col2:
        if st.button("✅ VISSZARAKÁS (+1 db)", use_container_width=True):
            uj_db = aktualis_keszlet + 1
            db.collection("keszlet").document(sku_id).update({"mennyiseg": uj_db})
            db.collection("naplo").add({"datum": ma_szoveg, "sku": sku_id, "tipus": "visszarakas", "darabszam": 1})
            st.success(f"Sikeresen visszarakva 1 db! Új készlet: {uj_db}")
            st.rerun()

# ==============================================================================
# B) ÉRTÉKESÍTŐ FELÜLET (Mindenki eléri, csak olvasható)
# ==============================================================================
elif funkcio == "📊 Értékesítő (Csak olvasható)":
    st.title("📊 Balettcipő Élő Készlet (Olvasó)")

    def apply_row_styles(row):
        kemenyseg = row.name
        color = kemenyseg_szinek.get(kemenyseg, "#FFFFFF")
        text_color = "#FFFFFF" if kemenyseg == "XST" else "#000000"
        return [f"background-color: {color}; color: {text_color}; font-weight: 500;" for _ in row]

    for w in widths:
        st.header(f"📦 \"{w}\" Szélességű Cipők")
        matrix_df = pd.DataFrame(0, index=hardnesses, columns=sizes)
        for m in sizes:
            for k in hardnesses:
                sku_id = f"{m}_{w}_{k}"
                matrix_df.at[k, m] = firebase_adatok.get(sku_id, 0)
        
        matrix_df["Keménység "] = matrix_df.index
        oszlop_osszegek = [matrix_df[m].sum() for m in sizes]
        teljes_vegosszeg = sum(oszlop_osszegek)
        
        formatted_df = matrix_df.copy()
        for m in sizes:
            formatted_df[m] = formatted_df[m].apply(lambda x: str(x) if x > 0 else "")
        
        display_df = formatted_df.copy()
        display_df.loc["ÖSSZ:"] = oszlop_osszegek + [f"🎁 VÉGÖSSZEG: {teljes_vegosszeg}"]

        styled_df = display_df.style.apply(
            lambda row: apply_row_styles(row) if row.name != "ÖSSZ:" else ['background-color: #365F91; color: white; font-weight: bold;'] * len(row), axis=1
        )
        st.dataframe(styled_df, use_container_width=True, height=390)

# ==============================================================================
# C) ADMIN FELÜLET (CSAK JELSZÓVAL!)
# ==============================================================================
elif funkcio == "🔐 Admin (Szerkeszthető)":
    st.title("🔐 Adminisztrátori Készletkezelés")
    
    # Jelszó bekérő mező az oldalsávban vagy a főoldalon
    bevitt_jelszo = st.sidebar.text_input("Írd be az Admin jelszót:", type="password")
    
    if bevitt_jelszo != ADMIN_JELSZO:
        st.warning("⚠️ Kérjük, add meg a helyes adminisztrátori jelszót az oldalsávban a hozzáféréshez!")
    else:
        st.success("🔓 Hozzáférés megadva!")
        
        for w in widths:
            st.header(f"🛠️ \"{w}\" Szélesség szerkesztése")
            matrix_df = pd.DataFrame(0, index=hardnesses, columns=sizes)
            for m in sizes:
                for k in hardnesses:
                    sku_id = f"{m}_{w}_{k}"
                    matrix_df.at[k, m] = firebase_adatok.get(sku_id, 0)
            
            matrix_df["Keménység "] = matrix_df.index
            
            edited_df = st.data_editor(matrix_df, key=f"editor_{w}", use_container_width=True, disabled=["Keménység "])
            if not edited_df.equals(matrix_df):
                if st.button(f"💾 \"{w}\" mentése", key=f"btn_{w}", type="primary"):
                    batch = db.batch()
                    valtozott_valami = False
                    for m in sizes:
                        for k in hardnesses:
                            sku_id = f"{m}_{w}_{k}"
                            uj_ertek = int(edited_df.at[k, m])
                            regi_ertek = firebase_adatok.get(sku_id, 0)
                            if uj_ertek != regi_ertek:
                                doc_ref = db.collection("keszlet").document(sku_id)
                                batch.set(doc_ref, {"mennyiseg": uj_ertek}, merge=True)
                                valtozott_valami = True
                    if valtozott_valami:
                        batch.commit()
                        st.success("Készlet sikeresen frissítve a felhőben!")
                        st.rerun()
                    else:
                        st.info("Nem történt változás, nincs mit menteni.")

        # --- EXCEL EXPORT (Csak az admin látja!) ---
        st.write("---")
        st.subheader("📊 Gyártástervezési Napló Mentése")
        
        def excel_keszites_memoriaban():
            ma_szoveg = datetime.now().strftime("%Y-%m-%d")
            wb = openpyxl.load_workbook("kiszedes_sablon.xlsx")
            ws = wb.active
            naplo_ref = db.collection("naplo").where("datum", "==", ma_szoveg).stream()
            
            adatok_kiszedes, adatok_visszarakas = {}, {}
            for doc in naplo_ref:
                data = doc.to_dict()
                sku, tipus, db_szam = data.get("sku"), data.get("tipus"), data.get("darabszam", 1)
                if tipus == "kiszedes": adatok_kiszedes[sku] = adatok_kiszedes.get(sku, 0) + db_szam
                elif tipus == "visszarakas": adatok_visszarakas[sku] = adatok_visszarakas.get(sku, 0) + db_szam

            def cella_kitoltes(sku_adatok, sor_eltolas):
                for sku, darab in sku_adatok.items():
                    meret, szelesseg, kemenyseg = sku.split("_")
                    try:
                        oszlop_idx = 2 + sizes.index(meret)
                        szelesseg_alap_sorok = {"M": 5, "W": 20, "XW": 35, "XXW": 50}
                        vegso_sor = szelesseg_alap_sorok[szelesseg] + hardnesses.index(kemenyseg) + sor_eltolas
                        ws.cell(row=vegso_sor, column=oszlop_idx, value=darab)
                    except: continue

            cella_kitoltes(adatok_kiszedes, sor_eltolas=0)
            cella_kitoltes(adatok_visszarakas, sor_eltolas=60)
            
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            return output.getvalue()

        try:
            excel_adat = excel_keszites_memoriaban()
            st.download_button(label="📥 Aznapi Gyártásterv Excel letöltése", data=excel_adat, file_name=f"Gyartasterv_{datetime.now().strftime('%Y-%m-%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        except:
            st.error("⚠️ A 'kiszedes_sablon.xlsx' hiányzik!")
