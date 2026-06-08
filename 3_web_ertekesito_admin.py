import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
from datetime import datetime

# --- JELSZÓ BEÁLLÍTÁSA ---
ADMIN_JELSZO = "admin123"

# --- 1. OLDAL BEÁLLÍTÁSAI ---
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

# --- 3. FIX ADATOK ---
widths = ["M", "W", "XW", "XXW"]
sizes = [str(i) for i in range(5, 15)] 
hardnesses = ["LGH", "SFT", "FLX", "SUP", "REG", "FRM", "STR", "XFR", "XST"]

kemenyseg_szinek = {
    "LGH": "#FFD1DC", "SFT": "#FFFFFF", "FLX": "#FF91A4", 
    "SUP": "#E0E0E0", "REG": "#FFC000", "FRM": "#CD7F32", 
    "STR": "#4682B4", "XFR": "#A6A6A6", "XST": "#CC0000"
}

# Színező függvény a táblázathoz
def szinezo(val, row_name):
    szin = kemenyseg_szinek.get(row_name, "#FFFFFF")
    return f'background-color: {szin}'

# --- 4. FUNKCIÓVÁLASZTÓ ---
st.sidebar.header("✨ Navigáció")
funkcio = st.sidebar.radio("Válassz felületet:", [
    "📱 Raktári Kiszedés (Gombos)",
    "📊 Értékesítő (Csak olvasható)",
    "🔐 Admin (Szerkeszthető)"
])

@st.cache_data(ttl=30)
def get_firebase_data():
    try:
        adatok = {}
        docs = db.collection("keszlet").stream()
        for doc in docs:
            adatok[doc.id] = int(doc.to_dict().get("mennyiseg", 0))
        return adatok
    except Exception as e:
        st.error(f"Adatbázis elérési hiba: {e}")
        return {}

# ==============================================================================
# SEGÉDFÜGGVÉNY: Mátrix generálás
# ==============================================================================
def get_matrix(adatok, w):
    # Létrehozzuk a mátrixot, majd kiegészítjük az oszlopokkal
    matrix = pd.DataFrame(0, index=hardnesses, columns=sizes)
    for m in sizes:
        for k in hardnesses:
            matrix.at[k, m] = adatok.get(f"{m}_{w}_{k}", 0)
    
    # Új oszlopok: Keménység , Méretek, Keménység 
    final_df = matrix.copy()
    final_df.insert(0, "Keménység ", hardnesses)
    final_df["Keménység "] = hardnesses
    
    # Összesítő sor (csak a méret-oszlopokra számolva)
    total_row = matrix.sum(axis=0).to_dict()
    total_row["Keménység "] = "ÖSSZESEN"
    total_row["Keménység "] = "ÖSSZESEN"
    
    # Sor hozzáadása a df-hez
    final_df = pd.concat([final_df, pd.DataFrame([total_row])], ignore_index=True)
    return final_df

# ==============================================================================
# A) RAKTÁRI GOMBOS FELÜLET
# ==============================================================================
if funkcio == "📱 Raktári Kiszedés (Gombos)":
    st.title("📱 Raktári Mozgás Rögzítése")
    adatok = get_firebase_data()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        valasztott_meret = st.radio("Méretek:", sizes, key="r_size")
    with col2:
        valasztott_szelesseg = st.radio("Szélességek:", widths, key="r_width")
    with col3:
        valasztott_kemenyseg = st.radio("Keménységek:", hardnesses, key="r_hard")
        
    sku_id = f"{valasztott_meret}_{valasztott_szelesseg}_{valasztott_kemenyseg}"
    aktualis_keszlet = adatok.get(sku_id, 0)
    
    st.write("---")
    st.info(f"Kiválasztott cipő: **{sku_id}** | 📦 Aktuális készlet: **{aktualis_keszlet} db**")
    
    b_col1, b_col2 = st.columns(2)
    ma_szoveg = datetime.now().strftime("%Y-%m-%d")
    
    if b_col1.button("❌ KISZEDÉS (-1 db)", type="primary") and aktualis_keszlet > 0:
        db.collection("keszlet").document(sku_id).set({"mennyiseg": aktualis_keszlet - 1}, merge=True)
        db.collection("naplo").add({"datum": ma_szoveg, "sku": sku_id, "tipus": "kiszedes", "darabszam": 1})
        st.rerun()
    if b_col2.button("✅ VISSZARAKÁS (+1 db)"):
        db.collection("keszlet").document(sku_id).set({"mennyiseg": aktualis_keszlet + 1}, merge=True)
        db.collection("naplo").add({"datum": ma_szoveg, "sku": sku_id, "tipus": "visszarakas", "darabszam": 1})
        st.rerun()

# ==============================================================================
# B) ÉRTÉKESÍTŐ FELÜLET
# ==============================================================================
# ==============================================================================
# B) ÉRTÉKESÍTŐ FELÜLET (Javított, színes verzió)
# ==============================================================================
elif funkcio == "📊 Értékesítő (Csak olvasható)":
    st.title("📊 Balettcipő Élő Készlet")
    adatok = get_firebase_data()
    
    for w in widths:
        st.subheader(f"📦 \"{w}\" Szélesség")
        df = get_matrix(adatok, w)
        
        # A stílus alkalmazása:
        # A 'Keménység ' és 'Keménység ' oszlopok értékei alapján színezzük a sorokat
        def szinezo_df(row):
            kemenyseg = row["Keménység "]
            szin = kemenyseg_szinek.get(kemenyseg, "#FFFFFF")
            return [f'background-color: {szin}'] * len(row)

        # Alkalmazzuk a stílust
        styled_df = df.style.apply(szinezo_df, axis=1)
        
        # Megjelenítés: a stílus objektumot adjuk át a st.dataframe-nek
        st.dataframe(styled_df, use_container_width=True)
# ==============================================================================
# C) ADMIN FELÜLET (Teljesen javított kód)
# ==============================================================================
elif funkcio == "🔐 Admin (Szerkeszthető)":
    st.title("🔐 Adminisztrátori Készletkezelés")
    
    # Jelszó védelem
    if st.sidebar.text_input("Jelszó:", type="password") == ADMIN_JELSZO:
        # Adatok lekérése a Firebase-ből
        adatok = get_firebase_data()
        
        for w in widths:
            st.subheader(f"🛠️ \"{w}\" szélesség szerkesztése")
            
            # Mátrix felépítése az adott szélességhez (összesítés nélkül a szerkesztéshez)
            matrix_df = pd.DataFrame(0, index=hardnesses, columns=sizes)
            for m in sizes:
                for k in hardnesses:
                    matrix_df.at[k, m] = adatok.get(f"{m}_{w}_{k}", 0)
            
            # Szerkeszthető tábla egyedi kulccsal (editor_ + szélesség)
            edited_df = st.data_editor(
                matrix_df, 
                use_container_width=True,
                key=f"editor_{w}"
            )
            
            # Mentés gomb egyedi kulccsal
            if st.button(f"💾 Mentés: {w}", key=f"mentes_{w}"):
                batch = db.batch()
                valtozas_tortent = False
                
                for m in sizes:
                    for k in hardnesses:
                        uj_ertek = int(edited_df.at[k, m])
                        regi_ertek = adatok.get(f"{m}_{w}_{k}", 0)
                        
                        # Csak azt frissítjük, ami változott
                        if uj_ertek != regi_ertek:
                            doc_ref = db.collection("keszlet").document(f"{m}_{w}_{k}")
                            batch.set(doc_ref, {"mennyiseg": uj_ertek})
                            valtozas_tortent = True
                
                if valtozas_tortent:
                    batch.commit()
                    st.success(f"✅ \"{w}\" szélesség sikeresen frissítve!")
                    st.rerun() # Frissítjük az oldalt a változások megjelenítéséhez
                else:
                    st.info("Nincs változás.")
