import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore

# --- ALAP BEÁLLÍTÁSOK ---
st.set_page_config(page_title="Balettcipő Raktár", layout="wide")

# Firebase inicializálás (ugyanaz, ami korábban működött)
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

# --- ADATLEKÉRÉS ---
def get_adatok():
    docs = db.collection("keszlet").stream()
    return {doc.id: int(doc.to_dict().get("mennyiseg", 0)) for doc in docs}

# --- MENÜ ---
menu = st.sidebar.radio("Navigáció", ["📱 Raktár", "📊 Értékesítő", "🔐 Admin"])

if menu == "📱 Raktár":
    st.title("📱 Raktári Mozgás")
    # Ide kerülhet a korábbi kiszedő/visszarakó logikád

elif menu == "📊 Értékesítő":
    st.title("📊 Értékesítői Nézet")
    adatok = get_adatok()
    st.write("Itt láthatod a készletet:")
    # Példa egy egyszerű táblázatra, ami biztosan megjelenik:
    df = pd.DataFrame.from_dict(adatok, orient='index', columns=['Mennyiség'])
    st.table(df)

elif menu == "🔐 Admin":
    st.title("🔐 Adminisztráció")
    jelszo = st.text_input("Jelszó:", type="password")
    if jelszo == "admin123":
        st.success("Sikeres belépés!")
        # Ide jöhet a heti riport gombja
    else:
        st.info("Kérlek, add meg a jelszót a belépéshez.")
