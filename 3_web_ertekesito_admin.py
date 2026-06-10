import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

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

@st.cache_data(ttl=0)
def get_firebase_data():
    try:
        adatok = {}
        docs = db.collection("keszlet").stream()
        for doc in docs:
            adatok[doc.id] = int(doc.to_dict().get("mennyiseg", 0))
        return adatok
    except: return {}

def get_matrix(adatok, w):
    # 1. Mátrix létrehozása számokkal
    matrix = pd.DataFrame(0, index=hardnesses, columns=sizes)
    for m in sizes:
        for k in hardnesses:
            matrix.at[k, m] = adatok.get(f"{m}_{w}_{k}", 0)
    
    # 2. ÖSSZEGZÉS: Ezt még akkor végezzük, amikor a táblázat tisztán számokból áll!
    total_sum = matrix.values.sum()
    total_row_values = matrix.sum(axis=0).to_dict()
    
    # 3. NULLÁK ELREJTÉSE: Csak most cseréljük le a 0-kat szövegre
    matrix = matrix.replace(0, "")
    
    # 4. Final DF építése
    final_df = matrix.copy()
    final_df.insert(0, "Keménység", hardnesses)
    final_df["Keménység "] = hardnesses 
    
    # 5. Összesítő sor hozzáadása
    total_row = total_row_values
    total_row["Keménység"] = "ÖSSZESEN"
    total_row["Keménység "] = str(total_sum)
    
    final_df = pd.concat([final_df, pd.DataFrame([total_row])], ignore_index=True)
    return final_df
    
def excel_export_gomb(df, nev):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Keszlet')
        
        # Hozzáférés az Excel fájlhoz és munkalaphoz
        workbook = writer.book
        worksheet = writer.sheets['Keszlet']
        
        # 1. Szegély formátum definiálása
        cell_fmt = workbook.add_format({'border': 1})
        
        # 2. Fejléc formátum (szegéllyel + háttérrel)
        header_fmt = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D3D3D3', 'align': 'center'})
        
        # 3. KÉZI KERETEZÉS: Minden cellára ráhúzzuk a keretet
        rows, cols = df.shape
        # Végigmegyünk az összes soron és oszlopon, és alkalmazzuk a cell_fmt-et
        for r in range(rows):
            for c in range(cols):
                worksheet.write(r + 1, c, df.iloc[r, c], cell_fmt)
        
        # 4. Fejléc újraírása (hogy a fejléc is szép legyen)
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)

    st.download_button(
        label=f"📥 {nev} exportálása Excelbe", 
        data=buffer.getvalue(), 
        file_name=f"{nev}.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
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

    # Frissítés gomb
    if st.button("🔄 Adatok frissítése"):
        st.cache_data.clear()
        st.rerun()

    # 1. Speciális készlet
    st.subheader("⚠️ ÉRTÉKESÍTHETŐ SPECIÁLIS KÉSZLET")
    col1, col2, col3 = st.columns(3)
    spec_data = {
        "V-LV": [["7W FLX", "5 pár"], ["6XXW REG", "1 pár"], ["8XW XTR", "1 pár"], ["11XW SUP", "1 pár"]],
        "U-LV": [["8W XFR", "1 pár"], ["8W REG", "2 pár"]],
        "U-DV": [["8M SFT", "8 pár"], ["8M STR", "1 pár"], ["9M STR", "3 pár"], ["9W STR", "3 pár"]],
        "U-DV-2": [["8W XST", "1 pár"], ["11XXW XST", "1 pár"], ["11W FLX", "1 pár"], ["11W STR", "1 pár"]],
        "V-DV": [["8W 1/2 XTR", "1 pár"], ["9XW 1/2 XTR", "2 pár"], ["10XW 1/2 XTR", "1 pár"], ["9XXW 2/3 REG", "1 pár"], ["9W REG H-CR", "1 pár"]]
    }
    with col1:
        st.info("### V-LV"); st.table(spec_data["V-LV"])
        st.info("### U-LV"); st.table(spec_data["U-LV"])
    with col2:
        st.success("### U-DV (1. rész)"); st.table(spec_data["U-DV"])
    with col3:
        st.success("### U-DV (2. rész) & V-DV")
        st.table(spec_data["U-DV-2"]); st.table(spec_data["V-DV"])
    
    st.divider()

    # 2. Színező függvény
    def szinezo(row):
        # Fontos: A függvény törzsének minden sora 4 vagy 8 szóközzel beljebb kell legyen!
        style = [f'background-color: {kemenyseg_szinek.get(row["Keménység"], "#FFFFFF")}; text-align: center'] * len(row)
        
    # 2. Adatok betöltése
    adatok = get_firebase_data()

    # 3. Összes export gomb (egy fülön, egymás alatt)
    if st.button("📥 Összes export egyetlen táblázatba"):
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet('Osszes_Keszlet')
            
            cell_fmt = workbook.add_format({'border': 1, 'align': 'center'})
            header_fmt = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D3D3D3', 'align': 'center'})
            title_fmt = workbook.add_format({'bold': True, 'font_size': 12})
            
            current_row = 0
            for w in widths:
                df_temp = get_matrix(adatok, w)
                worksheet.write(current_row, 0, f"Szélesség: {w}", title_fmt)
                current_row += 1
                for col_num, value in enumerate(df_temp.columns.values):
                    worksheet.write(current_row, col_num, value, header_fmt)
                current_row += 1
                rows, cols = df_temp.shape
                for r in range(rows):
                    for c in range(cols):
                        worksheet.write(current_row + r, c, df_temp.iloc[r, c], cell_fmt)
                current_row += rows + 2
        
        st.download_button(
            label="✅ Letöltés (Egyetlen fülön)",
            data=buffer.getvalue(),
            file_name="Osszes_Keszlet_Egyben.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # 4. Színező függvény
    def szinezo(row):
        k = row["Keménység"]
        szin = kemenyseg_szinek.get(k, "#FFFFFF")
        style = [f'background-color: {szin}; text-align: center'] * len(row)
        if k == "ÖSSZESEN":
            style = ['background-color: #f0f0f0; font-weight: bold; text-align: center'] * len(row)
        return style

    # 5. Készlet táblázatok megjelenítése
    for w in widths:
        st.subheader(f"📦 \"{w}\" Szélesség")
        df = get_matrix(adatok, w)
        
        st.dataframe(
            df.style.apply(szinezo, axis=1).hide(axis="index"),
            use_container_width=True
        )
        
        excel_export_gomb(df, f"Keszlet_{w}")
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
                    ws.merge_cells(start_row=kezdo_sor, start_column=1, end_row=kezdo_sor, end_column=16)
                    ws.cell(row=kezdo_sor, column=1, value=cim).font = Font(bold=True, size=14)
                    ws.cell(row=kezdo_sor, column=1).alignment = center
                    
                    # Napok fejlécei
                    napok = ["Hétfő", "Kedd", "Szerda", "Csütörtök", "Péntek"]
                    for i, nap in enumerate(napok):
                        col = i*3 + 1
                        ws.merge_cells(start_row=kezdo_sor+1, start_column=col, end_row=kezdo_sor+1, end_column=col+2)
                        ws.cell(row=kezdo_sor+1, column=col, value=nap).font = Font(bold=True)
                        ws.cell(row=kezdo_sor+1, column=col).alignment = center
                        ws.cell(row=kezdo_sor+2, column=col, value="MÉRET")
                        ws.cell(row=kezdo_sor+2, column=col+1, value="KEM.")
                        ws.cell(row=kezdo_sor+2, column=col+2, value="DB")

                    osszes_termek = sorted(list(set((k[1], k[2]) for k in adat_szotar.keys())))
                    data_start = kezdo_sor + 3
                    osszes_sor = data_start + len(osszes_termek)
                    
                    # 1. Adatok kiírása
                    for i, (msz, kem) in enumerate(osszes_termek):
                        sor = data_start + i
                        for nap_index in range(5):
                            c = nap_index * 3 + 1
                            val = adat_szotar.get((nap_index, msz, kem), 0)
                            if val > 0:
                                ws.cell(row=sor, column=c, value=str(msz).upper()).font = Font(bold=True)
                                ws.cell(row=sor, column=c+1, value=str(kem).upper()).font = Font(bold=True)
                                ws.cell(row=sor, column=c+2, value=int(val)).font = Font(bold=True)
                    
                    # 2. Képletek (Összesítők)
                    for nap_index in range(5):
                        c = nap_index * 3 + 3
                        r_str = f"{ws.cell(row=data_start, column=c).coordinate}:{ws.cell(row=osszes_sor-1, column=c).coordinate}"
                        ws.cell(row=osszes_sor, column=c, value=f"=SUM({r_str})").font = Font(bold=True)
                    
                    ws.cell(row=osszes_sor, column=16, value=f"=SUM(C{osszes_sor},F{osszes_sor},I{osszes_sor},L{osszes_sor},O{osszes_sor})").font = Font(bold=True)

                    # 3. Keretezés (Vastag választóvonalakkal)
                    thin = Side(style='thin')
                    thick = Side(style='thick')
                    
                    for r in range(kezdo_sor, osszes_sor + 1):
                        for c in range(1, 17):
                            # Melyik oszlop a nap vége? (3, 6, 9, 12, 15)
                            is_day_end = (c % 3 == 0) and (c < 16)
                            border_style = Border(
                                left=thin, top=thin, bottom=thin, 
                                right=thick if is_day_end else thin
                            )
                            ws.cell(row=r, column=c).border = border_style
                    
                    # Oszlopszélesség
                    for col_num in range(1, 17):
                        ws.column_dimensions[openpyxl.utils.get_column_letter(col_num)].width = 9
                        
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
