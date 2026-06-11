center = Alignment(horizontal="center", vertical="center")

def iras_blokkba(adat_szotar, kezdo_sor, cim):
                    # 1. Cím és fejlécek
                    ws.merge_cells(start_row=kezdo_sor, start_column=1, end_row=kezdo_sor, end_column=16)
                    ws.cell(row=kezdo_sor, column=1, value=cim).font = Font(bold=True, size=14)
                    ws.cell(row=kezdo_sor, column=1).alignment = center
                    
                    napok = ["Hétfő", "Kedd", "Szerda", "Csütörtök", "Péntek"]
                    for i, nap in enumerate(napok):
                        col = i*3 + 1
                        ws.merge_cells(start_row=kezdo_sor+1, start_column=col, end_row=kezdo_sor+1, end_column=col+2)
                        ws.cell(row=kezdo_sor+1, column=col, value=nap).font = Font(bold=True)
                        ws.cell(row=kezdo_sor+1, column=col).alignment = center
                        ws.cell(row=kezdo_sor+2, column=col, value="MÉRET"); ws.cell(row=kezdo_sor+2, column=col+1, value="KEM."); ws.cell(row=kezdo_sor+2, column=col+2, value="DB")
    # 1. Cím és Fejlécek (Fix pozíciók)
    ws.merge_cells(start_row=kezdo_sor, start_column=1, end_row=kezdo_sor, end_column=15)
    ws.cell(row=kezdo_sor, column=1, value=cim).font = Font(bold=True, size=14)
    ws.cell(row=kezdo_sor, column=1).alignment = center
    
    napok = ["Hétfő", "Kedd", "Szerda", "Csütörtök", "Péntek"]
    for i, nap in enumerate(napok):
        col_start = i * 3 + 1
        ws.merge_cells(start_row=kezdo_sor+1, start_column=col_start, end_row=kezdo_sor+1, end_column=col_start+2)
        ws.cell(row=kezdo_sor+1, column=col_start, value=nap).font = Font(bold=True)
        ws.cell(row=kezdo_sor+1, column=col_start).alignment = center
        ws.cell(row=kezdo_sor+2, column=col_start, value="MÉRET"); ws.cell(row=kezdo_sor+2, column=col_start+1, value="KEM."); ws.cell(row=kezdo_sor+2, column=col_start+2, value="DB")

    # 2. Összes termék keresése az adatbázisból (hogy a táblázat fix 10-15 soros legyen)
    osszes_termek = sorted(list(set((k[1], k[2]) for k in adat_szotar.keys())))
    data_start = kezdo_sor + 3
    
    # 3. KÉZI ÍRÁS - Itt nem számolunk, csak kitöltünk egy rácsot
    for i, (msz, kem) in enumerate(osszes_termek):
        row = data_start + i
        for nap_index in range(5):
            col = nap_index * 3 + 1
            val = adat_szotar.get((nap_index, msz, kem), 0)
            
            # Formázás: csak akkor írunk, ha van adat
            if val > 0:
                ws.cell(row=row, column=col, value=msz.upper())
                ws.cell(row=row, column=col+1, value=kem.upper())
                ws.cell(row=row, column=col+2, value=int(val))
    
    # 4. ÖSSZESÍTŐK (Ez a sor lesz a táblázat alja)
    osszes_sor = data_start + len(osszes_termek)
    for nap_index in range(5):
        col = nap_index * 3 + 3
        # Itt fixen összegezzük az oszlopot az első 15 sorban
        r_str = f"{ws.cell(row=data_start, column=col).coordinate}:{ws.cell(row=osszes_sor-1, column=col).coordinate}"
        ws.cell(row=osszes_sor, column=col, value=f"=SUM({r_str})").font = Font(bold=True)
    
    # 5. Keretezés (Fix rács 5 nap x 3 oszlop)
    for r in range(kezdo_sor, osszes_sor + 1):
        for c in range(1, 16):
            ws.cell(row=r, column=c).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

                    # 2. Összesített terméklista (hogy minden nap "tudja", hová kell írni)
                    osszes_termek = sorted(list(set((k[1], k[2]) for k in adat_szotar.keys())))
                    data_start = kezdo_sor + 3
                    
                    # 3. Adatok kiírása: CSAK HA VAL > 0
                    for i, (msz, kem) in enumerate(osszes_termek):
                        sor = data_start + i
                        for nap_index in range(5):
                            c = nap_index * 3 + 1
                            val = adat_szotar.get((nap_index, msz, kem), 0)
                            if val > 0: # <-- CSAK A POZITÍVOK KERÜLNEK BE
                                ws.cell(row=sor, column=c, value=str(msz).upper())
                                ws.cell(row=sor, column=c+1, value=str(kem).upper())
                                ws.cell(row=sor, column=c+2, value=int(val))
                    
                    # 4. Keretezés: Csak azokat a cellákat keretezzük, ahol adat van
                    osszes_sor = data_start + len(osszes_termek)
                    thin = Side(style='thin')
                    thick = Side(style='thick')
                    for r in range(kezdo_sor, osszes_sor + 1):
                        for c in range(1, 16):
                            is_day_end = (c % 3 == 0)
                            ws.cell(row=r, column=c).border = Border(left=thin, top=thin, bottom=thin, right=thick if is_day_end else thin)
                    
                    return osszes_sor + 2
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
