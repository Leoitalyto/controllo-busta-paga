import streamlit as st
import pandas as pd
import pdfplumber
import re

# Configurazione della pagina
st.set_page_config(page_title="Controllo Busta Paga FISE", page_icon="📊")

st.title("📊 Controllo Busta Paga FISE")
st.write("Carica il tuo file Excel delle timbrature e il PDF della busta paga per il controllo automatico.")

# --- CARICAMENTO FILE ---
excel_file = st.file_uploader("1. Carica il file Excel delle timbrature", type=["xlsx"])
pdf_file = st.file_uploader("2. Carica il PDF della busta paga", type=["pdf"])

if excel_file and pdf_file:
    # --- 1. ANALISI EXCEL (Calcolo spettante) ---
    df = pd.read_excel(excel_file)
    
    # Tariffe basate sulla tua busta paga
    paga_base = 13.22260
    tariffa_straord = paga_base * 1.20
    
    # Calcolo ore totali e straordinari
    df['Data'] = pd.to_datetime(df['Data'])
    df['Settimana'] = df['Data'].dt.strftime('%Y-%U')
    
    settimane = df.groupby('Settimana')['Ore_Lavorate'].sum().reset_index()
    
    ore_straord_calcolate = 0
    for index, row in settimane.iterrows():
        if row['Ore_Lavorate'] > 38:
            ore_straord_calcolate += (row['Ore_Lavorate'] - 38)
            
    euro_straord_spettanti = ore_straord_calcolate * tariffa_straord

    # --- 2. ANALISI PDF (Estrazione dati busta paga) ---
    # Usiamo l'OCR/estrazione testo per trovare le voci
    testo_pdf = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            testo_pdf += page.extract_text() or ""
            
    # Ricerca dei valori nel testo del PDF (Regex)
    # Cerca il numero dopo "STRAORDINARIO 20%" o simile
    match_ore_pagate = re.search(r'STRAORDINARIO\s*20%\s+\d+\s+([\d,\.]+)', testo_pdf)
    match_euro_pagati = re.search(r'STRAORDINARIO\s*20%\s+[\d,]+\s+[\d,]+\s+([\d,]+)', testo_pdf)
    
    # --- 3. CONFRONTO ---
    st.subheader("Risultato del Controllo")
    
    col1, col2 = st.columns(2)
    col1.metric("Straordinario Spettante (Calcolato)", f"{euro_straord_spettanti:.2f} €")
    
    if match_euro_pagati:
        euro_pagati = float(match_euro_pagati.group(1).replace(',', '.'))
        col2.metric("Straordinario Pagato (In Busta)", f"{euro_pagati:.2f} €")
        
        differenza = euro_straord_spettanti - euro_pagati
        
        if differenza > 0.5:
            st.error(f"⚠️ Attenzione! Mancano circa {differenza:.2f} € di straordinario.")
        elif differenza < -0.5:
            st.success(f"✅ Ti è stato pagato più del dovuto ({abs(differenza):.2f} € in più).")
        else:
            st.success("✅ Gli importi coincidono perfettamente!")
    else:
        st.warning("Non sono riuscito a leggere l'importo dello straordinario dal PDF. Controlla che il PDF sia leggibile (non scansionato come immagine).")
