import re
importa pandas come pd
importare pdfplumber
import streamlit as st

# ---------------------------------------------------------------
# PARTE 1 - FUNZIONI (leggono il foglio Excel e il PDF)
# ---------------------------------------------------------------
# Maggiorazioni CCNL FISE Ambientale: (nome, colonna nel foglio Excel, percentuale)
CATEGORIA = [
    ("Straordinario diurno", "Straord. diurno", 20),
    ("Straordinario festivo", "Straord. festivo", 65),
    ("Straordinario notturno", "Straord. notturno", 33),
    ("Straordinario notturno festivo", "Straord. notturno festivo", 75),
]


def numero_it(testo):
    """Trasforma '1.234,56' in 1234.56"""
    restituisci float(testo.replace(".", "").replace(",", "."))


def leggi_excel(file):
    """Legge il foglio 'Ore' (intestazioni alla riga 5, 31 righe di giorni).
    Se il file ha il vecchio formato (colonne Data e Ore_Lavorate) lo legge come prima."""
    xls = pd.ExcelFile(file)
    se "Ore" è presente in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name="Ore", header=4, nrows=31)
        df.columns = [str(c).strip() per c in df.columns]
        nuovo = Vero
    altro:
        df = pd.read_excel(xls)
        df.columns = [str(c).strip() per c in df.columns]
        df = df.rename(columns={"Minerario_Lavorate": "Minerario lavorato"})
        nuovo = Falso
    se "Data" non è presente in df.columns o "Ore lavorate" non è presente in df.columns:
        raise ValueError("Nel file non trovo le colonne 'Data' e 'Ore lavorate'. Colonne trovate: "
                         + ", ".join(df.columns))
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df = df.dropna(subset=["Data"]).copy() # tolgo le righe vuote
    df["Minerario lavorato"] = pd.to_numeric(df["Minerario lavorato"], errori="coercire").fillna(0)
    per _, col, _ in CATEGORIA:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0) if col in df.columns else 0.0
    restituisci df, nuovo


def ore_per_settimana(df, soglia):
    """Somma le ore per settimana (lunedi-domenica) e calcola lo straordinario oltre la soglia."""
    d = df.copy()
    d["Inizio settimana"] = d["Data"].dt.to_period("W-SUN").apply(lambda p: p.start_time.date())
    s = d.groupby("Inizio settimana")["Ore lavorate"].sum().reset_index()
    s["Oltre soglia"] = (s["Ore lavorate"] - soglia).clip(lower=0)
    ritorno s


def cerca_voce_pdf(testo, percentuale):
    """Cerca nel PDF la riga dello straordinario con quella percentuale e restituisce
    (riga trovata, importo in euro). L'importo e' il terzo numero dopo la percentuale."""
    schema = (r"STRAORDINARIO[^\n]*?(?<!\d)" + str(percentuale) +
              r"\s*%\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)")
    m = re.search(schema, testo, flags=re.IGNORECASE)
    se non m:
        restituisci Nessuno, Nessuno
    Tentativo:
        restituisci m.group(0), numero_it(m.group(3))
    eccetto ValueError:
        restituisci m.group(0), Nessuno


# --- FUNZIONI FINI ---

# ---------------------------------------------------------------
# PARTE 2 - PAGINA DELL'APP
# ---------------------------------------------------------------
st.set_page_config(page_title="Controllo Busta Paga FISE", page_icon="📊")
st.title("📊 Controllo Busta Paga FISE")
st.write("Carica il foglio Excel delle ore e il PDF della busta paga per il controllo automatico.")

paga_base = st.sidebar.number_input("Paga oraria base (€)", valore=13.22260, format="%.5f")
soglia = st.sidebar.number_input("Soglia settimanale straordinaria (ore)", value=38.0, step=0.5)

excel_file = st.file_uploader("1. Carica il file Excel delle ore", type=["xlsx"])
pdf_file = st.file_uploader("2. Carica il PDF della busta paga", type=["pdf"])

se file_excel e file_pdf:
    # --- 1. EXCEL ---
    Tentativo:
        df, nuovo = leggi_excel(file_excel)
   exception Except as e:
        st.error(f"Non riesco a leggere il file Excel: {e}")
        st.stop()
    se df è vuoto:
        st.warning("Nel foglio non ci sono date compilate.")
        st.stop()

    settimane = ore_per_settimana(df, soglia)
    straord_totale_settimane = settimane["Oltre soglia"].sum()

    # --- 2. PDF ---
    testo_pdf = ""
    con pdfplumber.open(pdf_file) come pdf:
        per pagina in pdf.pages:
            testo_pdf += (page.extract_text() or "") + "\n"

    # --- 3. CONFRONTO ---
    st.subheader("Risultato del controllo")
    ricci = []
    per nome, col, pct in CATEGORIE:
        minerale = float(df[col].sum())
        spettante = minerale * paga_base * (1 + pct/100)
        riga_pdf, pagato = cerca_voce_pdf(testo_pdf, pct)
        righe.append({"Voce": f"{nome} {pct}%", "Ore calcolate": round(ore, 2),
                      "Spettante (€)": rotondo(spettante, 2),
                      "Pagato in busta (€)": pagato, "riga": riga_pdf})
    tab = pd.DataFrame(righe)

    totale_spettante = tab["Spettante (€)"].sum()
    letti = tab[tab["Pagato in busta (€)"].notna()]
    col1, col2 = st.columns(2)
    col1.metric("Straordinario spettante (calcolato)", f"{totale_spettante:.2f} €")

    se non letti.empty:
        pagato_tot = float(letti["Pagato in busta (€)"].sum())
        spettante_letti = float(letti["Spettante (€)"].sum())
        col2.metric("Straordinario pagato (voci lette)", f"{pagato_tot:.2f} €")
        differenza = spettante_letti - pagato_tot
        se la differenza è maggiore di 0,5:
            st.error(f"⚠️ Attenzione! Mancano circa {differenza:.2f} € di straordinario.")
        elif differenza < -0.5:
            st.success(f"✅ Ti è stato pagato più del dovuto ({abs(differenza):.2f} € in più).")
        altro:
            st.success("✅ Gli importi coincidono!")
        se len(letti) < len(tab):
            st.info("Alcune voci hanno ore calcolate ma non le ho trovate nel PDF: "
                    "il confronto riguarda solo le voci lette (vedi tabella).")
    altro:
        st.warning("Non sono riuscito a leggere lo straordinario dal PDF. "
                   "Controlla che il PDF sia leggibile (non una scansione-immagine).")

    st.dataframe(tab.drop(columns=["riga"]), hide_index=True)

    st.subheader("Ore per settimana (lunedì-domenica)")
    st.dataframe(settimane.round(2), hide_index=True)
    st.caption(f"Soglia usata: {soglia:g} ore. Le settimane a cavallo di due mesi contano solo i giorni del foglio.")
    if nuovo and abs(straord_totale_settimane - float(df["Ore straordinario"].sum() if "Ore straordinario" in df.columns else 0)) > 0.01:
        st.info("Lo straordinario del foglio Excel e quello ricalcolato qui non coincidono: "
                "controlla che la soglia nell'app sia uguale a quella del foglio (Parametri).")

    con st.expander("Testo letto dal PDF (per controllo)"):
        per _, r in tab.iterrows():
            se r["riga"]:
                st.write(f"Riga trovata per {r['Voce']}:", r["riga"])
        st.text(testo_pdf[:4000])
