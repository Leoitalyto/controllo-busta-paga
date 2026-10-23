import re
import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import load_workbook

# ---------------------------------------------------------------
# PARTE 1 - FUNZIONI
# ---------------------------------------------------------------
# Le 4 righe della busta paga. "piena" = la riga paga ore x paga x (1 + %);
# se "piena" e' False la riga paga solo la maggiorazione (ore x paga x %).
VOCI = [
    {"chiave": "ord", "nome": "Straordinario 20%", "pct": 0.20, "piena": True,
     "schema": r"STRAORDINARIO\s*20\s*%"},
    {"chiave": "fest", "nome": "Straord. festivo diurno 65%", "pct": 0.65, "piena": True,
     "schema": r"STRAORD\.?\s*FEST\.?\s*DIURNO\s*65\s*%"},
    {"chiave": "nott", "nome": "Lavoro notturno 33%", "pct": 0.33, "piena": False,
     "schema": r"LAVORO\s+NOTTURNO\s*33\s*%"},
    {"chiave": "nottfest", "nome": "Lavoro notturno festivo 75%", "pct": 0.75, "piena": False,
     "schema": r"LAVORO\s+NOTT\.?\s*FEST\.?\s*75\s*%"},
]
NUM = r"([\d\.,]+)"


def numero_it(testo):
    """Trasforma '1.234,56' in 1234.56"""
    return float(testo.replace(".", "").replace(",", "."))


def tariffa(paga, voce):
    return paga * (1 + voce["pct"]) if voce["piena"] else paga * voce["pct"]


def leggi_excel(file):
    """Legge il foglio 'Ore' (intestazioni alla riga 5, 31 righe di giorni)."""
    xls = pd.ExcelFile(file)
    if "Ore" in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name="Ore", header=4, nrows=31)
        nuovo = True
    else:
        df = pd.read_excel(xls)          # vecchio formato
        df = df.rename(columns={"Ore_Lavorate": "Ore lavorate"})
        nuovo = False
    df.columns = [str(c).strip() for c in df.columns]
    if "Data" not in df.columns or "Ore lavorate" not in df.columns:
        raise ValueError("Non trovo le colonne 'Data' e 'Ore lavorate'. Colonne trovate: "
                         + ", ".join(df.columns))

    # --- CONTROLLI SULLE DATE E SUGLI ERRORI DEL FOGLIO ---
    primo_rigo = 6 if nuovo else 2          # numero della prima riga di dati in Excel
    if nuovo:
        try:
            if hasattr(file, "seek"):
                file.seek(0)
            ws_ctrl = load_workbook(file, data_only=True)["Ore"]
            errati = [c.coordinate for riga in ws_ctrl["A6:V36"] for c in riga if c.data_type == "e"]
        except Exception:
            errati = []
        if errati:
            raise ValueError("il foglio contiene errori di calcolo (#VALORE! o simili) nelle celle "
                             + ", ".join(errati[:8]) + (" ..." if len(errati) > 8 else "")
                             + ". Di solito una data non è una vera data: scrivila come gg/mm/aaaa.")
    # Se per caso le date sono testo, le leggo come giorno/mese/anno.
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True, format="mixed")
    ore_tmp = pd.to_numeric(df["Ore lavorate"], errors="coerce").fillna(0.0)
    senza_data = df.index[df["Data"].isna() & (ore_tmp != 0)] + primo_rigo
    if len(senza_data):
        raise ValueError("ci sono righe con orari ma senza una data valida (righe Excel: "
                         + ", ".join(str(int(r)) for r in senza_data) + "). Scrivi la data come gg/mm/aaaa.")
    df = df.dropna(subset=["Data"]).copy()

    def num(nome):
        if nome in df.columns:
            return pd.to_numeric(df[nome], errors="coerce").fillna(0.0)
        return pd.Series(0.0, index=df.index)

    df["Ore lavorate"] = num("Ore lavorate")
    lorde_presenti = "Ore lorde (senza pausa)" in df.columns
    df["Ore lorde"] = num("Ore lorde (senza pausa)") if lorde_presenti else df["Ore lavorate"]
    df["Notturne"] = num("di cui notturne")
    df["Ore assenza"] = num("Ore assenza")
    if "Festivo" in df.columns:
        df["Giorno festivo"] = df["Festivo"].fillna("").astype(str).str.strip().str.lower().isin(["sì", "si"])
    else:
        df["Giorno festivo"] = False
    return df, nuovo, lorde_presenti


def leggi_busta(testo):
    """Legge dal testo del PDF le 4 righe di straordinario/notturno e i dati del mese."""
    busta = {"voci": {}}
    for v in VOCI:
        m = re.search(v["schema"] + r"\s+" + NUM + r"\s+" + NUM + r"\s+" + NUM, testo, flags=re.IGNORECASE)
        try:
            busta["voci"][v["chiave"]] = (
                {"riga": m.group(0), "ore": numero_it(m.group(1)),
                 "tariffa": numero_it(m.group(2)), "importo": numero_it(m.group(3))} if m else None)
        except ValueError:
            busta["voci"][v["chiave"]] = None

    def uno(schema, gruppo=1):
        m = re.search(schema, testo, flags=re.IGNORECASE)
        try:
            return numero_it(m.group(gruppo)) if m else None
        except ValueError:
            return None

    busta["lavorabili"] = uno(r"ORE\s+LAVORABILI\s+" + NUM)
    busta["effettive"] = uno(r"ORE\s+EFFETT\.?\s+LAVORATE\s+" + NUM)
    busta["paga"] = uno(r"COMPETENZE\s+MESE\s+" + NUM + r"\s+" + NUM, 2)
    return busta


def ore_attese(df, lavorabili):
    """Calcola le ore delle 4 righe con il metodo MENSILE della busta:
    straordinario = ore effettive - (ore lavorabili - ore di assenza legge 104).
    Le ore festive vanno prima nella riga festiva, il resto al 20%.
    Il notturno e' pagato su tutte le ore notturne lavorate."""
    lorde = float(df["Ore lorde"].sum())
    assenze = float(df["Ore assenza"].sum())
    straord = max(0.0, lorde - (lavorabili - assenze))
    f = df["Giorno festivo"]
    fest_lorde = float(df.loc[f, "Ore lorde"].sum())
    notte_fest = float(df.loc[f, "Notturne"].sum())
    notte_fer = float(df.loc[~f, "Notturne"].sum())
    fest_tot = min(fest_lorde, straord)
    n_f = min(notte_fest, fest_tot)
    ore = {"ord": straord - fest_tot, "fest": fest_tot - n_f, "nott": notte_fer, "nottfest": n_f}
    return ore, lorde, assenze, straord


def ore_per_settimana(df, soglia):
    d = df.copy()
    d["Inizio settimana"] = d["Data"].dt.to_period("W-SUN").apply(lambda p: p.start_time.date())
    s = d.groupby("Inizio settimana")["Ore lavorate"].sum().reset_index()
    s["Oltre soglia"] = (s["Ore lavorate"] - soglia).clip(lower=0)
    return s


# --- FINE FUNZIONI ---

# ---------------------------------------------------------------
# PARTE 2 - PAGINA DELL'APP
# ---------------------------------------------------------------
st.set_page_config(page_title="Controllo Busta Paga FISE", page_icon="📊")
st.title("📊 Controllo Busta Paga FISE")
st.write("Carica il foglio Excel delle ore e il PDF della busta paga: l'app confronta ore e importi "
         "delle 4 righe di straordinario e lavoro notturno.")

st.sidebar.header("Solo se il PDF non viene letto")
paga_manuale = st.sidebar.number_input("Paga oraria base (€)", value=13.22260, format="%.5f")
lavorabili_manuale = st.sidebar.number_input("Ore lavorabili del mese", value=0.0, step=0.5)
soglia = st.sidebar.number_input("Soglia settimanale (solo tabella informativa)", value=38.0, step=0.5)

excel_file = st.file_uploader("1. Carica il file Excel delle ore", type=["xlsx"])
pdf_file = st.file_uploader("2. Carica il PDF della busta paga", type=["pdf"])

if excel_file and pdf_file:
    try:
        df, nuovo, lorde_presenti = leggi_excel(excel_file)
    except Exception as e:
        st.error(f"Non riesco a leggere il file Excel: {e}")
        st.stop()
    if df.empty:
        st.warning("Nel foglio non ci sono date compilate.")
        st.stop()

    testo_pdf = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            testo_pdf += (page.extract_text() or "") + "\n"
    busta = leggi_busta(testo_pdf)

    paga = busta["paga"] or paga_manuale
    lavorabili = busta["lavorabili"] or lavorabili_manuale
    if not lavorabili:
        st.error("Non trovo le ore lavorabili nel PDF. Scrivile nella barra laterale (menu in alto a sinistra).")
        st.stop()
    if not lorde_presenti:
        st.warning("Nel foglio Excel manca la colonna 'Ore lorde (senza pausa)': usa il foglio aggiornato, "
                   "altrimenti le ore lavorate risultano più basse della busta.")

    ore, lorde, assenze, straord = ore_attese(df, lavorabili)

    st.subheader("Controllo ore e importi")
    righe, ore_diverse, non_lette = [], [], []
    tot_busta = tot_atteso = 0.0
    for v in VOCI:
        b = busta["voci"][v["chiave"]]
        atteso_ore = ore[v["chiave"]]
        atteso_eur = atteso_ore * tariffa(paga, v)
        tot_atteso += atteso_eur
        if b:
            tot_busta += b["importo"]
            if abs(b["ore"] - atteso_ore) > 0.01:
                ore_diverse.append(f"{v['nome']}: busta {b['ore']:.2f} ore, foglio {atteso_ore:.2f} ore")
        else:
            non_lette.append(v["nome"])
        righe.append({
            "Voce": v["nome"],
            "Ore busta": b["ore"] if b else None,
            "Ore foglio": round(atteso_ore, 2),
            "Importo busta (€)": b["importo"] if b else None,
            "Importo atteso (€)": round(atteso_eur, 2),
        })
    tab = pd.DataFrame(righe)
    st.dataframe(tab, hide_index=True)

    c1, c2 = st.columns(2)
    c1.metric("Totale 4 voci in busta", f"{tot_busta:.2f} €")
    c2.metric("Totale calcolato dal foglio", f"{tot_atteso:.2f} €")
    diff = tot_busta - tot_atteso
    if non_lette:
        st.warning("Non ho trovato nel PDF queste righe: " + ", ".join(non_lette))
    if ore_diverse:
        st.warning("⚠️ Ore diverse tra busta e foglio (da controllare con l'ufficio paghe):\n\n- "
                   + "\n- ".join(ore_diverse))
    elif not non_lette:
        st.success("✅ Le ore delle 4 righe coincidono con il foglio.")
    st.caption(f"Differenza in euro rispetto al calcolo: {diff:+.2f} € "
               "(positivo = la busta è più alta). Dipende da come le ore sono divise tra le righe.")

    st.subheader("Altri controlli")
    if busta["effettive"] is not None:
        if abs(busta["effettive"] - lorde) <= 0.01:
            st.success(f"✅ Ore effettive lavorate: {busta['effettive']:.2f} in busta = {lorde:.2f} nel foglio (senza pausa).")
        else:
            st.warning(f"⚠️ Ore effettive lavorate: {busta['effettive']:.2f} in busta, {lorde:.2f} nel foglio (senza pausa).")
    for v in VOCI:
        b = busta["voci"][v["chiave"]]
        if not b:
            continue
        attesa = tariffa(paga, v)
        ok_tar = abs(b["tariffa"] - attesa) < 0.0001
        ok_imp = abs(b["ore"] * b["tariffa"] - b["importo"]) < 0.02
        if ok_tar and ok_imp:
            st.success(f"✅ {v['nome']}: tariffa {b['tariffa']:.5f} e importo corretti.")
        else:
            st.warning(f"⚠️ {v['nome']}: tariffa in busta {b['tariffa']:.5f} (attesa {attesa:.5f}); "
                       f"ore x tariffa = {b['ore'] * b['tariffa']:.2f} (in busta {b['importo']:.2f}).")

    st.info(f"Metodo della busta: straordinario = ore effettive {lorde:.2f} - (ore lavorabili {lavorabili:.2f} "
            f"- ore di assenza {assenze:.2f}) = {straord:.2f} ore. "
            "Le ore festive vanno nella riga festiva, le altre al 20%.")

    with st.expander("Calcolo settimanale oltre soglia (solo informativo)"):
        st.dataframe(ore_per_settimana(df, soglia).round(2), hide_index=True)
        st.caption("La busta non usa questo metodo: serve solo per confronto.")

    with st.expander("Testo letto dal PDF (per controllo)"):
        for v in VOCI:
            b = busta["voci"][v["chiave"]]
            st.write(v["nome"], "→", b["riga"] if b else "non trovata")
        st.text(testo_pdf[:4000])