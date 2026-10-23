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

    # ---------- CONFRONTO RIGA PER RIGA ----------
    # Gli importi sono arrotondati al centesimo, come fa la busta paga.
    st.subheader("Controllo ore e importi")
    righe, effetti, non_lette = [], [], []
    tot_busta = tot_atteso = 0.0
    for v in VOCI:
        b = busta["voci"][v["chiave"]]
        ore_att = ore[v["chiave"]]
        tar_att = tariffa(paga, v)
        D = round(ore_att * tar_att, 2)                  # importo atteso dal foglio
        diff_riga = None
        if b:
            A = b["importo"]                             # importo scritto in busta
            B = round(b["ore"] * b["tariffa"], 2)        # ore x tariffa della busta
            C = round(b["ore"] * tar_att, 2)             # ore della busta con la tariffa attesa
            tot_busta += A
            tot_atteso += D
            diff_riga = round(A - D, 2)
            cause = [
                (round(C - D, 2), f"{v['nome']}: la busta paga {b['ore']:.2f} ore, il foglio ne conta "
                                  f"{ore_att:.2f} ({b['ore'] - ore_att:+.2f} ore)"),
                (round(B - C, 2), f"{v['nome']}: tariffa in busta {b['tariffa']:.5f} invece di {tar_att:.5f}"),
                (round(A - B, 2), f"{v['nome']}: l'importo in busta ({A:.2f}) non è ore x tariffa ({B:.2f})"),
            ]
            for euro, testo in cause:
                if abs(euro) >= 0.01:
                    effetti.append((testo, euro))
        else:
            non_lette.append(v["nome"])
        righe.append({
            "Voce": v["nome"],
            "Ore busta": b["ore"] if b else None,
            "Ore foglio": round(ore_att, 2),
            "Importo busta (€)": b["importo"] if b else None,
            "Importo atteso (€)": D,
            "Differenza (€)": diff_riga,
        })
    st.dataframe(pd.DataFrame(righe), hide_index=True)
    if non_lette:
        st.warning("Non ho trovato nel PDF queste righe: " + ", ".join(non_lette)
                   + ". I totali qui sotto non le comprendono.")

    # ---------- VANTAGGI E SVANTAGGI ----------
    st.subheader("Vantaggi e svantaggi per te")
    vantaggio = sum(e for _, e in effetti if e > 0)
    svantaggio = -sum(e for _, e in effetti if e < 0)
    saldo = round(tot_busta - tot_atteso, 2)
    m1, m2, m3 = st.columns(3)
    m1.metric("A tuo vantaggio", f"{vantaggio:.2f} €")
    m2.metric("A tuo svantaggio", f"{svantaggio:.2f} €")
    m3.metric("Saldo", f"{saldo:+.2f} €")

    if not effetti:
        st.success("✅ Nessuna differenza: busta e foglio coincidono.")
    for testo, euro in effetti:
        if euro > 0:
            st.success(f"A tuo VANTAGGIO: la busta paga {euro:.2f} € in più. {testo}")
        else:
            st.error(f"A tuo SVANTAGGIO: la busta paga {abs(euro):.2f} € in meno. {testo}")

    # caso tipico: ore spostate tra la riga festiva diurna e la riga notturna festiva
    bf, bn = busta["voci"]["fest"], busta["voci"]["nottfest"]
    if bf and bn:
        d_f, d_n = bf["ore"] - ore["fest"], bn["ore"] - ore["nottfest"]
        if abs(d_f) > 0.01 and abs(d_f + d_n) <= 0.01:
            if d_f > 0:
                st.info(f"La busta ha spostato {abs(d_f):.2f} ore dalla riga 'notturno festivo' alla riga "
                        "'festivo diurno': quelle ore sono pagate come festivo diurno (tariffa più alta) "
                        "invece che con la sola maggiorazione notturna.")
            else:
                st.info(f"La busta ha spostato {abs(d_f):.2f} ore dalla riga 'festivo diurno' alla riga "
                        "'notturno festivo': quelle ore sono pagate con la sola maggiorazione notturna "
                        "(tariffa più bassa) invece che come festivo diurno.")
    if not non_lette:
        verso = "più alta" if saldo > 0 else "più bassa" if saldo < 0 else "uguale"
        st.caption(f"Totale delle 4 righe: busta {tot_busta:.2f} €, foglio {tot_atteso:.2f} € "
                   f"(busta {verso}). Il calcolo del foglio è una stima: per decidere se chiedere "
                   "una correzione confrontati con l'ufficio paghe o con il sindacato.")

    # ---------- ALTRI CONTROLLI ----------
    st.subheader("Altri controlli")
    if busta["effettive"] is not None:
        d_ore = busta["effettive"] - lorde
        if abs(d_ore) <= 0.01:
            st.success(f"✅ Ore effettive lavorate: {busta['effettive']:.2f} in busta = {lorde:.2f} nel foglio (senza pausa).")
        else:
            stima = abs(d_ore) * paga
            if d_ore < 0:
                st.error(f"A tuo SVANTAGGIO: la busta conta {abs(d_ore):.2f} ore lavorate in meno del foglio "
                         f"({busta['effettive']:.2f} contro {lorde:.2f}): circa {stima:.2f} € in meno "
                         "(stima, in parte già compresa nelle righe sopra).")
            else:
                st.success(f"A tuo VANTAGGIO: la busta conta {d_ore:.2f} ore lavorate in più del foglio "
                           f"({busta['effettive']:.2f} contro {lorde:.2f}): circa {stima:.2f} € in più "
                           "(stima, in parte già compresa nelle righe sopra).")
    for v in VOCI:
        b = busta["voci"][v["chiave"]]
        if b and abs(b["tariffa"] - tariffa(paga, v)) < 0.0001 and abs(b["ore"] * b["tariffa"] - b["importo"]) < 0.02:
            st.success(f"✅ {v['nome']}: tariffa {b['tariffa']:.5f} e importo corretti.")

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