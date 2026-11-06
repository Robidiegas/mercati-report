# -*- coding: utf-8 -*-
"""
Mercati Report
Genera in automatico una presentazione PowerPoint con grafici e tabelle di mercato
(inflazione, tassi, borse, valute, materie prime, cripto) da fonti gratuite:
Yahoo Finance (yfinance), FRED, SNB, Bundesbank, BIS.

Solo per uso personale. I dati Yahoo Finance sono soggetti ai termini d'uso di Yahoo.
"""

import io
import os
import re
import sys
import threading
import traceback
import subprocess
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd
import requests

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import FuncFormatter

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree


# =============================================================================
# CONFIGURAZIONE (ticker, colori)  — modificabile liberamente
# Ogni voce può avere una lista di ticker alternativi: si usa il primo che risponde.
# =============================================================================

GOLD = "#C6A15B"
GREY_DARK = "#5E5E5E"
GREY = "#8F8F8F"
GREY_LIGHT = "#BDBDBD"
BEIGE = "#E3D3B3"
BG_CHART = "#FFFFFF"
TITLE_GREY = "#6B6B6B"
TABLE_HEADER = "#D9D9D9"
TABLE_ROW = "#EDEDED"
TABLE_ROW_ALT = "#F5EFE3"
SECTION_COLORS = {"Dati Economici": "#E2E2E2", "Tassi e reddito fisso": "#D3D3D3",
                  "Azionario": "#EFEBE3", "Valute": "#D3D3D3",
                  "Materie prime e criptovalute": "#EFEBE3", "Titoli": "#E2E2E2"}

# Borse mondiali: (nome, [ticker alternativi])
BORSE = {
    "AMERICHE": [
        ("Dow Jones Industrial Avg", ["^DJI"]),
        ("S&P 500", ["^GSPC"]),
        ("Nasdaq 100", ["^NDX"]),
        ("S&P/TSX Composite", ["^GSPTSE"]),
        ("S&P/BMV IPC (Messico)", ["^MXX"]),
    ],
    "EUROPA": [
        ("Swiss Market Index", ["^SSMI"]),
        ("Euro Stoxx 50", ["^STOXX50E"]),
        ("FTSE 100", ["^FTSE"]),
        ("CAC 40", ["^FCHI"]),
        ("DAX", ["^GDAXI"]),
        ("FTSE MIB", ["FTSEMIB.MI"]),
    ],
    "ASIA/PACIFICO": [
        ("Nikkei 225", ["^N225"]),
        ("Hang Seng", ["^HSI"]),
        ("CSI 300", ["000300.SS", "ASHR"]),
        ("S&P/ASX 200", ["^AXJO"]),
    ],
}

VALUTE = [
    ("EUR/USD", ["EURUSD=X"]),
    ("USD/CHF", ["CHF=X"]),
    ("GBP/USD", ["GBPUSD=X"]),
    ("USD/JPY", ["JPY=X"]),
    ("USD/CAD", ["CAD=X"]),
    ("AUD/USD", ["AUDUSD=X"]),
    ("EUR/CHF", ["EURCHF=X"]),
    ("EUR/GBP", ["EURGBP=X"]),
    ("EUR/JPY", ["EURJPY=X"]),
]

COMMODITIES = [
    ("Oro ($/oz)", ["GC=F"]),
    ("Argento ($/oz)", ["SI=F"]),
    ("Platino ($/oz)", ["PL=F"]),
    ("Palladio ($/oz)", ["PA=F"]),
    ("Petrolio WTI (future)", ["CL=F"]),
    ("Metalli industriali (ETF DBB)", ["DBB"]),
    ("Agricoltura (ETF DBA)", ["DBA"]),
]

CRYPTO = [
    ("Bitcoin", ["BTC-USD"]),
    ("Ethereum", ["ETH-USD"]),
    ("Litecoin", ["LTC-USD"]),
    ("Bitcoin Cash", ["BCH-USD"]),
    ("XRP", ["XRP-USD"]),
    ("Dash", ["DASH-USD"]),
    ("EOS", ["EOS-USD"]),
]

# Indici obbligazionari Bloomberg non sono gratuiti: si usano ETF che li replicano
BOND_PROXY = [
    ("Global Aggregate (ETF BNDW)", ["BNDW"]),
    ("U.S. Aggregate (ETF AGG)", ["AGG"]),
    ("Euro Aggregate (ETF IEAG)", ["IEAG.L", "IEAG.AS", "EUNA.DE"]),
]

# Grafici di prezzo singoli: (titolo slide, sezione, etichetta, [ticker])
PRICE_CHARTS = [
    ("S&P 500", "Azionario", "S&P 500", ["^GSPC"]),
    ("STOXX Europe 600", "Azionario", "STOXX 600", ["^STOXX"]),
    ("MSCI World (ETF URTH)", "Azionario", "MSCI World", ["URTH"]),
    ("MSCI Emerging Markets (ETF EEM)", "Azionario", "MSCI EM", ["EEM"]),
    ("SMI", "Azionario", "SMI", ["^SSMI"]),
]
FX_CHARTS = [
    ("EUR/USD", "Valute", "EUR/USD", ["EURUSD=X"]),
    ("EUR/CHF", "Valute", "EUR/CHF", ["EURCHF=X"]),
    ("USD/CHF", "Valute", "USD/CHF", ["CHF=X"]),
]
COMMO_CHARTS = [
    ("Commodities Index", "Materie prime e criptovalute", "Bloomberg Commodity", ["^BCOM", "DJP"]),
    ("Petrolio Brent ($/barile)", "Materie prime e criptovalute", "Brent", ["BZ=F"]),
    ("Future Gas Europeo (TTF)", "Materie prime e criptovalute", "Gas TTF", ["TTF=F", "FRED:PNGASEUUSDM"]),
    ("Oro ($/oncia)", "Materie prime e criptovalute", "Oro", ["GC=F"]),
    ("BTC/USD", "Materie prime e criptovalute", "Bitcoin", ["BTC-USD"]),
]

WORLD_INDEX = ("Indice mondiale (ETF ACWI)", ["ACWI"])
DM_EM = [("Paesi sviluppati (MSCI World, ETF URTH)", ["URTH"]),
         ("Paesi emergenti (MSCI EM, ETF EEM)", ["EEM"])]
EM_REGIONS = [
    ("EM America Latina (ETF ILF)", ["ILF"]),
    ("EM Asia (ETF EEMA)", ["EEMA", "AAXJ"]),
    ("MSCI Emerging Markets (ETF EEM)", ["EEM"]),
]

# FRED (serie gratuite della Fed di St. Louis)
FRED_UST = {"3M": "DGS3MO", "6M": "DGS6MO", "1Y": "DGS1", "2Y": "DGS2", "5Y": "DGS5",
            "10Y": "DGS10", "20Y": "DGS20", "30Y": "DGS30"}
FRED_CPI_US = ["CPIAUCSL"]
FRED_CPI_EU = ["CP0000EZCCM086NEST", "CP0000EZ19M086NEST"]
FRED_POLICY = {"USA": ["DFEDTARU"], "Eurozona": ["ECBDFR"],
               "Svizzera": ["IRSTCB01CHM156N"], "Giappone": ["IRSTCB01JPM156N"]}
BIS_POLICY_CODES = {"Svizzera": "CH", "Giappone": "JP"}

# Bundesbank: rendimenti titoli federali tedeschi per scadenza residua
BUND_SERIES = "BBSIS/D.I.ZST.ZI.EUR.S1311.B.A604.R{m:02d}XX.R.A.A._Z._Z.A"
# SNB: rendimenti obbligazioni della Confederazione
SNB_CUBE = "rendoblim"

HTTP_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) MercatiReport/1.0"}


# =============================================================================
# UTILITÀ
# =============================================================================


# =============================================================================
# LINGUE  (italiano = testo originale; per aggiungere voci basta estendere TR)
# =============================================================================

LANG = "ita"   # "ita" | "eng" | "esp"

TR = {
    # sezioni
    "Dati Economici": ("Economic Data", "Datos Económicos"),
    "Tassi e reddito fisso": ("Rates & Fixed Income", "Tipos y Renta Fija"),
    "Azionario": ("Equities", "Renta Variable"),
    "Valute": ("Currencies", "Divisas"),
    "Materie prime e criptovalute": ("Commodities & Crypto", "Materias primas y criptomonedas"),
    "Titoli": ("Stocks", "Acciones"),
    # titoli slide
    "Mercati": ("Markets", "Mercados"),
    "Dati al": ("Data as of", "Datos al"),
    "Inflazione": ("Inflation", "Inflación"),
    "Rates e Bond Indexes": ("Rates & Bond Indexes", "Tipos e Índices de Bonos"),
    "Rendimenti governativi (%)": ("Government yields (%)", "Rendimientos soberanos (%)"),
    "Bond Indexes (via ETF)": ("Bond Indexes (via ETF)", "Índices de bonos (vía ETF)"),
    "Politica monetaria: principali banche centrali": ("Monetary policy: main central banks",
                                                       "Política monetaria: principales bancos centrales"),
    "Rendimento titoli di stato USA": ("US Treasury yields", "Rendimiento de los bonos del Tesoro de EE.UU."),
    "Curve dei titoli di Stato statunitensi": ("US Treasury yield curve", "Curva de los bonos del Tesoro de EE.UU."),
    "Borse mondiali": ("World stock markets", "Bolsas mundiales"),
    "Indice mondiale: performance ultimi {n} anni": ("World index: performance last {n} years",
                                                     "Índice mundial: rendimiento últimos {n} años"),
    "Mercati sviluppati ed emergenti YTD": ("Developed and emerging markets YTD",
                                            "Mercados desarrollados y emergentes YTD"),
    "Rendimento YTD delle regioni emergenti": ("YTD return of emerging regions",
                                               "Rendimiento YTD de las regiones emergentes"),
    "Variazione in % (USD)": ("Change in % (USD)", "Variación en % (USD)"),
    "Riepilogo valute": ("Currency overview", "Resumen de divisas"),
    "Note sui dati": ("Data notes", "Notas sobre los datos"),
    "Avvisi durante la generazione:": ("Warnings during generation:", "Avisos durante la generación:"),
    "Confronto": ("Comparison", "Comparación"),
    # tabelle
    "Scadenza": ("Maturity", "Vencimiento"),
    "Ultimo": ("Last", "Último"),
    "Valore": ("Value", "Valor"),
    "Indice": ("Index", "Índice"),
    "Cambio": ("Currency pair", "Tipo de cambio"),
    "Commodities": ("Commodities", "Materias primas"),
    "Criptovalute": ("Cryptocurrencies", "Criptomonedas"),
    "AMERICHE": ("AMERICAS", "AMÉRICAS"),
    "EUROPA": ("EUROPE", "EUROPA"),
    "ASIA/PACIFICO": ("ASIA/PACIFIC", "ASIA/PACÍFICO"),
    "USA": ("USA", "EE.UU."),
    "Germania": ("Germany", "Alemania"),
    "Svizzera": ("Switzerland", "Suiza"),
    "Eurozona": ("Eurozone", "Eurozona"),
    "Giappone": ("Japan", "Japón"),
    "Confederazione": ("Swiss Conf.", "Confederación"),
    "via": ("via", "vía"),
    "n.d.": ("n/a", "n.d."),
    "S&P/BMV IPC (Messico)": ("S&P/BMV IPC (Mexico)", "S&P/BMV IPC (México)"),
    "Oro ($/oz)": ("Gold ($/oz)", "Oro ($/oz)"),
    "Argento ($/oz)": ("Silver ($/oz)", "Plata ($/oz)"),
    "Platino ($/oz)": ("Platinum ($/oz)", "Platino ($/oz)"),
    "Palladio ($/oz)": ("Palladium ($/oz)", "Paladio ($/oz)"),
    "Petrolio WTI (future)": ("WTI Crude (future)", "Petróleo WTI (futuro)"),
    "Metalli industriali (ETF DBB)": ("Industrial metals (ETF DBB)", "Metales industriales (ETF DBB)"),
    "Agricoltura (ETF DBA)": ("Agriculture (ETF DBA)", "Agricultura (ETF DBA)"),
    # grafici
    "Andamento": ("Trend", "Evolución"),
    "Prezzo": ("Price", "Precio"),
    "Data inizio:": ("Start date:", "Fecha inicio:"),
    "Fonte:": ("Source:", "Fuente:"),
    "Ultimo dato:": ("Last data:", "Último dato:"),
    "Rendimento (%)": ("Return (%)", "Rendimiento (%)"),
    "Rendimento (%)|yield": ("Yield (%)", "Rendimiento (%)"),
    "Rendimento normalizzato (base 100)": ("Normalized return (base 100)", "Rendimiento normalizado (base 100)"),
    "Variazione annua (%)": ("YoY change (%)", "Variación interanual (%)"),
    "inflazione annua, ultimi 10 anni": ("annual inflation, last 10 years", "inflación anual, últimos 10 años"),
    "CPI USA": ("US CPI", "IPC EE.UU."),
    "HICP Eurozona": ("Eurozone HICP", "IPCA Eurozona"),
    "Tasso di riferimento (%)": ("Policy rate (%)", "Tasa de referencia (%)"),
    "Attuale": ("Current", "Actual"),
    "Un anno fa": ("One year ago", "Hace un año"),
    "Diff. (pb)": ("Diff. (bp)", "Dif. (pb)"),
    "Petrolio Brent ($/barile)": ("Brent Crude ($/barrel)", "Petróleo Brent ($/barril)"),
    "Future Gas Europeo (TTF)": ("European Gas Future (TTF)", "Futuro gas europeo (TTF)"),
    "Oro ($/oncia)": ("Gold ($/ounce)", "Oro ($/onza)"),
    "Oro": ("Gold", "Oro"),
    "(FRED, mensile $/MMBtu)": ("(FRED, monthly $/MMBtu)", "(FRED, mensual $/MMBtu)"),
    "Indice mondiale (ETF ACWI)": ("World index (ETF ACWI)", "Índice mundial (ETF ACWI)"),
    "Paesi sviluppati (MSCI World, ETF URTH)": ("Developed markets (MSCI World, ETF URTH)",
                                                "Mercados desarrollados (MSCI World, ETF URTH)"),
    "Paesi emergenti (MSCI EM, ETF EEM)": ("Emerging markets (MSCI EM, ETF EEM)",
                                           "Mercados emergentes (MSCI EM, ETF EEM)"),
    "EM America Latina (ETF ILF)": ("EM Latin America (ETF ILF)", "EM América Latina (ETF ILF)"),
    # note finali
    "NOTE_BODY": (
        "Sources: Yahoo Finance (indices, FX, commodities, crypto, ETFs, stocks), FRED – Federal Reserve "
        "Bank of St. Louis (Treasuries, CPI, Fed and ECB rates), Bundesbank (Bunds), Swiss National Bank "
        "(Confederation bonds), BIS (SNB and BoJ policy rates).\n\n"
        "Bloomberg/MSCI indices that are not freely available are approximated with ETFs tracking them, "
        "shown in brackets. ETF performance is computed on dividend-adjusted prices.\n\n"
        "For personal use only, generated automatically.",
        "Fuentes: Yahoo Finance (índices, divisas, materias primas, cripto, ETF, acciones), FRED – Federal "
        "Reserve Bank of St. Louis (Treasuries, IPC, tipos de la Fed y del BCE), Bundesbank (Bund), Banco "
        "Nacional Suizo (Confederación), BIS (tipos oficiales del BNS y del BoJ).\n\n"
        "Los índices Bloomberg/MSCI que no están disponibles gratuitamente se aproximan con ETF que los "
        "replican, indicados entre paréntesis. El rendimiento de los ETF se calcula sobre precios ajustados "
        "por dividendos.\n\n"
        "Documento para uso personal, generado automáticamente."),
    "RATES_NOTE": ("Data as of {d}. Rate changes in percentage points.\nSources: FRED, Bundesbank, SNB, Yahoo Finance.",
                   "Datos al {d}. Variaciones de tipos en puntos porcentuales.\nFuentes: FRED, Bundesbank, BNS, Yahoo Finance."),
}
TR_ITA = {
    "NOTE_BODY": (
        "Fonti: Yahoo Finance (indici, cambi, materie prime, cripto, ETF, titoli), FRED – Federal Reserve Bank "
        "of St. Louis (Treasury, CPI, tassi Fed e BCE), Bundesbank (Bund), Banca nazionale svizzera "
        "(Confederazione), BIS (tassi ufficiali BNS e BoJ).\n\n"
        "Gli indici Bloomberg/MSCI non disponibili gratuitamente sono approssimati con ETF che li "
        "replicano, indicati tra parentesi. Le performance degli ETF sono calcolate su prezzi "
        "rettificati per i dividendi.\n\n"
        "Documento ad uso personale, generato automaticamente."),
    "RATES_NOTE": "Dati al {d}. Variazioni dei tassi in punti percentuali.\nFonti: FRED, Bundesbank, SNB, Yahoo Finance.",
}
MONTHS = {
    "ita": ["Gen", "Feb", "Mar", "Apr", "Mag", "Giu", "Lug", "Ago", "Set", "Ott", "Nov", "Dic"],
    "eng": ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
    "esp": ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"],
}


def L(s):
    """Traduce un testo italiano nella lingua scelta (se non c'è traduzione resta com'è)."""
    if s is None:
        return s
    if LANG == "ita":
        return TR_ITA.get(s, s.split("|")[0])
    tr = TR.get(s)
    if tr is None:
        return s.split("|")[0]
    return tr[0] if LANG == "eng" else tr[1]


class Log:
    def __init__(self, fn=None):
        self.fn = fn or (lambda m: print(m))
        self.warnings = []

    def __call__(self, msg):
        self.fn(msg)

    def warn(self, msg):
        self.warnings.append(msg)
        self.fn("  ⚠ " + msg)


def _naive(idx):
    idx = pd.DatetimeIndex(pd.to_datetime(idx))
    if idx.tz is not None:
        idx = idx.tz_localize(None)
    return idx.normalize()


def _empty():
    """Serie vuota con indice di date (evita errori nei confronti con le date)."""
    return pd.Series(dtype=float, index=pd.DatetimeIndex([]))


def fmt_num(x, dec=2):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "n.d."
    s = f"{x:,.{dec}f}"
    return s.replace(",", "'")


def fmt_pct(x, dec=1):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "n.d."
    return f"{0.0 if abs(x) < 0.5 * 10 ** -dec else x:+.{dec}f}%"


def fmt_chg(x, dec=2):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "n.d."
    return f"{0.0 if abs(x) < 0.5 * 10 ** -dec else x:+.{dec}f}"


def perf(series, asof, kind="pct"):
    """Ritorna (ultimo, var MTD, var YTD). kind='pct' in %, 'diff' in punti."""
    s = series.dropna()
    s = s[s.index <= pd.Timestamp(asof)]
    if s.empty:
        return np.nan, np.nan, np.nan
    last = float(s.iloc[-1])
    asof = pd.Timestamp(asof)
    m0 = pd.Timestamp(asof.year, asof.month, 1)
    y0 = pd.Timestamp(asof.year, 1, 1)
    pm = s[s.index < m0]
    py = s[s.index < y0]
    def chg(base):
        if base.empty:
            return np.nan
        b = float(base.iloc[-1])
        if kind == "pct":
            return (last / b - 1) * 100 if b else np.nan
        return last - b
    return last, chg(pm), chg(py)


# =============================================================================
# DOWNLOAD DATI
# =============================================================================

class DataSource:
    def __init__(self, asof, years, log):
        self.asof = pd.Timestamp(asof)
        self.years = years
        self.log = log
        self.start = min(self.asof - pd.DateOffset(years=years),
                         pd.Timestamp(self.asof.year - 1, 11, 1)) - pd.Timedelta(days=10)
        self.yahoo = pd.DataFrame()
        self._chosen = {}

    # ---------------- Yahoo ----------------
    def load_yahoo(self, candidate_lists):
        import yfinance as yf
        tickers = sorted({t for c in candidate_lists for t in c if not t.startswith("FRED:")})
        self.log(f"Scarico {len(tickers)} ticker da Yahoo Finance…")
        end = (self.asof + pd.Timedelta(days=1)).strftime("%Y-%m-%d")
        frames = []
        # a blocchi per non stressare Yahoo
        for i in range(0, len(tickers), 25):
            chunk = tickers[i:i + 25]
            try:
                df = yf.download(chunk, start=self.start.strftime("%Y-%m-%d"), end=end,
                                 auto_adjust=True, progress=False, threads=True)
            except Exception as e:
                self.log.warn(f"Errore Yahoo sul blocco {chunk[:3]}…: {e}")
                continue
            if df is None or df.empty:
                continue
            if isinstance(df.columns, pd.MultiIndex):
                lvl0 = df.columns.get_level_values(0)
                close = df["Close"] if "Close" in lvl0 else df.xs("Close", axis=1, level=1)
            else:
                close = df[["Close"]].rename(columns={"Close": chunk[0]})
            frames.append(close)
        if frames:
            y = pd.concat(frames, axis=1)
            y.index = _naive(y.index)
            y = y[~y.index.duplicated(keep="last")].sort_index()
            self.yahoo = y
        # secondo tentativo, uno alla volta, per i ticker arrivati vuoti
        missing = [t for t in tickers
                   if t not in self.yahoo.columns or self.yahoo[t].notna().sum() <= 5]
        if missing:
            self.log(f"  Riprovo singolarmente {len(missing)} ticker: {', '.join(missing)}")
        for t in missing:
            try:
                h = yf.Ticker(t).history(start=self.start.strftime("%Y-%m-%d"), end=end,
                                         auto_adjust=True)
                if h is not None and not h.empty and "Close" in h:
                    s = h["Close"].copy()
                    s.index = _naive(s.index)
                    s = s[~s.index.duplicated(keep="last")]
                    self.yahoo = self.yahoo.drop(columns=[t], errors="ignore").join(s.rename(t), how="outer")
            except Exception:
                pass
        self.log(f"  Yahoo: ricevuti {self.yahoo.shape[1] if not self.yahoo.empty else 0} ticker.")

    def load_custom(self, tickers, start):
        """Prezzi di chiusura (non rettificati per dividendi) dei titoli scelti."""
        import yfinance as yf
        end = (self.asof + pd.Timedelta(days=1)).strftime("%Y-%m-%d")
        out = {}
        for t in tickers:
            try:
                h = yf.Ticker(t).history(start=pd.Timestamp(start).strftime("%Y-%m-%d"), end=end,
                                         auto_adjust=False)
                if h is not None and not h.empty and "Close" in h:
                    s = h["Close"].copy()
                    s.index = _naive(s.index)
                    out[t] = s[~s.index.duplicated(keep="last")].dropna()
            except Exception as e:
                self.log.warn(f"Yahoo {t}: {e}")
        return out

    def y(self, candidates, label=""):
        for t in candidates:
            if t.startswith("FRED:"):
                s = self.fred(t[5:], self.start)
                if len(s) > 5:
                    self._chosen[label or candidates[0]] = t
                    return s
                continue
            if t in self.yahoo.columns:
                s = self.yahoo[t].dropna()
                s = s[s.index <= self.asof]
                if len(s) > 5:
                    self._chosen[label or candidates[0]] = t
                    return s
        self.log.warn(f"Nessun dato Yahoo per {label or candidates} ({', '.join(candidates)})")
        return _empty()

    # ---------------- FRED ----------------
    def fred(self, series_id, start=None):
        url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series_id}"
        if start is not None:
            url += f"&cosd={pd.Timestamp(start):%Y-%m-%d}"
        try:
            r = requests.get(url, headers=HTTP_HEADERS, timeout=30)
            r.raise_for_status()
            df = pd.read_csv(io.StringIO(r.text))
            df.columns = [c.strip() for c in df.columns]
            s = pd.to_numeric(df.iloc[:, 1], errors="coerce")
            s.index = _naive(df.iloc[:, 0])
            s = s.dropna()
            return s[s.index <= self.asof]
        except Exception as e:
            self.log.warn(f"FRED {series_id}: {e}")
            return _empty()

    def fred_first(self, ids, start=None):
        best = _empty()
        for i in ids:
            s = self.fred(i, start)
            if not s.empty and (best.empty or s.index[-1] > best.index[-1]):
                best = s
        return best

    # ---------------- Parsing CSV generico (data, valore) ----------------
    @staticmethod
    def _parse_date_value_csv(text):
        rows = []
        for line in text.splitlines():
            line = line.strip().replace('"', "")
            if not line:
                continue
            sep = ";" if line.count(";") >= line.count(",") and ";" in line else ","
            parts = [p.strip() for p in line.split(sep)]
            if len(parts) < 2:
                continue
            try:
                d = pd.Timestamp(parts[0])
            except Exception:
                continue
            if len(parts[0]) < 8:
                continue
            for p in parts[1:]:
                try:
                    v = float(p.replace(",", ".")) if p not in (".", "") else None
                except ValueError:
                    v = None
                if v is not None:
                    rows.append((d, v))
                    break
        if not rows:
            return _empty()
        s = pd.Series([v for _, v in rows], index=_naive([d for d, _ in rows]))
        return s[~s.index.duplicated(keep="last")].sort_index()

    # ---------------- Bundesbank ----------------
    def bund(self, maturity):
        key = BUND_SERIES.format(m=maturity)
        url = (f"https://api.statistiken.bundesbank.de/rest/data/{key}"
               f"?format=csv&lang=en&startPeriod={(self.asof.year - 1)}-11-01")
        try:
            r = requests.get(url, headers=HTTP_HEADERS, timeout=30)
            r.raise_for_status()
            s = self._parse_date_value_csv(r.text)
            return s[s.index <= self.asof]
        except Exception as e:
            self.log.warn(f"Bundesbank {maturity}Y: {e}")
            return _empty()

    # ---------------- SNB ----------------
    def snb_yields(self):
        url = (f"https://data.snb.ch/api/cube/{SNB_CUBE}/data/csv/en"
               f"?fromDate={(self.asof.year - 1)}-11-01")
        out = {}
        try:
            r = requests.get(url, headers=HTTP_HEADERS, timeout=30)
            r.raise_for_status()
            lines = r.text.splitlines()
            hdr = next(i for i, l in enumerate(lines) if l.replace('"', "").startswith("Date"))
            df = pd.read_csv(io.StringIO("\n".join(lines[hdr:])), sep=";")
            df.columns = [c.strip().strip('"') for c in df.columns]
            dim = [c for c in df.columns if c not in ("Date", "Value")][0]
            df["Date"] = _naive(df["Date"])
            df["Value"] = pd.to_numeric(df["Value"], errors="coerce")
            for code, g in df.groupby(dim):
                digits = "".join(ch for ch in str(code) if ch.isdigit())
                if digits:
                    s = g.set_index("Date")["Value"].dropna().sort_index()
                    out[int(digits)] = s[s.index <= self.asof]
        except Exception as e:
            self.log.warn(f"SNB rendimenti Confederazione: {e}")
        return out

    # ---------------- BIS (tassi banche centrali) ----------------
    def bis_policy(self, codes):
        start = (self.asof - pd.DateOffset(years=self.years + 1)).strftime("%Y-%m-%d")
        urls = [
            f"https://stats.bis.org/api/v1/data/WS_CBPOL/D.{'+'.join(codes)}/all?format=csv&startPeriod={start}",
            f"https://stats.bis.org/api/v2/data/dataflow/BIS/WS_CBPOL/1.0/D.{'+'.join(codes)}?format=csv&startPeriod={start}",
        ]
        for url in urls:
            try:
                r = requests.get(url, headers=HTTP_HEADERS, timeout=40)
                r.raise_for_status()
                df = pd.read_csv(io.StringIO(r.text))
                cols = {c.upper().split(":")[0]: c for c in df.columns}
                ra, tp, ov = cols.get("REF_AREA"), cols.get("TIME_PERIOD"), cols.get("OBS_VALUE")
                if not (ra and tp and ov):
                    continue
                out = {}
                for code in codes:
                    g = df[df[ra].astype(str).str.split(":").str[0].str.strip() == code]
                    s = pd.Series(pd.to_numeric(g[ov], errors="coerce").values,
                                  index=_naive(g[tp])).dropna().sort_index()
                    if not s.empty:
                        out[code] = s[s.index <= self.asof]
                if out:
                    return out
            except Exception as e:
                self.log.warn(f"BIS tassi ufficiali: {e}")
        return {}


# =============================================================================
# GRAFICI
# =============================================================================

def _style():
    plt.rcParams.update({
        "font.family": ["Segoe UI", "Arial", "DejaVu Sans"],
        "font.size": 10,
        "axes.edgecolor": "#444444",
        "axes.linewidth": 0.8,
        "axes.spines.top": False,
        "axes.spines.right": False,
        "axes.titlesize": 11,
        "axes.titleweight": "normal",
        "legend.frameon": True,
        "legend.edgecolor": "#DDDDDD",
        "legend.fontsize": 8.5,
        "xtick.labelsize": 8.5,
        "ytick.labelsize": 8.5,
        "figure.facecolor": BG_CHART,
        "axes.facecolor": BG_CHART,
    })


def _footer(fig, s, asof, source):
    if s is not None and len(s):
        fig.text(0.06, 0.015, f"{L('Data inizio:')} {s.index[0]:%d/%m/%Y}", fontsize=7, color=GREY)
    fig.text(0.5, 0.015, f"{L('Fonte:')} {source}", fontsize=7, color=GREY, ha="center")
    fig.text(0.97, 0.015, f"{L('Ultimo dato:')} {pd.Timestamp(asof):%d/%m/%Y}", fontsize=7,
             color=GREY, ha="right")


def _date_axis(ax, ytd=False):
    if ytd:
        ax.xaxis.set_major_locator(mdates.MonthLocator())
        ax.xaxis.set_major_formatter(FuncFormatter(
            lambda x, _: MONTHS[LANG][mdates.num2date(x).month - 1]))
    else:
        loc = mdates.AutoDateLocator(minticks=4, maxticks=8)
        ax.xaxis.set_major_locator(loc)
        ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(loc, show_offset=False))


def _save(fig, path):
    fig.savefig(path, dpi=200, bbox_inches="tight", facecolor=BG_CHART)
    plt.close(fig)
    return path


def chart_price(s, title, label, asof, path, source="Yahoo Finance", dec=2):
    _style()
    label = L(label)
    fig, ax = plt.subplots(figsize=(11, 5.6))
    ax.plot(s.index, s.values, color=GOLD, lw=1.3, label=label)
    last = s.iloc[-1]
    ax.annotate(f"{fmt_num(last, dec)}  {label}", xy=(s.index[-1], last), xytext=(6, 0),
                textcoords="offset points", color=GOLD, fontsize=9, fontweight="bold", va="center")
    ax.set_title(f"{L('Andamento')} {label}", pad=10)
    ax.set_ylabel(L("Prezzo"))
    ax.legend(loc="upper left")
    ax.grid(axis="y", color="#EEEEEE", lw=0.6)
    _date_axis(ax)
    ax.set_xlim(s.index[0], s.index[-1] + (s.index[-1] - s.index[0]) * 0.13)
    _footer(fig, s, asof, source)
    return _save(fig, path)


def chart_multi(series, title, asof, path, ylabel, source, colors=None, pct=False, dec=2):
    _style()
    colors = colors or [GOLD, GREY, GREY_LIGHT, GREY_DARK]
    fig, ax = plt.subplots(figsize=(11, 5.6))
    lasts = []
    first = None
    for (name, s), c in zip(series.items(), colors):
        if s.empty:
            continue
        ax.plot(s.index, s.values, color=c, lw=1.3, label=L(name))
        lasts.append((s.iloc[-1], L(name), c, s.index[-1]))
        first = s if first is None or s.index[0] < first.index[0] else first
    _label_ends(ax, lasts, pct, dec)
    ax.set_title(L(title), pad=10)
    ax.set_ylabel(L(ylabel))
    ax.legend(loc="upper left")
    ax.grid(axis="y", color="#EEEEEE", lw=0.6)
    _date_axis(ax)
    if first is not None:
        x0, x1 = first.index[0], max(l[3] for l in lasts)
        ax.set_xlim(x0, x1 + (x1 - x0) * 0.14)
    _footer(fig, first, asof, source)
    return _save(fig, path)


def _label_ends(ax, lasts, pct, dec, with_name=True):
    """Etichette a fine linea, distanziate per non sovrapporsi."""
    if not lasts:
        return
    lasts = sorted(lasts, key=lambda x: -x[0])
    ymin, ymax = ax.get_ylim()
    gap = (ymax - ymin) * 0.045
    x_span = max(mdates.date2num(l[3]) for l in lasts) - min(ax.get_xlim()[0], mdates.date2num(lasts[0][3]))
    dx = max(x_span * 0.012, 1)
    placed = []
    for v, name, c, x in lasts:
        y = v
        if placed and placed[-1] - y < gap:
            y = placed[-1] - gap
        placed.append(y)
        txt = (f"{v:.{dec}f}%" if pct else fmt_num(v, dec))
        if with_name:
            txt += f"  {name.split(' (')[0]}"
        ax.annotate(txt, xy=(mdates.date2num(x), v), xytext=(mdates.date2num(x) + dx, y),
                    textcoords="data", color=c, fontsize=8.5, fontweight="bold",
                    va="center", ha="left")


def chart_rebased(series, title, asof, path, source, fill_between=False, ytd=True, dots=False):
    """Rendimento % ribasato a 0 all'inizio (YTD o periodo)."""
    _style()
    colors = [GREY, GOLD, BEIGE, GREY_DARK, GREY_LIGHT]
    fig, ax = plt.subplots(figsize=(11, 5.6))
    asof = pd.Timestamp(asof)
    reb = {}
    for name, s in series.items():
        if s.empty:
            continue
        if ytd:
            base = s[s.index < pd.Timestamp(asof.year, 1, 1)]
            if base.empty:
                continue
            b = base.iloc[-1]
            ss = s[s.index >= pd.Timestamp(asof.year, 1, 1)]
            ss = pd.concat([pd.Series([b], index=[pd.Timestamp(asof.year, 1, 1)]), ss])
        else:
            ss, b = s, s.iloc[0]
        reb[name] = (ss / b - 1) * 100
    lasts = []
    for (name, r), c in zip(reb.items(), colors):
        ax.plot(r.index, r.values, color=c, lw=1.6, label=L(name))
        if dots:
            ax.scatter([r.index[-1]], [r.iloc[-1]], s=40, color=c, edgecolor="#555555", zorder=5)
        lasts.append((r.iloc[-1], name, c, r.index[-1]))
    if fill_between and len(reb) == 2:
        (_, a), (_, b) = list(reb.items())
        j = pd.concat([a, b], axis=1).ffill().dropna()
        ax.fill_between(j.index, j.iloc[:, 0], j.iloc[:, 1], color="#DDDDDD", alpha=0.6, lw=0)
    ax.axhline(0, color="#444444", lw=0.9)
    ax.set_ylabel(L("Rendimento (%)"))
    ax.grid(axis="y", color="#EEEEEE", lw=0.6, ls=":")
    ax.legend(loc="upper left")
    if ytd:
        ax.set_xlim(pd.Timestamp(asof.year, 1, 1), pd.Timestamp(asof.year, 12, 31))
        _date_axis(ax, ytd=True)
        x_lab = pd.Timestamp(asof.year, 12, 31)
    else:
        _date_axis(ax)
        x0 = min(r.index[0] for r in reb.values())
        x1 = max(r.index[-1] for r in reb.values())
        ax.set_xlim(x0, x1 + (x1 - x0) * 0.1)
        x_lab = None
    _label_ends(ax, lasts, True, 1, with_name=False)
    if title:
        ax.set_title(L(title), pad=10)
    first = next(iter(reb.values())) if reb else None
    _footer(fig, first, asof, source)
    return _save(fig, path)


def chart_cpi(s, label, asof, path, source):
    _style()
    label = L(label)
    fig, ax = plt.subplots(figsize=(7.5, 5.4))
    ax.plot(s.index, s.values, color=GOLD, lw=1.2, label=label)
    ax.annotate(f"{s.iloc[-1]:.2f}%", xy=(s.index[-1], s.iloc[-1]), xytext=(6, 0),
                textcoords="offset points", color=GOLD, fontsize=9, fontweight="bold", va="center")
    ax.axhline(2, color=GREY_LIGHT, lw=0.8, ls="--")
    ax.set_title(f"{label} — {L('inflazione annua, ultimi 10 anni')}", pad=10)
    ax.set_ylabel(L("Variazione annua (%)"))
    ax.legend(loc="upper left")
    ax.grid(axis="y", color="#EEEEEE", lw=0.6)
    _date_axis(ax)
    ax.set_xlim(s.index[0], s.index[-1] + (s.index[-1] - s.index[0]) * 0.1)
    fig.text(0.06, 0.005, f"{L('Fonte:')} {source}", fontsize=7, color=GREY)
    fig.text(0.97, 0.005, f"{L('Ultimo dato:')} {s.index[-1]:%m/%Y}", fontsize=7, color=GREY, ha="right")
    return _save(fig, path)


def chart_policy(rates, asof, path, source):
    _style()
    colors = {"Eurozona": GOLD, "Giappone": GREY, "Svizzera": BEIGE, "USA": GREY_DARK}
    fig, ax = plt.subplots(figsize=(11, 5.6))
    x0 = None
    for name in ["Eurozona", "Giappone", "Svizzera", "USA"]:
        s = rates.get(name)
        if s is None or s.empty:
            continue
        s = pd.concat([s, pd.Series([s.iloc[-1]], index=[pd.Timestamp(asof)])])
        s = s[~s.index.duplicated(keep="first")]
        ax.step(s.index, s.values, where="post", color=colors[name], lw=1.6, label=L(name))
        x0 = s.index[0] if x0 is None else min(x0, s.index[0])
    ax.set_ylabel(L("Tasso di riferimento (%)"))
    ax.grid(axis="y", color="#EEEEEE", lw=0.6)
    ax.legend(loc="upper right")
    _date_axis(ax)
    if x0 is not None:
        ax.set_xlim(x0, pd.Timestamp(asof))
    fig.text(0.06, 0.015, f"{L('Data inizio:')} {x0:%d/%m/%Y}" if x0 is not None else "", fontsize=7, color=GREY)
    fig.text(0.5, 0.015, f"{L('Fonte:')} {source}", fontsize=7, color=GREY, ha="center")
    fig.text(0.97, 0.015, f"{L('Ultimo dato:')} {pd.Timestamp(asof):%d/%m/%Y}", fontsize=7, color=GREY, ha="right")
    return _save(fig, path)


def chart_curve(now, before, d_now, d_before, path, source):
    _style()
    mats = [m for m in now.index if m in before.index]
    fig, (ax, ax2) = plt.subplots(2, 1, figsize=(11, 6.6), gridspec_kw={"height_ratios": [3.2, 1]},
                                  sharex=True)
    x = np.arange(len(mats))
    ax.plot(x, now[mats].values, color=GOLD, marker="o", lw=1.6, ms=5, label=f"{L('Attuale')} ({d_now:%d/%m/%Y})")
    ax.plot(x, before[mats].values, color=GREY, marker="o", lw=1.6, ms=5,
            label=f"{L('Un anno fa')} ({d_before:%d/%m/%Y})")
    for i, m in enumerate(mats):
        a, b = now[m], before[m]
        up_a = a >= b
        ax.annotate(f"{a:.2f}", (i, a), xytext=(0, 9 if up_a else -14), textcoords="offset points",
                    ha="center", fontsize=8, color=GOLD, fontweight="bold")
        ax.annotate(f"{b:.2f}", (i, b), xytext=(0, -14 if up_a else 9), textcoords="offset points",
                    ha="center", fontsize=8, color=GREY)
    ax.set_ylabel(L("Rendimento (%)|yield"))
    ax.grid(axis="y", color="#EEEEEE", lw=0.6)
    ax.legend(loc="upper left")
    lo = min(now[mats].min(), before[mats].min())
    hi = max(now[mats].max(), before[mats].max())
    ax.set_ylim(lo - (hi - lo) * 0.15, hi + (hi - lo) * 0.18)
    diff = (now[mats] - before[mats]) * 100
    bars = ax2.bar(x, diff.values, color=[GOLD if v >= 0 else GREY_LIGHT for v in diff.values], width=0.45)
    for bx, v in zip(x, diff.values):
        ax2.annotate(f"{v:.0f}", (bx, v), xytext=(0, 3 if v >= 0 else -11), textcoords="offset points",
                     ha="center", fontsize=8)
    ax2.axhline(0, color="#444444", lw=0.8)
    ax2.set_ylabel(L("Diff. (pb)"))
    m = max(abs(diff.min()), abs(diff.max()), 10)
    ax2.set_ylim(-m * 1.4, m * 1.4)
    ax2.set_xticks(x)
    ax2.set_xticklabels(mats)
    ax2.set_xlabel(L("Scadenza"))
    fig.text(0.5, 0.005, f"{L('Fonte:')} {source}", fontsize=7, color=GREY, ha="center")
    return _save(fig, path)


def chart_base100(series, asof, path, source="Yahoo Finance"):
    """Confronto tra titoli, normalizzati a 100 alla prima data comune."""
    _style()
    df = pd.concat(series, axis=1).sort_index().ffill().dropna()
    norm = df / df.iloc[0] * 100
    colors = [GOLD, GREY, GREY_LIGHT, GREY_DARK, BEIGE, "#A88850"]
    fig, ax = plt.subplots(figsize=(11, 5.6))
    lasts = []
    for i, col in enumerate(norm.columns):
        c = colors[i % len(colors)]
        ax.plot(norm.index, norm[col], color=c, lw=1.5, label=col)
        lasts.append((norm[col].iloc[-1], col, c, norm.index[-1]))
    ax.axhline(100, color="#BBBBBB", lw=0.8, ls=":")
    ax.set_title(f"{L('Confronto')} — {' vs '.join(norm.columns)}", pad=10)
    ax.set_ylabel(L("Rendimento normalizzato (base 100)"))
    ax.legend(loc="upper left")
    ax.grid(axis="y", color="#EEEEEE", lw=0.6)
    _date_axis(ax)
    x0, x1 = norm.index[0], norm.index[-1]
    ax.set_xlim(x0, x1 + (x1 - x0) * 0.12)
    _label_ends(ax, lasts, False, 0, with_name=False)
    _footer(fig, norm.iloc[:, 0], asof, source)
    return _save(fig, path)


# Colori tabelle (da bloomberg_market.py)
T_GOLD        = "#CCAD79"
T_GOLD_LIGHT  = "#E8DCC8"
T_ROW_ODD     = "#FFFFFF"
T_ROW_EVEN    = "#F5F0E8"
T_DARK        = "#1A1A1A"
T_GREY        = "#888888"
T_BORDER      = "#C8B898"
T_BORDER_SOFT = "#DDD5C0"
T_POS         = "#2E7D32"
T_NEG         = "#C62828"


def table_image(header, rows, col_w, row_h, font, group_rows, num_cols_from, footer):
    """Disegna la tabella con matplotlib e la restituisce come PNG in memoria."""
    import matplotlib.patches as patches
    N = len(rows)
    HDR_H = row_h * 1.15
    FTR_H = row_h * 0.85 if footer else 0.0
    FIG_W = sum(col_w)
    FIG_H = HDR_H + N * row_h + FTR_H
    PAD = 0.10
    fs = font * 0.88

    fig = plt.figure(figsize=(FIG_W, FIG_H), facecolor="white")
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, FIG_W); ax.set_ylim(0, FIG_H); ax.axis("off")

    xs = np.concatenate([[0], np.cumsum(col_w)])
    def tx(j, y, s, **kw):
        right = j >= num_cols_from
        x = xs[j + 1] - PAD if right else xs[j] + PAD
        ax.text(x, y, s, ha="right" if right else "left", va="center", zorder=2, **kw)

    y_hdr = FIG_H - HDR_H
    ax.add_patch(patches.Rectangle((0, y_hdr), FIG_W, HDR_H, facecolor=T_GOLD_LIGHT,
                                   edgecolor=T_BORDER, linewidth=0.7, zorder=1))
    for j, h in enumerate(header):
        tx(j, y_hdr + HDR_H * 0.52, h, fontsize=fs, fontweight="bold", color=T_DARK)
    ax.axhline(y_hdr, color=T_BORDER, linewidth=1.0, zorder=3)

    di = 0
    for i, row in enumerate(rows):
        y0 = y_hdr - (i + 1) * row_h
        ym = y0 + row_h * 0.5
        if i in group_rows:
            ax.add_patch(patches.Rectangle((0, y0), FIG_W, row_h, facecolor=T_GOLD,
                                           edgecolor=T_BORDER, linewidth=0.4, zorder=1))
            ax.text(PAD, ym, str(row[0]), ha="left", va="center", fontsize=fs,
                    fontweight="bold", color=T_DARK, zorder=2)
        else:
            bg = T_ROW_ODD if di % 2 == 0 else T_ROW_EVEN
            di += 1
            ax.add_patch(patches.Rectangle((0, y0), FIG_W, row_h, facecolor=bg,
                                           edgecolor=T_BORDER_SOFT, linewidth=0.2, zorder=1))
            for j, v in enumerate(row):
                s = str(v)
                if j < num_cols_from:
                    tx(j, ym, s, fontsize=fs * 0.97, color=T_DARK)
                elif s.endswith("%"):
                    col = T_GREY if s in ("n.d.", "n/a") else (T_NEG if s.startswith("-") else T_POS)
                    tx(j, ym, s, fontsize=fs * 0.97, color=col, fontweight="bold",
                       fontfamily="monospace")
                else:
                    tx(j, ym, s, fontsize=fs * 0.97, color=T_GREY if s in ("n.d.", "n/a") else T_DARK,
                       fontfamily="monospace")
        ax.axhline(y0, color=T_BORDER_SOFT, linewidth=0.3, zorder=3)

    ax.add_patch(patches.Rectangle((0, y_hdr - N * row_h), FIG_W, HDR_H + N * row_h,
                                   fill=False, edgecolor=T_BORDER, linewidth=0.8, zorder=4))
    if footer:
        ax.text(FIG_W - PAD, FTR_H * 0.45, footer, ha="right", va="center",
                fontsize=fs * 0.8, color=T_GREY, style="italic", zorder=2)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=220, facecolor="white", edgecolor="none")
    plt.close(fig)
    buf.seek(0)
    return buf


# =============================================================================
# POWERPOINT
# =============================================================================

def _rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


class Deck:
    W, H = Inches(13.333), Inches(7.5)

    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = self.W, self.H
        self.blank = self.prs.slide_layouts[6]
        self.asof = None

    def _text(self, slide, txt, l, t, w, h, size=12, bold=False, color=TITLE_GREY,
              align=PP_ALIGN.LEFT, font="Arial"):
        tb = slide.shapes.add_textbox(l, t, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = 0
        lines = txt.split("\n")
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            r = p.add_run()
            r.text = line
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.name = font
            r.font.color.rgb = _rgb(color)
        return tb

    def slide(self, title, section=None, subtitle=None):
        s = self.prs.slides.add_slide(self.blank)
        title, subtitle = L(title), L(subtitle)
        self._text(s, title, Inches(0.6), Inches(0.35), Inches(11.2), Inches(0.7), size=28, bold=True)
        if subtitle:
            self._text(s, subtitle, Inches(0.6), Inches(0.98), Inches(11), Inches(0.4), size=14, bold=True)
        if section:
            bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, self.W - Inches(0.62), 0, Inches(0.62), self.H)
            bar.fill.solid()
            bar.fill.fore_color.rgb = _rgb(SECTION_COLORS.get(section, "#E2E2E2"))
            bar.line.fill.background()
            tb = self._text(s, L(section), self.W - Inches(0.62), Inches(0.3), Inches(0.62), Inches(6.9),
                            size=14, color="#6B6B6B", align=PP_ALIGN.LEFT)
            tb.text_frame._txBody.bodyPr.set("vert", "vert270")
            tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        return s

    def picture(self, slide, path, l=Inches(0.7), t=Inches(1.25), w=None, h=Inches(5.9)):
        if w is not None:
            return slide.shapes.add_picture(path, l, t, width=w)
        pic = slide.shapes.add_picture(path, l, t, height=h)
        max_w = self.W - Inches(0.62) - l - Inches(0.3)
        if pic.width > max_w:
            ratio = max_w / pic.width
            pic.width = int(pic.width * ratio)
            pic.height = int(pic.height * ratio)
        return pic

    def cover(self, title, subtitle):
        s = self.prs.slides.add_slide(self.blank)
        bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.W, self.H)
        bg.fill.solid(); bg.fill.fore_color.rgb = _rgb("#F4F1EA"); bg.line.fill.background()
        band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(2.6), Inches(0.35), Inches(2.2))
        band.fill.solid(); band.fill.fore_color.rgb = _rgb(GOLD); band.line.fill.background()
        self._text(s, title, Inches(0.9), Inches(2.6), Inches(11), Inches(1.2), size=48, bold=True)
        self._text(s, subtitle, Inches(0.9), Inches(3.8), Inches(11), Inches(0.8), size=20, color=GOLD)
        return s

    def table(self, slide, header, rows, l, t, col_w, row_h=Inches(0.27), font=10,
              group_rows=(), num_cols_from=1, footer=True):
        """Tabella come immagine, nello stile di bloomberg_market.py
        (intestazione oro chiaro, sezioni oro, righe bianco/crema, % verdi e rosse)."""
        header = [L(h) for h in header]
        rows = [[L(v) if (j == 0 or v == "n.d.") else v for j, v in enumerate(r)] for r in rows]
        buf = table_image(header, rows, [w / 914400 for w in col_w], row_h / 914400, font,
                          group_rows, num_cols_from,
                          f"{L('Ultimo dato:')} {pd.Timestamp(self.asof):%d/%m/%Y}" if footer and self.asof is not None else None)
        return slide.shapes.add_picture(buf, l, t, width=sum(col_w))

    def save(self, path):
        self.prs.save(path)


# =============================================================================
# GENERAZIONE REPORT
# =============================================================================

def parse_tickers(text):
    """'CAP.SN=CAP SA, NESN.SW' -> [('CAP.SN', 'CAP SA'), ('NESN.SW', 'NESN.SW')]"""
    out = []
    for part in re.split(r"[,;\n]+", text or ""):
        part = part.strip()
        if not part:
            continue
        if "=" in part and not part.endswith("=X") and not part.endswith("=F"):
            tk, lab = part.split("=", 1)
            out.append((tk.strip().upper(), lab.strip() or tk.strip().upper()))
        else:
            out.append((part.upper(), part.upper()))
    return out


def _sheet(name):
    return re.sub(r"[\[\]:*?/\\]", "-", str(name))[:31]


def build_report(asof, years, out_dir, log, ds=None, lang="ita", custom=None, custom_years=10):
    global LANG
    LANG = lang
    asof = pd.Timestamp(asof).normalize()
    stamp = f"{asof:%Y-%m-%d}"
    img_dir = os.path.join(out_dir, f"grafici_{stamp}")
    os.makedirs(img_dir, exist_ok=True)
    ds = ds or DataSource(asof, years, log)
    chart_start = asof - pd.DateOffset(years=years)
    excel = {}

    def img(name):
        return os.path.join(img_dir, name)

    # ------- download Yahoo in un colpo solo -------
    all_c = []
    for grp in BORSE.values():
        all_c += [c for _, c in grp]
    for lst in (VALUTE, COMMODITIES, CRYPTO, BOND_PROXY, DM_EM, EM_REGIONS):
        all_c += [c for _, c in lst]
    for lst in (PRICE_CHARTS, FX_CHARTS, COMMO_CHARTS):
        all_c += [c for *_, c in lst]
    all_c.append(WORLD_INDEX[1])
    ds.load_yahoo(all_c)

    deck = Deck()
    deck.asof = asof
    deck.cover(L("Mercati"), f"{L('Dati al')} {asof:%d.%m.%Y}")

    # ======================= ECONOMIE =======================
    log("Inflazione (FRED)…")
    try:
        cpi_start = asof - pd.DateOffset(years=11, months=2)
        us = ds.fred_first(FRED_CPI_US, cpi_start)
        eu = ds.fred_first(FRED_CPI_EU, cpi_start)
        s = deck.slide("Inflazione", "Dati Economici")
        paths = []
        for raw, lab, fn in ((us, "CPI USA", "cpi_usa.png"), (eu, "HICP Eurozona", "cpi_euro.png")):
            if raw.empty:
                paths.append(None); continue
            yoy = (raw / raw.shift(12) - 1).dropna() * 100
            yoy = yoy[yoy.index >= asof - pd.DateOffset(years=10)]
            excel[lab] = yoy.rename(lab).to_frame()
            paths.append(chart_cpi(yoy, lab, asof, img(fn), "FRED (BLS, Eurostat)"))
        for k, p in enumerate(paths):
            if p:
                deck.picture(s, p, l=Inches(0.5 + k * 6.15), t=Inches(1.5), w=Inches(6.0))
    except Exception as e:
        log.warn(f"Slide inflazione: {e}")

    # ======================= OBBLIGAZIONARIO =======================
    log("Tassi governativi (FRED, Bundesbank, SNB)…")
    ust = {}
    for m, sid in FRED_UST.items():
        ust[m] = ds.fred(sid, asof - pd.DateOffset(years=max(years, 1) + 1))
    try:
        s = deck.slide("Rates e Bond Indexes", "Tassi e reddito fisso")
        rows, groups = [], []
        groups.append(len(rows)); rows.append(["USA", "", "", ""])
        for m in ["30Y", "20Y", "10Y", "5Y", "2Y"]:
            l, mtd, ytd = perf(ust[m], asof, "diff")
            rows.append([f"US Treasury {m}", fmt_num(l, 3), fmt_chg(mtd), fmt_chg(ytd)])
        groups.append(len(rows)); rows.append(["Germania", "", "", ""])
        for m in [30, 20, 10, 5, 2]:
            l, mtd, ytd = perf(ds.bund(m), asof, "diff")
            rows.append([f"Bund {m}Y", fmt_num(l, 3), fmt_chg(mtd), fmt_chg(ytd)])
        groups.append(len(rows)); rows.append(["Svizzera", "", "", ""])
        snb = ds.snb_yields()
        for m in [30, 20, 10, 5, 2]:
            l, mtd, ytd = perf(snb.get(m, _empty()), asof, "diff")
            rows.append([f"{L('Confederazione')} {m}Y", fmt_num(l, 3), fmt_chg(mtd), fmt_chg(ytd)])
        deck._text(s, L("Rendimenti governativi (%)"), Inches(0.6), Inches(1.1), Inches(5), Inches(0.4), size=14, bold=True)
        deck.table(s, ["Scadenza", "Ultimo", "Var. MTD", "Var. YTD"], rows, Inches(0.6), Inches(1.55),
                   [Inches(2.6), Inches(1.1), Inches(1.1), Inches(1.1)], row_h=Inches(0.27),
                   font=9, group_rows=groups)
        excel["Rates"] = pd.DataFrame(rows, columns=["Scadenza", "Ultimo", "Var. MTD", "Var. YTD"])

        brows = []
        for name, c in BOND_PROXY:
            l, mtd, ytd = perf(ds.y(c, name), asof)
            brows.append([name, fmt_num(l), fmt_pct(mtd), fmt_pct(ytd)])
        deck._text(s, L("Bond Indexes (via ETF)"), Inches(7.0), Inches(1.1), Inches(5), Inches(0.4), size=14, bold=True)
        deck.table(s, ["Indice", "Ultimo", "%MTD", f"%YTD ({asof.year})"], brows, Inches(7.0), Inches(1.55),
                   [Inches(2.6), Inches(0.95), Inches(0.75), Inches(1.15)], row_h=Inches(0.3), font=9)
        excel["Bond Indexes"] = pd.DataFrame(brows, columns=["Indice", "Ultimo", "%MTD", "%YTD"])
        deck._text(s, L("RATES_NOTE").format(d=f"{asof:%d.%m.%Y}"),
                   Inches(7.0), Inches(3.4), Inches(5), Inches(0.6), size=9, color=GREY)
    except Exception as e:
        log.warn(f"Slide rates: {e}")

    log("Banche centrali…")
    try:
        rates = {}
        pol_start = chart_start
        for name in ["USA", "Eurozona"]:
            r = ds.fred_first(FRED_POLICY[name], pol_start - pd.DateOffset(years=1))
            if not r.empty:
                rates[name] = r
        bis = ds.bis_policy(list(BIS_POLICY_CODES.values()))
        for name, code in BIS_POLICY_CODES.items():
            if code in bis:
                rates[name] = bis[code]
            else:
                r = ds.fred_first(FRED_POLICY[name], pol_start - pd.DateOffset(years=1))
                if not r.empty:
                    rates[name] = r
        for k in list(rates):
            r = rates[k]
            before = r[r.index < pol_start]
            r = r[r.index >= pol_start]
            if not before.empty:
                r = pd.concat([pd.Series([before.iloc[-1]], index=[pol_start]), r])
            rates[k] = r[~r.index.duplicated(keep="last")].sort_index()
        if rates:
            s = deck.slide("Politica monetaria: principali banche centrali", "Tassi e reddito fisso")
            deck.picture(s, chart_policy(rates, asof, img("banche_centrali.png"), "FRED, BIS"))
            excel["Banche centrali"] = pd.DataFrame(rates)
    except Exception as e:
        log.warn(f"Slide banche centrali: {e}")

    try:
        trio = {f"US Treasury {m}": ust[m][ust[m].index >= chart_start] for m in ["2Y", "5Y", "10Y"]}
        if any(not v.empty for v in trio.values()):
            s = deck.slide("Rendimento titoli di stato USA", "Tassi e reddito fisso")
            deck.picture(s, chart_multi(trio, f"{L('Andamento')} US Treasury", asof, img("ust_storico.png"),
                                        "Rendimento (%)|yield", "FRED", dec=2))
            excel["US Treasury"] = pd.DataFrame(trio)
    except Exception as e:
        log.warn(f"Slide Treasury: {e}")

    try:
        d_before = asof - pd.DateOffset(years=1)
        now, before = {}, {}
        for m, srs in ust.items():
            a = srs[srs.index <= asof]; b = srs[srs.index <= d_before]
            if not a.empty and not b.empty:
                now[m], before[m] = a.iloc[-1], b.iloc[-1]
        if len(now) >= 3:
            s = deck.slide("Curve dei titoli di Stato statunitensi", "Tassi e reddito fisso")
            deck.picture(s, chart_curve(pd.Series(now), pd.Series(before), asof, d_before,
                                        img("curva_usa.png"), "FRED"), h=Inches(6.0))
            excel["Curva USA"] = pd.DataFrame({"Attuale": now, "Un anno fa": before})
    except Exception as e:
        log.warn(f"Slide curva: {e}")

    # ======================= AZIONARIO =======================
    log("Borse mondiali…")
    try:
        s = deck.slide("Borse mondiali", "Azionario")
        rows, groups = [], []
        for area, lst in BORSE.items():
            groups.append(len(rows)); rows.append([area, "", "", ""])
            for name, c in lst:
                l, mtd, ytd = perf(ds.y(c, name), asof)
                used = ds._chosen.get(name, c[0])
                shown = L(name) if used == c[0] else f"{L(name)} ({L('via')} ETF {used})"
                rows.append([shown, fmt_num(l), fmt_pct(mtd, 2), fmt_pct(ytd, 2)])
        deck.table(s, ["Indice", "Valore", "%MTD", f"%YTD ({asof.year})"], rows, Inches(2.2), Inches(1.2),
                   [Inches(3.6), Inches(1.5), Inches(1.2), Inches(1.5)], row_h=Inches(0.255), font=10,
                   group_rows=groups)
        deck._text(s, f"{L('Fonte:')} Yahoo Finance", Inches(2.2), Inches(7.0),
                   Inches(4), Inches(0.3), size=8, color=GREY)
        excel["Borse mondiali"] = pd.DataFrame(rows, columns=["Indice", "Valore", "%MTD", "%YTD"])
    except Exception as e:
        log.warn(f"Slide borse: {e}")

    try:
        w = ds.y(WORLD_INDEX[1], WORLD_INDEX[0])
        w = w[w.index >= chart_start]
        if not w.empty:
            s = deck.slide(L("Indice mondiale: performance ultimi {n} anni").format(n=years), "Azionario")
            deck.picture(s, chart_rebased({WORLD_INDEX[0]: w}, WORLD_INDEX[0], asof,
                                          img("indice_mondiale.png"), "Yahoo Finance", ytd=False))
    except Exception as e:
        log.warn(f"Slide indice mondiale: {e}")

    try:
        dm = {n: ds.y(c, n) for n, c in DM_EM}
        s = deck.slide("Mercati sviluppati ed emergenti YTD", "Azionario", "Variazione in % (USD)")
        deck.picture(s, chart_rebased(dm, None, asof, img("dm_vs_em_ytd.png"), "Yahoo Finance",
                                      fill_between=True), t=Inches(1.45), h=Inches(5.75))
    except Exception as e:
        log.warn(f"Slide DM/EM: {e}")

    try:
        em = {n: ds.y(c, n) for n, c in EM_REGIONS}
        s = deck.slide("Rendimento YTD delle regioni emergenti", "Azionario", "Variazione in % (USD)")
        deck.picture(s, chart_rebased(em, None, asof, img("em_regioni_ytd.png"), "Yahoo Finance",
                                      dots=True), t=Inches(1.45), h=Inches(5.75))
    except Exception as e:
        log.warn(f"Slide regioni EM: {e}")

    def price_slides(lst, prefix, dec=2):
        for title, section, label, c in lst:
            try:
                sr = ds.y(c, label)
                sr = sr[sr.index >= chart_start]
                if sr.empty:
                    continue
                if ds._chosen.get(label, "").startswith("FRED:"):
                    label = f"{L(label)} {L('(FRED, mensile $/MMBtu)')}"
                d = 4 if sr.iloc[-1] < 10 else dec
                p = chart_price(sr, title, label, asof,
                                img(f"{prefix}_{label.replace('/', '').replace(' ', '_').lower()}.png"),
                                dec=d)
                s = deck.slide(title, section)
                deck.picture(s, p)
                excel[_sheet(L(label))] = sr.rename(L(label)).to_frame()
            except Exception as e:
                log.warn(f"Grafico {title}: {e}")

    log("Grafici indici…")
    price_slides(PRICE_CHARTS, "az")

    # ======================= TITOLI SCELTI DALL'UTENTE =======================
    if custom:
        log(f"Titoli: {', '.join(t for t, _ in custom)}…")
        cstart = asof - pd.DateOffset(years=custom_years)
        cdata = ds.load_custom([t for t, _ in custom], cstart)
        chosen = {}
        for tk, lab in custom:
            try:
                sr = cdata.get(tk, _empty())
                sr = sr[(sr.index >= cstart) & (sr.index <= asof)].dropna()
                if len(sr) < 5:
                    log.warn(f"Nessun dato Yahoo per il titolo {tk}")
                    continue
                chosen[lab] = sr
                safe = re.sub(r"[^A-Za-z0-9]+", "_", tk).strip("_").lower()
                p = chart_price(sr, lab, lab, asof, img(f"titolo_{safe}.png"),
                                dec=4 if sr.iloc[-1] < 10 else 2)
                s = deck.slide(lab, "Titoli")
                deck.picture(s, p)
                excel[_sheet(lab)] = sr.rename(lab).to_frame()
            except Exception as e:
                log.warn(f"Grafico titolo {tk}: {e}")
        if len(chosen) >= 2:
            try:
                p = chart_base100(chosen, asof, img("titoli_confronto.png"))
                s = deck.slide(f"{L('Confronto')}: {' vs '.join(chosen)}", "Titoli")
                deck.picture(s, p)
            except Exception as e:
                log.warn(f"Grafico confronto titoli: {e}")

    # ======================= VALUTE =======================
    log("Valute…")
    try:
        s = deck.slide("Riepilogo valute", "Valute")
        rows = []
        for name, c in VALUTE:
            l, mtd, ytd = perf(ds.y(c, name), asof)
            rows.append([name, fmt_num(l, 4), fmt_pct(mtd), fmt_pct(ytd)])
        deck.table(s, ["Cambio", "Ultimo", "%MTD", "%YTD"], rows, Inches(0.6), Inches(1.4),
                   [Inches(3.0), Inches(1.6), Inches(1.6), Inches(1.6)], row_h=Inches(0.36), font=11)
        excel["Valute"] = pd.DataFrame(rows, columns=["Cambio", "Ultimo", "%MTD", "%YTD"])
    except Exception as e:
        log.warn(f"Slide valute: {e}")
    price_slides(FX_CHARTS, "fx")

    # ======================= MATERIE PRIME & CRYPTO =======================
    log("Materie prime e cripto…")
    try:
        s = deck.slide("Materie prime e criptovalute", "Materie prime e criptovalute")
        rows = []
        for name, c in COMMODITIES:
            l, mtd, ytd = perf(ds.y(c, name), asof)
            rows.append([name, fmt_num(l), fmt_pct(mtd), fmt_pct(ytd)])
        deck.table(s, ["Commodities", "Ultimo", "%MTD", "%YTD"], rows, Inches(0.6), Inches(1.3),
                   [Inches(2.9), Inches(1.2), Inches(0.95), Inches(0.95)], row_h=Inches(0.3), font=10)
        rows2 = []
        for name, c in CRYPTO:
            l, mtd, ytd = perf(ds.y(c, name), asof)
            rows2.append([name, fmt_num(l), fmt_pct(mtd), fmt_pct(ytd)])
        deck.table(s, ["Criptovalute", "Ultimo", "%MTD", "%YTD"], rows2, Inches(7.1), Inches(1.3),
                   [Inches(2.0), Inches(1.3), Inches(0.95), Inches(0.95)], row_h=Inches(0.3), font=10)
        excel["Commodities"] = pd.DataFrame(rows, columns=["Nome", "Ultimo", "%MTD", "%YTD"])
        excel["Crypto"] = pd.DataFrame(rows2, columns=["Nome", "Ultimo", "%MTD", "%YTD"])
    except Exception as e:
        log.warn(f"Slide commodities: {e}")
    price_slides(COMMO_CHARTS, "cmd")

    # ======================= CHIUSURA =======================
    s = deck.slide("Note sui dati")
    note = L("NOTE_BODY")
    deck._text(s, note, Inches(0.6), Inches(1.4), Inches(11.5), Inches(4), size=14, color="#444444")
    if log.warnings:
        deck._text(s, L("Avvisi durante la generazione:") + "\n" + "\n".join("• " + w for w in log.warnings[:12]),
                   Inches(0.6), Inches(4.6), Inches(11.5), Inches(2.6), size=9, color="#A23B3B")

    pptx_path = os.path.join(out_dir, f"Mercati_{stamp}_{LANG.upper()}.pptx")
    deck.save(pptx_path)
    log(f"Presentazione salvata: {pptx_path}")

    try:
        xlsx_path = os.path.join(out_dir, f"Mercati_dati_{stamp}_{LANG.upper()}.xlsx")
        with pd.ExcelWriter(xlsx_path) as xw:
            for name, df in excel.items():
                df.to_excel(xw, sheet_name=_sheet(name))
        log(f"Dati salvati: {xlsx_path}")
    except Exception as e:
        log.warn(f"Excel non salvato: {e}")
    return pptx_path


# =============================================================================
# INTERFACCIA GRAFICA
# =============================================================================

def open_path(p):
    try:
        if sys.platform.startswith("win"):
            os.startfile(p)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", p])
        else:
            subprocess.Popen(["xdg-open", p])
    except Exception:
        pass


def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    root = tk.Tk()
    root.title("Mercati Report")
    root.geometry("660x660")
    root.configure(bg="#FAF8F4")

    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("TFrame", background="#FAF8F4")
    style.configure("TLabel", background="#FAF8F4", font=("Segoe UI", 10))
    style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"), foreground=TITLE_GREY)
    style.configure("Gold.TButton", font=("Segoe UI", 11, "bold"), foreground="white",
                    background=GOLD, padding=8)
    style.map("Gold.TButton", background=[("active", "#B08B45"), ("disabled", "#D9CBB0")])

    frm = ttk.Frame(root, padding=18)
    frm.pack(fill="both", expand=True)
    ttk.Label(frm, text="Mercati Report", style="Title.TLabel").grid(row=0, column=0, columnspan=3, sticky="w")
    ttk.Label(frm, text="Genera la presentazione con grafici e tabelle di mercato aggiornati.",
              foreground=GREY).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 14))

    ttk.Label(frm, text="Dati al (gg.mm.aaaa)").grid(row=2, column=0, sticky="w")
    v_date = tk.StringVar(value=date.today().strftime("%d.%m.%Y"))
    ttk.Entry(frm, textvariable=v_date, width=14).grid(row=2, column=1, sticky="w", pady=4)

    ttk.Label(frm, text="Anni di storico nei grafici").grid(row=3, column=0, sticky="w")
    v_years = tk.IntVar(value=3)
    ttk.Spinbox(frm, from_=1, to=10, textvariable=v_years, width=6).grid(row=3, column=1, sticky="w", pady=4)

    ttk.Label(frm, text="Cartella di salvataggio").grid(row=4, column=0, sticky="w")
    default_out = os.path.join(os.path.expanduser("~"), "Documents", "Mercati Report")
    v_out = tk.StringVar(value=default_out)
    ttk.Entry(frm, textvariable=v_out, width=42).grid(row=4, column=1, sticky="we", pady=4)
    ttk.Button(frm, text="Sfoglia…", command=lambda: v_out.set(
        filedialog.askdirectory(initialdir=v_out.get()) or v_out.get())).grid(row=4, column=2, padx=6)

    ttk.Label(frm, text="Lingua / Language / Idioma").grid(row=5, column=0, sticky="w")
    langs = {"Italiano": "ita", "English": "eng", "Español": "esp"}
    v_lang = tk.StringVar(value="Italiano")
    ttk.Combobox(frm, textvariable=v_lang, values=list(langs), state="readonly", width=12).grid(
        row=5, column=1, sticky="w", pady=4)

    ttk.Label(frm, text="Titoli (ticker Yahoo)").grid(row=6, column=0, sticky="nw", pady=(8, 0))
    v_tick = tk.StringVar(value="")
    ttk.Entry(frm, textvariable=v_tick, width=42).grid(row=6, column=1, columnspan=2, sticky="we", pady=(8, 0))
    ttk.Label(frm, text="Separati da virgola. Esempio:  CAP.SN, NESN.SW, AAPL   "
                        "oppure con nome:  CAP.SN=CAP SA", foreground=GREY,
              font=("Segoe UI", 8)).grid(row=7, column=1, columnspan=2, sticky="w")
    ttk.Label(frm, text="Anni di storico titoli").grid(row=8, column=0, sticky="w")
    v_cyears = tk.IntVar(value=10)
    ttk.Spinbox(frm, from_=1, to=30, textvariable=v_cyears, width=6).grid(row=8, column=1, sticky="w", pady=4)

    btn = ttk.Button(frm, text="Genera presentazione", style="Gold.TButton")
    btn.grid(row=9, column=0, columnspan=3, sticky="we", pady=(16, 8))
    pb = ttk.Progressbar(frm, mode="indeterminate")
    pb.grid(row=10, column=0, columnspan=3, sticky="we")

    txt = tk.Text(frm, height=12, bg="white", relief="flat", font=("Consolas", 9), wrap="word")
    txt.grid(row=11, column=0, columnspan=3, sticky="nsew", pady=(10, 0))
    frm.columnconfigure(1, weight=1)
    frm.rowconfigure(11, weight=1)

    def write(msg):
        root.after(0, lambda: (txt.insert("end", msg + "\n"), txt.see("end")))

    def done(path, err):
        pb.stop()
        btn.state(["!disabled"])
        if err:
            messagebox.showerror("Errore", err)
        else:
            if messagebox.askyesno("Fatto", "Presentazione creata.\nVuoi aprirla?"):
                open_path(path)

    def go():
        try:
            asof = datetime.strptime(v_date.get().strip(), "%d.%m.%Y")
        except ValueError:
            messagebox.showwarning("Data", "Formato data non valido, usa gg.mm.aaaa")
            return
        out = v_out.get().strip()
        os.makedirs(out, exist_ok=True)
        txt.delete("1.0", "end")
        btn.state(["disabled"])
        pb.start(12)

        def work():
            try:
                p = build_report(asof, int(v_years.get()), out, Log(write),
                                 lang=langs.get(v_lang.get(), "ita"),
                                 custom=parse_tickers(v_tick.get()),
                                 custom_years=int(v_cyears.get()))
                root.after(0, lambda: done(p, None))
            except Exception as e:
                write(traceback.format_exc())
                root.after(0, lambda: done(None, str(e)))
        threading.Thread(target=work, daemon=True).start()

    btn.configure(command=go)
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--cli":
        # uso: python mercati_report.py --cli [gg.mm.aaaa] [anni] [cartella] [ita|eng|esp] ["CAP.SN, NESN.SW"]
        d = datetime.strptime(sys.argv[2], "%d.%m.%Y") if len(sys.argv) > 2 else datetime.today()
        yrs = int(sys.argv[3]) if len(sys.argv) > 3 else 3
        out = sys.argv[4] if len(sys.argv) > 4 else os.getcwd()
        lg = sys.argv[5] if len(sys.argv) > 5 else "ita"
        tks = parse_tickers(sys.argv[6]) if len(sys.argv) > 6 else None
        build_report(d, yrs, out, Log(), lang=lg, custom=tks)
    else:
        run_gui()
