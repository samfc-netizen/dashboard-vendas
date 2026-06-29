import re
import unicodedata
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from datetime import date
from urllib.parse import quote
import io
import zipfile
import json
from pathlib import Path

st.set_page_config(page_title="Dashboard de Vendas", layout="wide")

# =========================================================
# UX / VISUAL / PWA BÁSICO
# =========================================================

def aplicar_estilo_visual():
    st.markdown(
        """
        <style>
        :root {
            --app-bg: #f6f8fc;
            --card-bg: #ffffff;
            --text-main: #111827;
            --text-muted: #475569;
            --brand: #0f4c81;
            --brand-2: #2563eb;
            --border: #e5e7eb;
            --success: #0f766e;
            --warning: #b45309;
            --danger: #b91c1c;
        }

        .stApp { background: var(--app-bg); }

        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%);
            color: #ffffff;
        }
        section[data-testid="stSidebar"] * { color: #ffffff !important; }
        section[data-testid="stSidebar"] input,
        section[data-testid="stSidebar"] textarea,
        section[data-testid="stSidebar"] select,
        section[data-testid="stSidebar"] div[data-baseweb="select"] *,
        section[data-testid="stSidebar"] div[data-baseweb="input"] input {
            color: #111827 !important;
            background-color: #ffffff !important;
        }
        section[data-testid="stSidebar"] label,
        section[data-testid="stSidebar"] p,
        section[data-testid="stSidebar"] span {
            color: #f8fafc !important;
        }

        .block-container { padding-top: 1.25rem; max-width: 1500px; }

        .hero-card {
            background: linear-gradient(135deg, #0f4c81 0%, #1d4ed8 52%, #0f172a 100%);
            color: white;
            border-radius: 22px;
            padding: 28px 30px;
            margin-bottom: 20px;
            box-shadow: 0 18px 45px rgba(15, 23, 42, 0.18);
        }
        .hero-card h1 { margin: 0; font-size: 32px; line-height: 1.1; color: #ffffff; }
        .hero-card p { margin: 8px 0 0 0; color: #e2e8f0; font-size: 15px; }
        .section-card {
            background: var(--card-bg);
            border: 1px solid var(--border);
            border-radius: 18px;
            padding: 18px 20px;
            margin: 12px 0 18px 0;
            box-shadow: 0 8px 25px rgba(15, 23, 42, 0.06);
        }
        .mini-card {
            background: #ffffff;
            border: 1px solid #e5e7eb;
            border-radius: 16px;
            padding: 16px;
            min-height: 96px;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.05);
        }
        .mini-card .label { color: #64748b; font-size: 13px; font-weight: 600; margin-bottom: 6px; }
        .mini-card .value { color: #111827; font-size: 24px; font-weight: 800; }
        .mini-card .hint { color: #64748b; font-size: 12px; margin-top: 4px; }
        div[data-testid="stMetric"] {
            background: #ffffff;
            border: 1px solid #e5e7eb;
            border-radius: 16px;
            padding: 14px 16px;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.05);
            min-height: 88px;
            overflow: visible !important;
        }

        /* Corrige KPI cortado com reticências em cards estreitos */
        div[data-testid="stMetric"] label,
        div[data-testid="stMetric"] [data-testid="stMetricLabel"],
        div[data-testid="stMetric"] [data-testid="stMetricValue"],
        div[data-testid="stMetric"] [data-testid="stMetricDelta"] {
            overflow: visible !important;
            text-overflow: clip !important;
            white-space: normal !important;
        }
        div[data-testid="stMetric"] [data-testid="stMetricValue"] {
            color: #111827 !important;
            font-size: clamp(18px, 1.55vw, 28px) !important;
            line-height: 1.12 !important;
            letter-spacing: -0.02em !important;
            word-break: keep-all !important;
        }
        div[data-testid="stMetric"] [data-testid="stMetricValue"] * {
            overflow: visible !important;
            text-overflow: clip !important;
            white-space: normal !important;
            line-height: 1.12 !important;
        }
        div[data-testid="stMetric"] [data-testid="stMetricLabel"] {
            font-size: 12px !important;
            line-height: 1.25 !important;
            color: #334155 !important;
        }
        div[data-testid="stMetric"] [data-testid="stMetricDelta"] {
            font-size: 12px !important;
        }
        @media (max-width: 1200px) {
            div[data-testid="stMetric"] [data-testid="stMetricValue"] {
                font-size: 20px !important;
            }
        }
        .whatsapp-box {
            background: #f0fdf4;
            border: 1px solid #bbf7d0;
            border-radius: 16px;
            padding: 16px;
            margin: 10px 0 14px 0;
            color: #14532d;
        }
        .warning-box {
            background: #fff7ed;
            border: 1px solid #fed7aa;
            border-radius: 16px;
            padding: 16px;
            margin: 10px 0 14px 0;
            color: #7c2d12;
        }
        textarea, input { color: #111827 !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header():
    st.markdown(
        """
        <div class="hero-card">
            <h1>Dashboard de Vendas Dauto Tintas</h1>
            <p>Indicadores comerciais, comparativos e relatórios prontos para WhatsApp.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def instalar_pwa_basico():
    """Cria arquivos básicos de PWA e injeta manifest/service worker.
    Observação: em Streamlit, o comportamento depende do deploy, HTTPS e permissões do navegador.
    """
    try:
        manifest = {
            "name": "Dashboard de Vendas Dauto Tintas",
            "short_name": "Dauto Vendas",
            "start_url": "./",
            "display": "standalone",
            "background_color": "#f6f8fc",
            "theme_color": "#0f4c81",
            "description": "Dashboard comercial com relatórios de vendas.",
            "icons": [],
        }
        Path("manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        Path("sw.js").write_text(
            """
self.addEventListener('install', event => { self.skipWaiting(); });
self.addEventListener('activate', event => { event.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', event => { event.respondWith(fetch(event.request)); });
""".strip(),
            encoding="utf-8",
        )
        st.components.v1.html(
            """
<link rel="manifest" href="/manifest.json">
<meta name="theme-color" content="#0f4c81">
<script>
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('/sw.js').catch(function(e) { console.log('SW não registrado', e); });
}
</script>
""",
            height=0,
        )
    except Exception:
        pass


aplicar_estilo_visual()
instalar_pwa_basico()
render_header()

top_card = st.empty()

ARQUIVO_ZIP = "BASE.zip"
ARQUIVOS_BASE_SEPARADOS = ["BASE 1.xlsx", "BASE 2.xlsx", "BASE 3.xlsx"]
ARQUIVO_METAS = "METAS.xlsx"
ARQUIVO_COMPRAS = "COMPRAS.xlsx"


def _nome_arquivo_key(nome: str) -> str:
    """Normaliza nome de arquivo para localizar variações com espaços/acentos."""
    return re.sub(r"[^A-Z0-9]", "", unicodedata.normalize("NFKD", str(nome).upper()).encode("ASCII", "ignore").decode("ASCII"))


def resolver_arquivo_excel(candidatos: list[str]) -> Path | None:
    """Procura arquivo .xlsx na pasta do app aceitando variações de espaço no nome."""
    pasta = Path(".")

    for nome in candidatos:
        caminho = pasta / nome
        if caminho.exists():
            return caminho

    arquivos_xlsx = [p for p in pasta.glob("*.xlsx") if not p.name.startswith("~$")]
    mapa = {_nome_arquivo_key(p.name): p for p in arquivos_xlsx}
    for nome in candidatos:
        chave = _nome_arquivo_key(nome)
        if chave in mapa:
            return mapa[chave]

    return None


@st.cache_data(ttl=10)
def carregar_base_consolidada() -> pd.DataFrame:
    """Carrega a base de vendas.

    Prioridade:
    1) BASE 1.xlsx + BASE 2.xlsx + BASE 3.xlsx na pasta do app.
    2) BASE.xlsx direto na pasta do app.
    3) Primeiro .xlsx encontrado dentro de BASE.zip, mantendo compatibilidade com a versão antiga.
    """
    dfs = []
    arquivos_encontrados = []

    for nome in ARQUIVOS_BASE_SEPARADOS:
        caminho = resolver_arquivo_excel([nome])
        if caminho is not None:
            dfx = pd.read_excel(caminho, sheet_name=0, engine="openpyxl")
            dfx["_ARQUIVO_BASE"] = caminho.name
            dfs.append(dfx)
            arquivos_encontrados.append(caminho.name)

    if dfs:
        return pd.concat(dfs, ignore_index=True)

    caminho_base_unica = resolver_arquivo_excel(["BASE.xlsx"])
    if caminho_base_unica is not None:
        df_unica = pd.read_excel(caminho_base_unica, sheet_name=0, engine="openpyxl")
        df_unica["_ARQUIVO_BASE"] = caminho_base_unica.name
        return df_unica

    caminho_zip = Path(ARQUIVO_ZIP)
    if caminho_zip.exists():
        with zipfile.ZipFile(caminho_zip, "r") as z:
            arquivos_excel = [
                nome for nome in z.namelist()
                if nome.lower().endswith(".xlsx") and not Path(nome).name.startswith("~$")
            ]
            if arquivos_excel:
                nome_excel = arquivos_excel[0]
                with z.open(nome_excel) as f:
                    df_zip = pd.read_excel(io.BytesIO(f.read()), sheet_name=0, engine="openpyxl")
                    df_zip["_ARQUIVO_BASE"] = f"{ARQUIVO_ZIP}/{Path(nome_excel).name}"
                    return df_zip

    st.error(
        "Nenhuma base de vendas foi encontrada. Coloque na pasta do app os arquivos "
        "BASE 1.xlsx, BASE 2.xlsx e BASE 3.xlsx. Também aceito BASE.xlsx ou BASE.zip como fallback."
    )
    st.stop()


def ler_excel_base(sheet_name=0) -> pd.DataFrame:
    """Compatibilidade: retorna a base consolidada de vendas."""
    if sheet_name not in (0, "0", None):
        raise ValueError("A base de vendas separada é lida sempre pela primeira aba de cada arquivo BASE.")
    return carregar_base_consolidada()


def ler_excel_arquivo(nome_arquivo: str, sheet_name=0) -> pd.DataFrame:
    """Lê uma planilha auxiliar específica na pasta do app."""
    caminho = resolver_arquivo_excel([nome_arquivo])
    if caminho is None:
        raise FileNotFoundError(f"Arquivo {nome_arquivo} não encontrado na pasta do app.")
    return pd.read_excel(caminho, sheet_name=sheet_name, engine="openpyxl")


# ========= Helpers =========
def format_brl(v: float) -> str:
    try:
        return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "—"


def parse_number_any(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)

    s = str(v).strip()
    if s == "" or s.lower() in {"nan", "none"}:
        return None

    s = s.replace("\u00a0", " ")
    s = s.replace("R$", "").strip()
    s = s.replace(" ", "")

    s = re.sub(r"[^0-9\.\,\-]", "", s)
    if s in {"", "-", "-.", "-,"}:
        return None

    neg = s.startswith("-")
    s = s.lstrip("-")

    if "." in s and "," in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    else:
        pass

    try:
        out = float(s)
        return -out if neg else out
    except Exception:
        return None


def to_float_series(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    return series.apply(parse_number_any).astype("float64")


def normalize_dim(series: pd.Series, fallback: str) -> pd.Series:
    s = series.astype("string")
    s = s.fillna("")
    s = s.str.strip()
    s = s.replace(["nan", "NaN", "NONE", "None", "none"], "")
    s = s.where(s != "", other=fallback)
    return s


def canonical_key(text: str) -> str:
    if text is None:
        return ""
    s = str(text).strip().upper()
    s = re.sub(r"\s+", "", s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s


def pick_first_existing_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    cols_norm = {c: canonical_key(c) for c in df.columns}
    cand_norm = [canonical_key(x) for x in candidates]
    for c, cn in cols_norm.items():
        if cn in cand_norm:
            return c
    # tenta por "contains"
    for c in df.columns:
        ckey = canonical_key(c)
        for wanted in cand_norm:
            if wanted and wanted in ckey:
                return c
    return None


def month_to_num_pt(value) -> int | None:
    key = canonical_key(value)
    mapa = {
        "JAN": 1, "JANEIRO": 1,
        "FEV": 2, "FEVEREIRO": 2,
        "MAR": 3, "MARCO": 3, "MARÇO": 3,
        "ABR": 4, "ABRIL": 4,
        "MAI": 5, "MAIO": 5,
        "JUN": 6, "JUNHO": 6,
        "JUL": 7, "JULHO": 7,
        "AGO": 8, "AGOSTO": 8,
        "SET": 9, "SETEMBRO": 9,
        "OUT": 10, "OUTUBRO": 10,
        "NOV": 11, "NOVEMBRO": 11,
        "DEZ": 12, "DEZEMBRO": 12,
    }
    return mapa.get(key)


def format_decimal_pt(v, casas: int = 1) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    try:
        txt = f"{float(v):,.{casas}f}"
        return txt.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "—"


MESES = [
    (1, "JAN"),
    (2, "FEV"),
    (3, "MAR"),
    (4, "ABR"),
    (5, "MAI"),
    (6, "JUN"),
    (7, "JUL"),
    (8, "AGO"),
    (9, "SET"),
    (10, "OUT"),
    (11, "NOV"),
    (12, "DEZ"),
]

# Ordem desejada (UNAÍ por último)
LOJA_KEY_ORDER = ["ADE", "GAMA", "SOFNORTE", "CEILANDIA", "SIA", "AGLINDAS", "GUARA", "LUZIANIA", "UNAI"]
LOJA_KEY_RANK = {k: i for i, k in enumerate(LOJA_KEY_ORDER)}
DEFAULT_RANK = 10_000


def color_pos_neg(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    try:
        x = float(v)
    except Exception:
        return ""
    if x > 0:
        return "color: #1f77b4; font-weight: 700;"  # azul
    if x < 0:
        return "color: #d62728; font-weight: 700;"  # vermelho
    return ""


def styler_map_compat(styler, func, subset=None):
    """Compatibilidade entre versões do pandas/Styler.
    Em versões mais novas, Styler.applymap foi descontinuado/removido em favor de Styler.map.
    """
    if hasattr(styler, "map"):
        return styler.map(func, subset=subset)
    return styler.applymap(func, subset=subset)


def row_total_style(row):
    if str(row.get("LOJA", "")).upper() == "TOTAL":
        return ["background-color: #f2f2f2; font-weight: 900;"] * len(row)
    return [""] * len(row)


def month_block_style(df_view: pd.DataFrame):
    out = df_view.copy()
    out["MES"] = out["MES"].astype(str)
    last = None
    new_vals = []
    for m in out["MES"].tolist():
        if m == last:
            new_vals.append("")
        else:
            new_vals.append(m)
            last = m
    out["MES"] = new_vals
    return out


PLOT_CONFIG_INTERACTIVE_NO_ZOOM = {
    "displayModeBar": True,
    "displaylogo": False,
    "responsive": True,
    "scrollZoom": False,
    "doubleClick": "reset",
    "modeBarButtonsToRemove": ["zoom2d", "pan2d", "zoomIn2d", "zoomOut2d", "autoScale2d"],
}


# ========= Cache de dados =========
@st.cache_data(ttl=10)
def carregar_dados():
    df = ler_excel_base(sheet_name=0)
    df.columns = df.columns.astype(str).str.strip()

    # força object -> string (ajuda em filtros)
    obj_cols = df.select_dtypes(include=["object"]).columns
    if len(obj_cols) > 0:
        df[obj_cols] = df[obj_cols].astype("string")

    # DATA
    if "DATA" in df.columns:
        df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce", dayfirst=True)
    else:
        df["DATA"] = pd.NaT

    # Numéricos principais
    if "QTD" in df.columns:
        df["QTD_NUM"] = to_float_series(df["QTD"])
    else:
        df["QTD_NUM"] = pd.Series([None] * len(df), dtype="float64")

    if "UNIT" in df.columns:
        df["UNIT_NUM"] = to_float_series(df["UNIT"])
    else:
        df["UNIT_NUM"] = pd.Series([None] * len(df), dtype="float64")

    # Faturamento por linha
    df["FAT_LINHA"] = df["QTD_NUM"] * df["UNIT_NUM"]

    # Dimensões básicas
    df["LOJA_N"] = normalize_dim(df["LOJA"], "SEM LOJA") if "LOJA" in df.columns else pd.Series(["SEM LOJA"] * len(df), dtype="string")
    df["VENDEDOR_N"] = normalize_dim(df["VENDEDOR"], "SEM VENDEDOR") if "VENDEDOR" in df.columns else pd.Series(["SEM VENDEDOR"] * len(df), dtype="string")
    df["MARCA_N"] = normalize_dim(df["MARCA"], "SEM MARCA") if "MARCA" in df.columns else pd.Series(["SEM MARCA"] * len(df), dtype="string")
    df["SEGMENTO_N"] = normalize_dim(df["SEGMENTO"], "SEM SEGMENTO") if "SEGMENTO" in df.columns else pd.Series(["SEM SEGMENTO"] * len(df), dtype="string")
    df["LINHA_N"] = normalize_dim(df["LINHA"], "SEM LINHA") if "LINHA" in df.columns else pd.Series(["SEM LINHA"] * len(df), dtype="string")

    # PRODUTO (CÓD + DESCRIÇÃO)
    col_cod = pick_first_existing_col(df, ["CÓD", "COD", "CÓDIGO", "CODIGO", "COD.", "COD PROD", "CODPROD", "COD PRODUTO"])
    if col_cod is not None:
        df["COD_N"] = normalize_dim(df[col_cod], "SEM COD")
    else:
        df["COD_N"] = pd.Series(["SEM COD"] * len(df), dtype="string")

    col_desc = pick_first_existing_col(df, ["DESCRIÇÃO", "DESCRICAO", "DESCR", "DESCRI", "PRODUTO", "DESCRIÇÃO / REFERÊNCIA", "DESCRICAO / REFERENCIA", "DESCRIÇÃO/REFERÊNCIA", "DESCRICAO/REFERENCIA"])
    if col_desc is not None:
        df["DESC_N"] = normalize_dim(df[col_desc], "SEM DESCRIÇÃO")
    else:
        df["DESC_N"] = pd.Series(["SEM DESCRIÇÃO"] * len(df), dtype="string")

    # DESCRIÇÃO CANÔNICA POR CÓDIGO
    #
    # A base pode trazer descrições diferentes para o mesmo código de produto.
    # Para evitar duplicidade e falha de aglutinação nos relatórios, o produto
    # passa a ser agrupado pelo CÓDIGO, usando apenas uma descrição por código.
    # Critério: para cada código, utiliza a descrição que mais aparece na base;
    # em caso de empate, mantém a primeira encontrada na planilha.
    df["_ORDEM_LINHA"] = range(len(df))
    desc_por_codigo = (
        df[df["COD_N"].astype(str).str.upper().ne("SEM COD")]
        .groupby(["COD_N", "DESC_N"], dropna=False)
        .agg(QTD_DESC=("DESC_N", "size"), PRIMEIRA_LINHA=("_ORDEM_LINHA", "min"))
        .reset_index()
        .sort_values(["COD_N", "QTD_DESC", "PRIMEIRA_LINHA"], ascending=[True, False, True])
        .drop_duplicates("COD_N")
        .set_index("COD_N")["DESC_N"]
        .to_dict()
    )
    df["DESC_CANON_N"] = df["COD_N"].map(desc_por_codigo).fillna(df["DESC_N"])
    df["PRODUTO_KEY"] = df["COD_N"].astype(str)
    df["_ORDEM_LINHA"] = df["_ORDEM_LINHA"].astype("int64")

    # CLIENTE
    col_cliente = pick_first_existing_col(df, ["CLIENTE", "CLIENTE_NOME", "NOMECLIENTE", "RAZAOSOCIAL", "RAZAO SOCIAL"])
    if col_cliente is not None:
        df["CLIENTE_N"] = normalize_dim(df[col_cliente], "SEM CLIENTE")
    else:
        df["CLIENTE_N"] = pd.Series(["SEM CLIENTE"] * len(df), dtype="string")

    # VR TOTAL (se existir)
    col_vr = pick_first_existing_col(
        df,
        ["VR TOTAL", "VRTOTAL", "VR_TOTAL", "VALOR TOTAL", "VALOR_TOTAL", "VR TOTAL (NF)", "VALOR A FATURAR (NF)", "VALOR A FATURAR NF"],
    )
    if col_vr is not None:
        df["VR_TOTAL_NUM"] = to_float_series(df[col_vr])
    else:
        df["VR_TOTAL_NUM"] = pd.Series([None] * len(df), dtype="float64")

    # ✅ CUSTO TT + ST (coluna AA no Excel, buscamos pelo cabeçalho)
    col_custo = pick_first_existing_col(
        df,
        [
            "CUSTO TT + ST",
            "CUSTO TT+ST",
            "CUSTOTT+ST",
            "CUSTO TOTAL + ST",
            "CUSTO TOTAL+ST",
            "CUSTO+ST",
            "CUSTO + ST",
            "CUSTO_ST",
            "CUSTO ST",
            "CUSTO COM ST",
            "CUSTO COM ST (NF)",
        ],
    )
    if col_custo is not None:
        df["CUSTO_ST_NUM"] = to_float_series(df[col_custo])
    else:
        df["CUSTO_ST_NUM"] = pd.Series([None] * len(df), dtype="float64")

    # CMV (custo) específico da coluna T (conforme planilha BASE EXCEL)
    # Preferência: coluna cujo cabeçalho seja "CUSTO". Fallback: 20ª coluna (índice 19 = coluna T no Excel).
    col_cmv_t = pick_first_existing_col(df, ["CUSTO"])
    if col_cmv_t is None:
        try:
            col_cmv_t = df.columns[19]
        except Exception:
            col_cmv_t = None

    if col_cmv_t is not None:
        df["CUSTO_T_NUM"] = to_float_series(df[col_cmv_t])
    else:
        df["CUSTO_T_NUM"] = pd.Series([None] * len(df), dtype="float64")

    # Chave canônica da loja
    df["LOJA_KEY"] = df["LOJA_N"].astype("string").fillna("").apply(canonical_key)

    return df


@st.cache_data(ttl=10)
def carregar_movimentacoes_compras():
    """Lê as abas COMPRAS e DEVOLUÇÕES (ou variações do nome) e normaliza:
    - DATA (se existir)
    - LOJA / UNIDADE (se existir)
    - TOT. DOC (valor)
    Retorna: (df_compras, df_devol)
    """

    def _try_read_sheet(candidates: list[str]) -> pd.DataFrame:
        # Agora compras/devoluções ficam no arquivo COMPRAS.xlsx.
        # Mantém fallback para a base antiga, caso alguém ainda esteja usando BASE.zip/BASE.xlsx com abas internas.
        for sh in candidates:
            try:
                dfx = ler_excel_arquivo(ARQUIVO_COMPRAS, sheet_name=sh)
                dfx.columns = dfx.columns.astype(str).str.strip()
                obj_cols = dfx.select_dtypes(include=["object"]).columns
                if len(obj_cols) > 0:
                    dfx[obj_cols] = dfx[obj_cols].astype("string")
                return dfx
            except Exception:
                continue

        for sh in candidates:
            try:
                dfx = ler_excel_base(sheet_name=sh)
                dfx.columns = dfx.columns.astype(str).str.strip()
                obj_cols = dfx.select_dtypes(include=["object"]).columns
                if len(obj_cols) > 0:
                    dfx[obj_cols] = dfx[obj_cols].astype("string")
                return dfx
            except Exception:
                continue
        return pd.DataFrame()

    def _normalize_mov(dfm: pd.DataFrame) -> pd.DataFrame:
        if dfm is None or dfm.empty:
            return pd.DataFrame(
                {
                    "DATA": pd.to_datetime([], errors="coerce"),
                    "LOJA_N": pd.Series([], dtype="string"),
                    "LOJA_KEY": pd.Series([], dtype="string"),
                    "TOT_DOC_NUM": pd.Series([], dtype="float64"),
                }
            )

        # DATA (tenta achar coluna de data)
        col_data = pick_first_existing_col(
            dfm,
            ["DATA", "DATA EMISSAO", "DATA EMISSÃO", "DATA EMISSÃO NF", "DATA EMISSAO NF", "DT EMISSAO", "DT EMISSÃO"],
        )
        if col_data is not None:
            dfm["DATA"] = pd.to_datetime(dfm[col_data], errors="coerce", dayfirst=True)
        else:
            dfm["DATA"] = pd.NaT

        # LOJA (tenta achar coluna de loja / unidade)
        col_loja = pick_first_existing_col(dfm, ["LOJA", "UNIDADE", "FILIAL", "UNID", "UNID."])
        if col_loja is not None:
            dfm["LOJA_N"] = normalize_dim(dfm[col_loja], "SEM LOJA")
        else:
            dfm["LOJA_N"] = pd.Series(["SEM LOJA"] * len(dfm), dtype="string")

        dfm["LOJA_KEY"] = dfm["LOJA_N"].astype("string").fillna("").apply(canonical_key)

        # TOT. DOC (valor)
        col_tot = pick_first_existing_col(
            dfm,
            ["TOT. DOC", "TOT DOC", "TOT.DOC", "TOTAL DOC", "TOTAL DOCUMENTO", "VALOR DOC", "VALOR DOCUMENTO"],
        )
        if col_tot is not None:
            dfm["TOT_DOC_NUM"] = to_float_series(dfm[col_tot])
        else:
            dfm["TOT_DOC_NUM"] = pd.Series([None] * len(dfm), dtype="float64")

        return dfm[["DATA", "LOJA_N", "LOJA_KEY", "TOT_DOC_NUM"]].copy()

    df_compras_raw = _try_read_sheet(["COMPRAS"])
    df_devol_raw = _try_read_sheet(["DEVOLUÇÕES", "EDVOLUÇÕES", "DEVOLUCOES", "EDVOLUCOES"])

    df_compras = _normalize_mov(df_compras_raw)
    df_devol = _normalize_mov(df_devol_raw)

    return df_compras, df_devol


@st.cache_data(ttl=10)
def carregar_referencias_planejamento():
    def _read_first(candidates: list[str]) -> pd.DataFrame:
        # Agora Dias úteis, Meta Dauto Tintas e Ano-1 ficam no arquivo METAS.xlsx.
        # Mantém fallback para a base antiga com abas internas.
        for sh in candidates:
            try:
                dfx = ler_excel_arquivo(ARQUIVO_METAS, sheet_name=sh)
                dfx.columns = dfx.columns.astype(str).str.strip()
                return dfx
            except Exception:
                continue

        for sh in candidates:
            try:
                dfx = ler_excel_base(sheet_name=sh)
                dfx.columns = dfx.columns.astype(str).str.strip()
                return dfx
            except Exception:
                continue
        return pd.DataFrame()

    def _wide_to_long(dfw: pd.DataFrame, value_name: str) -> pd.DataFrame:
        if dfw is None or dfw.empty:
            return pd.DataFrame(columns=["LOJA_KEY", "MES_NUM", "MES", value_name])

        cols = list(dfw.columns)
        if not cols:
            return pd.DataFrame(columns=["LOJA_KEY", "MES_NUM", "MES", value_name])

        month_col = cols[0]
        loja_cols = [c for c in cols[1:] if canonical_key(c) not in {"ANO", "ANOBASE", "ANOATUAL"}]

        out = dfw[[month_col] + loja_cols].copy()
        out["MES_NUM"] = out[month_col].apply(month_to_num_pt)
        out = out[out["MES_NUM"].notna()].copy()
        if out.empty:
            return pd.DataFrame(columns=["LOJA_KEY", "MES_NUM", "MES", value_name])

        out = out.melt(id_vars=["MES_NUM"], value_vars=loja_cols, var_name="LOJA_RAW", value_name=value_name)
        out["LOJA_KEY"] = out["LOJA_RAW"].apply(canonical_key)
        out[value_name] = to_float_series(out[value_name])
        out["MES_NUM"] = out["MES_NUM"].astype(int)
        out["MES"] = out["MES_NUM"].map({m: nome for m, nome in MESES})
        out = out[["LOJA_KEY", "MES_NUM", "MES", value_name]].copy()
        return out

    df_meta_raw = _read_first(["META DAUTO TINTAS", "METAS DAUTO TINTAS", "METAS"])
    df_ano1_raw = _read_first(["ANO-1", "Ano-1", "ANO 1", "ANO1"])
    df_dias_raw = _read_first(["DIAS ÚTEIS", "DIAS UTEIS"])

    df_meta_long = _wide_to_long(df_meta_raw, "META")
    df_ano1_long = _wide_to_long(df_ano1_raw, "VENDAS_2025")

    dias_map = {}
    if df_dias_raw is not None and not df_dias_raw.empty:
        col_mes = pick_first_existing_col(df_dias_raw, ["MÊS", "MES"])
        col_dias = pick_first_existing_col(df_dias_raw, ["DIAS ÚTEIS EQUIVALENTES", "DIAS UTEIS EQUIVALENTES", "DIAS ÚTEIS", "DIAS UTEIS", "DIAS"])
        if col_mes is not None and col_dias is not None:
            aux = df_dias_raw[[col_mes, col_dias]].copy()
            aux["MES_NUM"] = aux[col_mes].apply(month_to_num_pt)
            aux["DIAS_EQ"] = to_float_series(aux[col_dias])
            aux = aux[aux["MES_NUM"].notna()].copy()
            dias_map = {
                int(r["MES_NUM"]): float(r["DIAS_EQ"])
                for _, r in aux.iterrows()
                if r["DIAS_EQ"] is not None and not (isinstance(r["DIAS_EQ"], float) and pd.isna(r["DIAS_EQ"]))
            }

    return df_meta_long, df_ano1_long, dias_map


# ========= Fixos =========

ANO_BASE = 2025
ANO_ATUAL = 2026

VENDAS_2025_FIXAS = {
    "ADE": {1: parse_number_any("173.495,62"), 2: parse_number_any("161.844,31"), 3: parse_number_any("156.461,49"),
            4: parse_number_any("160.042,31"), 5: parse_number_any("170.514,18"), 6: parse_number_any("127.951,77"),
            7: parse_number_any("171.267,64"), 8: parse_number_any("141.435,62"), 9: parse_number_any("173.581,08"),
            10: parse_number_any("129.854,69"), 11: parse_number_any("182.310,63"), 12: parse_number_any("110.305,86")},
    "GAMA": {1: parse_number_any("126.149,09"), 2: parse_number_any("112.627,88"), 3: parse_number_any("112.145,06"),
             4: parse_number_any("108.381,08"), 5: parse_number_any("113.956,72"), 6: parse_number_any("113.556,07"),
             7: parse_number_any("169.342,44"), 8: parse_number_any("147.321,43"), 9: parse_number_any("145.956,09"),
             10: parse_number_any("174.493,52"), 11: parse_number_any("143.130,29"), 12: parse_number_any("72.571,55")},
    "SOFNORTE": {1: parse_number_any("159.358,73"), 2: parse_number_any("134.693,97"), 3: parse_number_any("145.117,53"),
                 4: parse_number_any("173.621,70"), 5: parse_number_any("186.950,35"), 6: parse_number_any("156.154,16"),
                 7: parse_number_any("235.304,73"), 8: parse_number_any("217.834,81"), 9: parse_number_any("166.037,32"),
                 10: parse_number_any("161.261,52"), 11: parse_number_any("217.652,47"), 12: parse_number_any("98.613,61")},
    "CEILANDIA": {1: parse_number_any("194.390,32"), 2: parse_number_any("184.944,70"), 3: parse_number_any("185.770,54"),
                  4: parse_number_any("205.408,63"), 5: parse_number_any("228.472,84"), 6: parse_number_any("196.508,96"),
                  7: parse_number_any("253.148,05"), 8: parse_number_any("171.602,48"), 9: parse_number_any("216.952,27"),
                  10: parse_number_any("222.850,50"), 11: parse_number_any("251.524,47"), 12: parse_number_any("166.248,52")},
    "SIA": {1: parse_number_any("251.137,54"), 2: parse_number_any("230.566,96"), 3: parse_number_any("227.002,61"),
            4: parse_number_any("244.790,56"), 5: parse_number_any("237.861,78"), 6: parse_number_any("238.836,13"),
            7: parse_number_any("322.817,82"), 8: parse_number_any("222.557,71"), 9: parse_number_any("237.103,03"),
            10: parse_number_any("268.726,51"), 11: parse_number_any("261.805,42"), 12: parse_number_any("181.533,68")},
    "UNAI": {1: parse_number_any("105.139,34"), 2: parse_number_any("76.559,01"), 3: parse_number_any("89.579,81"),
             4: parse_number_any("110.549,31"), 5: parse_number_any("114.805,18"), 6: parse_number_any("106.100,04"),
             7: parse_number_any("132.666,72"), 8: parse_number_any("101.520,16"), 9: parse_number_any("115.579,91"),
             10: parse_number_any("94.022,74"), 11: parse_number_any("118.884,79"), 12: parse_number_any("86.025,19")},
    "AGLINDAS": {1: parse_number_any("93.433,31"), 2: parse_number_any("97.394,14"), 3: parse_number_any("111.490,33"),
                 4: parse_number_any("98.982,15"), 5: parse_number_any("113.704,23"), 6: parse_number_any("80.362,14"),
                 7: parse_number_any("116.011,78"), 8: parse_number_any("90.763,40"), 9: parse_number_any("100.906,32"),
                 10: parse_number_any("102.186,12"), 11: parse_number_any("99.542,13"), 12: parse_number_any("77.361,44")},
    "GUARA": {1: parse_number_any("224.115,64"), 2: parse_number_any("171.978,22"), 3: parse_number_any("184.169,25"),
              4: parse_number_any("175.469,30"), 5: parse_number_any("217.827,60"), 6: parse_number_any("211.533,27"),
              7: parse_number_any("237.466,11"), 8: parse_number_any("188.321,52"), 9: parse_number_any("185.359,79"),
              10: parse_number_any("216.227,79"), 11: parse_number_any("201.134,31"), 12: parse_number_any("145.289,81")},
    "LUZIANIA": {1: parse_number_any("238.008,58"), 2: parse_number_any("211.472,43"), 3: parse_number_any("231.544,50"),
                 4: parse_number_any("214.753,08"), 5: parse_number_any("283.003,12"), 6: parse_number_any("257.304,71"),
                 7: parse_number_any("319.198,39"), 8: parse_number_any("234.667,93"), 9: parse_number_any("239.019,20"),
                 10: parse_number_any("269.826,56"), 11: parse_number_any("348.121,51"), 12: parse_number_any("237.419,56")},
}

METAS_FIXAS = {
    "ADE": {1: parse_number_any("190.081,41"), 2: parse_number_any("161.771,41"), 3: parse_number_any("194.125,70"),
            4: parse_number_any("177.948,56"), 5: parse_number_any("181.992,84"), 6: parse_number_any("186.037,13"),
            7: parse_number_any("226.479,98"), 8: parse_number_any("190.081,41"), 9: parse_number_any("186.037,13"),
            10: parse_number_any("190.081,41"), 11: parse_number_any("195.338,98"), 12: parse_number_any("190.081,41")},
    "GAMA": {1: parse_number_any("186.541,42"), 2: parse_number_any("158.758,66"), 3: parse_number_any("190.510,39"),
             4: parse_number_any("174.634,52"), 5: parse_number_any("178.603,49"), 6: parse_number_any("182.572,45"),
             7: parse_number_any("222.262,12"), 8: parse_number_any("186.541,42"), 9: parse_number_any("182.572,45"),
             10: parse_number_any("186.541,42"), 11: parse_number_any("191.701,08"), 12: parse_number_any("186.541,42")},
    "LUZIANIA": {1: parse_number_any("304.446,33"), 2: parse_number_any("259.103,26"), 3: parse_number_any("310.923,91"),
                 4: parse_number_any("285.013,58"), 5: parse_number_any("291.491,16"), 6: parse_number_any("297.968,74"),
                 7: parse_number_any("362.744,56"), 8: parse_number_any("304.446,33"), 9: parse_number_any("297.968,74"),
                 10: parse_number_any("304.446,33"), 11: parse_number_any("312.867,18"), 12: parse_number_any("304.446,33")},
    "SOFNORTE": {1: parse_number_any("217.392,69"), 2: parse_number_any("185.015,06"), 3: parse_number_any("222.018,07"),
                 4: parse_number_any("203.516,56"), 5: parse_number_any("208.141,94"), 6: parse_number_any("212.767,32"),
                 7: parse_number_any("259.021,08"), 8: parse_number_any("217.392,69"), 9: parse_number_any("212.767,32"),
                 10: parse_number_any("217.392,69"), 11: parse_number_any("223.405,68"), 12: parse_number_any("217.392,69")},
    "CEILANDIA": {1: parse_number_any("239.467,06"), 2: parse_number_any("203.801,75"), 3: parse_number_any("244.562,10"),
                  4: parse_number_any("224.181,93"), 5: parse_number_any("229.276,97"), 6: parse_number_any("234.372,01"),
                  7: parse_number_any("285.322,45"), 8: parse_number_any("239.467,06"), 9: parse_number_any("234.372,01"),
                  10: parse_number_any("239.467,06"), 11: parse_number_any("246.090,61"), 12: parse_number_any("239.467,06")},
    "SIA": {1: parse_number_any("285.285,01"), 2: parse_number_any("242.795,75"), 3: parse_number_any("291.354,90"),
            4: parse_number_any("267.075,33"), 5: parse_number_any("273.145,22"), 6: parse_number_any("279.215,11"),
            7: parse_number_any("339.914,05"), 8: parse_number_any("285.285,01"), 9: parse_number_any("279.215,11"),
            10: parse_number_any("285.285,01"), 11: parse_number_any("293.175,87"), 12: parse_number_any("285.285,01")},
    "UNAI": {1: parse_number_any("121.856,43"), 2: parse_number_any("103.707,60"), 3: parse_number_any("124.449,12"),
             4: parse_number_any("114.078,36"), 5: parse_number_any("116.671,05"), 6: parse_number_any("119.263,74"),
             7: parse_number_any("145.190,64"), 8: parse_number_any("121.856,43"), 9: parse_number_any("119.263,74"),
             10: parse_number_any("121.856,43"), 11: parse_number_any("125.226,93"), 12: parse_number_any("121.856,43")},
    "AGLINDAS": {1: parse_number_any("118.232,45"), 2: parse_number_any("100.623,36"), 3: parse_number_any("120.748,04"),
                 4: parse_number_any("110.685,70"), 5: parse_number_any("113.201,28"), 6: parse_number_any("115.716,87"),
                 7: parse_number_any("140.872,71"), 8: parse_number_any("118.232,45"), 9: parse_number_any("115.716,87"),
                 10: parse_number_any("118.232,45"), 11: parse_number_any("121.502,71"), 12: parse_number_any("118.232,45")},
    "GUARA": {1: parse_number_any("226.846,86"), 2: parse_number_any("193.061,16"), 3: parse_number_any("231.673,39"),
              4: parse_number_any("212.367,28"), 5: parse_number_any("217.193,81"), 6: parse_number_any("222.020,33"),
              7: parse_number_any("270.285,62"), 8: parse_number_any("226.846,86"), 9: parse_number_any("222.020,33"),
              10: parse_number_any("226.846,86"), 11: parse_number_any("233.121,35"), 12: parse_number_any("226.846,86")},
}


def montar_df_2025_fixo(lojas_keys):
    rows = []
    for loja_key in lojas_keys:
        mapa_mes = VENDAS_2025_FIXAS.get(loja_key, {})
        for mes_num, mes_nome in MESES:
            val = float(mapa_mes.get(mes_num, 0.0) or 0.0)
            rows.append({"LOJA_KEY": loja_key, "MES_NUM": mes_num, "MES": mes_nome, "VENDAS_2025": val})
    return pd.DataFrame(rows)


def montar_df_metas_fixas(lojas_keys):
    rows = []
    for loja_key in lojas_keys:
        mapa_mes = METAS_FIXAS.get(loja_key, {})
        for mes_num, mes_nome in MESES:
            val = float(mapa_mes.get(mes_num, 0.0) or 0.0)
            rows.append({"LOJA_KEY": loja_key, "MES_NUM": mes_num, "MES": mes_nome, "META": val})
    return pd.DataFrame(rows)




# ============================================================
# Relatórios WhatsApp (Página Relatórios)
# ============================================================
TELEFONES_WHATSAPP_LOJAS = {
    "ADE": "556191604836",
    "GAMA": "556193767169",
    "LUZIANIA": "556195356524",
    "SOFNORTE": "556192568295",
    "CEILANDIA": "556199783739",
    "SIA": "556199783735",
    "UNAI": "553899292791",
    "AGLINDAS": "556191159390",
    "GUARA": "556182729009",
}
TELEFONE_WHATSAPP_GERAL_MENSAL = "556199783726"
TELEFONE_WHATSAPP_GERAL_DIARIO = "5561993215052"


def _money(v) -> str:
    return "R$ " + format_brl(float(v or 0))


def _pct(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    return f"{float(v):.2f}%".replace(".", ",")


def _status_meta_texto(realizado: float, meta: float, contexto: str = "loja") -> tuple[bool, str]:
    """Retorna sinalização textual quando a meta já foi batida pelo realizado."""
    realizado = float(realizado or 0.0)
    meta = float(meta or 0.0)
    if meta <= 0:
        return False, ""

    ating_real = (realizado / meta * 100) if meta else 0.0
    saldo = realizado - meta

    if saldo >= 0:
        if contexto == "geral":
            return True, (
                f"🏆 *META GERAL JÁ BATIDA:* o realizado está em *{_pct(ating_real)}* da meta, "
                f"com *{_money(saldo)}* acima do objetivo do mês."
            )
        return True, (
            f"🏆 *META JÁ BATIDA:* a loja já realizou *{_pct(ating_real)}* da meta, "
            f"ficando *{_money(saldo)}* acima do objetivo do mês."
        )

    return False, (
        f"🎯 Meta ainda em andamento: realizado em *{_pct(ating_real)}* da meta, "
        f"faltando *{_money(abs(saldo))}* para bater o objetivo."
    )


def _loja_ord_key(loja_key: str) -> tuple[int, str]:
    return (LOJA_KEY_RANK.get(str(loja_key), DEFAULT_RANK), str(loja_key))


def _periodo_texto(data_ini_rel, data_fim_rel) -> str:
    if data_ini_rel == data_fim_rel:
        return data_ini_rel.strftime("%d/%m/%Y")
    return f"{data_ini_rel.strftime('%d/%m/%Y')} a {data_fim_rel.strftime('%d/%m/%Y')}"


def _grupo_linha(nome) -> str:
    k = canonical_key(nome)
    if "THINNER" in k or "TINNER" in k:
        return "Thinner"
    if "AUTOMOT" in k:
        return "Automotivo"
    return "Decor/Industrial"


def _rank_text(df_base: pd.DataFrame, col: str, titulo: str, valor_col: str = "FAT_LINHA", top: int = 10) -> str:
    if df_base.empty or col not in df_base.columns:
        return f"*{titulo}*\nSem dados."
    total = float(df_base[valor_col].sum()) if valor_col in df_base.columns else 0.0
    tab = (
        df_base.groupby(col, dropna=False)[valor_col]
        .sum()
        .sort_values(ascending=False)
        .head(top)
    )
    linhas = [f"*{titulo}*"]
    for i, (nome, valor) in enumerate(tab.items(), start=1):
        perc = (float(valor) / total * 100) if total else 0.0
        linhas.append(f"{i}. {nome}: {_money(valor)} ({_pct(perc)})")
    return "\n".join(linhas)


def _linhas_texto(df_base: pd.DataFrame) -> str:
    if df_base.empty:
        return "*Segmentos*\nSem dados."
    aux = df_base.copy()
    aux["GRUPO_REL"] = aux["SEGMENTO_N"].apply(_grupo_linha)
    total = float(aux["FAT_LINHA"].sum())
    tab = aux.groupby("GRUPO_REL", dropna=False)["FAT_LINHA"].sum().reindex(["Automotivo", "Thinner", "Decor/Industrial"]).fillna(0.0)
    linhas = ["*Segmentos*" if total else "*Segmentos*"]
    for nome, valor in tab.items():
        perc = (float(valor) / total * 100) if total else 0.0
        linhas.append(f"- {nome}: {_money(valor)} ({_pct(perc)})")
    return "\n".join(linhas)


def _linhas_dataframe(df_base: pd.DataFrame) -> pd.DataFrame:
    aux = df_base.copy()
    if aux.empty:
        return pd.DataFrame(columns=["Linha", "Faturamento", "Participação"])
    aux["GRUPO_REL"] = aux["SEGMENTO_N"].apply(_grupo_linha)
    total = float(aux["FAT_LINHA"].sum())
    tab = aux.groupby("GRUPO_REL", dropna=False)["FAT_LINHA"].sum().reindex(["Automotivo", "Thinner", "Decor/Industrial"]).fillna(0.0).reset_index()
    tab.columns = ["Linha", "Faturamento"]
    tab["Participação"] = tab["Faturamento"].apply(lambda v: _pct((float(v) / total * 100) if total else 0.0))
    tab["Faturamento"] = tab["Faturamento"].apply(_money)
    return tab


def _peso_dia_comercial(data_val) -> float:
    """Peso comercial do dia.

    Regra:
    - Segunda a sexta = 1,0
    - Sábado = 0,5
    - Domingo = 0,0
    """
    try:
        dt = pd.to_datetime(data_val)
        if pd.isna(dt):
            return 0.0
        dow = int(dt.dayofweek)
        if dow <= 4:
            return 1.0
        if dow == 5:
            return 0.5
        return 0.0
    except Exception:
        return 0.0


def _coluna_valor_vendas(df_base: pd.DataFrame) -> str | None:
    """Escolhe a melhor coluna de valor para apurar movimento diário."""
    for col in ["FAT_LINHA", "VENDAS_2026", "VR_TOTAL_NUM", "VR_TOTAL"]:
        if col in df_base.columns:
            return col
    return None


def _dias_equivalentes_vendidos(df_base: pd.DataFrame) -> float:
    """Conta somente dias com venda real positiva.

    Uso principal: leituras históricas simples. Para previsão mensal, use
    _calcular_calendario_comercial_mes, que diferencia dia realizado, dia
    restante e feriado/dia sem expediente.
    """
    if df_base.empty or "DATA" not in df_base.columns:
        return 0.0

    valor_col = _coluna_valor_vendas(df_base)
    if valor_col is None:
        return 0.0

    dias = df_base[df_base["DATA"].notna()].copy()
    if dias.empty:
        return 0.0

    dias[valor_col] = pd.to_numeric(dias[valor_col], errors="coerce").fillna(0.0)
    diario = (
        dias.assign(DATA_DIA=dias["DATA"].dt.normalize())
        .groupby("DATA_DIA", as_index=False)[valor_col]
        .sum()
        .rename(columns={valor_col: "FAT_DIA"})
    )
    diario = diario[diario["FAT_DIA"] > 0].copy()
    if diario.empty:
        return 0.0

    diario["PESO_DIA"] = diario["DATA_DIA"].apply(_peso_dia_comercial)
    diario = diario[diario["PESO_DIA"] > 0].copy()
    return float(diario["PESO_DIA"].sum())


def _calcular_calendario_comercial_mes(df_base: pd.DataFrame, ano: int, mes: int, data_fim_rel=None) -> dict:
    """Calcula o calendário comercial válido para previsão.

    Lógica aplicada:
    - Segunda a sexta = 1 dia.
    - Sábado = 0,5 dia.
    - Domingo = 0.
    - Dia comercial já ocorrido com faturamento total da empresa igual a zero é
      removido do total do mês, pois normalmente representa feriado ou fechamento
      geral da operação.
    - Dias futuros continuam no total do mês para compor os dias restantes.
    """
    import calendar

    ano = int(ano)
    mes = int(mes)
    primeiro_dia = pd.Timestamp(date(ano, mes, 1))
    ultimo_dia = pd.Timestamp(date(ano, mes, calendar.monthrange(ano, mes)[1]))

    if data_fim_rel is None:
        data_corte = ultimo_dia
    else:
        data_corte = pd.to_datetime(data_fim_rel)
        if pd.isna(data_corte):
            data_corte = ultimo_dia
        data_corte = min(data_corte.normalize(), ultimo_dia)

    calendario = pd.DataFrame({"DATA_DIA": pd.date_range(primeiro_dia, ultimo_dia, freq="D")})
    calendario["PESO_DIA"] = calendario["DATA_DIA"].apply(_peso_dia_comercial)
    calendario = calendario[calendario["PESO_DIA"] > 0].copy()

    if calendario.empty:
        return {
            "dias_comerciais_base": 0.0,
            "dias_validos_mes": 0.0,
            "dias_realizados": 0.0,
            "dias_restantes": 0.0,
            "dias_ignorados": [],
            "datas_ignoradas_txt": "",
        }

    valor_col = _coluna_valor_vendas(df_base) if df_base is not None and not df_base.empty else None
    if df_base is not None and not df_base.empty and "DATA" in df_base.columns and valor_col:
        mov = df_base[df_base["DATA"].notna()].copy()
        mov = mov[(mov["DATA"].dt.year == ano) & (mov["DATA"].dt.month == mes)].copy()
        mov[valor_col] = pd.to_numeric(mov[valor_col], errors="coerce").fillna(0.0)
        diario = (
            mov.assign(DATA_DIA=mov["DATA"].dt.normalize())
            .groupby("DATA_DIA", as_index=False)[valor_col]
            .sum()
            .rename(columns={valor_col: "FAT_DIA_EMPRESA"})
        )
    else:
        diario = pd.DataFrame(columns=["DATA_DIA", "FAT_DIA_EMPRESA"])

    cal = calendario.merge(diario, on="DATA_DIA", how="left")
    cal["FAT_DIA_EMPRESA"] = pd.to_numeric(cal["FAT_DIA_EMPRESA"], errors="coerce").fillna(0.0)

    cal_ate_corte = cal[cal["DATA_DIA"] <= data_corte].copy()
    ignorados = cal_ate_corte[cal_ate_corte["FAT_DIA_EMPRESA"] <= 0].copy()
    datas_ignoradas = [d.strftime("%d/%m") for d in ignorados["DATA_DIA"].dt.to_pydatetime()]

    dias_comerciais_base = float(cal["PESO_DIA"].sum())
    peso_ignorado = float(ignorados["PESO_DIA"].sum()) if not ignorados.empty else 0.0
    dias_validos_mes = max(dias_comerciais_base - peso_ignorado, 0.0)

    realizados = cal_ate_corte[cal_ate_corte["FAT_DIA_EMPRESA"] > 0].copy()
    dias_realizados = float(realizados["PESO_DIA"].sum()) if not realizados.empty else 0.0
    dias_restantes = max(dias_validos_mes - dias_realizados, 0.0)

    return {
        "dias_comerciais_base": dias_comerciais_base,
        "dias_validos_mes": dias_validos_mes,
        "dias_realizados": dias_realizados,
        "dias_restantes": dias_restantes,
        "dias_ignorados": datas_ignoradas,
        "datas_ignoradas_txt": ", ".join(datas_ignoradas),
    }


def _previsao_por_calendario(realizado: float, dias_realizados: float, dias_restantes: float) -> tuple[float, float]:
    """Retorna (previsão, média diária) usando dias realizados e dias restantes."""
    realizado = float(realizado or 0.0)
    dias_realizados = float(dias_realizados or 0.0)
    dias_restantes = float(dias_restantes or 0.0)
    if dias_realizados <= 0:
        return realizado, 0.0
    media_dia = realizado / dias_realizados
    previsao = realizado + (media_dia * max(dias_restantes, 0.0))
    return float(previsao), float(media_dia)


def _referencias_mes(lojas_keys: list[str], mes: int):
    df_meta_ref, df_ano1_ref, dias_uteis_map = carregar_referencias_planejamento()
    if df_meta_ref is not None and not df_meta_ref.empty:
        metas = df_meta_ref[(df_meta_ref["MES_NUM"] == mes) & (df_meta_ref["LOJA_KEY"].isin(lojas_keys))].copy()
    else:
        metas = montar_df_metas_fixas(lojas_keys)
        metas = metas[metas["MES_NUM"] == mes].copy()
    if df_ano1_ref is not None and not df_ano1_ref.empty:
        ano1 = df_ano1_ref[(df_ano1_ref["MES_NUM"] == mes) & (df_ano1_ref["LOJA_KEY"].isin(lojas_keys))].copy()
    else:
        ano1 = montar_df_2025_fixo(lojas_keys)
        ano1 = ano1[ano1["MES_NUM"] == mes].copy()
    dias_uteis = float(dias_uteis_map.get(mes, 0.0) or 0.0)
    if dias_uteis <= 0:
        # fallback simples: considera dias de segunda a sexta = 1 e sábado = 0,5 dentro do mês
        import calendar
        ano = ANO_ATUAL
        ultimo = calendar.monthrange(ano, mes)[1]
        dias_uteis = sum(0.5 if pd.Timestamp(ano, mes, d).dayofweek == 5 else 1.0 for d in range(1, ultimo + 1) if pd.Timestamp(ano, mes, d).dayofweek <= 5)
    return metas, ano1, dias_uteis



def _mes_esta_fechado(df_base: pd.DataFrame, ano: int, mes: int, data_fim_rel) -> bool:
    """Detecta se o mês deve ser tratado como fechado no relatório.

    Critérios:
    - data final selecionada chegou ao último dia do mês; ou
    - já existem registros de meses posteriores na base, indicando que aquele mês ficou para trás.
    """
    try:
        import calendar
        data_fim_rel = pd.to_datetime(data_fim_rel).date()
        ultimo_dia_mes = date(int(ano), int(mes), calendar.monthrange(int(ano), int(mes))[1])
        if data_fim_rel >= ultimo_dia_mes:
            return True

        if df_base is not None and not df_base.empty and "DATA" in df_base.columns:
            datas = pd.to_datetime(df_base["DATA"], errors="coerce").dropna()
            if not datas.empty:
                max_data_base = datas.max().date()
                return max_data_base > ultimo_dia_mes
    except Exception:
        return False
    return False

def _criar_link_whatsapp(numero: str, texto: str | None = None) -> str:
    """Monta o link do WhatsApp.

    Quando o texto é enviado via parâmetro ?text=, navegadores podem cortar mensagens
    grandes por limite de URL. Para relatórios longos, a aplicação agora usa o link
    apenas com o número e copia o texto completo para a área de transferência.
    """
    numero = re.sub(r"\D+", "", str(numero or ""))
    if texto:
        return f"https://wa.me/{numero}?text={quote(texto)}"
    return f"https://wa.me/{numero}"


def _botao_copiar_e_abrir_whatsapp(label: str, numero: str, texto: str, key: str):
    """Botão visual: copia o relatório completo e abre o WhatsApp no contato.

    Observação técnica: por segurança, navegadores não permitem que uma página cole
    automaticamente dentro do campo de mensagem do WhatsApp Web. O app deixa o texto
    completo na área de transferência e abre o contato; se o navegador bloquear a cópia,
    o usuário ainda consegue copiar pelo text_area exibido na tela.
    """
    numero_limpo = re.sub(r"\D+", "", str(numero or ""))
    url = _criar_link_whatsapp(numero_limpo)
    texto_js = json.dumps(str(texto or ""), ensure_ascii=False)
    url_js = json.dumps(url, ensure_ascii=False)
    label_js = json.dumps(label, ensure_ascii=False)
    safe_id = re.sub(r"[^a-zA-Z0-9_-]", "_", str(key or "whatsapp"))

    st.components.v1.html(
        f"""
        <div style="font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;">
            <button id="btn_{safe_id}" style="
                width: 100%;
                border: 0;
                border-radius: 14px;
                padding: 13px 16px;
                background: linear-gradient(135deg, #16a34a 0%, #15803d 100%);
                color: #ffffff;
                font-weight: 800;
                font-size: 15px;
                cursor: pointer;
                box-shadow: 0 8px 20px rgba(22, 163, 74, 0.22);
            ">📲 {label_js.strip('"')}</button>
            <div id="status_{safe_id}" style="margin-top: 8px; color: #334155; font-size: 12px; line-height: 1.35;"></div>
        </div>
        <script>
        const texto_{safe_id} = {texto_js};
        const url_{safe_id} = {url_js};
        const status_{safe_id} = document.getElementById('status_{safe_id}');

        async function copiarTexto_{safe_id}(txt) {{
            if (navigator.clipboard && window.isSecureContext) {{
                await navigator.clipboard.writeText(txt);
                return true;
            }}
            const ta = document.createElement('textarea');
            ta.value = txt;
            ta.style.position = 'fixed';
            ta.style.left = '-9999px';
            ta.style.top = '0';
            document.body.appendChild(ta);
            ta.focus();
            ta.select();
            const ok = document.execCommand('copy');
            document.body.removeChild(ta);
            return ok;
        }}

        document.getElementById('btn_{safe_id}').addEventListener('click', async function() {{
            try {{
                await copiarTexto_{safe_id}(texto_{safe_id});
                status_{safe_id}.innerHTML = '✅ Texto completo copiado. O WhatsApp será aberto no contato; cole a mensagem no campo de envio se ela não aparecer automaticamente.';
            }} catch (e) {{
                status_{safe_id}.innerHTML = '⚠️ O navegador bloqueou a cópia automática. Use o campo de texto abaixo para copiar manualmente.';
            }}
            window.open(url_{safe_id}, '_blank');
        }});
        </script>
        """,
        height=86,
    )


def _botao_whatsapp(label: str, numero: str, texto: str, key: str | None = None):
    key = key or f"wa_{abs(hash((label, numero, len(str(texto or ''))))) }"
    _botao_copiar_e_abrir_whatsapp(label, numero, texto, key)


def _render_whatsapp_downloads(prefixo_key: str, numero: str, texto: str, nome_arquivo: str, label_whatsapp: str, label_download: str):
    st.markdown(
        f"""
        <div class="whatsapp-box">
            <b>Envio WhatsApp</b><br>
            Relatório único com <b>{len(str(texto or '')):,}</b> caracteres.<br>
            Ao clicar, o texto completo é copiado para a área de transferência e o WhatsApp abre direto no contato.
        </div>
        """.replace(",", "."),
        unsafe_allow_html=True,
    )
    c1, c2 = st.columns([1, 1])
    with c1:
        _botao_whatsapp(label_whatsapp, numero, texto, key=f"{prefixo_key}_wa_full")
    with c2:
        st.download_button(label_download, texto.encode("utf-8"), file_name=nome_arquivo, mime="text/plain", key=f"down_full_{prefixo_key}")

    st.caption("Se o navegador bloquear a cópia automática, copie pelo campo abaixo e cole no WhatsApp aberto.")

def _montar_relatorio_diario_loja(df_loja: pd.DataFrame, loja_nome: str, data_rel, valor_col: str) -> str:
    total = float(df_loja["FAT_LINHA"].sum()) if len(df_loja) else 0.0
    partes = [
        f"*{loja_nome} | Relatório diário*",
        f"Data: {data_rel.strftime('%d/%m/%Y')}",
        "",
        f"A loja vendeu *{_money(total)}* hoje.",
        "",
        _linhas_texto(df_loja),
        "",
        _rank_text(df_loja, "MARCA_N", "Top 10 marcas do dia", "FAT_LINHA"),
        "",
        _rank_text(df_loja, "CLIENTE_N", "Top 10 clientes do dia", valor_col),
    ]
    return "\n".join(partes)


def _montar_relatorio_mensal_loja(df_loja: pd.DataFrame, loja_nome: str, data_ini_rel, data_fim_rel, mes: int, meta: float, ano1: float, dias_uteis_mes: float, valor_col: str, mes_fechado: bool = False, info_dias: dict | None = None) -> str:
    realizado = float(df_loja["FAT_LINHA"].sum()) if len(df_loja) else 0.0
    if info_dias is not None:
        dias_venda = float(info_dias.get("dias_realizados", 0.0) or 0.0)
        dias_restantes = float(info_dias.get("dias_restantes", 0.0) or 0.0)
        dias_uteis_mes = float(info_dias.get("dias_validos_mes", dias_uteis_mes) or dias_uteis_mes or 0.0)
    else:
        dias_venda = _dias_equivalentes_vendidos(df_loja)
        dias_restantes = max(float(dias_uteis_mes or 0.0) - float(dias_venda or 0.0), 0.0)

    previsao, media_dia = _previsao_por_calendario(realizado, dias_venda, dias_restantes)
    resultado_ref = realizado if mes_fechado else previsao
    dif_ano1 = resultado_ref - float(ano1 or 0)
    cresc = (dif_ano1 / float(ano1) * 100) if ano1 else None
    dif_meta = resultado_ref - float(meta or 0)
    ating = (resultado_ref / float(meta) * 100) if meta else None
    meta_batida, texto_meta_realizada = _status_meta_texto(realizado, meta, contexto="loja")
    emoji_ano1 = "🟢" if dif_ano1 >= 0 else "🔴"
    emoji_meta = "🏆" if meta_batida else ("🟢" if dif_meta >= 0 else "🔴")

    if mes_fechado:
        titulo_rel = f"*{loja_nome} | Resultado final de {MESES[mes-1][1].title()}*"
        linha_resultado = f"A loja encerrou o mês com faturamento realizado de *{_money(realizado)}*."
        texto_ano1 = f"{emoji_ano1} Contra o Ano-1, o resultado ficou *{_money(abs(dif_ano1))} {'acima' if dif_ano1 >= 0 else 'abaixo'}*, variação de *{_pct(cresc)}*."
        texto_meta = f"{emoji_meta} Contra a meta, o resultado ficou *{_money(abs(dif_meta))} {'acima' if dif_meta >= 0 else 'abaixo'}*, com atingimento realizado de *{_pct(ating)}*."
        resumo_meta = (
            "mês fechado com meta batida. O resultado deve ser reconhecido, preservando aprendizado das linhas e clientes que mais puxaram a performance."
            if meta_batida
            else "mês fechado abaixo da meta. Vale revisar plano de ação, carteira de clientes, mix e ritmo comercial para o próximo ciclo."
        )
    else:
        titulo_rel = f"*{loja_nome} | Análise parcial de {MESES[mes-1][1].title()}*"
        linha_resultado = f"A loja está com *{_money(realizado)}* faturados até agora. Mantido o ritmo atual, a previsão de fechamento é de *{_money(previsao)}*."
        texto_ano1 = f"{emoji_ano1} Contra o Ano-1, a previsão fica *{_money(abs(dif_ano1))} {'acima' if dif_ano1 >= 0 else 'abaixo'}*, variação de *{_pct(cresc)}*."
        texto_meta = f"{emoji_meta} Contra a meta, a previsão fica *{_money(abs(dif_meta))} {'acima' if dif_meta >= 0 else 'abaixo'}*, com atingimento previsto de *{_pct(ating)}*."
        resumo_meta = (
            "Meta já batida. O foco agora é manter ritmo, defender margem e aproveitar oportunidades de venda adicional."
            if meta_batida
            else ("a loja está projetando fechamento acima da meta. O foco agora é sustentar o ritmo e proteger as linhas de maior peso." if dif_meta >= 0 else "a loja precisa acelerar para reduzir a distância da meta até o fechamento do mês.")
        )

    texto_dias = ""
    if not mes_fechado and info_dias is not None:
        ignorados_txt = str(info_dias.get("datas_ignoradas_txt", "") or "").strip()
        texto_dias = (
            f"Dias de venda considerados: *{format_decimal_pt(dias_venda, 1)}* realizados, "
            f"*{format_decimal_pt(dias_restantes, 1)}* restantes e "
            f"*{format_decimal_pt(dias_uteis_mes, 1)}* válidos no mês."
        )
        if ignorados_txt:
            texto_dias += f" Dias ignorados por venda zero na empresa: {ignorados_txt}."

    partes = [
        titulo_rel,
        f"Período: {_periodo_texto(data_ini_rel, data_fim_rel)}.",
        "",
        linha_resultado,
        texto_dias,
        texto_meta_realizada,
        "",
        texto_ano1,
        texto_meta,
        "",
        _linhas_texto(df_loja).replace("*Segmentos*", "*Linhas*"),
        "",
        _rank_text(df_loja, "MARCA_N", "Top 10 marcas da loja", "FAT_LINHA"),
        "",
        _rank_text(df_loja, "CLIENTE_N", "Top 10 clientes da loja", valor_col),
        "",
        "Resumo: " + resumo_meta,
    ]
    return "\n".join(partes)


def render_relatorios_whatsapp(df_base: pd.DataFrame, valor_col: str):
    st.markdown("## 📲 Relatórios para WhatsApp")
    st.markdown(
        """
        <div class="section-card">
            Gere relatórios diários ou parciais do mês, visualize o resumo na tela e envie pelo WhatsApp.
            Textos longos agora são mantidos em um único relatório: o app copia tudo para a área de transferência e abre o WhatsApp no contato selecionado.
        </div>
        """,
        unsafe_allow_html=True,
    )

    datas_validas_rel = df_base["DATA"].dropna()
    if datas_validas_rel.empty:
        st.warning("A base não possui datas válidas para gerar relatórios.")
        return

    lojas_map = df_base[["LOJA_KEY", "LOJA_N"]].drop_duplicates().groupby("LOJA_KEY")["LOJA_N"].first().to_dict()
    lojas_keys = sorted([k for k in lojas_map.keys() if k], key=_loja_ord_key)

    modo = st.radio("Tipo de relatório", ["Diário", "Parcial do mês / WhatsApp por loja"], horizontal=True)

    if modo == "Diário":
        data_padrao = datas_validas_rel.max().date()
        data_rel = st.date_input("Data do relatório diário", value=data_padrao, min_value=datas_validas_rel.min().date(), max_value=datas_validas_rel.max().date())
        df_rel = df_base[df_base["DATA"].notna() & (df_base["DATA"].dt.date == data_rel)].copy()
        total = float(df_rel["FAT_LINHA"].sum()) if len(df_rel) else 0.0
        st.metric("Faturamento do dia", _money(total))
        st.dataframe(_linhas_dataframe(df_rel), use_container_width=True, hide_index=True)

        blocos = []
        venda_loja = df_rel.groupby("LOJA_KEY")["FAT_LINHA"].sum().reindex(lojas_keys).fillna(0.0)
        texto_geral = [
            f"📊 *Relatório diário - {data_rel.strftime('%d/%m/%Y')}*",
            "",
            f"Hoje o faturamento total foi de *{_money(total)}*.",
            "",
            "*Venda por loja*",
        ]
        for k in lojas_keys:
            texto_geral.append(f"- *{lojas_map.get(k, k)}*: {_money(venda_loja.get(k, 0.0))}")
        texto_geral += ["", _linhas_texto(df_rel), "", _rank_text(df_rel, "MARCA_N", "Top 10 marcas do dia", "FAT_LINHA"), "", "Os blocos por loja trazem o detalhe de segmentos, marcas e clientes do dia.", "", "==============================", "RELATÓRIOS POR LOJA", "=============================="]
        for k in lojas_keys:
            df_loja = df_rel[df_rel["LOJA_KEY"] == k].copy()
            loja_nome = str(lojas_map.get(k, k))
            bloco = _montar_relatorio_diario_loja(df_loja, loja_nome, data_rel, valor_col)
            blocos.append((k, loja_nome, bloco))
            texto_geral += ["", "------------------------------", bloco]
        texto_full = "\n".join(texto_geral)

        _render_whatsapp_downloads(
            f"diario_{data_rel.strftime('%Y%m%d')}",
            TELEFONE_WHATSAPP_GERAL_DIARIO,
            texto_full,
            f"relatorio_diario_{data_rel.strftime('%Y%m%d')}.txt",
            "Abrir WhatsApp - Geral",
            "Baixar texto diário (.txt)",
        )
        st.text_area("Texto geral diário", texto_full, height=360)

        st.subheader("Envio por loja")
        for k, loja_nome, bloco in blocos:
            with st.expander(loja_nome, expanded=False):
                cols = st.columns([1, 1])
                numero = TELEFONES_WHATSAPP_LOJAS.get(k)
                if numero:
                    with cols[0]:
                        _botao_whatsapp(f"Abrir WhatsApp - {loja_nome}", numero, bloco)
                with cols[1]:
                    st.download_button(f"Baixar {loja_nome}.txt", bloco.encode("utf-8"), file_name=f"relatorio_diario_{k}_{data_rel.strftime('%Y%m%d')}.txt", mime="text/plain", key=f"down_daily_{k}")
                st.text_area(f"Texto {loja_nome}", bloco, height=260, key=f"txt_daily_{k}")
        return

    # Relatório parcial/mensal
    data_max_rel = datas_validas_rel.max().date()
    mes_padrao = data_max_rel.month
    mes = st.selectbox("Mês do relatório", options=[m for m, _ in MESES], index=mes_padrao - 1, format_func=lambda m: f"{m:02d} - {MESES[m-1][1]}")
    data_ini_rel = date(ANO_ATUAL, mes, 1)
    data_fim_rel = st.date_input("Data final do relatório parcial", value=data_max_rel, min_value=data_ini_rel, max_value=data_max_rel)
    df_rel = df_base[df_base["DATA"].notna()].copy()
    df_rel = df_rel[(df_rel["DATA"].dt.year == ANO_ATUAL) & (df_rel["DATA"].dt.month == mes) & (df_rel["DATA"].dt.date >= data_ini_rel) & (df_rel["DATA"].dt.date <= data_fim_rel)].copy()
    mes_fechado = _mes_esta_fechado(df_base, ANO_ATUAL, mes, data_fim_rel)

    metas, ano1, dias_uteis_mes_ref = _referencias_mes(lojas_keys, mes)
    info_dias_mes = _calcular_calendario_comercial_mes(df_base, ANO_ATUAL, mes, data_fim_rel)
    dias_uteis_mes = float(info_dias_mes.get("dias_validos_mes", dias_uteis_mes_ref) or dias_uteis_mes_ref or 0.0)
    meta_map = metas.groupby("LOJA_KEY")["META"].sum().to_dict() if not metas.empty else {}
    ano1_map = ano1.groupby("LOJA_KEY")["VENDAS_2025"].sum().to_dict() if not ano1.empty else {}

    realizado_total = float(df_rel["FAT_LINHA"].sum()) if len(df_rel) else 0.0
    dias_venda_total = float(info_dias_mes.get("dias_realizados", 0.0) or 0.0)
    dias_restantes_total = float(info_dias_mes.get("dias_restantes", 0.0) or 0.0)
    previsao_total, media_dia_total = _previsao_por_calendario(realizado_total, dias_venda_total, dias_restantes_total)
    meta_total = float(sum(meta_map.values())) if meta_map else 0.0
    ano1_total = float(sum(ano1_map.values())) if ano1_map else 0.0
    resultado_total_ref = realizado_total if mes_fechado else previsao_total
    dif_ano1_total = resultado_total_ref - ano1_total
    dif_meta_total = resultado_total_ref - meta_total
    ating_total = (resultado_total_ref / meta_total * 100) if meta_total else None
    cresc_total = (dif_ano1_total / ano1_total * 100) if ano1_total else None
    meta_geral_batida, texto_meta_geral_realizada = _status_meta_texto(realizado_total, meta_total, contexto="geral")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Faturamento realizado" if mes_fechado else "Faturamento atual", _money(realizado_total))
    m2.metric("Mês fechado" if mes_fechado else "Previsão fechamento", _money(resultado_total_ref))
    m3.metric(("Real x Ano-1" if mes_fechado else "Prev. x Ano-1"), _money(dif_ano1_total), _pct(cresc_total))
    m4.metric(("Real x Meta" if mes_fechado else "Prev. x Meta"), _money(dif_meta_total), _pct(ating_total))
    if not mes_fechado:
        st.caption(
            f"Dias considerados na previsão: {format_decimal_pt(dias_venda_total, 1)} realizados, "
            f"{format_decimal_pt(dias_restantes_total, 1)} restantes, "
            f"{format_decimal_pt(dias_uteis_mes, 1)} válidos no mês."
            + (f" Dias ignorados por venda zero na empresa: {info_dias_mes.get('datas_ignoradas_txt')}." if info_dias_mes.get("datas_ignoradas_txt") else "")
        )
    st.dataframe(_linhas_dataframe(df_rel), use_container_width=True, hide_index=True)

    blocos = []
    if mes_fechado:
        texto_geral = [
            f"📊 *Resultado final de {MESES[mes-1][1].title()}*",
            "",
            f"Período encerrado: *{_periodo_texto(data_ini_rel, data_fim_rel)}*",
            "",
            f"O mês foi fechado com faturamento realizado de *{_money(realizado_total)}*.",
            texto_meta_geral_realizada,
            "",
            f"{'🟢' if dif_ano1_total >= 0 else '🔴'} Na comparação com o Ano-1, o resultado ficou *{_money(abs(dif_ano1_total))} {'acima' if dif_ano1_total >= 0 else 'abaixo'}*, ou seja, *{_pct(cresc_total)}*.",
            f"{'🏆' if meta_geral_batida else ('🟢' if dif_meta_total >= 0 else '🔴')} Contra a meta do mês, o resultado ficou *{_money(abs(dif_meta_total))} {'acima' if dif_meta_total >= 0 else 'abaixo'}*, com atingimento realizado de *{_pct(ating_total)}*.",
            "",
            "*Visão por loja*",
        ]
    else:
        texto_geral = [
            f"📊 *Segue análise parcial de {MESES[mes-1][1].title()}*",
            "",
            f"Período analisado: *{_periodo_texto(data_ini_rel, data_fim_rel)}*",
            "",
            f"Até o momento, o faturamento total está em *{_money(realizado_total)}*.",
            f"Mantido o ritmo atual, a previsão de fechamento é de *{_money(previsao_total)}*.",
            f"Dias de venda considerados: *{format_decimal_pt(dias_venda_total, 1)}* realizados, *{format_decimal_pt(dias_restantes_total, 1)}* restantes e *{format_decimal_pt(dias_uteis_mes, 1)}* válidos no mês.",
            (f"Dias ignorados por venda zero na empresa: {info_dias_mes.get('datas_ignoradas_txt')}." if info_dias_mes.get("datas_ignoradas_txt") else ""),
            texto_meta_geral_realizada,
            "",
            f"{'🟢' if dif_ano1_total >= 0 else '🔴'} Na comparação com o Ano-1, a projeção indica *{_money(abs(dif_ano1_total))} {'acima' if dif_ano1_total >= 0 else 'abaixo'}*, ou seja, *{_pct(cresc_total)}*.",
            f"{'🏆' if meta_geral_batida else ('🟢' if dif_meta_total >= 0 else '🔴')} Contra a meta do mês, a projeção indica *{_money(abs(dif_meta_total))} {'acima' if dif_meta_total >= 0 else 'abaixo'}*, com atingimento previsto de *{_pct(ating_total)}*.",
            "",
            "*Visão por loja*",
        ]
    resumo_rows = []
    for k in lojas_keys:
        loja_nome = str(lojas_map.get(k, k))
        df_loja = df_rel[df_rel["LOJA_KEY"] == k].copy()
        meta = float(meta_map.get(k, 0.0) or 0.0)
        ano1_val = float(ano1_map.get(k, 0.0) or 0.0)
        realizado = float(df_loja["FAT_LINHA"].sum()) if len(df_loja) else 0.0
        dias_venda = dias_venda_total
        previsao, _media_loja = _previsao_por_calendario(realizado, dias_venda_total, dias_restantes_total)
        resultado_ref = realizado if mes_fechado else previsao
        ating = (resultado_ref / meta * 100) if meta else None
        dif_ano1 = resultado_ref - ano1_val
        cresc = (dif_ano1 / ano1_val * 100) if ano1_val else None
        meta_batida_loja, texto_meta_loja = _status_meta_texto(realizado, meta, contexto="loja")
        status = "🏆" if meta_batida_loja else ("🟢" if ating is not None and ating >= 100 else ("🟡" if ating is not None and ating >= 70 else "🔴"))
        if mes_fechado:
            if meta_batida_loja:
                texto_geral.append(f"{status} *{loja_nome}* fechou o mês com meta batida: *{_money(realizado)}* realizados, *{_pct((realizado / meta * 100) if meta else None)}* da meta e *{_money(realizado - meta)}* acima do objetivo.")
            else:
                texto_geral.append(f"{status} *{loja_nome}* fechou com *{_money(realizado)}* realizados, variação de *{_pct(cresc)}* contra Ano-1 e *{_pct(ating)}* da meta.")
        else:
            if meta_batida_loja:
                texto_geral.append(f"{status} *{loja_nome}* já bateu a meta: *{_money(realizado)}* realizados, *{_pct((realizado / meta * 100) if meta else None)}* da meta e *{_money(realizado - meta)}* acima do objetivo. Previsão: *{_money(previsao)}*.")
            else:
                texto_geral.append(f"{status} *{loja_nome}* está com *{_money(realizado)}* faturados. Previsão: *{_money(previsao)}*, variação de *{_pct(cresc)}* contra Ano-1 e *{_pct(ating)}* da meta.")
        resumo_rows.append({"Loja": loja_nome, "Realizado": _money(realizado), "Previsão/Resultado": _money(resultado_ref), "% Meta": _pct(ating), "Status Meta": "Meta batida" if meta_batida_loja else ("Fechado abaixo da meta" if mes_fechado else "Em andamento")})
        bloco = _montar_relatorio_mensal_loja(df_loja, loja_nome, data_ini_rel, data_fim_rel, mes, meta, ano1_val, dias_uteis_mes, valor_col, mes_fechado=mes_fechado, info_dias=info_dias_mes)
        blocos.append((k, loja_nome, bloco))

    titulo_marcas_geral = "Top 10 marcas do mês fechado" if mes_fechado else "Top 10 marcas até o momento"
    texto_geral += ["", _linhas_texto(df_rel).replace("*Segmentos*", "*Análise por linhas*"), "", _rank_text(df_rel, "MARCA_N", titulo_marcas_geral, "FAT_LINHA"), "", "==============================", "RELATÓRIOS POR LOJA", "=============================="]
    for _, _, bloco in blocos:
        texto_geral += ["", "------------------------------", bloco]
    texto_full = "\n".join(texto_geral)

    st.subheader("Visão por loja")
    st.dataframe(pd.DataFrame(resumo_rows), use_container_width=True, hide_index=True)
    _render_whatsapp_downloads(
        f"mensal_{ANO_ATUAL}_{mes:02d}",
        TELEFONE_WHATSAPP_GERAL_MENSAL,
        texto_full,
        f"relatorio_mensal_{ANO_ATUAL}_{mes:02d}.txt",
        "Abrir WhatsApp - Geral mensal",
        "Baixar texto mensal (.txt)",
    )
    st.text_area("Texto geral mensal", texto_full, height=360)

    st.subheader("Envio por loja")
    for k, loja_nome, bloco in blocos:
        with st.expander(loja_nome, expanded=False):
            cols = st.columns([1, 1])
            numero = TELEFONES_WHATSAPP_LOJAS.get(k)
            if numero:
                with cols[0]:
                    _botao_whatsapp(f"Abrir WhatsApp - {loja_nome}", numero, bloco)
            with cols[1]:
                st.download_button(f"Baixar {loja_nome}.txt", bloco.encode("utf-8"), file_name=f"relatorio_mensal_{k}_{ANO_ATUAL}_{mes:02d}.txt", mime="text/plain", key=f"down_month_{k}")
            st.text_area(f"Texto {loja_nome}", bloco, height=260, key=f"txt_month_{k}")


# =========================
# Carrega e prepara base
# =========================
df = carregar_dados()
df = df[df["FAT_LINHA"].notna()].copy()

# ========= Define a métrica de valor do cliente =========
use_vr_total = df["VR_TOTAL_NUM"].notna().any()
VAL_COL = "VR_TOTAL_NUM" if use_vr_total else "FAT_LINHA"

# =========================
# Navegação
# =========================
pagina_app = st.sidebar.radio("Página", ["Dashboard", "Relatórios"], index=0)
if pagina_app == "Relatórios":
    render_relatorios_whatsapp(df, VAL_COL)
    st.stop()

# =========================
# Sidebar filtros
# =========================
st.sidebar.header("Filtros")

lojas = sorted([x for x in df["LOJA_N"].dropna().astype(str).unique() if x.strip() != ""])
opcoes_loja = ["TODAS"] + lojas
lojas_sel = st.sidebar.multiselect("Lojas (LOJA)", opcoes_loja, default=["TODAS"])

if "TODAS" in lojas_sel and len(lojas_sel) > 1:
    lojas_sel = [x for x in lojas_sel if x != "TODAS"]

if "TODAS" in lojas_sel:
    lojas_sel_aplicadas = lojas[:]
else:
    lojas_sel_aplicadas = lojas_sel

vendedores = sorted([x for x in df["VENDEDOR_N"].dropna().astype(str).unique() if x.strip() != ""])
opcoes_vendedor = ["TODOS"] + vendedores
vendedores_sel = st.sidebar.multiselect("Vendedor (VENDEDOR)", opcoes_vendedor, default=["TODOS"])

if "TODOS" in vendedores_sel and len(vendedores_sel) > 1:
    vendedores_sel = [x for x in vendedores_sel if x != "TODOS"]

if "TODOS" in vendedores_sel:
    vendedores_sel_aplicados = vendedores[:]
else:
    vendedores_sel_aplicados = vendedores_sel

datas_validas = df["DATA"].dropna()
if len(datas_validas) > 0:
    data_min = datas_validas.min().date()
    data_max = datas_validas.max().date()
else:
    data_min = date.today()
    data_max = date.today()

st.sidebar.subheader("Período (DATA)")
data_ini = st.sidebar.date_input("Data inicial", value=data_min, min_value=data_min, max_value=data_max)
data_fim = st.sidebar.date_input("Data final", value=data_max, min_value=data_min, max_value=data_max)

st.sidebar.divider()
if st.sidebar.button("Recarregar agora (ignorar cache)"):
    st.cache_data.clear()
    st.rerun()

# =========================
# Aplica filtros (dashboard principal)
# =========================
df_f = df.copy()

if lojas_sel_aplicadas:
    df_f = df_f[df_f["LOJA_N"].isin(lojas_sel_aplicadas)]
else:
    df_f = df_f.iloc[0:0]

if vendedores_sel_aplicados:
    df_f = df_f[df_f["VENDEDOR_N"].isin(vendedores_sel_aplicados)]
else:
    df_f = df_f.iloc[0:0]

df_f = df_f[df_f["DATA"].notna()]
df_f = df_f[(df_f["DATA"].dt.date >= data_ini) & (df_f["DATA"].dt.date <= data_fim)]

fat_total = float(df_f["FAT_LINHA"].sum()) if len(df_f) else 0.0

st.divider()

# =========================
# Seleção de meses (multi)
# =========================
st.markdown("### Seleção de meses (para análise e tabelas)")
mes_opts = [nome for _, nome in MESES]
mes_sel_multi = st.multiselect(
    "Selecione 1 ou mais meses",
    options=mes_opts,
    default=[mes_opts[0]],
)

mes_nome_to_num = {nome: num for num, nome in MESES}
mes_nums_sel = [mes_nome_to_num[m] for m in mes_sel_multi if m in mes_nome_to_num]
if not mes_nums_sel:
    mes_nums_sel = [num for num, _ in MESES]

mes_sel_label = ", ".join([m for m in mes_opts if mes_nome_to_num[m] in mes_nums_sel])
st.markdown(f"**Meses selecionados:** {mes_sel_label}")

# ============================================================
# Comparativo: 2025 (fixo) x 2026 (Excel)
# ============================================================
st.subheader("Comparativo: 2025 (Ano-1) x 2026 (Ano Atual)")

lojas_keys_aplicadas = [canonical_key(x) for x in lojas_sel_aplicadas if canonical_key(x)]
df_meta_ref, df_ano1_ref, dias_uteis_map = carregar_referencias_planejamento()

# Ano-1 agora vem preferencialmente da planilha METAS.xlsx, aba ANO-1.
# Se a planilha auxiliar não existir ou estiver vazia, mantém os valores fixos antigos como fallback.
if df_ano1_ref is not None and not df_ano1_ref.empty:
    df_2025 = df_ano1_ref[df_ano1_ref["LOJA_KEY"].isin(lojas_keys_aplicadas)].copy()
else:
    df_2025 = montar_df_2025_fixo(lojas_keys_aplicadas)

df_2026 = df.copy()
df_2026 = df_2026[df_2026["DATA"].notna()].copy()
df_2026 = df_2026[df_2026["DATA"].dt.year == ANO_ATUAL].copy()

if lojas_sel_aplicadas:
    df_2026 = df_2026[df_2026["LOJA_N"].isin(lojas_sel_aplicadas)].copy()
else:
    df_2026 = df_2026.iloc[0:0].copy()

if vendedores_sel_aplicados:
    df_2026 = df_2026[df_2026["VENDEDOR_N"].isin(vendedores_sel_aplicados)].copy()
else:
    df_2026 = df_2026.iloc[0:0].copy()

df_2026 = df_2026[(df_2026["DATA"].dt.date >= data_ini) & (df_2026["DATA"].dt.date <= data_fim)].copy()
df_2026["MES_NUM"] = df_2026["DATA"].dt.month

df_2026_mensal = (
    df_2026.groupby(["LOJA_KEY", "MES_NUM"], dropna=False)["FAT_LINHA"]
    .sum()
    .reset_index()
    .rename(columns={"FAT_LINHA": "VENDAS_2026"})
)

df_comp = df_2025.merge(df_2026_mensal, on=["LOJA_KEY", "MES_NUM"], how="left")
df_comp["VENDAS_2026"] = df_comp["VENDAS_2026"].fillna(0.0)
df_comp["VAR_R$"] = df_comp["VENDAS_2026"] - df_comp["VENDAS_2025"]
df_comp["VAR_%"] = df_comp.apply(
    lambda r: (r["VAR_R$"] / r["VENDAS_2025"] * 100) if r["VENDAS_2025"] not in (0, None) else None,
    axis=1,
)

map_key_to_loja = (
    df[["LOJA_KEY", "LOJA_N"]].dropna().drop_duplicates().groupby("LOJA_KEY")["LOJA_N"].first().to_dict()
)
df_comp["LOJA"] = df_comp["LOJA_KEY"].map(lambda k: str(map_key_to_loja.get(k, k)))
df_comp["MES"] = df_comp["MES_NUM"].map({m: nome for m, nome in MESES})

df_mes = df_comp[df_comp["MES_NUM"].isin(mes_nums_sel)].copy()

total_2025 = float(df_mes["VENDAS_2025"].sum()) if len(df_mes) else 0.0
total_2026 = float(df_mes["VENDAS_2026"].sum()) if len(df_mes) else 0.0
var_abs_total = total_2026 - total_2025
var_pct_total = (var_abs_total / total_2025 * 100) if total_2025 != 0 else None

k1, k2, k3, k4 = st.columns(4)
k1.metric(f"Total {ANO_BASE} (R$)", "R$ " + format_brl(total_2025))
k2.metric(f"Total {ANO_ATUAL} (R$)", "R$ " + format_brl(total_2026))
k3.metric("Variação (R$)", "R$ " + format_brl(var_abs_total))
k4.metric("Variação (%)", (f"{var_pct_total:.2f}%".replace(".", ",")) if var_pct_total is not None else "—")

# ============================================================
# KPI: META x REALIZADO (velocímetro)
# ============================================================
st.markdown("### Meta × Realizado")

# Meta agora vem preferencialmente da planilha METAS.xlsx, aba META DAUTO TINTAS.
# Se a planilha auxiliar não existir ou estiver vazia, mantém os valores fixos antigos como fallback.
if df_meta_ref is not None and not df_meta_ref.empty:
    df_meta = df_meta_ref[df_meta_ref["LOJA_KEY"].isin(lojas_keys_aplicadas)].copy()
else:
    df_meta = montar_df_metas_fixas(lojas_keys_aplicadas)

df_meta_sel = df_meta[df_meta["MES_NUM"].isin(mes_nums_sel)].copy()

meta_total_sel = float(df_meta_sel["META"].sum()) if len(df_meta_sel) else 0.0
real_total_sel = float(df_mes["VENDAS_2026"].sum()) if len(df_mes) else 0.0

# ===== Previsão de fechamento por calendário comercial ajustado
df_prev = df_f.copy()
df_prev = df_prev[df_prev["DATA"].notna()].copy()
df_prev = df_prev[df_prev["DATA"].dt.month.isin(mes_nums_sel)].copy()

# O calendário da previsão é calculado pela empresa inteira, e não apenas pelo
# filtro de loja/vendedor. Assim, um dia útil sem venda em uma loja específica
# continua contando como dia comercial; só é removido quando a empresa inteira
# ficou com venda zero naquele dia.
infos_calendario = []
for _mes_calc in mes_nums_sel:
    try:
        import calendar
        _ultimo_mes = date(ANO_ATUAL, int(_mes_calc), calendar.monthrange(ANO_ATUAL, int(_mes_calc))[1])
        _corte_mes = min(data_fim, _ultimo_mes)
        infos_calendario.append(_calcular_calendario_comercial_mes(df, ANO_ATUAL, int(_mes_calc), _corte_mes))
    except Exception:
        pass

dias_venda_equiv = float(sum(float(x.get("dias_realizados", 0.0) or 0.0) for x in infos_calendario))
dias_uteis_equiv = float(sum(float(x.get("dias_validos_mes", 0.0) or 0.0) for x in infos_calendario))
dias_restantes_equiv = float(sum(float(x.get("dias_restantes", 0.0) or 0.0) for x in infos_calendario))
dias_ignorados_periodo = []
for _info in infos_calendario:
    dias_ignorados_periodo.extend(_info.get("dias_ignorados", []) or [])

previsao_fechamento, media_por_dia_venda = _previsao_por_calendario(real_total_sel, dias_venda_equiv, dias_restantes_equiv)
if dias_venda_equiv <= 0:
    media_por_dia_venda = None
    previsao_fechamento = None

clientes_base = df_prev.copy()
if "CLIENTE_N" in clientes_base.columns:
    clientes_base = clientes_base[clientes_base["CLIENTE_N"].astype(str).str.strip().ne("")]
    clientes_base = clientes_base[clientes_base["CLIENTE_N"].astype(str).str.upper().ne("SEM CLIENTE")]
if VAL_COL in clientes_base.columns:
    clientes_base = clientes_base[pd.to_numeric(clientes_base[VAL_COL], errors="coerce").fillna(0) > 0]
clientes_ativos = int(clientes_base["CLIENTE_N"].nunique()) if ("CLIENTE_N" in clientes_base.columns and len(clientes_base)) else 0

pct_meta = (real_total_sel / meta_total_sel * 100) if meta_total_sel != 0 else None
dif_meta_r = real_total_sel - meta_total_sel
dif_meta_p = (dif_meta_r / meta_total_sel * 100) if meta_total_sel != 0 else None

st.markdown("#### Realizado e Previsão")
pc1, pc2, pc3, pc4, pc5 = st.columns(5)
with pc1:
    st.metric("Faturamento Atual (R$)", "R$ " + format_brl(real_total_sel))
with pc2:
    st.metric("Dias de Venda", format_decimal_pt(dias_venda_equiv, 1), f"Restam {format_decimal_pt(dias_restantes_equiv, 1)}")
with pc3:
    st.metric("Média por Dia de Venda", ("R$ " + format_brl(media_por_dia_venda)) if media_por_dia_venda is not None else "—")
with pc4:
    st.metric("Previsão de Fechamento", ("R$ " + format_brl(previsao_fechamento)) if previsao_fechamento is not None else "—")
with pc5:
    st.metric("Clientes Ativos", f"{clientes_ativos:,}".replace(",", "."))

st.caption(
    "Dias de venda: segunda a sexta = 1, sábado = 0,5. "
    "Dias comerciais com venda zero na empresa inteira são tratados como feriado/dia sem expediente e saem da previsão."
    + (f" Dias ignorados no recorte: {', '.join(dias_ignorados_periodo)}." if dias_ignorados_periodo else "")
)

a1, a2, a3 = st.columns([1.0, 1.0, 1.6], gap="large")
with a1:
    st.metric("Meta (R$)", "R$ " + format_brl(meta_total_sel))
with a2:
    st.metric("Realizado (R$)", "R$ " + format_brl(real_total_sel))
with a3:
    if pct_meta is None:
        st.info("Meta zerada para o recorte selecionado (não é possível calcular %).")
    else:
        max_axis = 120
        val = max(0, pct_meta)
        fig_gauge = go.Figure(
            go.Indicator(
                mode="gauge+number",
                value=val,
                number={"suffix": "%", "valueformat": ".1f"},
                title={"text": "% da meta atingida"},
                gauge={
                    "axis": {"range": [0, max_axis]},
                    "bar": {"color": "#1f77b4"},
                    "steps": [
                        {"range": [0, 80], "color": "#f2f2f2"},
                        {"range": [80, 100], "color": "#e6e6e6"},
                        {"range": [100, max_axis], "color": "#d9d9d9"},
                    ],
                    "threshold": {"line": {"color": "#d62728", "width": 3}, "thickness": 0.75, "value": 100},
                },
            )
        )
        fig_gauge.update_layout(margin=dict(l=20, r=20, t=60, b=10), height=260)
        st.plotly_chart(fig_gauge, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

b1, b2 = st.columns(2)
with b1:
    st.metric("Diferença (Realizado − Meta) (R$)", "R$ " + format_brl(dif_meta_r))
with b2:
    st.metric("Diferença (%)", (f"{dif_meta_p:.2f}%".replace(".", ",")) if dif_meta_p is not None else "—")

# ------------------------------------------------------------
# Tabela Mês x Loja (drill)
# ------------------------------------------------------------
with st.expander("Abrir tabela comparativa Ano-1"):
    base_tbl = df_comp[
        ["MES", "MES_NUM", "LOJA", "LOJA_KEY", "VENDAS_2025", "VENDAS_2026", "VAR_R$", "VAR_%"]
    ].copy()

    base_tbl = base_tbl.sort_values(["LOJA_KEY", "MES_NUM"]).copy()
    base_tbl["ACUM_DIF_R$"] = base_tbl.groupby("LOJA_KEY")["VAR_R$"].cumsum()

    rows = []
    total_acum = 0.0

    for mes_num, mes_nome in MESES:
        bloco = base_tbl[base_tbl["MES_NUM"] == mes_num].copy()
        if len(bloco) == 0:
            continue

        bloco["_LOJA_ORD"] = bloco["LOJA_KEY"].map(lambda k: LOJA_KEY_RANK.get(k, DEFAULT_RANK)).astype(int)
        bloco = bloco.sort_values(["_LOJA_ORD", "LOJA"]).drop(columns=["_LOJA_ORD"])

        for _, r in bloco.iterrows():
            rows.append(
                {
                    "MES": mes_nome,
                    "LOJA": r["LOJA"],
                    "2025 (R$)": r["VENDAS_2025"],
                    "2026 (R$)": r["VENDAS_2026"],
                    "Dif (R$)": r["VAR_R$"],
                    "Dif (%)": r["VAR_%"],
                    "Acum Dif (R$)": r["ACUM_DIF_R$"],
                }
            )

        sum_2025 = float(bloco["VENDAS_2025"].sum())
        sum_2026 = float(bloco["VENDAS_2026"].sum())
        sum_var = sum_2026 - sum_2025
        sum_pct = (sum_var / sum_2025 * 100) if sum_2025 != 0 else None

        total_acum += sum_var

        rows.append(
            {
                "MES": mes_nome,
                "LOJA": "TOTAL",
                "2025 (R$)": sum_2025,
                "2026 (R$)": sum_2026,
                "Dif (R$)": sum_var,
                "Dif (%)": sum_pct,
                "Acum Dif (R$)": total_acum,
            }
        )

    df_tbl = pd.DataFrame(rows)
    df_tbl_disp = month_block_style(df_tbl)

    styler_tbl = (
        df_tbl_disp.style
        .apply(
            lambda row: ["background-color: #f2f2f2; font-weight: 900;"] * len(row)
            if str(row.get("LOJA", "")).upper() == "TOTAL"
            else [""] * len(row),
            axis=1,
        )
        .format(
            {
                "2025 (R$)": lambda v: "R$ " + format_brl(v),
                "2026 (R$)": lambda v: "R$ " + format_brl(v),
                "Dif (R$)": lambda v: "R$ " + format_brl(v),
                "Dif (%)": lambda v: (f"{v:.2f}%".replace(".", ",")) if v is not None and not (isinstance(v, float) and pd.isna(v)) else "—",
                "Acum Dif (R$)": lambda v: "R$ " + format_brl(v),
            }
        )
        .pipe(lambda s: styler_map_compat(s, color_pos_neg, subset=["Dif (R$)", "Dif (%)", "Acum Dif (R$)"]))
    )

    st.dataframe(styler_tbl, use_container_width=True, hide_index=True, height=560)


    # ------------------------------------------------------------
    # Tabela Meta x Realizado (drill)
    # ------------------------------------------------------------
with st.expander("Abrir Tabela Meta x Realizado"):
    # Meta e Realizado por loja no recorte selecionado (meses + filtros de loja + período)
    # Meta vem do dicionário fixo; Realizado vem do Excel (ANO_ATUAL) no mesmo recorte.
    meta_loja = (
        df_meta_sel.groupby("LOJA_KEY", dropna=False)["META"]
        .sum()
        .reset_index()
        .rename(columns={"META": "META_R$"})
    )

    real_loja = (
        df_mes.groupby("LOJA_KEY", dropna=False)["VENDAS_2026"]
        .sum()
        .reset_index()
        .rename(columns={"VENDAS_2026": "REALIZADO_R$"})
    )

    mkp_tbl = meta_loja.merge(real_loja, on="LOJA_KEY", how="outer").fillna(0.0)

    # Nome da loja
    mkp_tbl["LOJA"] = mkp_tbl["LOJA_KEY"].map(lambda k: str(map_key_to_loja.get(k, k)))

    mkp_tbl["DIF_R$"] = mkp_tbl["REALIZADO_R$"] - mkp_tbl["META_R$"]
    mkp_tbl["% ATINGIDO"] = mkp_tbl.apply(
        lambda r: (r["REALIZADO_R$"] / r["META_R$"] * 100) if r["META_R$"] not in (0, None) else None,
        axis=1,
    )

    # Ordena lojas (mesma ordem usada no comparativo)
    mkp_tbl["_LOJA_ORD"] = mkp_tbl["LOJA_KEY"].map(lambda k: LOJA_KEY_RANK.get(k, DEFAULT_RANK)).astype(int)
    mkp_tbl = mkp_tbl.sort_values(["_LOJA_ORD", "LOJA"]).drop(columns=["_LOJA_ORD"])

    # Linha TOTAL
    t_meta = float(mkp_tbl["META_R$"].sum()) if len(mkp_tbl) else 0.0
    t_real = float(mkp_tbl["REALIZADO_R$"].sum()) if len(mkp_tbl) else 0.0
    t_dif = t_real - t_meta
    t_pct = (t_real / t_meta * 100) if t_meta != 0 else None

    mkp_tbl = pd.concat(
        [
            mkp_tbl,
            pd.DataFrame(
                [
                    {
                        "LOJA_KEY": "",
                        "LOJA": "TOTAL",
                        "META_R$": t_meta,
                        "REALIZADO_R$": t_real,
                        "DIF_R$": t_dif,
                        "% ATINGIDO": t_pct,
                    }
                ]
            ),
        ],
        ignore_index=True,
    )

    view = mkp_tbl[["LOJA", "META_R$", "REALIZADO_R$", "DIF_R$", "% ATINGIDO"]].copy()

    sty = (
        view.style
        .apply(lambda row: ["background-color: #f2f2f2; font-weight: 900;"] * len(row) if str(row.get("LOJA", "")).upper() == "TOTAL" else [""] * len(row), axis=1)
        .format(
            {
                "META_R$": lambda v: "R$ " + format_brl(v),
                "REALIZADO_R$": lambda v: "R$ " + format_brl(v),
                "DIF_R$": lambda v: "R$ " + format_brl(v),
                "% ATINGIDO": lambda v: (f"{v:.2f}%".replace(".", ",")) if v is not None and not (isinstance(v, float) and pd.isna(v)) else "—",
            }
        )
        .pipe(lambda s: styler_map_compat(s, color_pos_neg, subset=["DIF_R$", "% ATINGIDO"]))
    )

    st.dataframe(sty, use_container_width=True, hide_index=True, height=520)

with st.expander("Abrir Tabela Previsão de Fechamento"):
    st.markdown("### Previsão de Fechamento por Loja")

    # ------------------------------------------------------------
    # Tabela de Previsão de Fechamento (drill)
    # ------------------------------------------------------------
    prev_base = df_2026.copy()
    prev_base = prev_base[prev_base["MES_NUM"].isin(mes_nums_sel)].copy()

    # Usa o mesmo contador de dias comerciais válidos para todas as lojas.
    # Se apenas uma loja vendeu zero em um dia útil, esse dia continua contando;
    # só sai do calendário quando a empresa inteira ficou sem venda.
    lojas_prev_lista = sorted(set(prev_base["LOJA_KEY"].dropna().tolist()) | set(df_mes["LOJA_KEY"].dropna().tolist()) | set(df_meta_sel["LOJA_KEY"].dropna().tolist()))
    dias_loja = pd.DataFrame({
        "LOJA_KEY": lojas_prev_lista,
        "DIAS_VENDA": [dias_venda_equiv] * len(lojas_prev_lista),
    })

    realizado_loja = (
        df_mes.groupby("LOJA_KEY", dropna=False)["VENDAS_2026"]
        .sum()
        .reset_index(name="REALIZADO_R$")
    )

    ano1_loja = (
        df_mes.groupby("LOJA_KEY", dropna=False)["VENDAS_2025"]
        .sum()
        .reset_index(name="ANO_1_R$")
    )

    meta_prev_loja = (
        df_meta_sel.groupby("LOJA_KEY", dropna=False)["META"]
        .sum()
        .reset_index(name="META_R$")
    )

    base_lojas_prev = pd.DataFrame({"LOJA_KEY": sorted(set(df_meta_sel["LOJA_KEY"].dropna().tolist()) | set(df_mes["LOJA_KEY"].dropna().tolist()))})
    prev_tbl = base_lojas_prev.merge(realizado_loja, on="LOJA_KEY", how="left")
    prev_tbl = prev_tbl.merge(dias_loja, on="LOJA_KEY", how="left")
    prev_tbl = prev_tbl.merge(ano1_loja, on="LOJA_KEY", how="left")
    prev_tbl = prev_tbl.merge(meta_prev_loja, on="LOJA_KEY", how="left")
    prev_tbl[["REALIZADO_R$", "DIAS_VENDA", "ANO_1_R$", "META_R$"]] = prev_tbl[["REALIZADO_R$", "DIAS_VENDA", "ANO_1_R$", "META_R$"]].fillna(0.0)

    prev_tbl["MEDIA_DIA_R$"] = prev_tbl.apply(
        lambda r: (r["REALIZADO_R$"] / r["DIAS_VENDA"]) if r["DIAS_VENDA"] not in (0, None) else None,
        axis=1,
    )
    prev_tbl["PREVISAO_R$"] = prev_tbl.apply(
        lambda r: (r["REALIZADO_R$"] + (r["MEDIA_DIA_R$"] * dias_restantes_equiv))
        if r["MEDIA_DIA_R$"] is not None and not (isinstance(r["MEDIA_DIA_R$"], float) and pd.isna(r["MEDIA_DIA_R$"]))
        else None,
        axis=1,
    )
    prev_tbl["DIF_PREV_X_ANO1_R$"] = prev_tbl.apply(
        lambda r: (r["PREVISAO_R$"] - r["ANO_1_R$"]) if r["PREVISAO_R$"] is not None and not (isinstance(r["PREVISAO_R$"], float) and pd.isna(r["PREVISAO_R$"])) else None,
        axis=1,
    )
    prev_tbl["CRESC_%"] = prev_tbl.apply(
        lambda r: (r["DIF_PREV_X_ANO1_R$"] / r["ANO_1_R$"] * 100) if r["ANO_1_R$"] not in (0, None) and r["DIF_PREV_X_ANO1_R$"] is not None and not (isinstance(r["DIF_PREV_X_ANO1_R$"] , float) and pd.isna(r["DIF_PREV_X_ANO1_R$"])) else None,
        axis=1,
    )
    prev_tbl["DIF_PREV_X_META_R$"] = prev_tbl.apply(
        lambda r: (r["PREVISAO_R$"] - r["META_R$"]) if r["PREVISAO_R$"] is not None and not (isinstance(r["PREVISAO_R$"], float) and pd.isna(r["PREVISAO_R$"])) else None,
        axis=1,
    )
    prev_tbl["ATING_META_%"] = prev_tbl.apply(
        lambda r: (r["PREVISAO_R$"] / r["META_R$"] * 100) if r["META_R$"] not in (0, None) and r["PREVISAO_R$"] is not None and not (isinstance(r["PREVISAO_R$"], float) and pd.isna(r["PREVISAO_R$"])) else None,
        axis=1,
    )

    prev_tbl["LOJA"] = prev_tbl["LOJA_KEY"].map(lambda k: str(map_key_to_loja.get(k, k)))
    prev_tbl["_LOJA_ORD"] = prev_tbl["LOJA_KEY"].map(lambda k: LOJA_KEY_RANK.get(k, DEFAULT_RANK)).astype(int)
    prev_tbl = prev_tbl.sort_values(["_LOJA_ORD", "LOJA"]).drop(columns=["_LOJA_ORD"])

    total_prev = {
        "LOJA_KEY": "",
        "LOJA": "TOTAL",
        "REALIZADO_R$": float(prev_tbl["REALIZADO_R$"].sum()) if len(prev_tbl) else 0.0,
        "DIAS_VENDA": float(dias_venda_equiv) if dias_venda_equiv is not None else 0.0,
        "ANO_1_R$": float(prev_tbl["ANO_1_R$"].sum()) if len(prev_tbl) else 0.0,
        "META_R$": float(prev_tbl["META_R$"].sum()) if len(prev_tbl) else 0.0,
    }
    total_prev["MEDIA_DIA_R$"] = (total_prev["REALIZADO_R$"] / total_prev["DIAS_VENDA"]) if total_prev["DIAS_VENDA"] != 0 else None
    total_prev["PREVISAO_R$"] = (total_prev["REALIZADO_R$"] + (total_prev["MEDIA_DIA_R$"] * dias_restantes_equiv)) if total_prev["MEDIA_DIA_R$"] is not None else None
    total_prev["DIF_PREV_X_ANO1_R$"] = (total_prev["PREVISAO_R$"] - total_prev["ANO_1_R$"]) if total_prev["PREVISAO_R$"] is not None else None
    total_prev["CRESC_%"] = (total_prev["DIF_PREV_X_ANO1_R$"] / total_prev["ANO_1_R$"] * 100) if total_prev["ANO_1_R$"] not in (0, None) and total_prev["DIF_PREV_X_ANO1_R$"] is not None else None
    total_prev["DIF_PREV_X_META_R$"] = (total_prev["PREVISAO_R$"] - total_prev["META_R$"]) if total_prev["PREVISAO_R$"] is not None else None
    total_prev["ATING_META_%"] = (total_prev["PREVISAO_R$"] / total_prev["META_R$"] * 100) if total_prev["META_R$"] not in (0, None) and total_prev["PREVISAO_R$"] is not None else None

    prev_tbl = pd.concat([prev_tbl, pd.DataFrame([total_prev])], ignore_index=True)

    prev_view = prev_tbl[[
        "LOJA",
        "REALIZADO_R$",
        "DIAS_VENDA",
        "MEDIA_DIA_R$",
        "PREVISAO_R$",
        "ANO_1_R$",
        "DIF_PREV_X_ANO1_R$",
        "CRESC_%",
        "META_R$",
        "DIF_PREV_X_META_R$",
        "ATING_META_%",
    ]].copy()

    prev_view_fmt = prev_view.copy()
    prev_view_fmt["REALIZADO_R$"] = prev_view_fmt["REALIZADO_R$"].apply(lambda v: "R$ " + format_brl(v))
    prev_view_fmt["DIAS_VENDA"] = prev_view_fmt["DIAS_VENDA"].apply(lambda v: format_decimal_pt(v, 1))
    prev_view_fmt["MEDIA_DIA_R$"] = prev_view_fmt["MEDIA_DIA_R$"].apply(lambda v: ("R$ " + format_brl(v)) if v is not None and not (isinstance(v, float) and pd.isna(v)) else "—")
    prev_view_fmt["PREVISAO_R$"] = prev_view_fmt["PREVISAO_R$"].apply(lambda v: ("R$ " + format_brl(v)) if v is not None and not (isinstance(v, float) and pd.isna(v)) else "—")
    prev_view_fmt["ANO_1_R$"] = prev_view_fmt["ANO_1_R$"].apply(lambda v: "R$ " + format_brl(v))
    prev_view_fmt["DIF_PREV_X_ANO1_R$"] = prev_view_fmt["DIF_PREV_X_ANO1_R$"].apply(lambda v: ("R$ " + format_brl(v)) if v is not None and not (isinstance(v, float) and pd.isna(v)) else "—")
    prev_view_fmt["CRESC_%"] = prev_view_fmt["CRESC_%"].apply(lambda v: (f"{v:.2f}%".replace(".", ",")) if v is not None and not (isinstance(v, float) and pd.isna(v)) else "—")
    prev_view_fmt["META_R$"] = prev_view_fmt["META_R$"].apply(lambda v: "R$ " + format_brl(v))
    prev_view_fmt["DIF_PREV_X_META_R$"] = prev_view_fmt["DIF_PREV_X_META_R$"].apply(lambda v: ("R$ " + format_brl(v)) if v is not None and not (isinstance(v, float) and pd.isna(v)) else "—")
    prev_view_fmt["ATING_META_%"] = prev_view_fmt["ATING_META_%"].apply(lambda v: (f"{v:.2f}%".replace(".", ",")) if v is not None and not (isinstance(v, float) and pd.isna(v)) else "—")

    st.dataframe(prev_view_fmt, use_container_width=True, hide_index=True, height=520)

# ------------------------------------------------------------
# Gráfico 2025 x 2026
# ------------------------------------------------------------
mensal_total = (
    df_comp.groupby(["MES_NUM", "MES"], dropna=False)[["VENDAS_2025", "VENDAS_2026"]]
    .sum()
    .reset_index()
    .sort_values("MES_NUM")
)

mensal_long = mensal_total.melt(
    id_vars=["MES_NUM", "MES"],
    value_vars=["VENDAS_2025", "VENDAS_2026"],
    var_name="Ano",
    value_name="Vendas",
)
mensal_long["Ano"] = mensal_long["Ano"].replace({"VENDAS_2025": str(ANO_BASE), "VENDAS_2026": str(ANO_ATUAL)})

fig_comp = px.line(
    mensal_long,
    x="MES",
    y="Vendas",
    color="Ano",
    markers=True,
    title=f"Vendas Mensais: {ANO_BASE} x {ANO_ATUAL} (Lojas do filtro)",
)
fig_comp.update_layout(yaxis_title="Vendas (R$)", xaxis_title="Mês")
st.plotly_chart(fig_comp, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

st.divider()

# =========================
# Linha 1: Loja + Marcas
# =========================
col1, col2 = st.columns([1.6, 1.0], gap="large")

with col1:
    st.subheader("Faturamento por Loja")

    vendas_loja = (
        df_f.groupby("LOJA_N", dropna=False)["FAT_LINHA"]
        .sum()
        .reset_index()
        .sort_values("FAT_LINHA", ascending=False)
        .rename(columns={"LOJA_N": "LOJA", "FAT_LINHA": "Faturamento"})
    )

    fig_loja = px.bar(vendas_loja, x="LOJA", y="Faturamento", title="Vendas por Loja (QTD × UNIT)")
    fig_loja.update_layout(xaxis_tickangle=-30, yaxis_title="Faturamento (R$)", xaxis_title="Loja")
    st.plotly_chart(fig_loja, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

    st.markdown(f"**Total (todas as lojas do filtro): R$ {format_brl(fat_total)}**")

    tbl_loja = vendas_loja.copy()
    tbl_loja["Faturamento (R$)"] = tbl_loja["Faturamento"].apply(lambda v: "R$ " + format_brl(v))
    st.dataframe(tbl_loja[["LOJA", "Faturamento (R$)"]], use_container_width=True, hide_index=True, height=260)

with col2:
    st.subheader("Top 10 Marcas por Faturamento")

    marcas = (
        df_f.groupby("MARCA_N", dropna=False)["FAT_LINHA"]
        .sum()
        .reset_index()
        .sort_values("FAT_LINHA", ascending=False)
        .rename(columns={"MARCA_N": "MARCA", "FAT_LINHA": "Faturamento"})
    )

    total_ref = fat_total if fat_total != 0 else 1.0
    marcas["% do Total"] = (marcas["Faturamento"] / total_ref) * 100

    top10 = marcas.head(10).copy()
    top10["Faturamento (R$)"] = top10["Faturamento"].apply(lambda v: "R$ " + format_brl(v))
    top10["% do Total"] = top10["% do Total"].apply(lambda p: f"{p:.2f}%".replace(".", ","))

    st.dataframe(top10[["MARCA", "Faturamento (R$)", "% do Total"]], use_container_width=True, hide_index=True, height=360)

    with st.expander("Ver todas as marcas (drill)"):
        full = marcas.copy()
        full["Faturamento (R$)"] = full["Faturamento"].apply(lambda v: "R$ " + format_brl(v))
        full["% do Total"] = full["% do Total"].apply(lambda p: f"{p:.2f}%".replace(".", ","))
        st.dataframe(full[["MARCA", "Faturamento (R$)", "% do Total"]], use_container_width=True, hide_index=True, height=520)


# ============================================================
# KPI: Marca → Linhas → Produtos
# ============================================================
st.subheader("KPI: Marca → Linhas → Produtos")

marcas_disp = sorted([x for x in df_f["MARCA_N"].dropna().astype(str).unique() if str(x).strip() != ""])
marca_opts = ["(Todos)"] + marcas_disp

if "marca_kpi_sel" not in st.session_state:
    st.session_state["marca_kpi_sel"] = "(Todos)"

marca_sel = st.selectbox("Marca (KPI)", options=marca_opts, index=marca_opts.index(st.session_state["marca_kpi_sel"]) if st.session_state["marca_kpi_sel"] in marca_opts else 0, key="marca_kpi_select")
st.session_state["marca_kpi_sel"] = marca_sel

if marca_sel == "(Todos)":
    df_marca = df_f.copy()
else:
    df_marca = df_f[df_f["MARCA_N"] == marca_sel].copy()

# Totais do recorte
val_total_marca = float(df_marca[VAL_COL].sum()) if len(df_marca) else 0.0
qtd_total_marca = float(df_marca["QTD_NUM"].sum()) if (len(df_marca) and "QTD_NUM" in df_marca.columns) else 0.0

km1, km2 = st.columns(2)
with km1:
    st.metric("Valor (R$)", "R$ " + format_brl(val_total_marca))
with km2:
    st.metric("Quantidade (QTD)", format_brl(qtd_total_marca))

# Linhas da marca
linhas_marca = (
    df_marca.groupby("LINHA_N", dropna=False)
    .agg(VALOR=(VAL_COL, "sum"), QTD=("QTD_NUM", "sum"))
    .reset_index()
    .rename(columns={"LINHA_N": "LINHA"})
    .sort_values("VALOR", ascending=False)
)

linhas_view = linhas_marca.copy()
linhas_view["VALOR (R$)"] = linhas_view["VALOR"].apply(lambda v: "R$ " + format_brl(v))
linhas_view["QTD"] = linhas_view["QTD"].apply(lambda v: format_brl(v))

st.markdown("#### Linhas (da marca selecionada)")
st.dataframe(linhas_view[["LINHA", "VALOR (R$)", "QTD"]], use_container_width=True, hide_index=True, height=420)

# Produtos da marca (CÓD + Descrição)
with st.expander("Ver produtos (CÓD + Descrição)"):
    if ("COD_N" not in df_marca.columns) or ("DESC_CANON_N" not in df_marca.columns):
        st.info("Colunas de produto (CÓD / Descrição) não foram encontradas no Excel.")
    else:
        prod = (
            df_marca.groupby(["COD_N", "DESC_CANON_N"], dropna=False)
            .agg(VALOR=(VAL_COL, "sum"), QTD=("QTD_NUM", "sum"))
            .reset_index()
            .rename(columns={"COD_N": "CÓD", "DESC_CANON_N": "DESCRIÇÃO"})
            .sort_values("VALOR", ascending=False)
        )
        prod_view = prod.copy()
        prod_view["VALOR (R$)"] = prod_view["VALOR"].apply(lambda v: "R$ " + format_brl(v))
        prod_view["QTD"] = prod_view["QTD"].apply(lambda v: format_brl(v))
        st.dataframe(prod_view[["CÓD", "DESCRIÇÃO", "VALOR (R$)", "QTD"]], use_container_width=True, hide_index=True, height=520)

st.subheader("Indicador: Marcas × Loja")

marcas_mxl = sorted([x for x in df_f["MARCA_N"].dropna().astype(str).unique() if str(x).strip() != ""])
marcas_sel_mxl = st.multiselect(
    "Selecione uma ou mais marcas",
    options=marcas_mxl,
    default=[],
    key="marcas_x_loja_sel",
)

if len(marcas_sel_mxl) == 0:
    st.info("Selecione ao menos uma marca para ver o ranking das lojas, as linhas e os produtos.")
else:
    df_mxl = df_f[df_f["MARCA_N"].isin(marcas_sel_mxl)].copy()

    c_mxl_1, c_mxl_2 = st.columns([1.1, 1.2], gap="large")

    with c_mxl_1:
        lojas_mxl = (
            df_mxl.groupby("LOJA_N", dropna=False)[VAL_COL]
            .sum()
            .reset_index()
            .rename(columns={"LOJA_N": "LOJA", VAL_COL: "Faturamento"})
            .sort_values("Faturamento", ascending=False)
        )
        lojas_mxl_view = lojas_mxl.copy()
        lojas_mxl_view["Faturamento (R$)"] = lojas_mxl_view["Faturamento"].apply(lambda v: "R$ " + format_brl(v))
        st.markdown("#### Ranking das Lojas")
        st.dataframe(lojas_mxl_view[["LOJA", "Faturamento (R$)"]], use_container_width=True, hide_index=True, height=360)

    with c_mxl_2:
        fig_mxl = px.bar(
            lojas_mxl.head(15),
            x="Faturamento",
            y="LOJA",
            orientation="h",
            title="Faturamento por Loja das Marcas Selecionadas",
        )
        fig_mxl.update_layout(
            yaxis={"categoryorder": "total ascending"},
            xaxis_title="Faturamento (R$)",
            yaxis_title="Loja",
        )
        st.plotly_chart(fig_mxl, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

    st.markdown("#### Ranking de Vendedores")
    vendedores_mxl = (
        df_mxl.groupby("VENDEDOR_N", dropna=False)[VAL_COL]
        .sum()
        .reset_index()
        .rename(columns={"VENDEDOR_N": "VENDEDOR", VAL_COL: "Faturamento"})
        .sort_values("Faturamento", ascending=False)
    )
    vendedores_mxl_view = vendedores_mxl.copy()
    vendedores_mxl_view["Faturamento (R$)"] = vendedores_mxl_view["Faturamento"].apply(lambda v: "R$ " + format_brl(v))
    st.dataframe(vendedores_mxl_view[["VENDEDOR", "Faturamento (R$)"]], use_container_width=True, hide_index=True, height=320)

    linhas_mxl = (
        df_mxl.groupby("LINHA_N", dropna=False)[VAL_COL]
        .sum()
        .reset_index()
        .rename(columns={"LINHA_N": "LINHA", VAL_COL: "Faturamento"})
        .sort_values("Faturamento", ascending=False)
    )
    linhas_mxl_view = linhas_mxl.copy()
    linhas_mxl_view["Faturamento (R$)"] = linhas_mxl_view["Faturamento"].apply(lambda v: "R$ " + format_brl(v))

    lm1, lm2 = st.columns([1.1, 1.3], gap="large")
    with lm1:
        st.markdown("#### Linhas das Marcas Selecionadas")
        st.dataframe(linhas_mxl_view[["LINHA", "Faturamento (R$)"]], use_container_width=True, hide_index=True, height=360)

    with lm2:
        linhas_opts_mxl = linhas_mxl["LINHA"].tolist()
        linha_sel_mxl = st.selectbox(
            "Selecione a linha para detalhar os produtos",
            options=linhas_opts_mxl,
            index=0 if len(linhas_opts_mxl) else None,
            key="linha_sel_marcas_x_loja",
        ) if len(linhas_opts_mxl) else None

        if linha_sel_mxl is not None:
            df_prod_mxl = df_mxl[df_mxl["LINHA_N"] == linha_sel_mxl].copy()
            prod_mxl = (
                df_prod_mxl.groupby(["COD_N", "DESC_CANON_N"], dropna=False)
                .agg(VALOR=(VAL_COL, "sum"), QTD=("QTD_NUM", "sum"))
                .reset_index()
                .rename(columns={"COD_N": "CÓD", "DESC_CANON_N": "DESCRIÇÃO"})
                .sort_values("VALOR", ascending=False)
            )
            prod_mxl_view = prod_mxl.copy()
            prod_mxl_view["VALOR (R$)"] = prod_mxl_view["VALOR"].apply(lambda v: "R$ " + format_brl(v))
            prod_mxl_view["QTD"] = prod_mxl_view["QTD"].apply(lambda v: format_brl(v))
            st.markdown(f"#### Produtos da Linha: {linha_sel_mxl}")
            st.dataframe(prod_mxl_view[["CÓD", "DESCRIÇÃO", "VALOR (R$)", "QTD"]], use_container_width=True, hide_index=True, height=360)

st.divider()

# =========================
# Linha 2: Vendedor + Segmento
# =========================
col3, col4 = st.columns([1.4, 1.0], gap="large")

with col3:
    st.subheader("Ranking de Vendas por Vendedor")

    vendedores = (
        df_f.groupby("VENDEDOR_N", dropna=False)["FAT_LINHA"]
        .sum()
        .reset_index()
        .sort_values("FAT_LINHA", ascending=False)
        .rename(columns={"VENDEDOR_N": "VENDEDOR", "FAT_LINHA": "Faturamento"})
    )

    fig_vend = px.bar(
        vendedores.head(15),
        x="Faturamento",
        y="VENDEDOR",
        orientation="h",
        title="Top Vendedores (QTD × UNIT)",
    )
    fig_vend.update_layout(yaxis={"categoryorder": "total ascending"}, xaxis_title="Faturamento (R$)", yaxis_title="Vendedor")
    st.plotly_chart(fig_vend, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

    with st.expander("Ver todos os vendedores"):
        vend_full = vendedores.copy()
        vend_full["Faturamento (R$)"] = vend_full["Faturamento"].apply(lambda v: "R$ " + format_brl(v))
        st.dataframe(vend_full[["VENDEDOR", "Faturamento (R$)"]], use_container_width=True, hide_index=True, height=520)

    # ============================================================
    # ✅ TABELA MARKUP (OCULTA EM DRILL)
    # ============================================================
    with st.expander("Abrir tabela: Faturamento × Custo TT + ST × Markup (por Vendedor)"):
        df_tmp = df_f.copy()
        df_tmp = df_tmp[df_tmp["FAT_LINHA"].notna()].copy()

        if "CUSTO_ST_NUM" not in df_tmp.columns or (df_tmp["CUSTO_ST_NUM"].notna().sum() == 0):
            st.info("Coluna **CUSTO TT + ST** não encontrada (ou sem valores) no Excel para o recorte filtrado.")
        else:
            vend_mkp = (
                df_tmp.groupby("VENDEDOR_N", dropna=False)[["FAT_LINHA", "CUSTO_ST_NUM"]]
                .sum()
                .reset_index()
                .rename(columns={"VENDEDOR_N": "VENDEDOR", "FAT_LINHA": "Faturamento", "CUSTO_ST_NUM": "Custo TT + ST"})
            )

            vend_mkp["MARKUP"] = vend_mkp.apply(
                lambda r: (r["Faturamento"] / r["Custo TT + ST"])
                if (r["Custo TT + ST"] not in (0, None) and not (isinstance(r["Custo TT + ST"], float) and pd.isna(r["Custo TT + ST"])))
                else None,
                axis=1,
            )

            vend_mkp = vend_mkp.sort_values("Faturamento", ascending=False).copy()

            tot_fat = float(vend_mkp["Faturamento"].sum()) if len(vend_mkp) else 0.0
            tot_cus = float(vend_mkp["Custo TT + ST"].sum()) if len(vend_mkp) else 0.0
            tot_mkp = (tot_fat / tot_cus) if tot_cus not in (0, None) else None

            vend_disp = vend_mkp.copy()
            vend_disp["Faturamento (R$)"] = vend_disp["Faturamento"].apply(lambda v: "R$ " + format_brl(v))
            vend_disp["Custo TT + ST (R$)"] = vend_disp["Custo TT + ST"].apply(lambda v: "R$ " + format_brl(v))
            vend_disp["Markup"] = vend_disp["MARKUP"].apply(
                lambda v: (f"{v:.2f}x".replace(".", ",")) if v is not None and not (isinstance(v, float) and pd.isna(v)) else "—"
            )

            vend_disp_tot = pd.concat(
                [
                    vend_disp[["VENDEDOR", "Faturamento (R$)", "Custo TT + ST (R$)", "Markup"]],
                    pd.DataFrame([{
                        "VENDEDOR": "TOTAL",
                        "Faturamento (R$)": "R$ " + format_brl(tot_fat),
                        "Custo TT + ST (R$)": "R$ " + format_brl(tot_cus),
                        "Markup": (f"{tot_mkp:.2f}x".replace(".", ",")) if tot_mkp is not None else "—",
                    }]),
                ],
                ignore_index=True,
            )

            st.dataframe(vend_disp_tot, use_container_width=True, hide_index=True, height=360)

with col4:
    st.subheader("Faturamento por Segmento")

    segmento = (
        df_f.groupby("SEGMENTO_N", dropna=False)["FAT_LINHA"]
        .sum()
        .reset_index()
        .sort_values("FAT_LINHA", ascending=False)
        .rename(columns={"SEGMENTO_N": "SEGMENTO", "FAT_LINHA": "Faturamento"})
    )

    fig_seg = px.pie(segmento, names="SEGMENTO", values="Faturamento", title="Participação por Segmento (QTD × UNIT)")
    st.plotly_chart(fig_seg, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

    segmento_tbl = segmento.copy()
    segmento_tbl["Faturamento (R$)"] = segmento_tbl["Faturamento"].apply(lambda v: "R$ " + format_brl(v))
    st.dataframe(segmento_tbl[["SEGMENTO", "Faturamento (R$)"]], use_container_width=True, hide_index=True, height=260)

    # ------------------------------------------------------------
    # Drill: Faturamento por Loja x Segmento Gerencial
    # ------------------------------------------------------------
    with st.expander("Abrir drill: Faturamento por Loja x Segmento Gerencial"):
        st.markdown("#### Faturamento por Loja x Segmento Gerencial")

        def classificar_segmento_gerencial(seg):
            key = canonical_key(seg)

            if key in {"", "SEMSEGMENTO", "NAN", "NONE"}:
                return "Sem Segmento"

            if "AUTOMOT" in key:
                return "Automotivo"

            if "THINNER" in key or "TINER" in key:
                return "Thinner"

            if (
                "IMOBILI" in key
                or "DECOR" in key
                or "INDUSTR" in key
                or "MOVELEIR" in key
                or "MOVEL" in key
            ):
                return "Decor/Industrial"

            return "Outros Segmentos"

        df_seg_loja = df_f.copy()
        df_seg_loja["SEGMENTO_GERENCIAL"] = df_seg_loja["SEGMENTO_N"].apply(classificar_segmento_gerencial)

        ordem_segmentos = ["Automotivo", "Decor/Industrial", "Thinner", "Sem Segmento", "Outros Segmentos"]

        seg_loja = (
            df_seg_loja.groupby(["LOJA_N", "SEGMENTO_GERENCIAL"], dropna=False)["FAT_LINHA"]
            .sum()
            .reset_index()
            .rename(columns={"LOJA_N": "LOJA", "SEGMENTO_GERENCIAL": "Segmento", "FAT_LINHA": "Faturamento"})
        )

        lojas_drill = sorted(
            [x for x in df_seg_loja["LOJA_N"].dropna().astype(str).unique() if x.strip() != ""],
            key=lambda x: (LOJA_KEY_RANK.get(canonical_key(x), DEFAULT_RANK), str(x)),
        )

        idx_completo = pd.MultiIndex.from_product(
            [lojas_drill, ordem_segmentos],
            names=["LOJA", "Segmento"],
        )

        seg_loja = (
            seg_loja.set_index(["LOJA", "Segmento"])
            .reindex(idx_completo, fill_value=0)
            .reset_index()
        )

        # Remove colunas/categorias opcionais quando não houver faturamento no recorte.
        segmentos_com_valor = seg_loja.groupby("Segmento")["Faturamento"].sum()
        segmentos_visiveis = [s for s in ordem_segmentos if float(segmentos_com_valor.get(s, 0.0) or 0.0) > 0]
        if not segmentos_visiveis:
            segmentos_visiveis = ["Automotivo", "Decor/Industrial", "Thinner"]

        seg_loja = seg_loja[seg_loja["Segmento"].isin(segmentos_visiveis)].copy()

        total_loja = (
            seg_loja.groupby("LOJA", dropna=False)["Faturamento"]
            .sum()
            .reset_index(name="Total Loja")
        )
        seg_loja = seg_loja.merge(total_loja, on="LOJA", how="left")
        seg_loja["% da Loja"] = seg_loja.apply(
            lambda r: (r["Faturamento"] / r["Total Loja"] * 100) if r["Total Loja"] not in (0, None) else 0.0,
            axis=1,
        )
        seg_loja["Label"] = seg_loja.apply(
            lambda r: ("R$ " + format_brl(r["Faturamento"]) + "<br>" + f"{r['% da Loja']:.1f}%".replace(".", ","))
            if r["Faturamento"] > 0 else "",
            axis=1,
        )

        fig_seg_loja = px.bar(
            seg_loja,
            x="LOJA",
            y="Faturamento",
            color="Segmento",
            barmode="group",
            text="Label",
            category_orders={"LOJA": lojas_drill, "Segmento": segmentos_visiveis},
            title="Faturamento por Loja: Automotivo x Decor/Industrial x Thinner",
            hover_data={"Faturamento": ":,.2f", "% da Loja": ":.2f", "Total Loja": ":,.2f", "Label": False},
        )
        fig_seg_loja.update_traces(textposition="outside", cliponaxis=False)
        fig_seg_loja.update_layout(
            xaxis_tickangle=-30,
            yaxis_title="Faturamento (R$)",
            xaxis_title="Loja",
            legend_title="Segmento Gerencial",
            height=520,
        )
        st.plotly_chart(fig_seg_loja, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

        st.markdown("##### Tabela detalhada")
        tabela_seg = seg_loja.pivot_table(
            index="LOJA",
            columns="Segmento",
            values="Faturamento",
            aggfunc="sum",
            fill_value=0,
        ).reset_index()

        pct_seg = seg_loja.pivot_table(
            index="LOJA",
            columns="Segmento",
            values="% da Loja",
            aggfunc="sum",
            fill_value=0,
        ).reset_index()

        tabela_final = pd.DataFrame({"LOJA": lojas_drill})
        for seg in segmentos_visiveis:
            valor_map = tabela_seg.set_index("LOJA")[seg].to_dict() if seg in tabela_seg.columns else {}
            pct_map = pct_seg.set_index("LOJA")[seg].to_dict() if seg in pct_seg.columns else {}
            tabela_final[f"{seg} (R$)"] = tabela_final["LOJA"].map(lambda l: float(valor_map.get(l, 0.0) or 0.0))
            tabela_final[f"{seg} (%)"] = tabela_final["LOJA"].map(lambda l: float(pct_map.get(l, 0.0) or 0.0))

        tabela_final["Total Loja (R$)"] = tabela_final[[f"{s} (R$)" for s in segmentos_visiveis]].sum(axis=1)

        total_row = {"LOJA": "TOTAL"}
        total_geral_seg = float(tabela_final["Total Loja (R$)"].sum()) if len(tabela_final) else 0.0
        for seg in segmentos_visiveis:
            total_seg = float(tabela_final[f"{seg} (R$)"].sum()) if len(tabela_final) else 0.0
            total_row[f"{seg} (R$)"] = total_seg
            total_row[f"{seg} (%)"] = (total_seg / total_geral_seg * 100) if total_geral_seg != 0 else 0.0
        total_row["Total Loja (R$)"] = total_geral_seg

        tabela_final = pd.concat([tabela_final, pd.DataFrame([total_row])], ignore_index=True)

        format_cols = {}
        for seg in segmentos_visiveis:
            format_cols[f"{seg} (R$)"] = lambda v: "R$ " + format_brl(v)
            format_cols[f"{seg} (%)"] = lambda v: f"{float(v):.2f}%".replace(".", ",")
        format_cols["Total Loja (R$)"] = lambda v: "R$ " + format_brl(v)

        sty_seg = (
            tabela_final.style
            .apply(lambda row: ["background-color: #f2f2f2; font-weight: 900;"] * len(row) if str(row.get("LOJA", "")).upper() == "TOTAL" else [""] * len(row), axis=1)
            .format(format_cols)
        )

        st.dataframe(sty_seg, use_container_width=True, hide_index=True, height=520)

# ===== MARKUP GERAL (apenas o número) =====
custo_total_mkp = float(df_f["CUSTO_ST_NUM"].sum()) if ("CUSTO_ST_NUM" in df_f.columns and len(df_f)) else 0.0
mkp_geral = (fat_total / custo_total_mkp) if custo_total_mkp not in (0, None) else None

st.markdown("#### Markup Geral")
if mkp_geral is None:
    st.metric("MKP", "—")
    st.caption("Sem dados de **CUSTO TT + ST** (ou custo zerado) no recorte selecionado.")
else:
    st.metric("MKP", (f"{mkp_geral:.2f}".replace(".", ",")))


st.divider()

# ============================================================
# TOP 10 LINHAS + DRILLS
# ============================================================
st.subheader("Top 10 Linhas por Faturamento")

linhas = (
    df_f.groupby("LINHA_N", dropna=False)["FAT_LINHA"]
    .sum()
    .reset_index()
    .sort_values("FAT_LINHA", ascending=False)
    .rename(columns={"LINHA_N": "LINHA", "FAT_LINHA": "Faturamento"})
)

total_ref = fat_total if fat_total != 0 else 1.0
linhas["% do Total"] = (linhas["Faturamento"] / total_ref) * 100

top10_linhas = linhas.head(10).copy()
top10_linhas["Faturamento (R$)"] = top10_linhas["Faturamento"].apply(lambda v: "R$ " + format_brl(v))
top10_linhas["% do Total"] = top10_linhas["% do Total"].apply(lambda p: f"{p:.2f}%".replace(".", ","))

st.dataframe(top10_linhas[["LINHA", "Faturamento (R$)", "% do Total"]], use_container_width=True, hide_index=True, height=360)

if "linha_sel" not in st.session_state:
    st.session_state["linha_sel"] = linhas["LINHA"].iloc[0] if len(linhas) else None

with st.expander("Ver todas as linhas (drill)"):
    full_linhas = linhas.copy()
    full_linhas["Faturamento (R$)"] = full_linhas["Faturamento"].apply(lambda v: "R$ " + format_brl(v))
    full_linhas["% do Total"] = full_linhas["% do Total"].apply(lambda p: f"{p:.2f}%".replace(".", ","))
    st.dataframe(full_linhas[["LINHA", "Faturamento (R$)", "% do Total"]], use_container_width=True, hide_index=True, height=520)

    st.markdown("**Detalhar uma LINHA (abre o detalhamento por MARCA logo abaixo):**")
    st.session_state["linha_sel"] = st.selectbox(
        "Escolha a LINHA para detalhar",
        options=full_linhas["LINHA"].tolist(),
        index=full_linhas["LINHA"].tolist().index(st.session_state["linha_sel"])
        if st.session_state["linha_sel"] in full_linhas["LINHA"].tolist()
        else 0,
        key="linha_sel_from_drill",
    )

st.divider()

st.subheader("Drill: LINHA → Marcas (valores e % sobre a LINHA)")

linhas_disponiveis = linhas["LINHA"].tolist()

if len(linhas_disponiveis) == 0:
    st.info("Não há linhas disponíveis no filtro atual.")
else:
    colS1, colS2, colS3 = st.columns([1.2, 0.6, 1.4], gap="medium")

    with colS1:
        termo = st.text_input("Pesquisar LINHA (digite parte do nome)", value="", placeholder="Ex.: verniz, primer, massa...")

    with colS2:
        if st.button("Pesquisar"):
            t = (termo or "").strip().lower()
            if t:
                matches = [x for x in linhas_disponiveis if t in str(x).lower()]
                if matches:
                    st.session_state["linha_sel"] = matches[0]
                else:
                    st.warning("Nenhuma LINHA encontrada com esse termo. Tente outro.")
            else:
                st.info("Digite um termo para pesquisar.")

    t_live = (termo or "").strip().lower()
    linhas_filtradas = [x for x in linhas_disponiveis if (t_live in str(x).lower())] if t_live else linhas_disponiveis

    linha_padrao = st.session_state.get(
        "linha_sel",
        linhas_filtradas[0] if linhas_filtradas else (linhas_disponiveis[0] if linhas_disponiveis else None),
    )
    if linhas_filtradas and linha_padrao not in linhas_filtradas:
        linha_padrao = linhas_filtradas[0]

    with colS3:
        linha_sel = st.selectbox(
            "LINHA selecionada",
            options=linhas_filtradas if linhas_filtradas else linhas_disponiveis,
            index=(linhas_filtradas if linhas_filtradas else linhas_disponiveis).index(linha_padrao)
            if linha_padrao in (linhas_filtradas if linhas_filtradas else linhas_disponiveis)
            else 0,
            key="linha_sel_main",
        )
        st.session_state["linha_sel"] = linha_sel

    df_linha = df_f[df_f["LINHA_N"] == st.session_state["linha_sel"]].copy()
    total_linha = float(df_linha["FAT_LINHA"].sum()) if len(df_linha) else 0.0

    cA, cB = st.columns([1, 2])
    with cA:
        st.metric("Total da LINHA (R$)", "R$ " + format_brl(total_linha))
    with cB:
        pct_linha_no_total = (total_linha / (fat_total if fat_total != 0 else 1.0)) * 100
        st.metric("% da LINHA sobre Total", f"{pct_linha_no_total:.2f}%".replace(".", ","))

    # Mantém a visão agregada por MARCA para o gráfico lateral
    marcas_linha = (
        df_linha.groupby("MARCA_N", dropna=False)["FAT_LINHA"]
        .sum()
        .reset_index()
        .sort_values("FAT_LINHA", ascending=False)
        .rename(columns={"MARCA_N": "MARCA", "FAT_LINHA": "Faturamento"})
    )

    denom_linha = total_linha if total_linha != 0 else 1.0
    denom_total = fat_total if fat_total != 0 else 1.0

    marcas_linha["% da LINHA"] = (marcas_linha["Faturamento"] / denom_linha) * 100
    marcas_linha["% do Total"] = (marcas_linha["Faturamento"] / denom_total) * 100

    # Nova visão detalhada: LINHA → MARCA → CÓDIGO/DESCRIÇÃO DO PRODUTO
    produtos_linha = (
        df_linha.groupby(["MARCA_N", "COD_N", "DESC_CANON_N"], dropna=False)
        .agg(Faturamento=("FAT_LINHA", "sum"), QTD=("QTD_NUM", "sum"))
        .reset_index()
        .sort_values("Faturamento", ascending=False)
        .rename(
            columns={
                "MARCA_N": "MARCA",
                "COD_N": "CÓDIGO",
                "DESC_CANON_N": "DESCRIÇÃO",
            }
        )
    )

    produtos_linha["% da LINHA"] = (produtos_linha["Faturamento"] / denom_linha) * 100
    produtos_linha["% do Total"] = (produtos_linha["Faturamento"] / denom_total) * 100
    produtos_linha["Faturamento (R$)"] = produtos_linha["Faturamento"].apply(lambda v: "R$ " + format_brl(v))
    produtos_linha["QTD"] = produtos_linha["QTD"].apply(lambda v: format_brl(v))
    produtos_linha["% da LINHA"] = produtos_linha["% da LINHA"].apply(lambda p: f"{p:.2f}%".replace(".", ","))
    produtos_linha["% do Total"] = produtos_linha["% do Total"].apply(lambda p: f"{p:.2f}%".replace(".", ","))

    cX, cY = st.columns([1.2, 1.0], gap="large")
    with cX:
        st.markdown(f"**Produtos dentro da LINHA: {st.session_state['linha_sel']}**")
        st.dataframe(
            produtos_linha[["MARCA", "CÓDIGO", "DESCRIÇÃO", "Faturamento (R$)", "QTD", "% da LINHA", "% do Total"]],
            use_container_width=True,
            hide_index=True,
            height=520,
        )

    with cY:
        fig_ml = px.bar(
            marcas_linha.head(15),
            x="Faturamento",
            y="MARCA",
            orientation="h",
            title=f"Top Marcas na Linha: {st.session_state['linha_sel']}",
        )
        fig_ml.update_layout(
            yaxis={"categoryorder": "total ascending"},
            xaxis_title="Faturamento (R$)",
            yaxis_title="Marca",
        )
        st.plotly_chart(fig_ml, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

st.divider()

# ============================================================
# Ranking de Clientes + Drill (mantido embaixo)
# ============================================================
st.subheader("Ranking de Clientes")

df_cli_base = df_f.copy()
df_cli_base["VAL_CLIENTE"] = df_cli_base[VAL_COL]
df_cli_base = df_cli_base[df_cli_base["VAL_CLIENTE"].notna()].copy()

total_cliente_ref = float(df_cli_base["VAL_CLIENTE"].sum()) if len(df_cli_base) else 0.0
den_total_cli = total_cliente_ref if total_cliente_ref != 0 else 1.0

clientes = (
    df_cli_base.groupby("CLIENTE_N", dropna=False)["VAL_CLIENTE"]
    .sum()
    .reset_index()
    .sort_values("VAL_CLIENTE", ascending=False)
    .rename(columns={"CLIENTE_N": "CLIENTE", "VAL_CLIENTE": "VR_TOTAL"})
)
clientes["% do Total"] = (clientes["VR_TOTAL"] / den_total_cli) * 100

cC1, cC2 = st.columns([1.15, 1.0], gap="large")

with cC1:
    st.markdown("#### Top 15 Clientes")
    top15 = clientes.head(15).copy()
    top15["VR TOTAL (R$)"] = top15["VR_TOTAL"].apply(lambda v: "R$ " + format_brl(v))
    top15["% do Total"] = top15["% do Total"].apply(lambda p: f"{p:.2f}%".replace(".", ","))
    st.dataframe(top15[["CLIENTE", "VR TOTAL (R$)", "% do Total"]], use_container_width=True, hide_index=True, height=380)

    with st.expander("Ver todos os clientes (drill)"):
        full = clientes.copy()
        full["VR TOTAL (R$)"] = full["VR_TOTAL"].apply(lambda v: "R$ " + format_brl(v))
        full["% do Total"] = full["% do Total"].apply(lambda p: f"{p:.2f}%".replace(".", ","))
        st.dataframe(full[["CLIENTE", "VR TOTAL (R$)", "% do Total"]], use_container_width=True, hide_index=True, height=520)

with cC2:
    st.markdown("#### Gráfico: Top 15 Clientes")
    fig_cli = px.bar(
        top15.sort_values("VR_TOTAL"),
        x="VR_TOTAL",
        y="CLIENTE",
        orientation="h",
        title="Top 15 Clientes (VR TOTAL)",
    )
    fig_cli.update_layout(
        yaxis={"categoryorder": "total ascending"},
        xaxis_title="VR TOTAL (R$)",
        yaxis_title="Cliente",
    )
    st.plotly_chart(fig_cli, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

st.divider()

st.subheader("Drill: Cliente → Linhas → Marcas")

clientes_lista = clientes["CLIENTE"].tolist()
clientes_disp = ["(Todos)"] + clientes_lista
if len(clientes_lista) == 0:
    st.info("Não há clientes na base filtrada.")
else:
    s1, s2, s3 = st.columns([1.2, 0.6, 1.4], gap="medium")
    with s1:
        termo_cli = st.text_input("Pesquisar CLIENTE (digite parte do nome)", value="", placeholder="Ex.: construtora, oficina, nome...")

    with s2:
        if st.button("Pesquisar Cliente"):
            t = (termo_cli or "").strip().lower()
            if t:
                matches = [x for x in clientes_lista if t in str(x).lower()]
                if matches:
                    st.session_state["cliente_sel"] = matches[0]
                else:
                    st.warning("Nenhum cliente encontrado com esse termo.")
            else:
                st.info("Digite um termo para pesquisar.")

    t_live = (termo_cli or "").strip().lower()
    clientes_filtrados_base = [x for x in clientes_lista if (t_live in str(x).lower())] if t_live else clientes_lista
    clientes_filtrados = ["(Todos)"] + clientes_filtrados_base

    if "cliente_sel" not in st.session_state:
        st.session_state["cliente_sel"] = "(Todos)"

    cliente_padrao = st.session_state["cliente_sel"]
    if cliente_padrao not in clientes_filtrados:
        cliente_padrao = "(Todos)" if "(Todos)" in clientes_filtrados else (clientes_filtrados[0] if clientes_filtrados else "(Todos)")
        st.session_state["cliente_sel"] = cliente_padrao

    with s3:
        cliente_sel = st.selectbox(
            "CLIENTE selecionado",
            options=clientes_filtrados,
            index=clientes_filtrados.index(cliente_padrao) if cliente_padrao in clientes_filtrados else 0,
            key="cliente_sel_main",
        )
        st.session_state["cliente_sel"] = cliente_sel

    if st.session_state["cliente_sel"] == "(Todos)":
        df_cliente = df_cli_base.copy()
    else:
        df_cliente = df_cli_base[df_cli_base["CLIENTE_N"] == st.session_state["cliente_sel"]].copy()

    total_cliente = float(df_cliente["VAL_CLIENTE"].sum()) if len(df_cliente) else 0.0

    m1, m2 = st.columns([1, 2])
    with m1:
        st.metric("VR TOTAL do Cliente (R$)", "R$ " + format_brl(total_cliente))
    with m2:
        pct_cli_total = (total_cliente / den_total_cli) * 100
        st.metric("% do Cliente sobre Total", f"{pct_cli_total:.2f}%".replace(".", ","))

    linhas_cli = (
        df_cliente.groupby("LINHA_N", dropna=False)["VAL_CLIENTE"]
        .sum()
        .reset_index()
        .sort_values("VAL_CLIENTE", ascending=False)
        .rename(columns={"LINHA_N": "LINHA", "VAL_CLIENTE": "VR_TOTAL"})
    )
    den_cli = total_cliente if total_cliente != 0 else 1.0
    linhas_cli["% do Cliente"] = (linhas_cli["VR_TOTAL"] / den_cli) * 100
    linhas_cli["% do Total"] = (linhas_cli["VR_TOTAL"] / den_total_cli) * 100

    if "linha_cli_sel" not in st.session_state:
        st.session_state["linha_cli_sel"] = "(Todas)"

    cL1, cL2 = st.columns([1.2, 1.0], gap="large")
    with cL1:
        st.markdown("#### Linhas compradas por este Cliente")
        linhas_view = linhas_cli.copy()
        linhas_view["VR TOTAL (R$)"] = linhas_view["VR_TOTAL"].apply(lambda v: "R$ " + format_brl(v))
        linhas_view["% do Cliente"] = linhas_view["% do Cliente"].apply(lambda p: f"{p:.2f}%".replace(".", ","))
        linhas_view["% do Total"] = linhas_view["% do Total"].apply(lambda p: f"{p:.2f}%".replace(".", ","))
        st.dataframe(linhas_view[["LINHA", "VR TOTAL (R$)", "% do Cliente", "% do Total"]], use_container_width=True, hide_index=True, height=420)

        st.markdown("**Selecione uma LINHA para ver as MARCAS dentro dela (para esse cliente):**")
        linhas_opcoes = ["(Todas)"] + linhas_cli["LINHA"].tolist()
        linha_padrao = st.session_state.get("linha_cli_sel", "(Todas)")
        if linha_padrao not in linhas_opcoes:
            linha_padrao = "(Todas)"
        st.session_state["linha_cli_sel"] = st.selectbox(
            "LINHA do Cliente",
            options=linhas_opcoes,
            index=linhas_opcoes.index(linha_padrao),
            key="linha_cli_sel_box",
        )

    with cL2:
        fig_lc = px.bar(
            linhas_cli.head(15).sort_values("VR_TOTAL"),
            x="VR_TOTAL",
            y="LINHA",
            orientation="h",
            title="Top Linhas deste Cliente",
        )
        fig_lc.update_layout(yaxis={"categoryorder": "total ascending"}, xaxis_title="VR TOTAL (R$)", yaxis_title="Linha")
        st.plotly_chart(fig_lc, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

    st.divider()

    linha_label = st.session_state["linha_cli_sel"]
    st.markdown(
        f"### Drill: Cliente **{st.session_state['cliente_sel']}** → Linha **{linha_label}** → Marcas"
    )

    if st.session_state["linha_cli_sel"] == "(Todas)":
        df_cli_linha = df_cliente.copy()
    else:
        df_cli_linha = df_cliente[df_cliente["LINHA_N"] == st.session_state["linha_cli_sel"]].copy()

    total_cli_linha = float(df_cli_linha["VAL_CLIENTE"].sum()) if len(df_cli_linha) else 0.0

    cM1, cM2 = st.columns([1, 2])
    with cM1:
        st.metric("Total da Linha no Cliente (R$)", "R$ " + format_brl(total_cli_linha))
    with cM2:
        pct_linha_cli = (total_cli_linha / (total_cliente if total_cliente != 0 else 1.0)) * 100
        st.metric("% da Linha sobre Cliente", f"{pct_linha_cli:.2f}%".replace(".", ","))

    marcas_cli = (
        df_cli_linha.groupby("MARCA_N", dropna=False)["VAL_CLIENTE"]
        .sum()
        .reset_index()
        .sort_values("VAL_CLIENTE", ascending=False)
        .rename(columns={"MARCA_N": "MARCA", "VAL_CLIENTE": "VR_TOTAL"})
    )

    denom_linha_cli = total_cli_linha if total_cli_linha != 0 else 1.0
    marcas_cli["% da Linha"] = (marcas_cli["VR_TOTAL"] / denom_linha_cli) * 100
    marcas_cli["% do Cliente"] = (marcas_cli["VR_TOTAL"] / (total_cliente if total_cliente != 0 else 1.0)) * 100
    marcas_cli["% do Total"] = (marcas_cli["VR_TOTAL"] / den_total_cli) * 100

    mc1, mc2 = st.columns([1.2, 1.0], gap="large")
    with mc1:
        marcas_view = marcas_cli.copy()
        marcas_view["VR TOTAL (R$)"] = marcas_view["VR_TOTAL"].apply(lambda v: "R$ " + format_brl(v))
        marcas_view["% da Linha"] = marcas_view["% da Linha"].apply(lambda p: f"{p:.2f}%".replace(".", ","))
        marcas_view["% do Cliente"] = marcas_view["% do Cliente"].apply(lambda p: f"{p:.2f}%".replace(".", ","))
        marcas_view["% do Total"] = marcas_view["% do Total"].apply(lambda p: f"{p:.2f}%".replace(".", ","))
        st.dataframe(
            marcas_view[["MARCA", "VR TOTAL (R$)", "% da Linha", "% do Cliente", "% do Total"]],
            use_container_width=True,
            hide_index=True,
            height=520,
        )

    with mc2:
        fig_mc = px.bar(
            marcas_cli.head(15).sort_values("VR_TOTAL"),
            x="VR_TOTAL",
            y="MARCA",
            orientation="h",
            title="Top Marcas nesta Linha (Cliente)",
        )
        fig_mc.update_layout(yaxis={"categoryorder": "total ascending"}, xaxis_title="VR TOTAL (R$)", yaxis_title="Marca")
        st.plotly_chart(fig_mc, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)


# ============================================================
# KPI: Evolução de Faturamento por Cliente (Jan → Dez)
# - A seleção "LINHA do Cliente" (no Drill acima) filtra a evolução.
# - Se LINHA do Cliente = (Todas), considera todas as linhas.
# - Mostra TODOS os clientes (linhas) no modelo Cliente x Meses.
# ============================================================
st.subheader("Evolução de Faturamento por Cliente (Jan → Dez)")

linha_kpi = st.session_state.get("linha_cli_sel", "(Todas)")

df_evo = df_f.copy()
df_evo = df_evo[df_evo["DATA"].notna()].copy()

# Mantém o comportamento de "Jan → Dez" do ANO_ATUAL
df_evo = df_evo[df_evo["DATA"].dt.year == ANO_ATUAL].copy()

# Linha (opcional)
if linha_kpi != "(Todas)":
    df_evo = df_evo[df_evo["LINHA_N"] == linha_kpi].copy()

if len(df_evo) == 0:
    if linha_kpi != "(Todas)":
        st.info(f"Sem dados para o recorte atual na linha **{linha_kpi}** no ano **{ANO_ATUAL}**.")
    else:
        st.info(f"Sem dados para o recorte atual no ano **{ANO_ATUAL}**.")
else:
    df_evo["MES_NUM"] = df_evo["DATA"].dt.month

    # Pivot: Cliente x Mês (VAL_COL já definido no app: VR_TOTAL_NUM ou FAT_LINHA)
    tbl_evo = (
        df_evo.groupby(["CLIENTE_N", "MES_NUM"], dropna=False)[VAL_COL]
        .sum()
        .reset_index()
        .pivot(index="CLIENTE_N", columns="MES_NUM", values=VAL_COL)
        .fillna(0.0)
    )

    # Garante colunas 1..12
    for m in range(1, 13):
        if m not in tbl_evo.columns:
            tbl_evo[m] = 0.0
    tbl_evo = tbl_evo[[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]]

    meses_full = {
        1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL",
        5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO",
        9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO",
    }
    tbl_evo = tbl_evo.rename(columns=meses_full)

    # Ordena por total do ano (maior → menor)
    tbl_evo["_TOTAL_ANO"] = tbl_evo.sum(axis=1)
    tbl_evo = tbl_evo.sort_values("_TOTAL_ANO", ascending=False).drop(columns=["_TOTAL_ANO"])

    # Formatação (R$)
    tbl_disp = tbl_evo.copy()
    for c in tbl_disp.columns:
        tbl_disp[c] = tbl_disp[c].apply(lambda v: "R$ " + format_brl(v))

    tbl_disp = tbl_disp.reset_index().rename(columns={"CLIENTE_N": "CLIENTE"})

    if linha_kpi != "(Todas)":
        st.caption(f"Filtro aplicado: **LINHA do Cliente = {linha_kpi}**")
    else:
        st.caption("Filtro aplicado: **LINHA do Cliente = (Todas)** (todas as linhas)")

    st.dataframe(tbl_disp, use_container_width=True, hide_index=True, height=520)

st.divider()

with st.expander("Ver base filtrada (conferência)"):
    st.dataframe(df_f, use_container_width=True, height=520)

# ============================================================
# Indicador de Compras (Base + COMPRAS + DEVOLUÇÕES) - DRILL
# ============================================================
def _safe_div(num: float, den: float):
    try:
        if den in (0, None) or (isinstance(den, float) and pd.isna(den)):
            return None
        return float(num) / float(den)
    except Exception:
        return None


def _fmt_num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    try:
        return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "—"


def _indicador_style(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    try:
        x = float(v)
    except Exception:
        return ""

    # > 0,00 azul
    if x > 0:
        return "background-color: #D9E8FF; color: #0B3D91; font-weight: 700;"
    # entre -0,01 e -0,09 verde
    if -0.09 <= x <= -0.01:
        return "background-color: #DFF5E1; color: #1B5E20; font-weight: 700;"
    # entre -0,10 e -0,15 amarelo
    if -0.15 <= x <= -0.10:
        return "background-color: #FFF4CC; color: #7A5A00; font-weight: 700;"
    # abaixo de -0,15 vermelho
    if x < -0.15:
        return "background-color: #FFD6D6; color: #8A1C1C; font-weight: 700;"
    return ""


def _calc_indicador_compras_por_loja(df_f: pd.DataFrame, df_compras: pd.DataFrame, df_devol: pd.DataFrame,
                                    lojas_sel_aplicadas, data_ini, data_fim) -> pd.DataFrame:
    # ===== Base (Vendas) por loja: FATURAMENTO + CMV (coluna T -> CUSTO_T_NUM)
    if df_f is None or df_f.empty:
        base = pd.DataFrame({"LOJA": [], "FATURAMENTO": [], "CMV": []})
    else:
        base = (
            df_f.groupby("LOJA_N", dropna=False)
            .agg(FATURAMENTO=("FAT_LINHA", "sum"), CMV=("CUSTO_T_NUM", "sum"))
            .reset_index()
            .rename(columns={"LOJA_N": "LOJA"})
        )

    # ===== Recorte compras/devoluções (loja + data)
    def _aplicar_recorte_mov(dfm: pd.DataFrame) -> pd.DataFrame:
        if dfm is None or dfm.empty:
            return pd.DataFrame(columns=["LOJA_N", "TOT_DOC_NUM", "DATA"])

        out = dfm.copy()
        if lojas_sel_aplicadas:
            out = out[out["LOJA_N"].isin(lojas_sel_aplicadas)]

        if "DATA" in out.columns and out["DATA"].notna().any():
            out = out[out["DATA"].notna()]
            out = out[(out["DATA"].dt.date >= data_ini) & (out["DATA"].dt.date <= data_fim)]
        return out

    comp_f = _aplicar_recorte_mov(df_compras)
    dev_f = _aplicar_recorte_mov(df_devol)

    comp_by = (
        comp_f.groupby("LOJA_N", dropna=False)["TOT_DOC_NUM"].sum().reset_index().rename(columns={"LOJA_N": "LOJA", "TOT_DOC_NUM": "COMPRAS"})
        if not comp_f.empty else pd.DataFrame({"LOJA": [], "COMPRAS": []})
    )
    dev_by = (
        dev_f.groupby("LOJA_N", dropna=False)["TOT_DOC_NUM"].sum().reset_index().rename(columns={"LOJA_N": "LOJA", "TOT_DOC_NUM": "DEVOLUÇÕES"})
        if not dev_f.empty else pd.DataFrame({"LOJA": [], "DEVOLUÇÕES": []})
    )

    # ===== Consolida tudo por loja (outer para não perder loja que só exista em uma fonte)
    out = base.merge(comp_by, on="LOJA", how="outer").merge(dev_by, on="LOJA", how="outer")

    # Preenche nulos
    for c in ["FATURAMENTO", "CMV", "COMPRAS", "DEVOLUÇÕES"]:
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0.0)

    out["COMPRAS LÍQ"] = out["COMPRAS"] - out["DEVOLUÇÕES"]
    out["MKP VENDAS X CUSTO"] = out.apply(lambda r: _safe_div(r["FATURAMENTO"], r["CMV"]), axis=1)
    out["MKP COMPRAS X VENDAS"] = out.apply(lambda r: _safe_div(r["FATURAMENTO"], r["COMPRAS LÍQ"]), axis=1)
    out["DIF COMPRAS X CMV"] = out["CMV"] - out["COMPRAS LÍQ"]
    out["INDICADOR DE COMPRAS"] = out["MKP COMPRAS X VENDAS"] - out["MKP VENDAS X CUSTO"]

    # Ordena (opcional)
    out = out.sort_values("LOJA", key=lambda s: s.astype(str)).reset_index(drop=True)

    # ===== Linha TOTAL (soma valores e recalcula razões no total)
    total_fat = float(out["FATURAMENTO"].sum()) if len(out) else 0.0
    total_cmv = float(out["CMV"].sum()) if len(out) else 0.0
    total_comp = float(out["COMPRAS"].sum()) if len(out) else 0.0
    total_dev = float(out["DEVOLUÇÕES"].sum()) if len(out) else 0.0
    total_comp_liq = total_comp - total_dev

    total_mkp_vc = _safe_div(total_fat, total_cmv)
    total_mkp_cv = _safe_div(total_fat, total_comp_liq)
    total_dif = total_cmv - total_comp_liq
    total_ind = None
    if (total_mkp_vc is not None) and (total_mkp_cv is not None):
        total_ind = total_mkp_cv - total_mkp_vc

    total_row = pd.DataFrame([{
        "LOJA": "TOTAL",
        "FATURAMENTO": total_fat,
        "CMV": total_cmv,
        "MKP VENDAS X CUSTO": total_mkp_vc,
        "COMPRAS": total_comp,
        "DEVOLUÇÕES": total_dev,
        "COMPRAS LÍQ": total_comp_liq,
        "MKP COMPRAS X VENDAS": total_mkp_cv,
        "DIF COMPRAS X CMV": total_dif,
        "INDICADOR DE COMPRAS": total_ind,
    }])

    out = pd.concat([out, total_row], ignore_index=True)

    # Reordena colunas
    cols = [
        "LOJA",
        "FATURAMENTO",
        "CMV",
        "MKP VENDAS X CUSTO",
        "COMPRAS",
        "DEVOLUÇÕES",
        "COMPRAS LÍQ",
        "MKP COMPRAS X VENDAS",
        "DIF COMPRAS X CMV",
        "INDICADOR DE COMPRAS",
    ]
    out = out[[c for c in cols if c in out.columns]]
    return out


def _render_indicador_compras_drill(df_f: pd.DataFrame, lojas_sel_aplicadas, data_ini, data_fim):
    # Drill (expander) + computação sob demanda para não pesar a abertura do dashboard
    with st.expander("Abrir Indicador de Compras", expanded=False):
        st.caption("Tabela por loja (com TOTAL ao final), usando o mesmo recorte de lojas e período do dashboard.")
        show_tbl = st.checkbox("Calcular/mostrar tabela", value=False, key="show_indicador_compras_tbl")
        if not show_tbl:
            st.info("Marque a opção acima para calcular e exibir a tabela. (Isso deixa o carregamento inicial mais leve.)")
            return

        df_compras, df_devol = carregar_movimentacoes_compras()
        tbl = _calc_indicador_compras_por_loja(df_f, df_compras, df_devol, lojas_sel_aplicadas, data_ini, data_fim)

        sty = (
            tbl.style
            .format(
                {
                    "FATURAMENTO": lambda v: "R$ " + format_brl(v),
                    "CMV": lambda v: "R$ " + format_brl(v),
                    "COMPRAS": lambda v: "R$ " + format_brl(v),
                    "DEVOLUÇÕES": lambda v: "R$ " + format_brl(v),
                    "COMPRAS LÍQ": lambda v: "R$ " + format_brl(v),
                    # DIF é número normal (sem R$)
                    "DIF COMPRAS X CMV": _fmt_num,
                    "MKP VENDAS X CUSTO": _fmt_num,
                    "MKP COMPRAS X VENDAS": _fmt_num,
                    "INDICADOR DE COMPRAS": _fmt_num,
                }
            )
            .pipe(lambda s: styler_map_compat(s, _indicador_style, subset=["INDICADOR DE COMPRAS"]))
        )

        st.dataframe(sty, use_container_width=True, hide_index=True, height=520)


# Render no final do dashboard
_render_indicador_compras_drill(df_f, lojas_sel_aplicadas, data_ini, data_fim)