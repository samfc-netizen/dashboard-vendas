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
import html
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


        /* Navegação multipainel, sem alterar o menu lateral */
        div[data-baseweb="tab-list"] {gap: 6px; border-bottom: 1px solid #e2e8f0; padding-bottom: 8px;}
        button[data-baseweb="tab"] {border-radius: 10px 10px 0 0; padding: 10px 15px; font-weight: 650;}
        button[data-baseweb="tab"][aria-selected="true"] {background: #eaf2ff; color: #0f4c81;}
        div[data-testid="stDataFrame"] {border-radius: 12px; overflow: hidden;}

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
        /* DAUTO BI | identidade visual unificada */
        .stApp {background:#f4f7fc !important; color:#142443;}
        .block-container {max-width:1680px; padding-top:1.1rem; padding-bottom:3rem;}
        h1,h2,h3 {color:#10234a !important; letter-spacing:-.035em; font-weight:760 !important;}
        h3 {font-size:1.28rem !important; margin-top:1.2rem !important;}
        div[data-testid="stMetric"] {border:1px solid #e0e9f5 !important; border-radius:15px !important; box-shadow:0 4px 18px rgba(25,61,110,.045) !important; min-height:100px; padding:16px 18px !important;}
        div[data-testid="stMetricValue"] {color:#10244e !important; font-weight:780 !important;}
        div[data-testid="stMetricLabel"] {color:#61738e !important; font-weight:650 !important;}
        div[data-testid="stDataFrame"] {border:1px solid #dce6f2 !important; border-radius:14px !important; box-shadow:0 4px 18px rgba(25,61,110,.04); background:#fff;}
        div[data-testid="stDataFrame"] [role="columnheader"] {background:#183454 !important; color:#fff !important; font-weight:700 !important;}
        div[data-testid="stDataFrame"] [role="gridcell"] {border-color:#eaf0f7 !important;}
        div[data-testid="stExpander"] {background:#fff; border:1px solid #dce6f2; border-radius:14px !important; overflow:hidden; margin-bottom:12px;}
        div[data-testid="stVerticalBlockBorderWrapper"]:has(> div > div[data-testid="stVerticalBlock"]) {border-radius:16px;}
        div[data-testid="stTabs"] button[data-baseweb="tab"] {border-radius:10px 10px 0 0; font-size:13px; padding:12px 16px;}
        div[data-testid="stTabs"] button[aria-selected="true"] {color:#1261d6 !important; background:#eaf2ff !important;}
        div[data-testid="stPlotlyChart"] {background:#fff; border:1px solid #e0e9f5; border-radius:15px; padding:8px; box-shadow:0 4px 18px rgba(25,61,110,.04);}
        .hero-card {background:linear-gradient(115deg,#123a73 0%,#165ed3 60%,#102448 100%); border-radius:18px; padding:23px 28px; margin-bottom:15px; box-shadow:0 12px 30px rgba(16,58,117,.12);}
        .hero-card h1 {font-size:27px; color:white !important;}
        .hero-card p {font-size:13px;}
        .bi-heading {display:flex; gap:12px; align-items:center; padding:10px 0 14px; margin-top:14px;}
        .bi-icon {display:grid; place-items:center; width:42px; height:42px; border-radius:12px; background:linear-gradient(145deg,#1263df,#0e448f); color:#fff; font-size:22px; font-weight:800; box-shadow:0 5px 12px rgba(24,94,200,.15);}
        .bi-heading h2 {margin:0; font-size:22px; color:#10234a; line-height:1.15;}
        .bi-heading p {margin:4px 0 0; font-size:12px; color:#71819a;}
        .bi-goal {background:white; border:1px solid #e0e9f5; border-radius:16px; padding:20px 22px; box-shadow:0 4px 18px rgba(25,61,110,.045); margin:8px 0 15px;}
        .bi-goal-title {font-size:14px; font-weight:750; color:#15294e; margin-bottom:16px;}
        .bi-goal-values {display:flex; justify-content:space-between; gap:12px; color:#183454; font-size:16px; font-weight:760;}
        .bi-goal-values small {display:block; color:#71819a; font-size:11px; font-weight:500; margin-top:4px;}
        .bi-progress {height:24px; background:#e5edf7; border-radius:99px; overflow:hidden; margin-top:16px;}
        .bi-progress span {height:100%; display:flex; align-items:center; justify-content:center; background:linear-gradient(90deg,#1253b8,#2475ed); border-radius:99px; color:white; font-size:12px; font-weight:750; min-width:0; white-space:nowrap;}
        .bi-scale {display:flex; justify-content:space-between; color:#8a9ab1; font-size:10px; margin-top:9px;}
        .bi-alert {background:#fff; border:1px solid #e4eaf3; border-radius:14px; padding:16px; margin:7px 0; box-shadow:0 4px 18px rgba(25,61,110,.04);}
        .bi-alert-label {font-size:12px; color:#687b97; margin-bottom:6px;}
        .bi-alert-value {font-size:clamp(16px,1.5vw,24px); font-weight:800; color:#c62828; overflow-wrap:anywhere;}
        .bi-alert-value.positive {color:#1565c0;}
        @media(max-width:850px){.block-container{padding-left:1rem;padding-right:1rem}.hero-card{padding:18px}.hero-card h1{font-size:23px}.bi-heading h2{font-size:19px}}

        /* Alternância automática pelo viewport: sem depender de user-agent ou seleção manual. */
.bi-mobile-only {display:none;}
        @media(max-width:767px){.st-key-previsao_desktop{display:none!important;}}
        .bi-desktop-only {display:block;}
        .bi-mobile-grid {display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px;}
        .bi-mobile-store {background:#fff;border:1px solid #dce6f2;border-radius:16px;padding:15px;margin:10px 0;box-shadow:0 4px 15px rgba(25,61,110,.05);}
        .bi-mobile-store-head {display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:12px;}
        .bi-mobile-store-name {font-weight:800;color:#142b52;font-size:15px;}
        .bi-mobile-cell {background:#f5f8fc;padding:10px;border-radius:10px;min-width:0;}
        .bi-mobile-cell small {display:block;color:#65758c;font-size:10px;margin-bottom:4px;}
        .bi-mobile-cell strong {font-size:13px;color:#142b52;overflow-wrap:anywhere;}
        .bi-positive {color:#1565c0!important;}
        .bi-negative {color:#c62828!important;}
        .bi-track {height:7px;background:#e7eef8;border-radius:8px;overflow:hidden;margin-top:6px;}
        .bi-fill {height:100%;border-radius:8px;background:#c62828;}
        .bi-fill.positive {background:#1565c0;}
        .bi-pill {font-size:11px;font-weight:800;padding:5px 9px;border-radius:99px;background:#fff0f0;color:#c62828;white-space:nowrap;}
        .bi-pill.positive {background:#eaf2ff;color:#1565c0;}
        @media(max-width:767px){
          .bi-desktop-only {display:none!important;}
          .bi-mobile-only {display:block!important;}
          .block-container {padding-left:.75rem!important;padding-right:.75rem!important;padding-top:.5rem!important;}
          .hero-card {padding:17px!important;border-radius:15px!important;}
          .hero-card h1 {font-size:20px!important;}
          .hero-card p {font-size:12px!important;}
          div[data-testid="stHorizontalBlock"] {flex-wrap:wrap!important;}
          div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {min-width:min(100%,155px)!important;flex:1 1 155px!important;}
          div[data-testid="stMetric"] {padding:12px!important;min-height:85px!important;}
          div[data-testid="stMetricValue"] {font-size:20px!important;}
          div[data-testid="stTabs"] [data-baseweb="tab-list"] {overflow-x:auto;flex-wrap:nowrap;scrollbar-width:thin;}
          div[data-testid="stTabs"] button[data-baseweb="tab"] {white-space:nowrap;font-size:12px;padding:10px;}
          div[data-testid="stDataFrame"] {max-width:100%;}
          .bi-goal {padding:13px!important;}
          .bi-goal-values {flex-wrap:wrap;font-size:13px!important;}
          .bi-heading h2 {font-size:18px!important;}
        }

        textarea, input { color: #111827 !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def bi_plotly_chart(fig, *args, **kwargs):
    """Aplica visual institucional aos gráficos, sem modificar os dados."""
    try:
        fig.update_layout(
            paper_bgcolor="#ffffff", plot_bgcolor="#ffffff",
            font=dict(family="Arial, sans-serif", color="#344968", size=12),
            title_font=dict(size=17, color="#10234a"),
            legend=dict(bgcolor="rgba(255,255,255,0)"),
            margin=dict(l=25, r=20, t=55, b=35),
        )
        fig.update_xaxes(showgrid=False, linecolor="#e5edf6", tickfont=dict(color="#64748b"))
        fig.update_yaxes(gridcolor="#edf2f8", zerolinecolor="#e5edf6", tickfont=dict(color="#64748b"))
    except Exception:
        pass
    return st.plotly_chart(fig, *args, **kwargs)


def render_header():
    st.markdown(
        """
        <div class="hero-card">
            <h1>DAUTO TINTAS | Inteligência Comercial</h1>
            <p>Performance de vendas, metas, clientes e resultados em uma única plataforma.</p>
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
        def _pwa_write_if_changed(path, content):
            path = Path(path)
            try:
                if path.exists() and path.read_text(encoding="utf-8") == content:
                    return
            except Exception:
                pass
            path.write_text(content, encoding="utf-8")

        _pwa_write_if_changed("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
        _pwa_write_if_changed(
            "sw.js",
            """
self.addEventListener('install', event => { self.skipWaiting(); });
self.addEventListener('activate', event => { event.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', event => { event.respondWith(fetch(event.request)); });
""".strip(),
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
# Bases de vendas organizadas em blocos de 4 meses.
# Cada bloco pode ser dividido em arquivos complementares quando ficar grande:
# BASE 1.xlsx, BASE 1-1.xlsx, BASE 1-2.xlsx...
# BASE 2.xlsx, BASE 2-1.xlsx, BASE 2-2.xlsx... etc.
# A descoberta dos arquivos é automática; não é necessário cadastrar novos nomes aqui.
ARQUIVOS_BASE_SEPARADOS = []
ARQUIVO_METAS = "METAS.xlsx"
ARQUIVO_COMPRAS = "COMPRAS.xlsx"


# Cache maior para evitar releitura pesada dos Excel a cada interação.
# Use o botão "Recarregar agora" no menu lateral quando trocar as planilhas.
CACHE_TTL_SEGUNDOS = 60 * 60

COLUNAS_CATEGORIA_VENDAS = [
    "LOJA_N", "VENDEDOR_N", "MARCA_N", "SEGMENTO_N", "LINHA_N", "LOJA_KEY"
]


def _write_text_if_changed(path: str | Path, content: str, encoding: str = "utf-8") -> None:
    """Evita regravar arquivos estáticos em todo rerun do Streamlit."""
    path = Path(path)
    try:
        if path.exists() and path.read_text(encoding=encoding) == content:
            return
    except Exception:
        pass
    path.write_text(content, encoding=encoding)


def _read_excel_fast(source, sheet_name=0) -> pd.DataFrame:
    """Leitura Excel com parâmetros mais leves.

    na_filter=False reduz o custo de detecção automática de NaN em bases grandes.
    A normalização numérica/data já é feita manualmente nas funções do app.
    """
    return pd.read_excel(source, sheet_name=sheet_name, engine="openpyxl", na_filter=False)


def _otimizar_memoria_vendas(df: pd.DataFrame) -> pd.DataFrame:
    """Reduz memória sem alterar as colunas usadas nas telas."""
    if df is None or df.empty:
        return df
    for col in COLUNAS_CATEGORIA_VENDAS:
        if col in df.columns:
            try:
                df[col] = df[col].astype("category")
            except Exception:
                pass
    for col in ["QTD_NUM", "UNIT_NUM", "FAT_LINHA", "VR_TOTAL_NUM", "CUSTO_ST_NUM", "CUSTO_T_NUM"]:
        if col in df.columns:
            try:
                df[col] = pd.to_numeric(df[col], errors="coerce").astype("float32")
            except Exception:
                pass
    if "_ORDEM_LINHA" in df.columns:
        try:
            df["_ORDEM_LINHA"] = pd.to_numeric(df["_ORDEM_LINHA"], errors="coerce").fillna(0).astype("int32")
        except Exception:
            pass
    return df


def _otimizar_memoria_mov(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    for col in ["LOJA_N", "LOJA_KEY"]:
        if col in df.columns:
            try:
                df[col] = df[col].astype("category")
            except Exception:
                pass
    if "TOT_DOC_NUM" in df.columns:
        try:
            df["TOT_DOC_NUM"] = pd.to_numeric(df["TOT_DOC_NUM"], errors="coerce").astype("float32")
        except Exception:
            pass
    return df


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


def carregar_base_consolidada() -> pd.DataFrame:
    """Carrega e consolida automaticamente todas as bases de vendas.

    Convenção recomendada (4 meses por bloco):
    - BASE 1.xlsx = Jan-Abr
    - BASE 2.xlsx = Mai-Ago
    - BASE 3.xlsx = Set-Dez

    Se um bloco ficar grande, podem existir complementos, por exemplo:
    BASE 1-1.xlsx, BASE 1-2.xlsx, BASE 2-1.xlsx, BASE 2-2.xlsx etc.

    Todos os arquivos que seguirem o padrão BASE <n>.xlsx ou BASE <n>-<n>.xlsx
    serão encontrados e concatenados automaticamente em ordem numérica.

    Fallbacks mantidos:
    1) BASE.xlsx
    2) primeiro .xlsx encontrado dentro de BASE.zip
    """
    pasta = Path(".")

    def chave_ordenacao_base(path: Path):
        # Aceita BASE 1.xlsx, BASE1.xlsx, BASE 1-1.xlsx, BASE 2-2.xlsx etc.
        nome = path.stem.strip().upper()
        m = re.match(r"^BASE\s*(\d+)(?:\s*-\s*(\d+))?$", nome)
        if not m:
            return (10_000, 10_000, nome)
        bloco = int(m.group(1))
        complemento = int(m.group(2) or 0)
        return (bloco, complemento, nome)

    arquivos_base = []
    for p in pasta.glob("*.xlsx"):
        if p.name.startswith("~$"):
            continue
        nome = p.stem.strip().upper()
        if re.match(r"^BASE\s*\d+(?:\s*-\s*\d+)?$", nome):
            arquivos_base.append(p)

    arquivos_base = sorted(arquivos_base, key=chave_ordenacao_base)

    if arquivos_base:
        dfs = []
        for caminho in arquivos_base:
            dfx = _read_excel_fast(caminho, sheet_name=0)
            dfx["_ARQUIVO_BASE"] = caminho.name
            dfs.append(dfx)
        return pd.concat(dfs, ignore_index=True)

    caminho_base_unica = resolver_arquivo_excel(["BASE.xlsx"])
    if caminho_base_unica is not None:
        df_unica = _read_excel_fast(caminho_base_unica, sheet_name=0)
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
                    df_zip = _read_excel_fast(io.BytesIO(f.read()), sheet_name=0)
                    df_zip["_ARQUIVO_BASE"] = f"{ARQUIVO_ZIP}/{Path(nome_excel).name}"
                    return df_zip

    st.error(
        "Nenhuma base de vendas foi encontrada. Use arquivos como BASE 1.xlsx, BASE 1-1.xlsx, "
        "BASE 2.xlsx, BASE 2-1.xlsx etc. Também aceito BASE.xlsx ou BASE.zip como fallback."
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
    return _read_excel_fast(caminho, sheet_name=sheet_name)


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

MES_NOME_TO_NUM = {nome: num for num, nome in MESES}
MES_NUM_TO_NOME = {num: nome for num, nome in MESES}


def meses_no_periodo(data_ini_val, data_fim_val) -> list[int]:
    """Meses existentes no intervalo de datas selecionado."""
    try:
        ini = pd.Timestamp(data_ini_val).to_period("M")
        fim = pd.Timestamp(data_fim_val).to_period("M")
        if fim < ini:
            ini, fim = fim, ini
        meses = [int(p.month) for p in pd.period_range(ini, fim, freq="M")]
        return [m for m in range(1, 13) if m in set(meses)]
    except Exception:
        return [num for num, _ in MESES]

# Ordem desejada (UNAÍ por último)
LOJA_KEY_ORDER = ["ADE", "GAMA", "SOFNORTE", "CEILANDIA", "SIA", "AGLINDAS", "GUARA", "LUZIANIA", "UNAI"]
LOJA_KEY_RANK = {k: i for i, k in enumerate(LOJA_KEY_ORDER)}
DEFAULT_RANK = 10_000


def tabela_previsao_mobile(df):
    """Cards mobile com os mesmos números calculados para a tabela desktop."""
    def moeda(v):
        return "R$ " + format_brl(v) if pd.notna(v) else "—"
    def pct(v):
        return (format_decimal_pt(v, 2) + "%") if pd.notna(v) else "—"
    def celula(label, value, cls="", barra=None):
        track = ""
        if barra is not None and pd.notna(barra):
            width = max(0, min(100, float(barra)))
            track = f'<div class="bi-track"><div class="bi-fill {"positive" if float(barra)>=100 else ""}" style="width:{width:.2f}%"></div></div>'
        return f'<div class="bi-mobile-cell"><small>{html.escape(label)}</small><strong class="{cls}">{html.escape(str(value))}</strong>{track}</div>'
    blocks=[]
    for _, r in df.iterrows():
        ating = r['ATING_META_%']
        positive = pd.notna(ating) and ating >= 100
        pill = 'positive' if positive else ''
        dif_meta = r['DIF_PREV_X_META_R$']
        dif_ano = r['DIF_PREV_X_ANO1_R$']
        crescimento = r['CRESC_%']
        def classe(v):
            return 'bi-positive' if pd.notna(v) and v > 0 else 'bi-negative' if pd.notna(v) and v < 0 else ''
        cells = [
            celula('Realizado', moeda(r['REALIZADO_R$'])),
            celula('Previsão', moeda(r['PREVISAO_R$'])),
            celula('Meta', moeda(r['META_R$'])),
            celula('Atingimento', pct(ating), 'bi-positive' if positive else 'bi-negative', ating),
            celula('Ano anterior', moeda(r['ANO_1_R$'])),
            celula('Crescimento', pct(crescimento), classe(crescimento)),
            celula('Diferença x meta', moeda(dif_meta), classe(dif_meta)),
            celula('Diferença x ano anterior', moeda(dif_ano), classe(dif_ano)),
            celula('Média por dia', moeda(r['MEDIA_DIA_R$'])),
            celula('Dias de venda', format_decimal_pt(r['DIAS_VENDA'],1)),
        ]
        blocks.append(f'<div class="bi-mobile-store"><div class="bi-mobile-store-head"><span class="bi-mobile-store-name">{html.escape(str(r["LOJA"]))}</span><span class="bi-pill {pill}">{pct(ating)} da meta</span></div><div class="bi-mobile-grid">{"".join(cells)}</div></div>')
    st.markdown('<div class="bi-mobile-only">'+''.join(blocks)+'</div>', unsafe_allow_html=True)


def color_pos_neg(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    try:
        x = float(v)
    except Exception:
        return ""
    if x > 0:
        return "color: #1565c0; font-weight: 700;"  # azul
    if x < 0:
        return "color: #c62828; font-weight: 700;"  # vermelho
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
        return ["background-color: #e9f1ff; color: #10234a; font-weight: 900;"] * len(row)
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
@st.cache_resource(ttl=CACHE_TTL_SEGUNDOS, show_spinner="Carregando base de vendas...")
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

    return _otimizar_memoria_vendas(df)


@st.cache_resource(ttl=CACHE_TTL_SEGUNDOS, show_spinner="Carregando compras e devoluções...")
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

    return _otimizar_memoria_mov(df_compras), _otimizar_memoria_mov(df_devol)


@st.cache_resource(ttl=CACHE_TTL_SEGUNDOS, show_spinner="Carregando metas e referências...")
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
        mask_mov = (
            df_base["DATA"].notna()
            & (df_base["DATA"].dt.year == ano)
            & (df_base["DATA"].dt.month == mes)
        )
        mov = df_base.loc[mask_mov, ["DATA", valor_col]].copy()
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

    modo = st.radio(
        "Tipo de relatório",
        ["Diário", "Geral / Parcial do mês"],
        horizontal=True,
        key="tipo_relatorio_whatsapp",
    )

    if modo == "Diário":
        data_padrao = datas_validas_rel.max().date()
        data_rel = st.date_input("Data do relatório diário", value=data_padrao, min_value=datas_validas_rel.min().date(), max_value=datas_validas_rel.max().date())
        data_rel_ts = pd.Timestamp(data_rel)
        mask_dia = df_base["DATA"].notna() & (df_base["DATA"].dt.normalize() == data_rel_ts)
        cols_rel = [
            c for c in [
                "DATA", "LOJA_KEY", "LOJA_N", "FAT_LINHA", valor_col,
                "SEGMENTO_N", "MARCA_N", "CLIENTE_N"
            ] if c in df_base.columns
        ]
        df_rel = df_base.loc[mask_dia, cols_rel].copy()
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
        # Renderiza apenas uma loja por vez. Criar componentes para todas as lojas em
        # cada rerun aumenta muito o uso de memória do navegador e do Streamlit.
        blocos_map = {k: (loja_nome, bloco) for k, loja_nome, bloco in blocos}
        if blocos_map:
            loja_detalhe_key = st.selectbox(
                "Selecionar loja para visualizar/enviar",
                options=list(blocos_map.keys()),
                format_func=lambda k: blocos_map[k][0],
                key="loja_detalhe_relatorio_diario",
            )
            loja_nome, bloco = blocos_map[loja_detalhe_key]
            cols = st.columns([1, 1])
            numero = TELEFONES_WHATSAPP_LOJAS.get(loja_detalhe_key)
            if numero:
                with cols[0]:
                    _botao_whatsapp(f"Abrir WhatsApp - {loja_nome}", numero, bloco, key=f"wa_daily_{loja_detalhe_key}")
            with cols[1]:
                st.download_button(
                    f"Baixar {loja_nome}.txt",
                    bloco.encode("utf-8"),
                    file_name=f"relatorio_diario_{loja_detalhe_key}_{data_rel.strftime('%Y%m%d')}.txt",
                    mime="text/plain",
                    key=f"down_daily_{loja_detalhe_key}",
                )
            st.text_area(f"Texto {loja_nome}", bloco, height=260, key=f"txt_daily_{loja_detalhe_key}")
        return

    # Relatório parcial/mensal
    data_max_rel = datas_validas_rel.max().date()
    mes_padrao = data_max_rel.month
    mes = st.selectbox("Mês do relatório", options=[m for m, _ in MESES], index=mes_padrao - 1, format_func=lambda m: f"{m:02d} - {MESES[m-1][1]}")
    import calendar
    data_ini_rel = date(ANO_ATUAL, mes, 1)
    ultimo_dia_mes = date(ANO_ATUAL, mes, calendar.monthrange(ANO_ATUAL, mes)[1])
    data_max_mes = min(data_max_rel, ultimo_dia_mes)
    # Se o mês selecionado ainda não existe na base, evita date_input com limites inválidos.
    if data_max_mes < data_ini_rel:
        st.warning("O mês selecionado ainda não possui dados disponíveis na base.")
        return
    chave_data_mes = f"data_fim_relatorio_{ANO_ATUAL}_{mes:02d}"
    valor_data_mes = st.session_state.get(chave_data_mes, data_max_mes)
    if valor_data_mes < data_ini_rel or valor_data_mes > data_max_mes:
        valor_data_mes = data_max_mes
        st.session_state[chave_data_mes] = valor_data_mes
    data_fim_rel = st.date_input(
        "Data final do relatório parcial",
        value=valor_data_mes,
        min_value=data_ini_rel,
        max_value=data_max_mes,
        key=chave_data_mes,
    )
    data_ini_rel_ts = pd.Timestamp(data_ini_rel)
    data_fim_rel_ts = pd.Timestamp(data_fim_rel) + pd.Timedelta(days=1)
    mask_rel = (
        df_base["DATA"].notna()
        & (df_base["DATA"] >= data_ini_rel_ts)
        & (df_base["DATA"] < data_fim_rel_ts)
    )
    # Seleciona somente as colunas necessárias. Evita copiar a base inteira ao trocar o filtro.
    cols_rel = [
        c for c in [
            "DATA", "LOJA_KEY", "LOJA_N", "FAT_LINHA", valor_col,
            "SEGMENTO_N", "MARCA_N", "CLIENTE_N"
        ] if c in df_base.columns
    ]
    df_rel = df_base.loc[mask_rel, cols_rel].copy()
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
    blocos_map = {k: (loja_nome, bloco) for k, loja_nome, bloco in blocos}
    if blocos_map:
        loja_detalhe_key = st.selectbox(
            "Selecionar loja para visualizar/enviar",
            options=list(blocos_map.keys()),
            format_func=lambda k: blocos_map[k][0],
            key="loja_detalhe_relatorio_mensal",
        )
        loja_nome, bloco = blocos_map[loja_detalhe_key]
        cols = st.columns([1, 1])
        numero = TELEFONES_WHATSAPP_LOJAS.get(loja_detalhe_key)
        if numero:
            with cols[0]:
                _botao_whatsapp(
                    f"Abrir WhatsApp - {loja_nome}",
                    numero,
                    bloco,
                    key=f"wa_month_{loja_detalhe_key}_{mes:02d}",
                )
        with cols[1]:
            st.download_button(
                f"Baixar {loja_nome}.txt",
                bloco.encode("utf-8"),
                file_name=f"relatorio_mensal_{loja_detalhe_key}_{ANO_ATUAL}_{mes:02d}.txt",
                mime="text/plain",
                key=f"down_month_{loja_detalhe_key}_{mes:02d}",
            )
        st.text_area(
            f"Texto {loja_nome}",
            bloco,
            height=260,
            key=f"txt_month_{loja_detalhe_key}_{mes:02d}",
        )


# =========================================================
# CAMPANHAS
# =========================================================
# Para cadastrar/alterar campanhas futuras, centralize as regras neste bloco
# e crie/ajuste a respectiva função de renderização.
CAMPANHAS = {
    "SUPER_SETEMBRO_2026": {
        "nome": "Super Setembro",
        "inicio": date(2026, 7, 1),
        "fim": date(2026, 9, 30),
        "lojas": ["ADE", "GAMA", "SOFNORTE", "CEILANDIA", "LUZIANIA"],
        "meta_desbloqueio": 525000.00,
        "metas_gerentes": {
            "ADE": 75000.00,
            "GAMA": 75000.00,
            "SOFNORTE": 45000.00,
            "CEILANDIA": 105000.00,
            "LUZIANIA": 300000.00,
        },
        "premio_gerente_pct": 0.005,
        "marca_vendedores": "CORAL",
        "meta_vendedor": 35000.00,
        "premio_vendedor_pct": 0.01,
    }
}


CAMPANHA_LOJA_NOMES = {
    "ADE": "ADE",
    "GAMA": "Gama",
    "SOFNORTE": "Sofnorte",
    "CEILANDIA": "Ceilândia",
    "LUZIANIA": "Luziânia",
}


def classificar_segmento_gerencial(seg):
    """Classificação oficial usada no drill Faturamento por Loja x Segmento Gerencial."""
    key = canonical_key(seg)

    if key in {"", "SEMSEGMENTO", "NAN", "NONE"}:
        return "Sem Segmento"

    # Decor/Industrial gerencial = Imobiliário + Decorativo + Industrial + Moveleira.
    # A checagem vem antes de Automotivo para tratar cadastros híbridos com Moveleira.
    if (
        "IMOBILI" in key
        or "DECOR" in key
        or "INDUSTR" in key
        or "MOVELEIR" in key
        or "MOVEL" in key
    ):
        return "Decor/Industrial"

    if "THINNER" in key or "TINER" in key:
        return "Thinner"

    if "AUTOMOT" in key:
        return "Automotivo"

    return "Outros Segmentos"


def _mask_decor_industrial_campanha(df: pd.DataFrame) -> pd.Series:
    """Usa exatamente a mesma regra do Segmento Gerencial do dashboard.

    Para o Super Setembro, entram Imobiliário/Decorativo, Industrial e Moveleira,
    conforme a classificação de SEGMENTO_N exibida no drill
    "Faturamento por Loja x Segmento Gerencial".
    """
    if df is None or df.empty:
        return pd.Series(False, index=getattr(df, "index", None), dtype=bool)
    if "SEGMENTO_N" not in df.columns:
        return pd.Series(False, index=df.index, dtype=bool)
    return df["SEGMENTO_N"].apply(classificar_segmento_gerencial).eq("Decor/Industrial")


def _status_campanha(cfg: dict) -> tuple[str, str]:
    hoje = date.today()
    if hoje < cfg["inicio"]:
        return "AGUARDANDO INÍCIO", "⚪"
    if hoje <= cfg["fim"]:
        return "EM ANDAMENTO", "🟠"
    return "FINALIZADA", "🟢"


def _fmt_pct_campanha(v: float) -> str:
    try:
        return f"{float(v):.1f}%".replace(".", ",")
    except Exception:
        return "—"


def _montar_texto_whatsapp_super_setembro(
    cfg: dict,
    status_nome: str,
    data_referencia: date,
    desbloqueado: bool,
    realizado_desbloqueio: float,
    tabela_gerentes: pd.DataFrame,
    tabela_vendedores: pd.DataFrame,
) -> str:
    titulo_resultado = "RESULTADO FINAL" if status_nome == "FINALIZADA" else "RESULTADO PARCIAL"
    falta_desbloqueio = max(cfg["meta_desbloqueio"] - realizado_desbloqueio, 0.0)
    pct_desbloqueio = (realizado_desbloqueio / cfg["meta_desbloqueio"] * 100.0) if cfg["meta_desbloqueio"] else 0.0

    linhas = [
        f"🚀 *SUPER SETEMBRO | {titulo_resultado}*",
        f"📅 Dados até {data_referencia.strftime('%d/%m/%Y')}",
        "",
        "*🔓 DESBLOQUEIO GERAL*",
        f"Realizado: *R$ {format_brl(realizado_desbloqueio)}*",
        f"Meta: *R$ {format_brl(cfg['meta_desbloqueio'])}*",
        f"Atingimento: *{_fmt_pct_campanha(pct_desbloqueio)}*",
    ]
    if desbloqueado:
        linhas.append("Status: ✅ *Campanha desbloqueada*")
    else:
        linhas.append(f"Falta: *R$ {format_brl(falta_desbloqueio)}*")
        linhas.append("Status: 🔒 *Premiações ainda bloqueadas*")

    linhas += ["", "*👔 GERENTES | Decor/Industrial gerencial Q3*"]
    for _, r in tabela_gerentes.iterrows():
        loja = str(r["LOJA"])
        realizado = float(r["REALIZADO"])
        meta = float(r["META"])
        pct = float(r["ATINGIMENTO"])
        falta = max(meta - realizado, 0.0)
        bateu = bool(r["META_ATINGIDA"])
        premio = float(r["PREMIO_CALCULADO"])
        status_meta = "✅ Meta atingida" if bateu else f"⏳ Falta R$ {format_brl(falta)}"
        linhas.append(
            f"• *{loja}*: R$ {format_brl(realizado)} / R$ {format_brl(meta)} "
            f"({_fmt_pct_campanha(pct)}) — {status_meta}"
        )
        if bateu:
            if desbloqueado:
                linhas.append(f"  Premiação: *R$ {format_brl(premio)}* ✅")
            else:
                linhas.append(f"  Prêmio calculado: *R$ {format_brl(premio)}* — aguardando desbloqueio geral")

    linhas += ["", "*🎯 VENDEDORES | Coral Q3*", f"Meta individual: *R$ {format_brl(cfg['meta_vendedor'])}*"]
    if tabela_vendedores.empty:
        linhas.append("Sem vendas Coral registradas nas cinco lojas da campanha.")
    else:
        for _, r in tabela_vendedores.iterrows():
            nome = str(r["VENDEDOR"])
            loja = str(r["LOJA"])
            realizado = float(r["REALIZADO"])
            pct = float(r["ATINGIMENTO"])
            falta = max(cfg["meta_vendedor"] - realizado, 0.0)
            bateu = bool(r["META_ATINGIDA"])
            premio = float(r["PREMIO_CALCULADO"])
            status_meta = "✅ Meta atingida" if bateu else f"⏳ Falta R$ {format_brl(falta)}"
            linhas.append(
                f"• *{nome}* ({loja}): R$ {format_brl(realizado)} "
                f"({_fmt_pct_campanha(pct)}) — {status_meta}"
            )
            if bateu:
                if desbloqueado:
                    linhas.append(f"  Premiação: *R$ {format_brl(premio)}* ✅")
                else:
                    linhas.append(f"  Prêmio calculado: *R$ {format_brl(premio)}* — aguardando desbloqueio geral")

    linhas += [
        "",
        "_Premiações somente são liberadas após o atingimento coletivo de R$ 525.000,00 no segmento gerencial Decor/Industrial (Imobiliário/Decorativo + Industrial + Moveleira) nas cinco lojas da campanha._",
    ]
    return "\n".join(linhas)


def _botao_copiar_texto_campanha(texto: str, key: str = "copiar_campanha"):
    safe_id = re.sub(r"[^a-zA-Z0-9_-]", "_", str(key))
    texto_js = json.dumps(str(texto or ""), ensure_ascii=False)
    st.components.v1.html(
        f"""
        <div style="font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;">
            <button id="btn_{safe_id}" style="
                width: 100%; border: 0; border-radius: 14px; padding: 13px 16px;
                background: linear-gradient(135deg, #16a34a 0%, #15803d 100%);
                color: #fff; font-weight: 800; font-size: 15px; cursor: pointer;
                box-shadow: 0 8px 20px rgba(22,163,74,.22);">
                📋 Copiar texto para WhatsApp
            </button>
            <div id="status_{safe_id}" style="margin-top:8px;color:#334155;font-size:12px;"></div>
        </div>
        <script>
        const texto_{safe_id} = {texto_js};
        document.getElementById('btn_{safe_id}').addEventListener('click', async function() {{
            const status = document.getElementById('status_{safe_id}');
            try {{
                await navigator.clipboard.writeText(texto_{safe_id});
                status.innerHTML = '✅ Texto copiado. Agora é só colar no WhatsApp.';
            }} catch (e) {{
                status.innerHTML = '⚠️ O navegador bloqueou a cópia automática. Use o campo de texto abaixo.';
            }}
        }});
        </script>
        """,
        height=82,
    )


def render_super_setembro(df_base: pd.DataFrame | None = None):
    cfg = CAMPANHAS["SUPER_SETEMBRO_2026"]
    status_nome, status_emoji = _status_campanha(cfg)

    # IMPORTANTE: campanhas usam SEMPRE a base completa de vendas,
    # independentemente dos filtros de mês/período/loja/vendedor do Dashboard.
    # Isso garante que o Super Setembro considere todo o Q3 (jul + ago + set).
    df_base_campanha = carregar_dados().copy()
    df_base_campanha = df_base_campanha.loc[df_base_campanha["FAT_LINHA"].notna()].copy()

    st.markdown("## 🏆 Campanhas")
    st.markdown(
        f"""
        <div class="section-card">
            <div style="font-size:12px;font-weight:800;color:#b45309;letter-spacing:.08em;">CAMPANHA DE INCENTIVO | Q3 2026</div>
            <div style="font-size:28px;font-weight:900;color:#0f172a;margin-top:4px;">SUPER SETEMBRO</div>
            <div style="color:#475569;margin-top:4px;">A arrancada final para fecharmos o Q3 com força nos segmentos Decorativo, Industrial e Moveleira.</div>
            <div style="margin-top:12px;font-weight:800;">{status_emoji} {status_nome}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Recorte FIXO da campanha: 01/07/2026 a 30/09/2026.
    # Não depende de nenhum filtro selecionado na barra lateral do Dashboard.
    df_c = df_base_campanha[df_base_campanha["DATA"].notna()].copy()
    ini_ts = pd.Timestamp(cfg["inicio"])
    fim_exclusivo = pd.Timestamp(cfg["fim"]) + pd.Timedelta(days=1)
    df_c = df_c[(df_c["DATA"] >= ini_ts) & (df_c["DATA"] < fim_exclusivo)].copy()
    df_c = df_c[df_c["LOJA_KEY"].isin(cfg["lojas"])].copy()

    st.caption(
        "Base da campanha: Q3 completo (01/07/2026 a 30/09/2026), "
        "independentemente dos filtros do Dashboard. O resultado usa exatamente o mesmo "
        "Segmento Gerencial do drill Faturamento por Loja x Segmento Gerencial: "
        "Imobiliário/Decorativo + Industrial + Moveleira."
    )

    if not df_c.empty:
        max_data_base = pd.to_datetime(df_c["DATA"], errors="coerce").dropna().max().date()
        data_referencia = min(max_data_base, cfg["fim"])
    else:
        data_referencia = min(date.today(), cfg["fim"])

    # Decor/Industrial gerencial: base da condição coletiva e campanha dos gerentes.
    # Mesma regra do drill Faturamento por Loja x Segmento Gerencial:
    # Imobiliário/Decorativo + Industrial + Moveleira, usando SEGMENTO_N.
    mask_di = _mask_decor_industrial_campanha(df_c)
    df_di = df_c[mask_di].copy()
    realizado_desbloqueio = float(df_di["FAT_LINHA"].sum()) if len(df_di) else 0.0
    desbloqueado = realizado_desbloqueio >= cfg["meta_desbloqueio"]
    pct_desbloqueio = (realizado_desbloqueio / cfg["meta_desbloqueio"] * 100.0) if cfg["meta_desbloqueio"] else 0.0
    falta_desbloqueio = max(cfg["meta_desbloqueio"] - realizado_desbloqueio, 0.0)

    st.markdown("### Condição para desbloqueio")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Realizado Decor/Industrial", "R$ " + format_brl(realizado_desbloqueio))
    c2.metric("Meta coletiva", "R$ " + format_brl(cfg["meta_desbloqueio"]))
    c3.metric("Atingimento", _fmt_pct_campanha(pct_desbloqueio))
    c4.metric("Saldo", "R$ " + format_brl(falta_desbloqueio))
    st.progress(min(max(pct_desbloqueio / 100.0, 0.0), 1.0))
    if desbloqueado:
        st.success("✅ Campanha desbloqueada: as premiações podem ser liberadas para quem atingir a regra individual.")
    else:
        st.warning(f"🔒 Premiações ainda bloqueadas. Faltam R$ {format_brl(falta_desbloqueio)} para o grupo das cinco lojas atingir R$ 525.000,00.")

    # -------------------- Gerentes --------------------
    st.markdown("### Campanha 1 | Gerentes")
    st.caption("Meta acumulada do segmento gerencial Decor/Industrial (Imobiliário/Decorativo + Industrial + Moveleira) por loja no Q3. Ao atingir a meta, o gerente recebe 0,5% sobre todo esse faturamento da loja no trimestre, condicionado ao desbloqueio geral.")

    # Consolidação explícita mês a mês. Isso impede que qualquer seleção de mês do
    # dashboard interfira no Q3 e também deixa visível a composição do acumulado.
    if len(df_di):
        df_di["MES_CAMPANHA"] = pd.to_datetime(df_di["DATA"], errors="coerce").dt.month
        realizado_loja_mes = (
            df_di.groupby(["LOJA_KEY", "MES_CAMPANHA"], dropna=False)["FAT_LINHA"]
            .sum()
            .to_dict()
        )
    else:
        realizado_loja_mes = {}

    rows_g = []
    for loja_key, meta in cfg["metas_gerentes"].items():
        julho = float(realizado_loja_mes.get((loja_key, 7), 0.0) or 0.0)
        agosto = float(realizado_loja_mes.get((loja_key, 8), 0.0) or 0.0)
        setembro = float(realizado_loja_mes.get((loja_key, 9), 0.0) or 0.0)
        realizado = julho + agosto + setembro
        ating = (realizado / meta * 100.0) if meta else 0.0
        falta = max(meta - realizado, 0.0)
        bateu = realizado >= meta
        premio = realizado * cfg["premio_gerente_pct"] if bateu else 0.0
        if bateu and desbloqueado:
            status_premio = "✅ Liberado"
        elif bateu:
            status_premio = "🔒 Aguardando desbloqueio"
        else:
            status_premio = "—"
        rows_g.append({
            "LOJA": CAMPANHA_LOJA_NOMES.get(loja_key, loja_key),
            "JULHO": julho,
            "AGOSTO": agosto,
            "SETEMBRO": setembro,
            "META": meta,
            "REALIZADO": realizado,
            "ATINGIMENTO": ating,
            "FALTA": falta,
            "META_ATINGIDA": bateu,
            "PREMIO_CALCULADO": premio,
            "STATUS_PREMIO": status_premio,
        })
    tbl_g = pd.DataFrame(rows_g)
    tbl_g_disp = tbl_g.copy()
    tbl_g_disp["Julho"] = tbl_g_disp["JULHO"].apply(lambda v: "R$ " + format_brl(v))
    tbl_g_disp["Agosto"] = tbl_g_disp["AGOSTO"].apply(lambda v: "R$ " + format_brl(v))
    tbl_g_disp["Setembro"] = tbl_g_disp["SETEMBRO"].apply(lambda v: "R$ " + format_brl(v))
    tbl_g_disp["Meta Q3"] = tbl_g_disp["META"].apply(lambda v: "R$ " + format_brl(v))
    tbl_g_disp["Realizado Q3"] = tbl_g_disp["REALIZADO"].apply(lambda v: "R$ " + format_brl(v))
    tbl_g_disp["Atingimento"] = tbl_g_disp["ATINGIMENTO"].apply(_fmt_pct_campanha)
    tbl_g_disp["Falta"] = tbl_g_disp["FALTA"].apply(lambda v: "R$ " + format_brl(v))
    tbl_g_disp["Situação"] = tbl_g_disp["META_ATINGIDA"].map({True: "✅ Meta atingida", False: "⏳ Em andamento"})
    tbl_g_disp["Prêmio calculado"] = tbl_g_disp["PREMIO_CALCULADO"].apply(lambda v: "R$ " + format_brl(v) if v > 0 else "—")
    tbl_g_disp["Status prêmio"] = tbl_g_disp["STATUS_PREMIO"]
    st.dataframe(
        tbl_g_disp[["LOJA", "Julho", "Agosto", "Setembro", "Meta Q3", "Realizado Q3", "Atingimento", "Falta", "Situação", "Prêmio calculado", "Status prêmio"]].rename(columns={"LOJA": "Loja"}),
        use_container_width=True,
        hide_index=True,
    )

    # -------------------- Vendedores Coral --------------------
    st.markdown("### Campanha 2 | Vendedores Coral")
    st.caption("Campanha individual nas cinco lojas do Super Setembro. Meta: R$ 35.000,00 em Coral no Q3. Ao atingir, recebe 1% sobre todo o faturamento Coral do trimestre, condicionado ao desbloqueio geral.")

    marca_alvo = canonical_key(cfg["marca_vendedores"])
    mask_coral = df_c["MARCA_N"].astype(str).apply(lambda x: marca_alvo in canonical_key(x))
    df_coral = df_c[mask_coral].copy()
    rows_v = []
    if not df_coral.empty:
        agr = (
            df_coral.groupby("VENDEDOR_N", dropna=False)
            .agg(
                REALIZADO=("FAT_LINHA", "sum"),
                LOJAS=("LOJA_N", lambda s: " / ".join(sorted(set(str(x) for x in s if str(x).strip())))),
            )
            .reset_index()
        )
        for _, r in agr.iterrows():
            realizado = float(r["REALIZADO"] or 0.0)
            ating = (realizado / cfg["meta_vendedor"] * 100.0) if cfg["meta_vendedor"] else 0.0
            falta = max(cfg["meta_vendedor"] - realizado, 0.0)
            bateu = realizado >= cfg["meta_vendedor"]
            premio = realizado * cfg["premio_vendedor_pct"] if bateu else 0.0
            if bateu and desbloqueado:
                status_premio = "✅ Liberado"
            elif bateu:
                status_premio = "🔒 Aguardando desbloqueio"
            else:
                status_premio = "—"
            rows_v.append({
                "VENDEDOR": str(r["VENDEDOR_N"]),
                "LOJA": str(r["LOJAS"]),
                "REALIZADO": realizado,
                "ATINGIMENTO": ating,
                "FALTA": falta,
                "META_ATINGIDA": bateu,
                "PREMIO_CALCULADO": premio,
                "STATUS_PREMIO": status_premio,
            })
    tbl_v = pd.DataFrame(rows_v, columns=["VENDEDOR", "LOJA", "REALIZADO", "ATINGIMENTO", "FALTA", "META_ATINGIDA", "PREMIO_CALCULADO", "STATUS_PREMIO"])
    if not tbl_v.empty:
        tbl_v = tbl_v.sort_values(["REALIZADO", "VENDEDOR"], ascending=[False, True]).reset_index(drop=True)
        tbl_v_disp = tbl_v.copy()
        tbl_v_disp["Coral Q3"] = tbl_v_disp["REALIZADO"].apply(lambda v: "R$ " + format_brl(v))
        tbl_v_disp["Meta"] = "R$ " + format_brl(cfg["meta_vendedor"])
        tbl_v_disp["Atingimento"] = tbl_v_disp["ATINGIMENTO"].apply(_fmt_pct_campanha)
        tbl_v_disp["Falta"] = tbl_v_disp["FALTA"].apply(lambda v: "R$ " + format_brl(v))
        tbl_v_disp["Situação"] = tbl_v_disp["META_ATINGIDA"].map({True: "✅ Meta atingida", False: "⏳ Em andamento"})
        tbl_v_disp["Prêmio calculado"] = tbl_v_disp["PREMIO_CALCULADO"].apply(lambda v: "R$ " + format_brl(v) if v > 0 else "—")
        tbl_v_disp["Status prêmio"] = tbl_v_disp["STATUS_PREMIO"]
        st.dataframe(
            tbl_v_disp[["VENDEDOR", "LOJA", "Coral Q3", "Meta", "Atingimento", "Falta", "Situação", "Prêmio calculado", "Status prêmio"]].rename(columns={"VENDEDOR": "Vendedor", "LOJA": "Loja"}),
            use_container_width=True,
            hide_index=True,
            height=min(640, 80 + 35 * max(len(tbl_v_disp), 1)),
        )
    else:
        st.info("Nenhuma venda Coral encontrada no período da campanha para as cinco lojas participantes.")

    # -------------------- Exportação WhatsApp --------------------
    st.markdown("### 📲 Exportar parcial para WhatsApp")
    texto_wa = _montar_texto_whatsapp_super_setembro(
        cfg=cfg,
        status_nome=status_nome,
        data_referencia=data_referencia,
        desbloqueado=desbloqueado,
        realizado_desbloqueio=realizado_desbloqueio,
        tabela_gerentes=tbl_g,
        tabela_vendedores=tbl_v,
    )
    cc1, cc2 = st.columns([1, 1])
    with cc1:
        _botao_copiar_texto_campanha(texto_wa, key="super_setembro_whatsapp")
    with cc2:
        st.download_button(
            "⬇️ Baixar texto (.txt)",
            texto_wa.encode("utf-8"),
            file_name="super_setembro_parcial_whatsapp.txt" if status_nome != "FINALIZADA" else "super_setembro_resultado_final_whatsapp.txt",
            mime="text/plain",
            use_container_width=True,
            key="down_super_setembro_whatsapp",
        )
    st.text_area("Prévia do texto para WhatsApp", texto_wa, height=420, key="preview_super_setembro_whatsapp")
    st.caption(f"Período da campanha: {cfg['inicio'].strftime('%d/%m/%Y')} a {cfg['fim'].strftime('%d/%m/%Y')} | Dados considerados até {data_referencia.strftime('%d/%m/%Y')}")


# =========================
# Carrega e prepara base
# =========================
df = carregar_dados()
df = df.loc[df["FAT_LINHA"].notna()]

# ========= Define a métrica de valor do cliente =========
use_vr_total = df["VR_TOTAL_NUM"].notna().any()
VAL_COL = "VR_TOTAL_NUM" if use_vr_total else "FAT_LINHA"

# =========================
# Navegação
# =========================
pagina_app = st.sidebar.radio("Página", ["Dashboard", "Campanhas", "Relatórios"], index=0)
if pagina_app == "Campanhas":
    render_super_setembro()
    st.stop()
if pagina_app == "Relatórios":
    render_relatorios_whatsapp(df, VAL_COL)
    st.stop()

# =========================
# Sidebar filtros
# =========================
st.sidebar.header("Filtros")

lojas = sorted([x for x in df["LOJA_N"].dropna().astype(str).unique() if x.strip() != ""])
opcoes_loja = ["TODAS"] + lojas

vendedores = sorted([x for x in df["VENDEDOR_N"].dropna().astype(str).unique() if x.strip() != ""])
opcoes_vendedor = ["TODOS"] + vendedores

datas_validas = df["DATA"].dropna()
if len(datas_validas) > 0:
    data_min = datas_validas.min().date()
    data_max = datas_validas.max().date()
else:
    data_min = date.today()
    data_max = date.today()

# Os filtros ficam dentro de um formulário para impedir que cada clique
# dispare o recálculo completo do dashboard. O processamento ocorre uma
# única vez, somente ao clicar em "Aplicar filtros".
with st.sidebar.form("form_filtros_dashboard", clear_on_submit=False):
    lojas_sel = st.multiselect("Lojas (LOJA)", opcoes_loja, default=["TODAS"])
    vendedores_sel = st.multiselect("Vendedor (VENDEDOR)", opcoes_vendedor, default=["TODOS"])

    st.subheader("Período (DATA)")
    data_ini = st.date_input("Data inicial", value=data_min, min_value=data_min, max_value=data_max)
    data_fim = st.date_input("Data final", value=data_max, min_value=data_min, max_value=data_max)

    st.subheader("Meses")
    mes_padrao_nome = MES_NUM_TO_NOME.get(data_max.month, "JAN")
    mes_sel_multi = st.multiselect(
        "Selecione 1 ou mais meses",
        options=[nome for _, nome in MESES],
        default=[mes_padrao_nome],
    )

    aplicar_filtros = st.form_submit_button("Aplicar filtros", use_container_width=True, type="primary")

if "TODAS" in lojas_sel and len(lojas_sel) > 1:
    lojas_sel = [x for x in lojas_sel if x != "TODAS"]
lojas_sel_aplicadas = lojas[:] if "TODAS" in lojas_sel else lojas_sel

if "TODOS" in vendedores_sel and len(vendedores_sel) > 1:
    vendedores_sel = [x for x in vendedores_sel if x != "TODOS"]
vendedores_sel_aplicados = vendedores[:] if "TODOS" in vendedores_sel else vendedores_sel

if data_ini > data_fim:
    data_ini, data_fim = data_fim, data_ini

mes_nums_sel = [MES_NOME_TO_NUM[m] for m in mes_sel_multi if m in MES_NOME_TO_NUM]
if not mes_nums_sel:
    mes_nums_sel = meses_no_periodo(data_ini, data_fim)
if not mes_nums_sel:
    mes_nums_sel = [data_max.month]
mes_nums_sel = sorted(set(mes_nums_sel))

st.sidebar.divider()
st.sidebar.caption("As bases ficam em cache por até 1 hora. Ao trocar os arquivos Excel, use o botão abaixo.")
if st.sidebar.button("Recarregar agora (ignorar cache)", use_container_width=True):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.rerun()

# Cabeçalho institucional: logo incorporada ao próprio arquivo Python.
# Não é necessário enviar a imagem separadamente para o Streamlit Cloud.
st.markdown("""
<style>
.dauto-central-header {display:flex; align-items:center; justify-content:space-between; gap:24px; padding:4px 0 14px;}
.dauto-central-title {font-size:1.35rem; font-weight:800; color:#10234a; margin:0 0 7px;}
.dauto-central-subtitle {font-size:.82rem; color:#71819a; line-height:1.5; margin:0;}
.dauto-central-logo {width:clamp(210px, 29vw, 390px); max-width:42%; height:auto; object-fit:contain; flex-shrink:0;}
@media(max-width:767px) {
 .dauto-central-header {align-items:flex-start; flex-direction:column-reverse; gap:14px;}
 .dauto-central-logo {width:225px; max-width:85%; align-self:center;}
 .dauto-central-title {font-size:1.2rem;}
}
</style>
<div class="dauto-central-header">
  <div>
    <div class="dauto-central-title">Central de análise comercial</div>
    <p class="dauto-central-subtitle">Navegue entre a visão executiva, os detalhamentos e as análises específicas. Os filtros laterais valem para todas as áreas.</p>
  </div>
  <img class="dauto-central-logo" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAvgAAAC/CAYAAACL4KgeAAC7dElEQVR42uy9eaAsV1Xv/1m7+twp80TIRAKEMCMzOD0QRBREQIwIKg48QfEhD5Axw8kJCYOICKg/wBEVfYIDIAgoMogiAjIJYQoQICRhyJzc6XTt7++PvatqV9Wu7j733iR9klra5Nxzuqur9vhd373Wd8Foo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTbaaKONNtpoo4022mijjTZgNjbBaOMcWI3zYE2AxibZ5zVEmb9pH65T/exv3Oc43TX/fHM6FmY8g+K9mvaj3WZ9h81+jwzOGVjH17Tgd2zgflft5jVPnrICR+wAfxAU22CyAkUBZRH+Xq6D3w3ueti7Ez62Bz4wHaf9aKONtkk25+c+ENx9wUow4QCsxMcN1sX/9XEzc/h6Q/MAFn9f7ccOvBTe1zFfXy9uTi7dxn34m8eFD7vwbifwJbAb73dBsRP89eCvArsa1q+Bl187dueNbc8+CA4+FvaeAMVx4I4Af3h4uYPBtmA4QyasGiAe4cH78G8JkwEOWRF/LsK/61EafzILLxmShdFbASsXf45jtP6s1Iw1CcljCuNcTEDbwQqM76I9L4aXXzL3sd2ZP4Psh0Hr8Xs7AMsEEqgM/wXDTOBABVAgXHyUMtxPPQsm4C5D3345/MF1y9v3qxMofx0r7hfuvXpmp7iuxLmLB6bAFPn4Ph/+ZrE9BOBDv1Io/K3CzVyIt5fB2u4F7uk0jGeFNvSu/no5sOrfEhQlqEQ+roDOgVZAW8M9mMIHnYcy/hxfsjK4LWbhM3E1MxPVk7cXudbIacPt+vtBSr7HFO/XxbZKxlQ1hyTMCvCT8L66vUH1+Evmhvdhfa+fpZkvZmbgVDsrcf6Y980+oQJZEdrHe6zYjdZ/C176xU3ixB8D/mScuy34W+PtOHC3xnQc6BjEweC2YEzAHPgI8JkivxfjOrArwb6D/MXgLgL/VdBFML0YXnrFSA6MNtpoy2ST+J/HwtbnhvU/Q0opJZIEUkRLWNiIFUCOXAfjZBY8o9l3AjjKcFSGxwIsorqMKeyIrsQowZXgdgLXIHc5nPkNsAvAPgvrn4XLL4LX7xy7+EDaC45hMrkTnu9B9j3AnYCTYeXoAI5cBaJiP7rwo3WISAOjoBkgFRYJ77dk4FQ4TzUZXF3QxfFhyd/7rmu950phCFUATIq+gioAdCVMXgvMB/jefS+27ckBt3bHdeJPxLFt9RNZ++ZUz53kmR0wvQiK3wOWGOBf4LC7PNzY9mMBk8Y+kHpdoKrfzNrrS9VsloLd9JMTTOUFYuvvAgsAfH+SMXkqsoi/LWC1avwJsCL5DksgWYXrrRm/Zs0SmZL56tynxfFY+yTRZ1E63qv+tc73ddZWq8bQpJlD6fc3jk9s66JeQlW1rVXPmrZN0bpO5QqbhBRHZ+U7G02btb67+ueE4LwWfwYsKcB//lEUW+/uZA/w8vcD3QXsBKSDYeLqbUeRR4gUhKqtJj5saFeXtKdhthJ7060bukZMvgGrn8P8x/B8ENynYG3vuF+MNtpoNz3AN7MITjqAPAX4vr1jtwCUAmtk1oHqam1yvUtXRFPysYC5GqCkNoQqAjKsV9tDQMdidgdkDwyrsS9hcgV26y/BmR9H9l5w/wlrl43dvS/2opOh/AHwP4K5+1P6E8EObrokObUJP6vVt9nO7/SuKuCtBBBaPbQaJzKOUatGig1QZqYOYLT0+qhsBnYA/A5sPbLHC5ibRlLat+dI8jw1uO+Aui7bHxhcxfYSlAVifXOMDe0V08D0StbvYxd/6a03kzuwuXbyAuKsmsmFtt61YCjhFFFMsbJoBl+FdisfsWzh+x4QrsaitRyx+INXKxKnxvwOzNe9XQ+F1nNWY7oD1MlNkQRQKwHzFr2kinAxr9ZMkusspOSDMGun20vtHqiex5L1vf/AZgbFnuUbj6tHh7XKPRJz34fnFA874glJ0tY+sk9VZzuQT56/wztZd9WqDzYmwo4CO8pw95TcEzBdDe6/sXPehXfvgfKzI9gfbbTRbjqAj3NmFhe8ai+03NJfL3iqqI4KbJkBppoUrUFOa9mssZpyG1pKgtYUEx0nocWkqTkSrze+AuwYcMfA5PswPQ24CFbfh/zfwc5/g1dcP3b9LHveYRTbHoq3x0P5/cAJxkrY2GTx6F5ddGLNyQ+dHbLv6LXGV0PTd5jD7jgJANms74p22FILwyF3gGTIuht+EUJlKBeL+TZrvA91x2X1PCnYq8axa+aDdYCpqisXDJx9LZkdIaAMN142DlkKnPGmgf6nC4ItYc5J208G2xYE+BPfOJzV0Uk1DuP6YA0mTon9xuGzZC3p9K1Zm4ivCI2G8s4ulw0gV3JkQf796d+T744cjNX3qJbnG50Qy2Qy5DmbejW25EhVjeNMz1NIPFl5MGfBcbvJQb0Dfz/MfgbxYzA5FSsK8Bglwvt4qmF19JXSRSo9WEkjqHzn1K0bcuWJIX9qhhsm3GHgHmIqHiLTNcCnsLP/Ab/7zfBbF497y2ijjXYjA3yLkcrOAoOWrPVymIWY0sBcCGRmpoa0SzYk+S5D5VpQTdWC2CKGusAwPSJPlmEpHLmT7uEgRYa2YgCpwlmrg2d3O1i5ndz0l/GHfALO+mvw/wDnf3kcAqk9/1Tclscjnoh3dwnH+mXoNJsm4MCsdY7d+U8fuMQOTMC7ehQl+ZTMDAspT/77engpBV0N+WkUFYMZ3iVDrdjkOSYVLVI+jtU08ESqnOQqDpwMsus+Z4o2V5Yc4n9BcNx6m920TgeoORzMdHe//1wyb2milxY3H3MAatyq7vqiNOZLSk6eImjtgu3mJiqo3fzDaIUedQetujR+ZrooidTqTBeS9BFZeiCRsiTxfox5k6E/xyzxrjrf3wrOUWw4qydxXAO23IQA/zmH4LY+GunnwX4Aih0xfUFoGjcCs7CnJc6nuh5Qd5+qGqhorRt9q8J2ZJ0h78N+CWCHouIHkftBzD0Tzv47pL+EF/33uNeMNtpoNxLAr/YJQ70QggrIV4x8FRXvBjaODvZLNzdLd5Uu49kBOkP6Gx0mUD2gYD2aUJQK+Y4YVtwbinvD+jNh9R1o+pcw+RCs3YKVEc68s5k9WdgT8JPjw8GMjzEjsj4gs+qQp9c3ZgOZZq1Nljya6OJykQ+jsDlOQDY0wdrh8RAYdXzlzC4OqA2XjOf6O2UpWLRkjtgcIOTipPENiGTrkitc3Ur1w2qgH4knfWnIQwb3JkiyPUAsC1FntaNSf643VnprRxyUmsG+VyS9kntU47BG57AVltVb4nL3oDkP1gWjGj7lGrz3ztBrRc0xMOfo+WjNvUgpCWNo6m78cbe6BfzDMXsOcj8Yk4GBqa9PX6zLOmXax7oNm+kgrNkOwwCwudpHJhMlmBPCY9P4Ze4k0+T/ytZ/KQD96WvgxZ8cIchoo412wwL8KHIi6+yxCabXAFjKbxQ2sGGRPzZmsU1ncBOrQwNmqPJZhTamMRLJHY8VvwJ6Aub/Eb3w1fDiD9+yuv+Fx2ErTwV7snAnBgZs6nueUg6QWf73gz0rZcIPFuhzzR8iWoS9bMdZB9GQhJAMV9myKMJXw6SmX+STZ7W5zmn7kpXySh0r4JZ77Hzb4C6uRvcLuCNdIr/P5leZui1nS7FxFjPNWH4WGXNGHxAa/YBsJdE56jIOs/o50xAH6qzGZreHcg55bpoNtUd1whE8hcCO32h2egGnPdRM/0fwI8htxbzCia2RzW/JnRwpzyPkHjd6NRubxz2Htf6vF1OQHYYVv4zxSHT2G+D618IrvjpCkdFGG+2GAfhkNhrLLIDdGNuhkIrBxU+Z42HmHNsvsImlLNwgsKqO48sqsTLE65odDJMngD0MzvxzKF4DaxfdvLt91VHocXieb7h7hy6ZRlVS8smrQyC/+7ch1j3dK20YiPXGGW1nM+c3DI6ReQ5iGncrDMrFAItUkoQ8q0anTnFMNcf3s+ZI6x5beN7B+rIz+A7YkuR42mwvLMMN2NDYSpO3KWHXosnPClH3Yq4AvDboZNocp4B5SHmDDm02r2PgYboHnlpwDlj7pMO6c6o7dmvlKbMQe1XeSE7omXfF7Lmgn5Ts4OCZlL6RzbW+06g5DlvuOW3gIKnry88D9mF/saYTq4t5wHljHWHHYu65aMej4axXw843jLlho4022oE019poMrliC+1X8xK79pV5Sm7AZn7E+oAwazkiUD4c73I0NnkWpn+G1f8NT9lx8+zyM24L7rXm+UujuHeMaPcNezxAf9pw3+T7VNmjcbNhHw3NAMFD97BRGNy61j5jaB+aTTHPu44pswR01KcfNbdoc5577vhdNjMb7tDsu/snhLn+bGe/CnYs2CJeqM7ymbumZH8/NK67Za/mzQGb8b7sOFDGAbY2odJto+RvYo5Hs8CcVW6+26AjciMw+A+a4M76Vax4N0yeFGprlD7JwqlvTBpwbDS8cXTTR4aWBS2y3pEZH9kFqqxkY+O+4+4I7vfhoL+FF9xjhCSjjTbagQX4ZtZlOBbat9Ow/NZmYPnNJcU+XWYofbVCJ2iS5ZLrpbt4715zzG9P8aJ6U1HxPx5Kj9wdQK+DW/89rN73ZtXbxdmnYyvvwuxXhFup9N4qASTwNgiAhoB31sEjybmgl+sqLcKAZUCJ5m/aqTS5zXJK6hvxyScXZPDNty4+jx2W2pBkGAz6avxPYaVc7sF0RLsnNQfw7KvfIgyuWJApnvgG7vp8n2ScxVaAh+XZBOuuc5nSZr2Tp+wYVbtB6n9af07FsO9qLpkNOEc27JBbZ13PAve0xMSs+r3W0uu/gVV0zrg99qA3mNzvgZ0QFa58qGsQa6SY5UmIoefNpYVZO2VioRPBIYfQcveRtFFy781np7GYmvtRmLwD94InMlaYH2200Q4YwK+WuE4NFc1iKVp1rwaYqFmMVVeJzTJeRWtT6zgAA+Ta4ALf/yJDzmIsKUE1oYg1tsywycMxezuc/eTNv+CuHgpn/Rbe/SW406LOXahWFsIhzOrCsAwmBfZA0mAHDDhd88IacgIX3RpA3b8NMGvq8HxzQywMg2KhfrbBSdEBYfPmTv2zqQNEprC7XPphJZV1Maa+UNL8UKxMO2Vyox0cvziDb74PoOeAsx5B23tPLAY14wQgqWvWdixNfQYey4TgMBySNjSXFmXlu+8fCkGyDCjuOcSRmg55TDfMulis/gS28i5YeWKUTWuKmlfKOLXy2wIEwQwHNHdwNDhGc+00y6GzDsivi/8VwTlpCqQJm3ozd6Jp5c+xs34ffuPYEZ6MNtpoBwDgSzOZtqGNJMt22XyWaW6AbBMkqbQwzEAuVe+T80KJOhtdLMzbUGWYQky6HYsVr8XOeTU8/6jN2cVn3BbTX5lNngNsCWyRt9arKvvCjI0qBypmsJZigbCBWaAvM3bEjDE4y4HIg1NLg443glSUVmzuIDUNhZ3AjBCcVCKlesqtyx+oY6auI6dF2OUZ/RdFKYMUyr7AR80ZhTlW12Y4l0abfOjmIHUZ8Ohc9tRuumuXGD4mtT6w7Dkgmfa0RfKXujnuGljnew59p66FOYE/wGP0KSu4s56F1xvBTg15UkG6xlrzZM6elPu3zR82cwmtXHfNDVVM5rjFPcZkaoVEhpdC+JFDxa9hh74Fnv+AEaKMNtpo+wnwTcwDxi1Wp6HtzeYAuFlHninL1Vu41T8Gr482lQ8DSkHgjIXXLMPI0LAtVmeReYEm2OT/YNs3YYzkWffBJn8Hk0fG6o1q+2EL6hAu4tgNSh/mN8rWMBgCWENAsOtPzmOIc45IP85Wi1eyrazoFyDNHdlrABS2HE6ftlPB0qvoJMtHVTPKZgtZZRn1jHPevK262BUL3o8P1KhpGPHOEgHogFtjIz5GWvhMw35m7rixm68y67QsB1ot71wN1BLsJ6TOO2HtJ2OFQg+4A3jK9OyDzG7928heHipll76SWxbeJNmGYrwWTTQecmyGHDgtuEbm2q4dIlglsFQnEvEtklnpYeWB2Ja/oXjho0aYMtpoo+3HDu3bi1QuvrB1TGltfejcppA7cs4upmoHP1pHDSfZjVocc47OnRkL2XECZmxq7UeX0LrAPRjb+ndw1sM2R9euPgIr/haKe4H3UlvnXfPKpXY3uoUcvxkgaoDlHxxvCzodrciBoTG7AIPHhuBDNVhjhErtJ88BS4PtaKoLuIUDgQmsF0s/xOTdwl22SFgVSSx79RvhN5BkG0LulAyMDiExtCb07tEG8kWG5kQrmDupjTDLOa5i7GedVmgOUGT4d0MCCVpkTsxtccskOuyrPe8wOOj3JPuNEDLpfVVISjG8xRyzQ/wWCQtlhrM/L+xmXijOomtk8rN6A8jFhcgIJ8jFyWjrH1G88HEjVBlttNH2DeBLmhn6MLRBa85+nktmq3+n9sZYlV9XB0O1mJbkJrqxOBuh23Kb9FzGfypkp0Lx53DWo5e5U4virMdh7k/BTgHvI+CxhZUgBuRSFwJxuc11nkpJZpxkPK3ZG/i8UItZe7U2Gg3ipo38uxvc1LthEJYdV2Fgh9OUIpmXW5cf4Nf6gdZ2toZQpebMVSVFsGSh9NhGCpAFmczM2mV5mUkxHHY2j91WcsNq0+ODybgZZ8JsjiOUftcshbJF60fYDDTfmy9azDneL3v+UbDl9zH7xaBrX4XkWPix6yjlToEWyAka7NN5qkq5flvQwcquX93Dpexa6eK8mnq8bkU5eT2c+dMjXBlttNE2YpO42mgm8M38O3sS39W67si+KWXmhnSdTW3R4VR1oBU4P2fzTO8zVwqezjNkQj+ktq+BTT1yt4bi9XCmwXlvWT6X7ZwnlJ7XII4CeWqFnA2C9I2QefNyLYaSdjt91dOnTvuCGZ+b9QzKY81gvh5jG8P4isUUXHMapP79tqT71B9rasVYC6xkU0TmtCYrA6FuyQ+asZ4M9ZtVDtS+qu/Y7PGQ+96OJvxQ4eWZC4nRHhN2AOaWdS88p/0GgKlZ368ZDFnqVGm+YWz1VhivRfZYVIo4DcN9ecuC93k5HfPWBOXJiMG9QPP3mLlrYS6fKJurYhZzG5rN2KYeiiOheC32ggn+JX81wpacPe8wKB4Ie7bAlirE0YXwp5KwuFoJvqRwRrnzY/C7l47tNtrNH+DXQaAa3iQ7C5sWWWyHQFq3wkp93VSKgw7Y0QAFMnsB7xUuSdV/mMEkZp0IZyE2tLiVMXmddNYueNG7lwfcn/ELyF4NHIrJH9Br54pZzdrIZiCZXlXROeNI8+5n6B7njMnMty6Iri3EHsuDOawT/6R5UqKzwIfti7j/TQjwcwot85y8hfpTUhglDnYu2CBOYTOvOIsq3GGBNSr5/pZzP6vPNOdPtoE5NM/5WbDtbEaRcSkzpxauB9Ed1H4/x+mzjsbc6wx7jKyUYfEYeV/VeSxuCWqv+wMkwlATasgJmDeOF5nf3S+z3Jola8WTmQgSocURaOX3cWdO8ee9aYQuvUZ+DFb8MdqeZrUHt9YmoXVD3ojDczHs+YGxzUa7uVs7Bj+loDRnMxnSGV9ge5iPsza6wWVibLv3NC8WdgYD0z5yN0ISlG5lVvwRrH3vcnTlGT9qrLwKOJRFY2M140VHPGOGnnQPZKS4IQOEZqrraAPDJJfUCoslvPX/sLgOPlIX284tkjSv7fFtxY2lr2SbesJ5h37g2GR4Pua/wjZ6S21NIy323hQcasE5M5BjoqH3twgKFiuA1QWcQzHfQ0mgGwrJ04J5MNUb9lUH/3mHYYe+xnCPiexqUk/BbEg9tHu/lvOebSDZx4aXEcs5PUPymIs6i4uG8eS35ZB4ay7E5csCyJcdjorfgzN+ZIQuqZ1ewJbHwqRIjvxDsn2d1OMDWWAOKN8Hr/ja2G6j3UIAfifOtSuLlotZzYHhDS33mbK5ue9TsnBbR0+v97eB3bn+mOaDv9z+0FGQCLCh9BInYvozeOE9b9puXL0dVrxKcodFBfj07q2LK3rqNZ1ci27hFw3Fzw/El6oLesN4Ui+eepYud+b7bGhjtRlgKNu3ihJ/6uTWLqqiEyaHFoj3z8d5MCw/Wl9/ZfkBvqKMkJKmqICmp0lytT627bVJa01pxbZv4GRl6lq3MRRCPgi08g7JzKUtU1Cqm2bU7mdb/FRjqObEkINiQ47Loo5Vp3JuLZVG+2jAWrTHxsGY7Vg1ip8R69ED8wgzw1lXjnQ219RuDFX1Cmo1o/kVpGeub/NqfmgDpMRQHP8sedbK4cGFmi14gR2DFa+CF9xthC+V3fZOOPuBcKIaZanMhGpljkSqauqR3ja22Wi3QIAvYV6WMl9DIJ8DEZoZJ6JpYxrmM3dda8tsbuBa5mZs+fUGkErmlx44DZu8Hs446abpwmcfBO4VsHIalLKQnZjcZPtBBvW00/e0gIoN982Q7GS/ekyMKy2JwFrzjswZqEJrme+3RSU2Oy3QYMxquO9ddNxaKwYfp8ENPXcMPzdJXQbTJQf494lFqARpiJLNB5u9WOdc6zXJ9xsA+JP+W2u1KFWbvGaOlQE1muzzdYpl9ebYImosC3y/caDWRgbj0617E1Y5wBb3g2reppNusg/j5q5PBJ6mwNzTyLItXiajr+GfFpqzNqjX8DYxVwFnMSd3Mccst14uEj3WqkPg4jie3Anb8juwevgIYQC39VHgjgrUWzU2CzBn7fLUE4AL4Yp/HxtttFsOwJcyhWCrRd0LUwPC90Vara/9HK8Zqo/3r6E2mzS0SA6VC88gWsuhXGVYxwzDk7JBlX58XSZdUw92P6z4LXjawTd+Dx7yLIxHR2k55cqtZ/WwF9BzbsURz+r7LtvZY8rLNr1JKKS7UXDSOxkg6bf2H9Uav9W/5WMcptpF1OpY1wUZ/F7SsmYDqVksnRpft13gdLLkha4uMRTXD2uTAWYzwMwiiko1tQpBUmTPBjIpfILsFMdZFd5dVjg/G/hT37frOPuZOlfxuWXmhXmlVHC3PtbM9TKXgGwDS6Ly861XV2IeIK2o7mpOIKklV+wTv9e3D5bqmAeXWbxn25Yt594Zs/NMbI0lxPNr8ozQFuVIhdTR6kRlyeaF0MW2qE4YzdQum5473Z7hzA4tAwtki6s7FsmTHKgUKh6G+d8cIczqDlQ8xmqipapKb3T7zsxh2DvhDy4b2220Ww7ATymRtIgVJKWQEsCkakObt2lF9sfiS9WG0vu+FHzF/0ZgFki8ZPFNXlYtzPXnmo09AZTZxLN5zkqG4WnWWLMgZVa9tRTip+HoZ9643XfWjyGe3VdUtvzGMRRTmjuWrs90NABS0n6o/kuNHuJYyXNTnWpI+WJp3c2WfLJcE34VgFb9vT7Z2X3MOe4kGJD4AhsqoiPrhkEZzM/pyIG6jkMaB3wJO/zSrx614pXr1ViYW0NhEbaztkWr+noF3fwk07I+jlJ3nQtjVsk4VnW0H+oqkVkW4/yv/6UGB4ZBkZx/9sI/ZjmBmuMA5SrWmiRTuoiq5WTnOqyquWCOTiHz+m6tNbyr9wVJ2KZAnpXAdPHB8rJD9q7730LcRnhtaD3urFWaM5dyf+8B7VYHRdKhXq9S787HCa+qZ0Xyc0ckNp9vtGhxrNSRyTh+tZNplSPr/u8Yjz/9AeCeKGUp67j7ZF6ZifJ6ib8fYd9otyyA3+wN+cp9XVo4HBeqikZvy86k7EcZ4yFrBbR+cOpgKdyodBXJpUw1xXj7ldShJ6FlI+g3tR+OuZJ5Q5VJ+zmD3ppKiwiTw+zZcOYjbpyue+EJmHsJ2GFmUf87Jr1pqLqOLdYG3b/1ZStN1olXDlyJSa3N0VkNDHI0dqy/kNntG5Be19PxjQOXasxVLGo1DpTESxg0xWPS7x1g4ygWY4rNQgC+wj2azZEcHBxgaVtXIAoM87C+5Az+pQLKbBXq7uPOqyydY2ObAehh2+IAP8Z0p+xdRdg3J2/duRBV1+PYiO9VZAQD8E8dT6BdmKxzz4uqSinjDc4ruFQP/0R0pkOqBGfFJSFJqvWEGh/Hx/FbhzMkryK8MMLpWziBk8pwChK8GA9+feGNxu1+BvBITD7GEObBr21g/swLs2mRM6FdTAlpYJG9MBe1510egNe7XHPUq4YwaMil1mlRcgs2IFY0q8ZC17mr7roOvJSgOCicHD/7VrdM+CLDiseDbRVlQuz42hNv2tUBfBgu/dgI+0a7pdikmSgGKlyILaepRj5YiUXNCT0VblJ7UiVno+qCq8wi3PC6KcM7yayIYRJbizpJ6A25uCn5/mKdLSJDX1azC+iVQSRKM1PxyA7D3DnoBf8NL/nWDddtqw7jOVB8j1k3qbbaDTqVlmZp0s/Tkm/92qmOa07USmRydTVGm6TXj4ioIJs1WwGldoiNNUWflGW6rUqx81jLqahxnavfb0SiNvtg1Rj3E9i7WKx3rEtlnSLMQ3H3xoy8h46/rRqoLjvAJ8lytF4kQ4t8dDPaiI4meFc1d6NmFKhwDRERuFczl/Db/c5QcgNKtedbR0ZG7QVUCp6JZFS4wrSp2tGtw6F0UKgdb2LVSM23kbUyamLiZYpf6hOgTs3CCFrj+Lfm9Mx1iJJuqyiJU2pYjqSGRAFFudAmM1n7wWlpzwgEhJfkzDK8Q5b1TsCx1Ge0NavGQmtZ9ElbCVXJq5EA6INvl8idtmKAPEl/qlo/qpsMbaaqsKC6sqhDa4R16mZ0nymdF/VzlYLie7Adv4l47i0PvrzgtrDlh+uyJGF/SjP+Y8iOXIwKeCu8fucI+0a7hQF83o38brACcyFTLTAd6/HIOybUuajZZTHZSoYlCXYyXy8/qkGWr1kNVGBFET0DDypjDIjV1bakin2vmCIXNlKtgD8EOAqKIzFOEBwLOjJESJRgilVbfWbRTryVWYWZbA7D3cKnLlJRPn6yFHL3g8mvA6vcYCVi/MPA/XLYLPPo3FwHhG8AxDOP1K8TmGUNqOd6ZN+GyVdQ+S2wizF3Tehn70BFBPkVhrVWlLISGCMJhw8n+dV4UiNzpoq+NIVNWlaPNVOJN9XXcmZCSny9SFHKQrgCHu89rFwLexeMzVQBnc27y+bmStLP7QtPExKxZfkBfn0s42tdlWzUns+ws4Nz0NS02UYHqPsacs8DP4nqIw5TgXCq6Puw9kxiH3qwdcxNoYxjRBbpa1fHkjl8GFNKAv7Mg5Xhe1RgKsAOBfs5oaOYE52R+4MGkWo6hswDf4dzn43P7CNDbUhFPSqtWndrVlk1kqxOz+pfFHHQyeoxGNY1l/Zv4pmU4K6D8psLkBGHTks7Azi6BsfWFkDTnPAlddj0HumiAaLIchdwDiYR8Ns14C7B+ArSN8P6Zdfi3DqUE5nbhsojQLcCTsDsVLBbIVZiM+SCgCrvQzVZkct76obyzFLbmjlsPGBPhbPeAS/6wC0Lvmx5KLITQv5ZInog3zitVlbk0sWw992MNtotD+C/6APAB2YyjBv9/QY2NrTR96xuAw4HTgT7fuDHMfcAZIfUChDdOJXuGWmueMmiChW9W3LWjudxv4pW3w5rHznwXfacQzD3AigOCQvbQENpTvbXEOCcJ69n3gLEKaqt+TPg3on59+HtAvjOt+GVuxYeG0Pv8fs43vyC10mIyX0wN7PBhqpc2ryB5SpREcGeJQf4xxl1TJOTAgLu41PNeeSZUrsbBfhrFwO/m+1/vx9r0sLj+E+3mX39h4QdVcX/tgBsmntS+adVvHtGrrINVKvTUZtC8Tr8We89oPNqf9f1wZlij0N6SBRroHskPCQIoKGhYjP6y5Klr+dAFpFat8vAPoK59+PLj4L/CuLbsDad3UbPPggddiL4+2N6aNh3ituCFYFBVlLIoqpInRBNixYxm1NjpP83CdyhmH8eevbH4BXX3zKgy+oWM/txiSJ4a66/PKcFM618L/ryl0fIN9otEOBvNlvbDVwWXx+D1T9E3B/cozF+HLlTA3Pim7Nlq9hBWbZC4awFNafZnVu0I50PxTGYnote9wR46vqBffYdjwX7gaiJ3Nyg1HE+tBiQn7fB9MIFXNC/MfsEcn8G638Lq5dQ3pKmjVmjuhhD8vd7oyan/7f0pozQyCDIH6q4ajc0yryx7CInbJpSyUL5jNdWTEYaYjPr8Suk7zdJw5x3nGn664KV5hzHFvMrhqRVu15AK0oyozhkFfnCp5D9P7B3wDc+j16/wXX5FdcDXwC+gPhLWD0JeDDw01A8COzgEDJTKRD52eSRBogmsXiOVG0lwENh208Af33LWIP93cXkgcGxavKX1NNFdQbai/NvpXzzLWqXGm20yc3jMdZ2Au8H3o9e9Hvgfxbjl43JycKHbMgqX09OlcxhNqRAGVw/UIzJWoyRJcClFLhHUVzySErecuCe8zmHYDyNcOboW3eZxgzPOubN5TfnPpKGCNdJq4UZ+prg99HWv4Dn3kLlxnwT6mDt0OpcG6uzgaeHSWohGh9jeE3LH6JzqeDWPkSzdOrMLXAipHQO2azqsQZcvZnGRhni3jOLRrVOxNIfTfOkGQBJATUXU0JjArlZjPvQJgH4bvpEiXsFKWEwK7oZAu35kagjMzSeZpASzXtdED0I8ddfRvYH4P8KzjlQ65Vg7evAn8PqmyLQ/zWj+FHQFvBetY8X4sCFr3N39tuHbbWThQRisQXTM9EL3nvD5n8tC3KZPNI8x8jLg7WF45SGmZmBLqBc/7cR7o02AvxNb2d9GTgXveiNMv/rwFOQHdRkWXlCUHYV88+gZrnmEE7qAjmqSH9J2BY8z4LV98DadQdmwzzoJ5Hdr45lbQErXwvYqcsKWZKcZZ11MHuc3WHEKora/F9JK2vwwi/eAudKJy3UhYSuyhvKfcAyCaQ2wEaGDt5sbTINeZtq14tIp5XNCEbP/M2QSVHKUQb7V4boRrZrZBxWyuoScY23XCXvRm9a1k2E74JUSR0R9/AZJyiWX0KVFx6H+GWQw8wHJ8bnWZV5ceYdAbc00Kc/l0yGczK/C/HHyL8S1r5yA5JLu4F3oae/TxzzaExnw+SuMPVNkrJhciEa34j5R2RPkM3mJW8kQyp9bpPQ5H44fhHPy27eS/Hzj6C0HwVZs+92NjerNjcH5v8JvfTym/8WdeaTzNxDBXtBk5hT4w2tg/eS/1t4+b+OsHcE+DcPoC+eBWsfwHgV2MkhZsdFibMZ7AhtkNtjuLsqPGqHG4SNvBQU3wt6BPCm/X+e1R14PTXKO/gKPUoDm0MqtzZn08ypvajGJWbg1kEvRv7F8MK9t/hZYy42f2SKOoofSmTtWJDRrqnd+ue9mwDYxqMw+UQypqN+Y3klyS62a+Uhmgv58qEwzQx2f/lMjaBAy9GuA3DSLu4uHLXUQPXEvn/S0yT1L7e57Y8Bf2fkk5PGfiybNDAQBkJblKYwpGtc/b7CCX0Vb78BZ7/9xnvg1+wB3oRWP4zx22hyOpSxQqTam4Nl1ux03SWTb9BZr6WBhdwXvwCrb4C1m/Hp6vb7hxCdWIWtJthi1UOLajryDiuvR7zz5r8pPWgCk8dLxSOiZHhUmov/GzQF/nOEvCPAv5nZ6lvReZdi5WuF3SvQSEVQtcxI91lH9aSfFpYH+c1Rs1mT5Osm4J4Mq2+LTM/+7Jg/hOl+UPrAHtNImXap4RmbZs9hkdWKRRGjBX5RMSgAt9fMVr3OfhmbLyj6hoG18Yymh1I7Q2Noc846Wu2MSoMtmwHgx+GkqD1ohg07lS3C0gYY/uBZWqU9EmjsrbZ5xkY+jrzGdbJGTz3JKjZEn0tQIpOZvGPpDzVefhDa9XNRoUdNror6yjeQLeiVMvSWq5WRrM3NnyZmxn/LF78IZ3zmpnn2ta+j1SdhK19APNfwWwCfyEhXDRGGg4thWYnnoizaz60famIpDYKSnN0JTR8KvPFmup87TD9lsoOFLytHWBTJHBKiBJtg6OPikP8+8PdxegG3Pgh2xB7ZbXC4D//dmdknj6R5HzS1Pa5Ift7R+dwVLlx/m8I1d8Q4zmsEhwrW9jSj43sPMfOnhOf3PhXOigPuerCvj7v3LRLgn3sfKO8Ftg6uBF+GY+CSkKSCC1JpLrKMJc3fK3m1SlfYWzxCnlZ1zwMDvP5BWPvuTfOYZ34Erf40FH+E2YPM5MNRt29r9ncJtRhKOzOuuMv4S0Gyrq7mWnw/8H3Ae/cPU/J4ZJMgf+cbBbuEqhe9X/U2y7QGlKKmfZ2dVCmghqq0DtxeydakT7/8pgP35z0Q/ENxFgsb+ChZ6Kn/DdMoFxjlOIOwfTOep1XGwST+t1Oq0gPFFXDtP8LLr53L0sq7cJASi2h1aii2gAoDeRwiKUtZbdKueovrVAxbWjwbnr1oKQNmS0ak88hmAJka9MXiaJYr3La8FuaO0ST5qA7NaRVb6yRUN+VD1CoBIGetMm1oM3jZO+9tcC+lMpGaEbaVAfotGc0esM15kRMD/Zc8Pw9nfOmmff613Yiz4NyrhJ2HsY1EVsmq0526P9UOOcrkSfWTr6siI1WBtlhoS4VBcTqs/k1LGehmY8UdMH5UoXp8q1Hqk/lkERbl38Czdx34+zjt+7HiXHBbIziPZ407pBqwi3aFClwkK+LCYIIjglRzU0GtStBxcNgkHGN6YPs0enDrcIQMfUz6hXPhDZE4POhwoSONKXWByFCERCZzwq6CXZdujj5+/qlQnNr8ez3u55qCXQ/HfuqG6dObLcAvHxekF1WCZGaSyhi37eOihKku/WyYlXFTSvXJfNybpjTVQcygvBp4DPDvN+GieyG86Mng3yzZvaJ2rhLd6B6AV7cuUzf8oFHrR63N11WTTMBBUDwO6X2DwdrzWYvbgn6o4bFcRze9v/WZDRDIFhhE9VCnBVnvgDHMcFNhL4FvvBxuSvUB/QDmzgsqdE2h4g4yrJBEFTMSfjQpbARVudJO0QDVgUgO8xeiHR8Crh3Ebq2RwMDRTmcMJYm1ylXQ7ezkTUTH+uZgrS0NJ+5XcR0Ecp1LtEISWhVInYdD/eZYTg9VlUlZJxGn8RRN4mUb7LeQbZUd5FoAMFdNe4ntkeC2m5Veor2+5vKdqmmYmU+9MLfYsGZpbbyJgT4OxS/e9OA+tfKVsLIFyjXQSthzinZOFMnWaf12GQxhijXL2qdfFa/BgxF3AT59s0MsTg9DOoFaZCI6kJbsiyoJNTDKb6LybTfQwncv4EHBNbekwJywXA2LpAxPGqVnrW1YmZyknENvCC6CN+xpfrf3RJgcJjV1l6sLB3Vd+xbauymSr80m/xfcU0HTMP63GFZWVeW+Bt/+MeCiEb4vMF1qGiCoeBiUTioL8IXwhZkvoCyEd6h0EF6SdzAN/46/N/POkJnJzHmH+SK+tgDbbvrHPevLqPhV8BfHmuxU5ehbFc87MZKt2OIeAOnSapZBN/5hcM5J+95L7hHACWZ9B8HM6kKK1jrct0yYdwzfVZ9JDNcoYuVLZ8L+BKYvhtev37R95idQIk29fOmR95j3WBn/6z3mBSUyH05N5D2Ukfj1Bt4sxCXGUxUvkDcjfDYEJG0FtsxadxKIth7W7OBw2JArkEpCdjbvNsZXXY85vnxSHmrTcNe5H+mUArKMgzO7xIb3cNUmaYNLHchV7qalLEDC2KfUYysdN85LycV3JeMijiEt/bA4/xhMPy58ylO3SBNlY7ds9mDorMfNNQrD/OegfDKc8fnlaos1D+svR/YyZNM6UjxutzZLijmdJ5ZfhdRNDAukj9DksEio3czsmduRHh0miQ8zwlw8avfJKx7WqngbnP+NGwKDgt0hdFhZirIEXxq+rP4tlaXwZQgj8iVGaaiE0iu+jNILX2IqI7kaf7aSEB5Rqvq94ovSh5wnXdIeMe72YNsxZ1jhQt2HSSzqWID0Vbjk2uXv46dvFe5OyE2EtmDaAqwgtwWbbDcr9sI1VzPaBgC+zOLuEdcPl6zLrllN6gWnVbo7+JOi3pRCImvra4rmKOumtjM/gtyvW2BqXeIRt0MGuhg9jQ3thdB2VucQrRTjkUshTokFufbBXreC9Ig6glWRoZBvcGFzCp6cOqgVc98vJpY4JfHnoGleOMz+HXadCWtLkFBbrAS2ysVXBhTJ+nrSVZROEh7RcOQeKE3yYaMIDbcSvmshp2PakNXWHzeWB7fdpOwWzIuOWrx3DwdvDoAfwYo07A6RJB1LM1B95m8BBB25SZydQ6o4RSrVGEn9+Vd5fNWci7FLZnVxZ2sdjaTIeNkZ/GL6YLA71nrwlnf20mHS4hpyiaXZkx8TFGboWxQ8FdY+uZwNsjaF8nyMv0DOGOrObhV1pcQLixfli3FyZno0PO3gmxdcOfRe4B8APsjnVEpbTc0DYgFqw8o9UL7phrmPp28B3b6drOdrOdvAzfikw8I+E5Tumk6tQ9jq7jWsOwFqUlFAadE7nII6jovtAX0I7MNQfBwrPoPZBWCfwfiEcyv/fNOexC9qRx4Dum1QOzTVCUvxhCYUqvvuNSN03wjAT5iWsPD4hDzwzZ7TQpENsGo+XdErvr583KYc+GJ5Hvvst8n4rSylFNcNyxW36qjW9Dfu5Jc+dQfcCthDYXUf4qq/dRrSvaPXnqflM4HPlmL4Kp/A2utGGhgbjjDMgKso9Hx4yXeWoquMiVFpAjTdpVqTVGDODGcW/4vCQVJATWqfttRrqtUqanFVdbDXLUZRF+rKXbaY6UyRnn6oinU2J1OrNgHXbY4VxLr06sDfZxX8yUiJBudMmyr+vnkcTz/JQP0FowpLGWyTRHUmKbuwxIWujJKHE7JdZbOSy6F3sqPu76zvOCv1LIWXtMp09YPLPSrWdiN/NsYnkbk09rMniZmEYllnezLLTDFl/mUlEneDo+51s0Ir5h4Dk0Pinhr94JQ48XHsFGbwP7Dzv2+YG9l6BOikFpi30qrT4vZxVZLnZ9aQp8k+pOozUrN2WAfcq4x7mQzze6D8Snt26e+Qf3h8PQytPATtfCha/2F01Q97/40/2xyd7E8w3FEV2xy6uMI4BthXN4ejshw2ye7MFgpzWIsMrqvSWKO9ngSIpixENfnq0WcxwWSZxtHO12DbH4a5/xUTNcMj+Aw+cel+rXbWl3XEv5NnjmOy8ozuh/YeBWwIODvK7/dmx1AlQKi1WvTie1u3mE5/tRPcVHtonhCmBFAY6P9juvofy+SECodVpXLT0A41MKCvg5zCBKl22KyIiygxHMk3pKq2LApjXMXC9hwpklCuemg0Y0WpJ01b2lRRtyjc2KZQ0VEKVC3KOrafMROTlJNoTdN4EFZ52YZJV2ySAgEHhx3f6hWg/biJUkqdMBuS2uPmbvXG3j0dMilx5d2SAvzVYzA9MH3Uaqrl0g1Iw2004Cdq4GebGPh/RJf8yeYYG2sX4849i5K/xvzB6dNYckieVW5LZXc7XFqqMBQXQoR8jP3/AeCD3Cxs9XDwj7D6tNpXSaQZv9GBlf+AXn4DhaRsPwHzx8fOikHuLu0zUxXqYC1xn06l3YSplxw9R9cSldUYZRHyDK4EXdKeGWt7gc0vYe10R/nykKrWTBzktSfnzL7oRy2/DeDHLMVSJsPPxQqKziod49ppxcA5a5I+28lifepymexlVyOdj/w1lYhBLolNLRbF2oCl226duGJZKJobcm25LWy54wb5QPNmPxip6TpEpcX6WYYdm1G4S12mqD4WNGf4C5H9/pJxopYKCyZDK0RD1PGs1ujj9ORBzRSLUql+k+tMA6dQu2AuQRt2kOpEJ1xfKVtfux01WaN+CfUWuLEk/pp4ILk5dPBrz0Yxn6Wu/BXj+Ezz6sU1Kju+BngKCd+KMpnbNolM5jUyXJkGg4WYeR9jGZvJKqvScZPKTUnGfhNBVxWZazH9yxmyVLi7Y5ycCAb3RYOsDfblM+tUUkOiiVpKP2tm+KtRcf5NnyO0ASvPfrsZf4GKllpD+qz0SJjsCkSOa2qfEArg/myqQnEz7X8Z3EnmRQzXrBkB+ciMu4j4/XfxxT/egB15YkjAYjdmJeZKzErJTcFNZW6KuWlgPIrSzMpsPkU92A3MleBKyZUQrhOOYly8vptibi/GFOwLsOvimyUilbs9OJcO8jquyXl5X36D0Ra2tg6+mTWpYc2qoW7MeaomU4eiGGY2EGIrW07W6dL3YSe8DeznVIE7m6Njni6g1kHPpj6TXi+27mDQvdmQktA5R4Hdox2m4Oho+cyWxp4h6dk8nje0gkxvhDO/uVwTvnr6EDvfkFWWZbda5QlaIUmypgCZ6+nObWhwWtQisvbg6Ah9DBe5yiJe38wR2+gN3WR9k4SKV8fRTl0UUmu4Z2Rb6f0+6jspLVm6e9NwNsJH5kn9GPT01C+2XSOLng7rvofaVpBfUg6r9PcBdxB4hRodc5ahHDHR/Yy6VeSIp3Dlm+GMj8KZmwu/iNdg+knkjk3zFFp7bC7c0ubQDpaKsdQ18+6CnnUU/M532dxmmB4rJgVMff3IzkUnJ6lpExTsPgT+Czfg/XwM7HFghzqxwwccFao1S8JV0cjmnPd4s0eBPb5d9K1aBybOzD4ola+N9x/ndyXj7VxUBSpD7D0e+Br89k54xU3fLzxzGxy6EvT8j9x9APL27hQI47pgoKmK0cWugfLCA/8YT98KB+0ACzUrmO6Gl19/0xEppxdwl61wzTYotoTT/BXBNXvh0N1wwZ5Fw5Qm9UhLtanN+trUndjiXrW9cNxsZq45imqt1H4JmYTXr6PVPzSzRwo7oqJkRQ7Q5xbgDO1iOY+g+ou7L6xOFtcndncAndzJ3MyD+y6YHCrOlf2cM0yXIP3Nki7v+V8MxLvnHJ+GUgxoKU2yUN2sC+MmH3edmRr4vXsQAwmD3YFlgi2bA9R2Y6Or0hfNZtsPF4NsRd/we9c9HfOwfTMdyqrx0Wx4QA45l3XMrdHVQazU9bWUZMnpBebuHo4lnIYE/nJFwGYTFd0BYmb4yyX32n2XHb4p7ezPwdo/gvvfjdJjWihxmKOZvz5WGbbV+NOJsO1UYJMD/DNOAfcQKHsiBuGETGA+Hs9qitnfoRtSIOLF3wS+yRAC9J0fdc4PBA601kK1dmyvfRDO+6uN3cNa8vPqraA8AooE9O0BCsGkBC7LF9p85nbYelwIza3Ir0IwEazvgRdndPNXJ8A9Qd8P7m6GbivcQdhRQv47sPZZmH4Qdv4bvOL62c+wuoNQASwev/ijwO4c86+sxQZhgN9JocMon39ahK9pRrMH7YXzL1tsXVg9MSRs2/1BdwMdg7kI8CfXozO/DPogTP8VXnZR/hrPPyLkY1xvdazY1kLswYfaO+uXw6uuWqw/n3sibP0esHuC7gI6Dg47BNPWmHsg7LB15K+DO10CZ34O+ASUH4WXfGsOwO8vrjXLllunu0eDKcatkkpa1NRSnxJ+WPAv4H66ZlSqZNYcGOsFSHZLKs4Q/TPdDW09DLh8QRx5d3CH1ud4mgHeM2DfhlyN+hpxIoRqOu+B8gvL1z2VJ1/FhA0kcnbkjmYxXjJrj3PFQHkViw1Uea8q+cc6k8TybGQWzDAA/kwe9ix5ItFxVd5zA05iTJrhpDruaINzv+ekysOuTQLkDo0V59LTHUuEr2lXy07XEmsoXMN6665V4WhLW+fqPgfD7juSUtLWqRwihnB/85zK0Bgt+r9A6O3wP5/ctHhV7m8wfs6wbVUI20K55DmCrXsqKDNMkryHYgfO3QPPhzc3wC9+2PAnhkwrVx8NNadaBXVgpPQlVPzr8tz707cCp7ZCl8MCUK0EU4n923fNzgjFzTRtBsqKoizPVah8AvDZ/gcPejjmXkPjOcnACz+B4t2Ip7ZH2It+ACufATwIOKYahTUwsQJhPwGTZ8DBH0RnnAvnf2jGnT8F09Pr75S2o/KoOpkrJU1NHnEk3v015nws8SYFr6kMseL6T+mcX2BmLsKzbwU7ngLlz4LdFrS1mUw+2brtB0A/h239Mrb6Bvye/w9eemV7BdvyLKz8JbRlGjXGzdZNMjxsBW17DjBHyen5p2LF08A9EjgJ2N6QXFHs3JJShzYhUWHYZbgLpbPeAtf+GfzuVwYAfiiYZF02zhjeo5UnUdtslLe+luCy2dpedM5bMD02toesEw+aZMR16ca2N2RdByBtQA+yE2HPSQsDfLM7k9QQymTttdjs7uGCmAMw6/Ncvxf5fwy6zctnIcQjwqKhk4pcAsUgO+gtVDKuGraS09xYLmcoAgd1Rnk3LMdmsJU9AKOYU2UL7/fL0DW9m1V9qFoDtlmy53kmt6zBXFg5dmwaptbwplYIWFqTtWFN8mEX1vYPk5OMpMDbkrIle28FdmuzzKMvSkfnDjxa88jMzK9L9pbNraQx/QhM/ke4+4XwC9qFzIYOflLmeqhgWGutc5j83TZ3TuJTVjD3E5JcDLZPyktYCAmsU5gcyL9nucJMjzwK43a9vaue4G4n6Cv7fv3VbeDvDe44Kj4kAdyxJszACY67D3InQtn4HCiuu+Xu5kKrW3A8C5XPQRwZ8ol80OO34GCFN1Zj2XZg9nBwd4czfwXO+6eB+X5/KG6nUBg1He2JGkHKOmsCduvqdDMeVUWhiwmi/ACcMwzuizMeTFm8BIIQAJQ+1hyI31N2g24dcBqy87Ft34tWnwprl9SOm/n7gTshVTIMISAOTHugvGJGvznc9MmoeD6428VixhKlT8NjZF1cEvU9gqriNpm7O1bcHQ55AnreKrysdRLkuotrukZogZRD0QHD6aLT2orcEtP4+i+wi8M9WlvQLpdMC8Oyf8PsukBHgLvtwgsb3L4WS8t9T0diT10aLNcXrWQfRdlu+zpM/mtJ+8bqgj/ZEo8ZAKEZ4CKNgmmJcG8E3FvRKDi42fH2Q6EIXTURLNQzMR/3+q1Lnhx3qVXKD8kCbPlxOsP3stkEwuaya0y9Z7Jm7bP+MOy2SJ3Un22japwuYbhjsX4b0JHqTkzLjAGb0e82vIeEREr7Euz+Lza1rV2D2TtaSN4YTq612U6R9dq1AhweodvFsIpNasfeA/wDMSk+V0d81lvNNarcCXrrkj3ArYFjau1xpe6cAf4qmF66H87ischOMJUKaoA+Zh3Lx9OCr8N1V+U/6+8cb6guqCVFNl8Wk1mffRDwO4jzwB8RinE1SQ+iQK3B64GpR9MSOB6z18EL7tL/7uccApwaWfOqsJc3LB7+tlQqkg73PhT7msb/qpRUhhN+fXnGnHssvngzZg+MknxlyHNIAZQFcC5XaaXHNp16sB/H9OrYHsCRRyB/u0BG+TLcl3xToMh9FyYDjubqBNOL0OT/A3e7UPxMXrXWYfX9jrYagQemIQotPK9g6o293uBUWPkTeN7T+wDfXEIhJBhxACjVKiE2vE63wYwt56bU2NfB/qflKXVrytRJtEq8n5RuUc02k3cMBMUEdMpit3TcYYiTk8mo/s4/A8gu2trBSfgETL+5nF3TSVbMlEI1ZjmWAyA722rlgrDST8K882RDEHITZlaBp2pA1Soq3lg2Wdk8uHBpwbvcs+UOm7o5++15FiXRWtIpmymEuEgmljOw2avkDJCbJmtbfaJvy0mWaMutwbZBPGwzmw3ge2zSMCdR83khtf0jIc52k5sv3w+6LtWTXmjxyawx/dosIcwrXvk4YBMXvCoeAe4oq8ZVxSTE2jxWTwVnmH0MrvvIUt2+0x3AH1QrrdXRAJWq1uRi2Pvtff+CyQmgoxo9zQqTWGTw7Svwmj39zz3t4AAufXuDbHLTLgt1ew5eA54Wo4R9C9eYtX3Kfo2eMpwQbHluvwbQoUeDnQSVgrqLSqCWqFtnIFSq1qhQsDIUAPPCDwH81Qdj5WvBjg7VgFNiL2agKlZFrspoqi4dHifYugc9Dg56Wvj8+olgt6rkPC3dzQKe/i4UeUl0Z78E9py4WcRjkIK6HVSEsGlcqEiswoWfowwFvil2X9dOmHqMrbByPjzvIfXoCO3mk42pcWZmASR1Fpfs0WzdGX6Zi7MQkl7P/TTYT1SzEiutyvNoRjCpQkFnF7Y8c1v/uR48xy+2khdHYDqiduqCgyFraY0MAQTL/23QOdDXYfch8FLgm9q3yqEHd55qu8HXDHZYyKPZZSFZ8gpCyMVBnfd/Ld7ZDms+v0vAwW2R+Q5ozhScEp2S9kPH2VbF3FXRP8ViY3QWA8kAY5+L2lLOTatihfZuGnk7w8eYcetVfe7UkBuWNe+CWjZJ5dZZA8Rq+YfhKZ+Jqatj8OtGq2L6bW4M+02IWI+rNyzloxqHnN7c3qHMfLVA7n0C2c1ACdt/yZh8E+yOqlqql2s/vK70YvCtvec2oYd2GEwPA67afG20ejjmfxTF5OFY0bwSV6EJ/Y1ayXrrDad9v6+Or93Z5Jwoy7QYTazOAugiePl+VDUsTsX8jqYQTBo27AA/oDpzxDGgEzPrkYVkVfdV4Ocw+/Vw+uBES1/axROi0jff1XUUAHmBPRzsdkByL9OTwQ6r+7EiR6zMHF+S1ALo7q21p7cLTb+YGUOnYf73kbsVFu+1Domt8vUSBqEqrVWvO151Rc2Q0f0bcMabwJ1gpkNUV4lKn9kB+hpcenXmfu6G988HVjDvEwWOqgCKC16rF7JrMdZjGakJ8tuh2BKiirxPV8owlkoPK4fAym/CM/8TXrlr0jD4VkGd/sJhA3sSAzi3VegnbvNyyx0zaXyBcMxliYZ3U02y8mzTgadODD5JDk3L63RKQkOOWHBlOBx0SDcUSMxmo2fu/9mcilKYPR5tfzC228MRZR2X1oKgPqn6RRMPHyRAhV3XlN+TDHYadnBcEK50oUF2h1qX7BLa1VSYQgYHWXXWZ4YTu6uV4sRaR7wu7kU28bt2TDUQB9wC+rJaNck2mAwaggbzXtPQJq0BcC9T5Wg048sZ2hSFrsqkmoINJnQD8wRlWo5ALLTXzKujNxEoUdEqxjAEcrMbVactKkDQq3q0hCo60nHNJq+8zutAuNqQExBBUHU06oSux/MFbh52Fea/IdkdmxM8s9Y6PWNZ0gCZUNc+rJXwtD0Ag03pBH0/FN+D+drtCwRsUaUfKkg8m4EuQfaOJXNQHNj9lCbPqw0dwD69fy677oaYUCVzNKcchkoP7hMDzP/J4A+LxGvDHFhhiKuBIzHODEmopoi+rwe+i9gJ2gLu6ADSPQGsqj62qCMeAkN/JOgebYBvtwVXRF1/mtiYSGrXxGY1mF2sFVDJtFX3G/8uXQyuE4XwskOw618O7i6xfUKl4Xp+WRN2oaJSBvg2uOsQK6BDsclh4XNlwEXmTkTu8VCWwrm63SkUTt5rwPyZfo2O0wvwvxqcHZVWF5Kp1/54vfJ9yP4Bpp9GKztDzrBtCfH+9n2oeDy4EwOgh6Stqzzr7wd3P+DfJgwQ0NmNqJPF32X6WyyctXbuGDy0zGuJXYRpJ9iOLBObU8/pNlpHDjGUuEobQoAO6x8DZG/oCNCOFlxdAIdq3oaakvuN/vKJmJ1YS/MlbzRS6dRE3StNgOmeYlia2qGkpoJ1gGyXvrPMGPTU8YsbXvuYmShuVFOs44TNt2m7tHHm6nM0vpum8GS0Ig3WlxzgH5dkJ5vNy6ecJQPY/5xn4QG/VHaoBYA/42iidbrWqfRLu9p0O4W5ieVQTAZbMjuyt1HYAovUDMZa3QQv4xrQzaTQzdou6dwvAz+cbJp1ln2PGBjak1PisFdtHQVgUGzdhOz9BOPxYAdFqriDParl2hTCc/z7UPml5XqGLceg9TsmOeId+SztBf/Z/WijLRh3TijclDCIsY1+qE1OBbcFUxdcGOhKcI8j5AAacCXYm1Dx97D3q8B1wBbwJ0DxRMw9BfmVEIbUy48TaCWG46S//iDisQ1z778X/HPBr2CFatwc/m5g70X2GmA9Eok+If4E/mp4fyfUadcvgz3SMC/z1gd03qBwUVr1XzB7E774BJRXQ7EC5VGY3RPxOLD/Bbj4dQ8DvtU/iXeBrjC/juwz/SY/7TaYPQI1miFNDKYc2G6MF+Gve/UMidG/h3PfCP6PoLh3DfJbUkDuUFh5UAPwM1uvWUb9oruQLMDsN3/yS656oG8FD1UHNWvHEGDradj1Ft7sSUjoyyPgnK2Bzp5lK4eh9aLSzu2pbgwlb3bvaRYD1FzDRy1hm8sUafiZZ35PS1fFN0v1DA3oGOCgWhGiywvaAj/PagM5Qm5P7bgsGJrkOuEjXomWFe2aCJn2U+da9T2mOsnrmwDZOt96rplM7ECfZD9XlX13m0ZOKNh1BgfbsGNt5GWWBkCt1W5oRdHEDLRly2c6fQtw2Fwi0jb6t6pomq9Ig50wuf5mwuAL6bs5UmSwXbpqOt0xE8M42+ONlVjAZ5PZ3jsZWx4SYo3buqA1AaB4mmy+RPaW5VOBu/4kbOXW6pQMalZFdxXs/fJ+fMHhoNs1RRfTceNAfBMYyFexO4d9yqsdBuc9ZkdJ/lEh9lv/DcVz4az3ZSDe1+H0j8Bd9mDu2ahUzfWY0Qkk3tFxcC8kZfR1/pFBarKKxqrmQxn3SPsAnPWWDTg/t8F4KtFbsIxrCIXD7FvInQH6a7S2s3ORLwWJ2dW/wvEMxAtjBu73IHaD9+nJa7y4A7s2r4zk7ob8icFhdUloEDHmnnejK34nnzOR2tmfgNXnYfb3sYCqGpWmOrLiDnB64XrsUReYdaKDNbCBZ5FdvvrNklp5PTGGpE48UZJAOLg/qc8WqzO0217ejuD9zl3/d4QkRusoQGr+gV4SA22DfZK+V62gItsIfLAFx0NCLYSwJ8uDmvQxVTET4figV1F51qAcCuNpJY/71PHwLBqDj9+Sz0K3+feRURWJKXEN2McIleuW3bybO77meP7Dzlct1OBg2yah8Q9WfTLXOdkcnBOa1zjVmPDNurR0SbYnFuC2kRnm3XFhA3Hli605tr78+8hG/GNdjy24dg38rbXGp6eitaqJTWAzhuhMflhwfIyBpieAUUMKM+ALUL53+fp326khIkCJJHMVte0I4Hv67f2ZeIijQ7H3GO0hJRjBfx3umnGIV52Zv3MSlZ3AGCE4IuYA/jvop+Cs9w6PyDeXiN9BXGQ4Fw6gklPLEB/PfM2I8u5AERNnQ6RKfbqt6calRO3xwJ2qWHXhJCmRG3YOuBLxJDjnj6EH7lNn5Br8516M2Z+GpFd/FMbx9VpUz0lZdEa+C1suydzTaVCspEy5tYC2//h8cF/ZpR8CPhmr3DpYcdjEwcSF1Fp3FLBl0mWDzRbcdzJ7VDektEFpVqU8L7GtyOJxh6wg0XRNKZcBnfNOwSOzPLALn52gPcUCJPGWKotbqTPcKmk/AKLSOL8caTjIli/Q9/OYOM1IbO06IAzEkuakVzJ/b+c6ZFiuoQTPKvRBLU92Ayo6MQxD6ns1lslEn6V6ZLk2kJZfJvJ4g8tcNtFvVkGeRcI28NG3VeSmdm2iOJ3O7FETotak3Ha84UEfMZx2iS4HqCVU0ZGFICLf4XZsWN3WmoQZDbSBVaTYzdG829vI9c6YIxnm3lKBj9b8iZkLTckFh9hkDP4zt2PFIwP8lA+qQD6ulbVUbGTwC4zyrWLtiuWbE+VdwJlhXjGJug4LDfP5S7D1yv3wIG6PlUc0EQeWxMM6gM/DT+ciJw4X7jYp5VuvTYqKIFZehOxXYO2i+fexdil2zgVScQqUdRhvzJgLBckcu4Zd81UH7uQQy28KpF6Vc2omdD0qLly8Xc4/BtZ/JsbE+9Y+VD+oeWAV1v55sWu+uURnvhImjwZ3XCyCY/36O9XJyTSjoKNjG4Ujb/VZQuOsHr/4M75+J1p9MW793ngfQwpUgtsLU1HwVcojppMe0FgE3Cm/ALVxWUtL0i9lYlgXUmOuV5nHcohFqZzUDDTTYR7CTjaFrQuEK9lKGECdGzJrBT1rCBzPigTIhFQMqfJ0i8QaZIsTdybQTAcgC2w37Enk72PIscgXVFbqkG0A4OMGb6/rYFjGUWHG56pY3M3A4KvTDgNqW8ZsNZW+X9fyb9wmA/hWyw5TFWnL+K1dqnveEKkJk/nz4sa3Q0VQKEjGgGt8lKGRbFpsuWieeQI3I7Tv/KRWnqpPEzPpPJm1Upq1xqvJmwpftMlUhw6+J/j7V7tSeP4K1Ie2krmIIcvrVfi3LF+G36oDu2MzfZ3AW6OGZYScgf0KK7pjSAZtQkXqcEjzwvj8AKg+HvyxDSZJtp2gJFMinQNrn194J4Dr2iCgdhsMMcUVM3Jnth0G6yfH5bMlFRZcVfs2aAO1AsrvxfxdSE77as9HIJsY+PehnX+yseY+70vY2vvQ5GdhamkSYx3mGwDSV2EtE4Kt9WauOjWhljWq+ik4+/1w7ptYKPF67V143pVvgvCfqKJT2EAWYr9i7SBbzwwU6Wz5K9esHyLc9pmoUR2UmwVvXZrZ04qzlq6DI+cfw3hcnslJ1cFmg9qZQHIIS3eTmawNWoecwJn3o9nD1WZ9fuO4f/Fnr2saVL9YNESnE1CQFs/KTggbjJ+tM5frhb86zd2y/JtyV/VjniM2kDCYTQCimjdmsHuzAHwHbktbGGE2EztzfMq3wH2QzzSqKszLYxesY3e7tr2O+Mazm/G8Cz1Is85uRdOVmw3Alx3SzB/X4JGuaMNG18G2hsMU2Lm5Gqb4CXCHonXfPH86Fyofz5ngQ0z55PI9w54jYMtdosxEW8YfBd123Kf2z4HQndsa3lRlxAuJa5AGEnjt7sBhdfhTK1x+Ypjei678fxu8oW09QRarADvXUa7PyDVYPx7zJ9f5ssT5UMuL6+voqkWLoRiUP4psG5Zk6sqQuVhm1u9Ges2MRNZZGOIjGD/bKyVtrUp+A4nNwcmx+lTXKTD5Vg2To0Gvx875CaS3Ah8DLsk7C4vZJFkR8qojlmFAuwIynfhKJeoP4XKbIQbfnUBdEKSkJ8hsA/HVZmSX4W74Tv1LfzU8Y+8CXmjni2qm2QZBeW7hz4YU5eFqNzSnC+YtRyAOfHcvZSB3wpA6iMrDZ81yBBYF9flnN7NKijj6TFpUncRl4tLU/N462n9DShgE9b/2MxqbTf29H/405FAx2zOcpwG+9LbDYPfE5gGzgaZqy6jWFa4CuK8Gq7SENUXeXKK7XtOqCG0+Hwo4BF7VXz9aJ4cGktsGxSYu2tQFaO7YpJIZbS3xLlAamBfZ31tDfhtTod2bp12efxTwiOaUvNZzT9rDV2EWAv8mWNu7fM+x9bbgT0yFIdR4+GZmV0nl5/f9+nsOw7bdvo1Dkggd06VofYA19/cOEQKhoqJaWEcg/mzxWHCA1R2IW+U89hC3Z5dKu2cw+LoT4ugghxqFaKK2flTO/jK8YkEn9cVHYnse2I6J95hZrBTgDPRR2P2v+9jw34hj01rdWodvl1MYKrjlPwHuSmGHYxZjGYuEyfces0PAnohxukmXAJ8X53wU9AHg07C2oZyNyVyAOIuB6uh698nvyjmxmNywzPjejkN+W780zwxR4jrk3uo6BT00WaujUOnpX8lCEdbFlFg1elhWcgEwNCdeHhuOfx3EYkOhOsr8OCv0YBGN5wV+v1B138yXK/UsJBfkuBYynz12kFr+cvdeml+lDJvMaukLNWW/N0WhK1n8/35/DhYXW6SfXFJ3QQZ7Ngnc32l1AqzNX1fVOvnpxJNZu4GkJfd8zK4Om2lUpZKrI99soP+zIUvKOf2GvBfGQRT+eMplZGw3atdth0NObickLubodp2f/BusGkq7odi1aZql2P59eN2lycOp5kmiqiKAwjAuQtN/WdK18fbAoX3+KtY0M10Ke762Hw7EsaFQVRIfbdVXOZC+DCuX9z93+hZwd6rWGLVCNMyBvzwA4I3YyhHY3uO77HDzv+XXZ+ca2KnASi2HbVV4jo8VwSZfXJzb8ycBt+lqHai1kOpdsK8F0fz1TS2gqoqHKYl32AkM9Ovhn8aueT8qHttI+cs6RepkTBWwuTsZ3MmGHi54BuirsPo/wIeh/A8oPjeP3XctHNvBki2Y23mPyKTNWlckxNVFmcEVS760nBT1VtWjnKwD9JU0WjVJUuAthitSyS3ogWlPnWad+hjKlRPOvBahDm020BpStrBBOaUFkHmuCqw6v9eiyJ5WyJjNpUjTN6pzsmah5PVik3zaiZlPpFL6IQkpo6sWiqlUC7ylRF4oobzsMplHxFronTaeNQYt0w9J0zXrTK1GEF+32iQxxDtk0fnTUDv05oJoM0CpdFoyqFTXerPlU9EBjKsSp7WVYxMrqtep7BpKqrUMgVSPGSeD7ZR2e24WdsThUNymtXHOWnu7icoz1apcol6iXWymEB2vn8DcJOg3+lg8tBF7qRYJswKDf4XJctZFcP5OQamlimRQsm8WIPsCvOSq/fiC24MdGdCVt1riul473QX5k427HA46uR1O2iq8cTHsumRj97Ln1MjgK+5fEbt6C8o1xYxcAxmoyVVQS8rSghzldAMJtuUdgMOrLN3OWuNAu8F9bD/afXu9cAd5y3AeU1fa1ZX9gluVPXsXci8PpyuuiJtfG8GoUpCTZL4U01KUJfiDzbiHGT+L+VcD/4yV/4g749fgBcfOBvhdInKIkRXZEIt0AWrqGAa92rg7LaF2c4/OvWsdJ9cq0pXTu7MBQDMQxoNL2+viBQH+9bQkFhKElACjxS41B/gPShXOvlzW0ZgHzJVhteepSMwC/akiXNdZGHIokh10Hw+WpnXcrA1UsU0OfRaKoa2eQ0JmW2BlM8QauyAkkNAlNsdxtPza0h+itsx89YDtFMhrRhXwdkSf2utKOmar9ANcZDIXjPu56YbCt5CiBu4M8LnB0zwap4+IG+65lCpCG99yTsP8iRWJY52Kxcqtl8bsRd8q+qUKXXdgdi3suXZztMnqrcE9hLquoav3zm4FFOH3SvzD8mnfxyErO82SBUAJlo799Mn5xS5nAtk7IW0nUd9sQIgX8JmBD54IOrZf66fOl/wK/PYGHUJ3V3DbUPhioeiY1SeRM6pPn7MVdFJNClhHLRCuhMkGCpi5k6AommChkiQ63MCugHJ/nMLj2pt62Z6bVlwM110+Y4z/J+JpwLdQURDikBLWIwVT3hIy0gtfSmUZcic4CvHDeH4f3PvgBT9HtqJMDXM0WG1TOfbVmHHcnhnYy8g6NY1+FOjudbZAz5HJxNdb2qldJyDBJ01ojWHaC35BgL9+LfLlYMQQzE20NYaBRu8C+5Ze2ge3G7ye5gH53PNpDoDUgCeSO+mw+QAjs4hMW6c51rmoLeoZpX/zaWGAHbBlx9Lvx7HijOY5dloIx+X97v3aBG9su06Ccm7CcZIsP3Be2kJ46iG+ZWRep5ci9oTHs6wTjuasRwOLQwPuPKD7wouP3vT43vnvN9xBhski05hjG8x1C+lohiNtajFx4aKXwdbNAfCdPRx0CsKHIoQpk5Uu1IUFALvrw0uKJw6TdAcpSNzKWhuWobKEvftTwdaBu1NK+FkVlhW0Lq+D9a8ONPLtMQ7v6nRUZAKUX9iw42HcqSEo0vEpQ34dyi/O+PSRAeD7zDJogF0KfGcDd3N8C96mJ/aBNfk2bP/uPje96bbhmmXIBYkqaZJZSAzWRfOTd9feAvZYzL8H3DpMHOYcZq4PdLylEL2e58ELjt6L3Rncn2JnvByeuT0D8AcYNs0BdbNATH3sbJuAcCn+F3CaUapJsNUApZL4gjNAaJt/rD2Ba6BYUO7JXY3Z7pbHUC9685ByB0wwAHo3or4z65obGRcb2dyHAPrQPaffMQT+O11irYssKhetsi2BCf04o85tasb9tp5JAg4Bv+Qg5j3C2NvLuNcMZ1Kd0BXNcB5rv1mCqzcJajs4lFCft2CmQdQacDypEs0ydSCWLskWYMvXQ36RWXa90b7M+c6pr3lhOg3KB2xudL+6A/Ej9b6oYQJFi+RbWe7zFdFgF8LadPnb5JnbkT2RIBwfHJtko21qMlkIQZF7B7z08uV8lr0ngp1cL3hSoCqovF9dC8UX9/36Ow7CdOe081UDEgfwLZhcNDDJ7oGi0hc02ExVjHHxhY3dy+lFC6C3yBxnGN8GPyOZ2J8MdpwhX+/GSpKrzS6EOyxa48Bg/bhmk/FVSnM6Rb4Le/fR4V11iDsFudZ2AVRrihAuGE60+p9Ij0J6NOgPwT4HthMrXA34VRUOI0OsWwNRjRKsQMWzYfvL4CkrbYDfzYDtADFtJLSjFT5SYK5q32UN0VndhvEksG0tLsSctSjwnmPVASuW2b8xGUVa5uYK2H3Zgjd2FeK6drCyz7f7AsfeXawwL9bdbF4FzjZYtlkl1W0GSJ/FqNvsMdZzMhaN4U9q9jZDXxuIwVfZrhpZR/v1c7S7yeeWcbLq+44TxdwEuNdyb8hXemBv0uGa218zTgCtV+CpqjqIh8M3UfVSr5lxX91jtaHxXpUGrHM8XHTw0XIC/L2XYHyrdXxY7U65/iYThGWdZrHkCDQcBQvcNozHBBZz09q9Dd1LlC2+xEzWHzrqVyMfXJZ8MvRqdveCzdEkh98X7IHVEys9sWqpGTsDfwX4dyzvs6zcGeyIgOxjkHKlCGMOzH0T9l6679ffeRzYbVpH+SG2X+H69tV8mMjqBNNda86wpY7gDGMXsEHH48TDgJPqxcoaFkdyIPdluGZG3mFxl8iMqJXwS1GRIBcOFOvKYbkC49Bmjw17SMgBqgHN9bCvVRPWjwXdtVE4jIVIY3G/eJK/AQdpbTesvQvWnoLWH4rKR2OcBfwD4gKwa4VzUBTIFWFikFF29DEsqwTsqXDYj1XfMElx/jBQ7Swulmdl++S/T1btpc2xvb+hB4sqdLagiQTTfCYOsqx49KtNobpcLE2tL4tdi3qjV4X4SY6jjsHVQNng2eC7mxc8+GjW2it61+gWs7IBzN7Vzp95vzl99I3Al1mOxRx/pu35bOSYSWrfq+VTMDQD0GfbImqOBLz8g0i/x9KGqJwOfE4tr6o7B3Jh40Na+DkFrrrxtm+iMB2zwXE3FBI2VwUrzbRf9AjsxrYLrkZ3/wLYPXvH7QOnbnUVx6HUJdoqH6EFSsAeCeWdgc+yCc2seKLMHRxqg9WVrtp+oVoYIq9c11v3XGvQgHbj90eK8cY090iC6oxPx4SZ2qfGVgD6CFz56SV+mFODDGVdvSzpKwN0EWzZjwq2nAgc1WzCqRKTAfbFfJjI3iNg6+1Dg1aKVw2BYPBtoYs3disHHQs6vq7nJ1kTklQA+tJsyU3dISQjW5k6tCH60zxmX1ocDxxvcMnWzDLchLgYDi7dxwV05T7htCKed9Tt7WLH2k6YfmXfrv3iS4FL8bwHVicwPQrcHWByb/APBLs3xh2iWE3Z7LcViSJCwbOVLTB5HPCPVJRQrUc1h/ict7e0cj1UHbBXOdXLaKsOcz8n7LDQQrGj0qTZQRQ73Da1UkS9apdV2vHH4ZULSpbtvBa4oiWuaIm4yDyGdBaonIV6NQNwZ+o77RPyyiSizrzfWWz+gqFks+9ToSjiwgy+NW6cFXlaeh9OWZpflADfC+fdZnn3sPc4WjFN1uujdu5Xpu9p7S/pYtypsnEFm8Oui1mxlvejFyj6lgPDZkGcp5dzv1T25hLTp+uwzK7/nBkb2YWkn8FfT6h4ROSRuzWs/BKb0lbvIfRTTfKf8utWTo66SxoMruMuhEfgvrUxRvGmspccYaaHWkbFI0oZx1SFUAgC0zsW30dvbJMBpzVBRdX/tQD45/crbMpxKrCtv5LIonN0YX53KW4N3LqZm4ZRRLDiEHwV+PbGbsZORDpcgcm2Sus/ST783AySqAggNnHnKxxuZuZsN768eMPNLxqqu64n4qNEDYcA+yRgYWaPDoUMfS1rXVehDof4V8HWi/d/DK1N4SXfgvP/HdZeDef+LEx/CBW/hHQBsiIP9Kp70e3hKduh1sHXXMmK7JqSFnbtMaimpPCP9lmv5Aa1lftD+ZNQKhvrOkebuFOZOYskY2yWSX4P2H8ufm87doF9q+k0lxX10QaVcQbfpzyMTQ8MpPnZknVNgAwTNeRTtKrnWnAKm+fuQPcZGtGaBZgs115N5oraJRLn2d5aq33WDRizcwZoh9NVAskxFusEnH8knj9Yzo3sSoMTtjQP5RdgCDJ/z4UB2kYH8RJu8/MqRi86L5tr2aBS17KYdx/BtBP8jqoa0cwT4MEFodeY1kLC5gX8Ajr3DXD2/2yeUXF6YeaeLdwxxlQzHb5uW80jDCxR+0WgAtCn4UuXLX2zFLvvI+/uHNCpw5yaGiWVTnhQLbPAcu5+z/I+zLO2YYee3DilHflbPOD3L2xKFlhvqyt+1TVDhPZg5ZfyM2zlFINDQb6umUeZ4ET7LJyz0aJotzfcNiVhg1ZPWb8ObgbAP2UHlLcj0mtpjb8gPG5XQfHNDZBOHrvLzkqYpkn8sqTt3Ulw3NHA1zfomN8NeGT/yFFNzoD5y9D65e0Z+pQJHLc11EjZkcz06xR8ret2zVeCMgGXAn8Oqx/D7B3AyeBVg5p6v/FAcRActA3Y6WoAZcOgz8wWI51mbmjLVl59dRtMnwt2VKLNZF3gNbgx5doqE8IU3BtnGF+DvRs4VlybIl0Ubimh7ROCQzPYQBtiyAYHUPDE0rj7Fk5XezwwcP1eQbMh5q4jHd8Ger7zWWkwWW8DEQs96bkWYbShjOGrm8aISdk5pjIXUdF5BnXaNonIMWRPgtXDl3MjuwstrdGN1mKAvAItKWvdnpPLbwc38RaZis5zdfFnjedNEaQ0/QzY1xo/2dprxdBp6EJzOJ6sVpVVZEdjOitNKFt6HFvc/bFCpxMK2bRPervjJNcmmVOOtsCOou54lcvJB+HN5ZI3i+HtkcBBbcUsZbBIAeiDsOXC5X2cHbeKYRzNJt2gCoe4FrQfCbarW5Ddqb+W+Ehc25WU7msDK+0dgG21uIR12dly4yFvZnekzles8g0qTGlXwN4ZISs7jgJ3dF0qia5Msl0P7qrFb+bNJXIXZz3jRhDjRPB339hDvm4Fs+cJjiNm7lotqa4Gr8h9PhaeSkbt8T+F2Qdh+wfBfxD8v4XXjg/CzvfDyr03di9rF4C9Mzi7uUS+6qv3poylywM2GwBtHYbKugs5TXKUQUepZEnMuV81s5+I5YFn76fzqgnmFuRWaLcD6T/hixssIKHPmfkG4YSjNC3CBIphhy1xs0M9kYqJtlhgQXXCaMx9sZA9Y07hMEK9x6znj0xtTBY/i1MoVmMhSE8W9qPoWITPSU2oFCGCLBaTYF6k15AaSQocPZk5sdGiAgC6HDRtx8N2ckyTE7TByIqOU1gf5xqxX7gfzv3Mcm5kFwDeN3k2MwB8zlHuMLdKlWOsRVYbbN80IF9Y2albmQ89y5EG3RoRxuIhgktha98C+7f0pmdW4M0l3NuAQxTYKWu8Qe8Rj8Md/7+XHb9GYHZq6f25YTC7ZvscyGcbdJg7heHUa6Nqwyl3g/59+WfM6slgj2jGgEdeTYCL6kxtA7+OeNNyqwL5U0C3jut358TZCIno27+6H19weIhbDyF7QY7Ax3nmALsIdn8jM9MMdFfFQ2LRkl10se7OBgH+L2wDu6t6e5yqo/xvwHUzMI8/JibFJimPFn0Dh2AF1jfqwH89n9RV0du2DSY/G+LcF5zA7pLnIP/TqCwrQK9YkCpBwNWm2G36E0JeEt+D2T3N7F6Y3QvsXsB9obz/PnjEe1uFaVssaYEZ18PVuxNkn8SdzyqRneKgFiurrBBPzcWB7XPi8g2zET0S2TmqM3+1AEremAPQcXSmwD9vnE3RZwXXqoqto19od+ZmkLvZFnfkQhOEzH5q3cgqHKBCW3UYYaVk1RQAUCxjpnqGK5NVpKCCpYjs6yv6cIEqxjKWumhGq5IR54YlGHNObKceWVbFUvtKjbpvgF2HrDlbTCp1kgH6Yhjc5uqmxaZ1SP8HVk9cvo3siIa6CNWelU2qHkomzfSlouqBtfvFYNcmAfjXyIJmA61TpyG5+9wc7Sa6p1Fq5pYd5AtXvsPw64MnL8b8BOOZdVbCehKjQB2ycynWHr7MTQLPPwKbvBzcnWsJJO3D8pNfJwbK29hnYcenl3/O2MNDzHBUAelyLfU/HMDnYMt7l/t5Vu4GdpClCje1GpQD3Ffh8v1JKjoes2OaU+6EtQoD4rOwLSMDuXpQYPDVLCyxfkvY5nUZrHx5Y7dy4rGgO1jk9FoKTsHP/AK8ZpYk5TZg0t4cC6CIYe66Nbj7bNDB+jxovWace/KBXqCfDGoz8+wFx2BnvRz5sww/aZ+3pUyMTOYVEoq71IWujgSqByuFlchKGiLoV+C8ExZ/vqcfKuz7mv5PyTAX1DLRV+Evd0GtouOpse4cMq4L+DOyXEleVIzHDkV8lkTWbPVOGK9Cdhjm/cwY4n2BFd3wLMzM+Ko0/Y99uNpFoVCJdbzcOZtjCqB6CbONGlXjHHjA+UbzsV41ArNupNkWlgcnVVEJ61d7kizWZqiDcsJXu+haehlOMhfdCkWKy9oSTEMqNbO6Y9ECoILFk2z3XgZbrgI7PLZP4k73nYpWc3fb3tK5pGTyuIAFZHcF+w3gecvH4VZenQ+YywbGYG6IupS5z3RFjQ+dA79pJBEV82V6zve8sTtvLrcWXVteh6fc/WHZ1q+A3REpqmvb/Gfr5Kd3t5WuHJGagNOj8Xo1rJ4Oa0sIaF+1Fbt6DfyjAzXt8gM+eb7B5886grklwQHl2+E51y/3bHn6Vsx+PCgVeG9VzoZZkxtZr58O0DvhBVcu+RJwGjgTU994K0Vc0xxQXrifCcK3Q+XhgUAI4h0N5hIBZK5ljlWnR2ErxzesQZ3NGuQs0Vdh+wYdD3cK6Gi1KPwWAP7CnD1rGpKmXezsomHJDIG2mXGO3Oo65ZaPwnfXYVsBbgdsOwn2XA3ndXJwpp+ClYtNdlvMq1cf0LyQ24rZy+CcU5D+FPhKDK0hhPwddyzOHor0a8D9kZfMpqApYmt7kopYxHVXwGu9yf15zK5H7Aj7JGCuyiAvA7s//RM48/lw3ifmYNdDMc4E7tOceVPlpoSREI50/qOihqpjCr9QyIfapdaz4vl1k3pLyFiDYgk26dVbYe73QzU374dQSE/abgPgsrc44xB6J5z/tX244W8DF4KdNqfxZ3TarN3Ex5gIVyJ7DfCB8LNKYBrQihOaxkHsrNFcrSd5zDIpIoCPwLT7nuDpGSocTJ3qoReohMDbeoMJCtcJ32P8IvBzC2l/q8+Qzwytb52GbCTWe3J5kDqzU2hV2bT8YJijoDP4xbLYtvZUOPff4ey3Lc8+dh/gW9ZUCER1QtyceTLEXEZaNjo6NeIrloccmGeHBi5sKPxmAdZ+cDzTZrCXtw1e8h140bvA7pgm2Q4KAmQUYtRbSIfmkUV2wp2G+dei1Z+Dta8sT1ucXsBVLwB7WjvTUsOOjjI5bRsqSGgG5ZXI/+3yz5cj7wjcvzmmctDSwI+kh5mBdqK9b13u53nmdow7duILg7SrKnBk+5sUfifMOVj3CbmmGBKzp1NUKmnJLcdDecxwIZryC/DsDToe/naIg6LgfKu+DGgKkzkSrSvfDVV3/Y76piyd+N4L7oHX38D6N7DDdgIr4A+G8ghnkzO8p9Oe538DnfNusF9tJzcnlXbxHrmDMPtNTD8H7vNwzmUR8xwBnIZ0CmIFNDXcROITmD8Y465QCd+31uXLYZpJCHafAvsMFA8gTayv57YvwX4EirvD6ttw5dsp+QIU1wZQMdkKe28F3A/T45H9YHTQEg4xXSjKi+C6f6qRSmzUciY73I1l7v4eEno5FeRRUuxLN/EmvXo0VvwB8BBMPq9IYek999rBLPN3MptTjaHNQNdA+df7ds9re+GcTwKP6KFVGwCSQ7UKupR2iwxzBbgJnLl8C6he9L3h6NarJU851H2ZcdyLVc0DbUGxIEO+dh1a/SzmHtw6HrVOVdcOUJHlmr8DeFtVPE1RhulQ0CvQud+Asz+xHB1zSYwZKepGNuvU+rI5/mVGqrid5C0s0B2biMG3WsxkMfGBWVhtkDlb9kCdv8X8r2DsIB7vyM94vpm1MIx8padGik/II/e9oD+G1Z+HtYuXoh3c3Z+F5wVB0sZ8FRLRc9CGcqQGHKDu+mIpRaEJMH0/XPC5pZ8szj0Yz62qtlFfaCA6hg6s/AR8/RPL/UDFYWAn1sEMceOJZ9YmdB1i//rF/Gk1lErjQGUuVJJe+WKe3vO3Bw4KSQ413xtb2APuMxu/F50WiTjf2sQC7rkO/JyQn/JSxNfBjm2iP1JHwcdAd9uO6Y7tcuise1Uqg+0HhfW/kBWPD8XGEsWO1poSq8KZ3dqwW1eKRyEAwYfjEdMU2URBweY1iDN7x2yBkQf0dbj2uxmscA12zh8b9gAFpl9h4ZKSU4AS4zhwT8XrFzAuB78rnPiVE+AQxBHICqz03X00aMQE0VOp/F14da0Q5Br2kWwlwe4GrJxSTO88US0p+SVg7k/Fir8EPQ4rM/S3b6EyMwulsrMxwgtswnW0UgHwL8DH9n2z1EeDZ05esYROXw1Vc23eaHU8e51p44Hyx+FFJy8hWiiamL4I0zeYfKghpZJWW210oBYfRnFhk29Tb0Pfp0z35XJf6pcsLkwecSroL2D13svTN65FyUszQO2s+gqZz/QKHW0aU4z+dHlSZE4+t3WSKFuJazVp45e8UQ79qMF/mE3oPsW84KKWNIMNHXnENduS0RHCLR8M/D8494438X6zA150HtKLwW8JCZfKz4dEhGOmck6O1EnrrlCtjH6K9FfLr56zug3px1shS0ZUJateHqkMeU6mv4c37F7uZ9p+kuGPM7yImhHVNht2LvdtmH5t36//zO1Id8gff1YJtuWArKS/UxIDkwIGw7STYu8G6yWsTpC7fT/UoVrk3DdhyzfnEmWmd9chOkxj/4doGPAxg1MefAm+RGUZwrn8tbBn4PqTD4N7Q12jpq0CEsOJHOZMBqXkS6ks0bQU61NRlk3cMdfjyufB9LOgW7UwQzvR7suDBb381X8J03cYE2s7Gq15HJ4PW0E6HtmpmDsVdDLSEXGhKBuRlRbgs4A1yz+D617b36Hr1cHq2One/ec2ItddlJRjmpRPvrwx7PQCd+5PY+7twMMDEquoAh8SPyO4N7P61UoaVoZEyrAuzTpliqElyPxOKP9g/7L+p5/A7NJK92xm0ZhBicKudKusQRAubDHmTsZpCZPVKh3karC5fpGseeo65N5rDRwNw34Dha4Ayo9hdgkWpX4CIFd2o87d61Bhr778WdVXHrO7Yu5tcM4TlsR1drNp6qG2J68W4hMSwVc0nq2zXBn6M+y66iiDViJlizjqHMR1SQTflk9U3jPwy90Oz9gj2RsklTE5I3g9lqmo3SYfIhDqVnbVgOOu9nGZTT3mvh/z74bVJ8KbboLy6efeHexvQWcgTWr1L5Oh+GLOurUhQa+0bQrAfwR2/8smmCz3BPcAo+ykJTQsbdQZMOS/gd/zluV/JDtN2GGVMly9ZVUFEU3fgMl39v36W0/AuD3mE1avUDOx/Gdh7ZosDjK7S5vBdvHE2gF8m9JvUNnn2sMxd7ueLka9qPmLYO/8fAn5P4fpV7GqeFOyJ6dUcbUwhud2gqugGFDoWfNoz4uR/xdUOGTWKxOtUkHoKJadMd8VMJmA9oKdhT//L4CTgIOQVwtPmYu3OKug1yt3SfwfMf2vUL8gE2bZMH7CrMSYxlDpJh+ydZBdOysOfCn2/Kl01TO6Tkajg19vRr7RxdTsdSVbWjuAMPXlXvyNDEjOOBm72x+C/irEg/okZq1SYVh0EV0MQLY368IslAv+t/17ji9eDHyym5zVUjTSLGY6/qET3mMVk9+MFYf4afjTbUu4cFIFuBjzi7LNdnZIEourSyn6cFs2cE/HfzkULnMNO5DK6IgDCMF9XBBKITsBszfA2ivgxUfddH1yV8C7hRRibIHubbEsaYEE/PID2v4DWUtrypr8gnlrSuf0zTZtra/JPxn6mFFY1GXXoFTm4MI2e15bRdTUeQlmUHrgZIw/xz73p7B6lxuJkT4Uzn46+HdixY8FBrpTv8P2YU2Yp5BWjy1C0p70x/Cyq5d/irjHIDtU8j7gmOpVLfWR+aYgFPZ56VeW/5n8PVBRxPnv+udx7gJY24/EZ93V4FbRS3bgLCqtuZDKxicH9qpDQKcmAyZVqQPsmzDdoOMxORZ0fLhkAVYYFIYKi2GbXwwhxvPsvC8FhtxfHvM0EzBeMz9RaS8SfFaAuW/BdTPG+Uu+g/b8b2z6zvi8jrqGpG9idns6s1aAm5jxbWf2dDj3lZGrPh6zreE5U21txXyIck5tg7WL0PrjjembQaVhiXhEdGabSnVxt3CYiui4T9KjiPA8cg6VFznb++vou0+F11zTZ+AqUJ8K9qextJrDivY2cWX2djNwNxKb8vyj4JwnY1v+CdwvxcHv6+dUJ5olifetXj1MMiunNZUukwWpL5mB/67kf3v/NXvfXGL6ty5SkoZit8mcwnXkXGo/xzeOiU0Bvhe+/v1LtmpazejM8bUshxMsdWryp1JRJkiwdwP39dR1zN5NkDLI3owNFY+bk+vSf2/zIbNSwArmngnrfwtrPwSrN0GM+qUdqlr75mh1Q8uS2BSripyxfZPE6dw6zLYqIsO64Vjq5choA+FlzVBxxfK3xQuulHhtrBdhzUJkfRCf9XoGtH5z8rc9Kz3IYfbzZvYu3NrZsHrKDQTsD4bVnzSztxruVWAnxGQ62VDxO5h9qjcQsmcD0qGKhBLmPwZ73rr8Y2P1aOBRAWzRPtZp5SI7B343KjfDM03ATjbze4BdUniZ2W5gD6zvxvb+D/ulhFacLLQX/M6gW+93gr8Optdh5ZX46QDI3H4k2EGY32PGHvDhJb8Lre+F9S/BKzaYYLtyEqgw07Wg3WDxxW403YPKxYt5lee/Ga3/IpQfA0qziUMTFwB/fLmJg6KI+YIA34RLrp194Zd9HV37RFSuYbqkvpacAwosvlQUEF5m7nrD/lHeHuP9eX/UzLXJbYG9SLsQe4A9BrsN7Qa+PbugV2Xnf03iScg/WSo/hLQLKJArJFdIFgUlzJlZIcOpQvhGgVl4oRL8V7C9vwd7H+n9b70OXr+edcN6i2hFkmYSf7rgqBt004Blb7TVrF3/uw7o5HKw9VTc3h9DPAG4b3B7qoSEREzaBsqn5wB96qfNKScetD8qB6kwKP8Kzv3YAXk8r/dhXAkcPuBBZVhAdWLLcwmgZkmUs8B2YPwKOv39SxTD6Qa7ygb6jJnh3TH5yQV1g1bIzpaNMcWef8b4Irg7o9InxfhsEPMOJRR2T1263ohVhGBV8ccbVjxYlPeC4u1o9U+ADzVyXze0XQgc3czpGIWhrmrbLIWYFLe1iupFBq8+JtwsOviXAYdEKsZbU3p9hmpKNvzR2oVoWs6AHEwnm6M9yreI4kngfsgqDz2ukXV0Tc9r1zCzsoiaTCIeB1NJ7iRgDbOfh7W/Qf7vgc8sxi7OsjNOxk0ehoqfQ/57hdsShACinJaS/TE5lBpM9+kmn6s/Tob3LTPQOuIP4KWXL/2wKNwPUeqOZkG9L2B610QQNG8E06fRnv9a/rH+ftADXya2/H8hR8YZhTeVOHCiKKHc8tn9nE9vQ+5TQT450MpRxU7Il7A+cP2t30b+STDdIqYOtrSV7jT5Bhs+JfX/jSaPEb4I69GkKpoq2CuYXrCx67307ehZH6bY/kPy/vsx7gB2pGHbomjmXuAKvC7GrV+I14IY5VVXAeegF7wJ7HGY+0GMU4BDY9yvx3RdqMCt/5aV78Jf/KFevof0BijfHZ8xMoKR0KXcA+ufX+w513YDfw6rb4Hp/WH6YCjujnEicDS4g7GqNoCEmAZHjsuRXQz+sxT+w5T+4/Cyr889Z4nrQ6JZbj1mrdq4lQMqKf5oLTzWAWk6QJvSqoMrVuDYI2H9ZOD+mHsQ7H0AIhYM8D7oule34pIdZQ64zxyjaqgAWPLQNbincKCvIP97B3D1uADxUcz9SCzUMJsNHfLIusoMvcuUYO5RcI8fgje/Z2nWTnU6ZV/hXgqUzNORuNkggw9w1jdg7e8xO6NTljXKluX9q4XUfzIep9Xx3aFmWEjmssOQ+1nMPxr4CLb2Tjz/Af5L8P6r4AM3UPz6ngKsiAjW0iiw7MnSHCWZdqK+NRT4vGObpbKD1UKloi8TqQXGsM2Yz5KBLzZHe6xdhdZeiXF/iR2AzJwCAaTZ65d12g0WC3uLVXtqzelK4cPcqcAZmP0q6ONo9d+BjwBfgSsug8uuHwYMT1mBow4HOxlW7ovxYLAHIG4DuKB0JR/WKW/ZOa18alRvz+lwa3Uwxay8jZBg96/A328CptvheRxYEdWPmjDusA4HqWX5qmDK271eeuXyP9cHpvCBT3S30+zP+2wvu4is1vrceXgd8B8H9nlf8h3gOwf2mr/zXUreDLw5zLnDtoki4satJVyxB16zZ98CNl9yQcBRT98KRx4B5SEBkxYlrF8Lu66CV+4a7qfzPgccQGWqtWuA9yDeExKW9xwCxaGw53BUbBeFwZ4S/G4orkF2NVxxHbx+fSNjaRJXiJk6+Fkm0rc3o/5bfH0sa6F66f9Fa48ioKgyfqqqlBqzHFTFycdEhCgrpMqdZwvGwejoo2B6Ama3Ag5JvtPHanE0cVZFTKDsbxItKbt5IMI6DZL9t5lheyW9CM770gEcDHtxa/+I9CO9XIZFNj3NYIvazydgB6Yz0fM+uiSxnK616w2BYZv9jMoggbomesMYbRxKav3N2OR/I3dsEwoWB0VQB7BFmdshGf1WN8ZiWBaLZYRiASXAQcgeAnoIZtcgdzH80DfhId8Bro+V8yxOR2tK2lffa5W2r7UUFqir8IWaCOYnoIPBHwr+HvFm2mo3Yqa61OxnNS3ghi+x+Sj8ZCmHO1s6dIEFtwa6IalrE7XJN98FJ70Zc78IpQZPtQbWrF6S8kBoU+oIWFV9LkjrxmGuMkrPHgXuYRgPC0l0uhyOvAyOvAStXR7k/dgd+s5tAR0BHB8ZtlsBh9bHU4YPcz6zKM2KuVdnfcpJxs4IQ5K6Ho2uknR+BHJLbuUdkT243jS73opFGVEzE/5y+fV3Mtot0F6/Dqwf+Ou+Zg/hqPWyJSJCpsCV8fW1A3nlSY/gtc6isq9saXutM8weiLkHzqaaU93HLmWRFNJqYgAUFlgldEeCC8kwRT3A1H/O7CaiPolcV5ELR2RgE5P0x3DBXxx43ODfjtkZQbdV6gKGnrNiNvvmWz+nXpo82IPg4GcDqzc9wKoTgWKhrAFSd16F325f9+VNYoLQRq34H8PeJtyvtPV7XTwlcMqeHs0alHOr74aqwGoXzlFTYskfCu4umLtLM52sX1G0mzfTyTyvqg23ZQ5drGvigp9u5tV1Ogemt83IZekVa63f79hcObbOKhJ5Vr0MqsK/6kuBaqBGQIKAppunPV6/js59MZQ/iNntExIHc9aEL3VPFa3takqz+QllC0UVCXVaHU0rBeQrwHFgx4Hdq1ViuabNY6XmJjg+Ckg40oIH6jL3OT/GMnltmcVMWSe5zVZYcjprppdLa/++OabHyqORHUtV6bXzoM1PBcJ/GK79H0YbbbR9ZUcJ4Qm1CMgMvV7tC6emSr/dh5j4UmgqmCr+O3n55KXq5xLKMvmvR/JYOOe1lGysEwiUZIh32JQh3Js8nzQcc99zguo1f+KAT8H0JTdM/PraRaB3I5e9t8HwqZbHMsfLMackyfHpcN6PLMUQtaqsjUtIbGZo/XdlYgdAZytrTQZ79yFZdc1LxWsxf3lQIqoy/SvmMKpKBc23VrKc5RLvcv2UjrMEjJt5M/MZz03p/PFi6mWlD0UyvMd883Ock6ZpeNG80LoPr70+bMhhnooyXiMtoEG+eEbSxt0wA6v71fpgVpuOuk8QuOrSCH3N44Y9CDrZmboWs2VfBSvTzdUkZ38JeLGZX48ns8JMLSlin4wLl4B10WrLhcdGyIuy4G2Vie5okXaGDHlUejSNr3UfVDHWPVb9HOaKVTLL5kLUJzLJm/A2q8ZDw6HEsZ6TjKXt0DTrmNUFFpMRJuElChP8k/dXvXJzjIPnHYbcT4TtuehwCzLDm6oQxECavWNQW3y00UZbEOAnR8lt7f7Ma1EWMncddSuQDjCbSm+vJTVl9SaZLv5YVWY2v/AbeTWPWSAwg1MsBy6DRrszuBrZ2XD+N26w3pL7u3DUPKftNePfZPqmJrhqsOiBwzG/BqvHL8NAbcqn5eov9MdOtjZA90Q4La5sjo3JZKZ25seBvwyModqAW4rFO2rs1xeeyY3DWeM1OuKqg9pkFUyuVasqrV0z6mqzPU/W17rcsrDFtvz5ujhEg7jka0YUa+n5poG0OdkQeiBealzx3OC0mj3dZAg/BXE24EzDTMZ6FnDdnFb+JeKNRmEN4g0nUNZRxtGAdKY56x9Mas6aERxu68vvNKM4iFaEVz2OlSwYRk8oqtbg10AXWc5hVayWabV8KkPbUSvpnEyDFA50CbLnwyt3bYohUGz5PtD3mCWhkdky42bAxWD/OkK00UbbX4BvETiL/jH+IrqEs6T/2pRKw9c12nFmFl+1prE1K71yl0wLPlki1zcAYMVw8aEZu0N6n3WB8dar0fWT/CvgnH+8Ybtr14eAL8CMUJIusLAh4GgDCdH1GzzwACied9MUjKHxNRTBaGTKpIxkXAZQWX/49Z3YChJLgnLfSWP534LyU0FBKakv0K1NFGPnyTkgzH6eVr2yTriRCEyi6kqq9X2ZpTkoVMV2ZHPnROKIW+Mw1CdkSgCrZWXOE7DUChvrluHsVCitnIn+arTkdl2dG9EiTZIieq12zkgfaoZL1CDKchPC/LW9kj9L0iejTB3V6ZZqNZAU4qVgvmkVaWB9H9p76lORYE2RmeYAuB0dVHVUZ3EJRwuV5EQb3Hf+OxiaVZ9CxKrPnTfKdw7rGAT5BlZi/nw4e5OEsJxe4IufDEptUXKrOr+zpEpgExb1b3DBl0eINtpo+wvwpQhKitlSxPNA5TD6qRfEOkOsQ/K1y9x761bI7eeVdbg/DYAmkdcgnsMAd5/N8g9s2ATD3gjX/3YmQfAA20svB/cPdOKiN0gvdh2UWp2lqThch1UJ9Cvw+dNvWkpUtdKTJeNloZCxOQCqCuqPklf70X9rlyA919DVFiuQNOgtyttmNREHgMoMVrL3i8F47RQgeGNekbDufWj4mvmbGKAzm7GV8U5mLyLaVKWeDpbUh2PdGhsDZPKcBAU1VR3ZZBE6zRy5GOn/hsRW52Y9dAWC28B5o+ph7XYeKrqs3GfbhQUbEknznfDeRQf2mq4bMfieNBEhCkcYeiP+ij/ePH1/+9sCD2s2+2aDFmWQLK6YJ/m9OP8PSyTVPNpomxjgZxjBhdQ+NB8Yp4C8FS/cSorqojU1jK2llQqbn22okNC8e11E2WDW49UnCz40n9b/Wdr7zBvtmFR7/gYrL48KQ/mdSgOOziB+yJXC9RZp2e0Y58O5d70JEX41XCUb2EAzzzhTjiUbL1vsp4O29s8S52OuFAMZrWEPU69/us8xD0gs4AXMKwq60DUWAaAk+vWdB8gFzlnW46reX0kNqk0ObCbbKEkyC9hv0kSEYVv9AHJnGrbHmjLQ9fzOK6sssDTM+vWMNTDnMFs97tIqpDPA/CxfNbcvKkNqDdxQ+4DLCBUt/fslvWBTxae7rQ8BOxFKiWmQ966BvU8evADTFyiv+jdGG220AwDwawAh63AWszchm8Gi5Dap6vg6kbgejKmcs6038fcDnx2KqZ8XEpEDN02Co4glhS0Us3ov2vULURP2RrLzPoPnr6w+3uxEPqSbyEL5E9bKcmiBsyCr6MHdDuP3Q5XgmxY42ZCT1mGfLZfEmtlIF5ZJXdjKV8nrDy0mfFvnJKoq5plzOC1X9VIbb6NZVVLNhufXUATNonPSMqBdUgPo1V1blBSTs+GCcpvGrgvkg+u388y1dEYieJ1vUr/hBq4ZeKPYGa+T+TVh6zHzQ63DWo/219HJKqQt4irmlG4WCOdcyAGxTr9m5p/NSvO3iZnx32j3r8DaJZvIqduGpo9GKqBUcxTiaUIwSWTp7K3wmu8w2mijHQCATz7bZ3BzV2bhywHnViiMkPeWzTHrxKRa7npDoJwZIJ4ZID7D+qaRK2Y5DeJYNA5nQh9A7hfgt29sPVWBvU74y4IUhUnKhAbNCvnoPJNEO4Qg2WtDc3gP9iCz7a+GV2y/8YF943FpHjNXxaT7RH1jcEwkHqLJDkxs89pemLxQ+L8XmFS2UGudDJshKltM3RDoG3JilXfMNaShreFmnjXFZneTss5y6/e0x1oVm94KyWh9s9tEy+nBBlZUz2IDuu0951MM5hInY8LSrNHNve2Y8OXLwb1SOMX8EbWYn6F1exFQrRngOiUKZoHyjGzy4OSwDMmgzPd0lNpysfvDDu7EgM9IxS/CSy/cXP29fhewB5hNse6i3SYHHfhrKHa/bYRmo412QAG+zQ+zEXlmeIgptsWZql6spJE9rV2IKckt+LPUc9K3qb3QWoL6g9xCYch/EE2fDGsX3zTddvZngb9ZaJ9XxrOxeTtkp+FCiXkv7InY9S+GZ96IIF/tsTkUvsKMsTZ4iuEWAwwbthdcifyvg70LCpdmzdVq3AxUfF2EbhxiF4eS/DJFgebOlSqMwJhR+XiB4TPzRMJqsF8ri2z6JVXFLMBmdJIo54ZeOdW5PfV73eZvJtamaPs5iN+PKk+BO8+oy2x4vZ/3cbeBoaxkCV3Yf2nvZ4PhoUNTvTenCzPjQlT8Mpzxmc3mzTlnj4CVo2CCWHFixcGKg4kL8tKFCzkZBUb5SaZ7PsNoo4223zYhhR3zCgQli5GRYbgP1LaTK/Zjc5gazXEsMlKJ2ROKLKsal3gzoHwXuF+D8y+6aXHE+h9iK6eDHU91ztl6FuvETMwqpapWAZ5W+LOlzL4X2G/Akevw9LNutBjQtHDkgGDShjb42pIcLjvQwutrl6HVp0LxJ2APravcKnmgGSE0GqLyFkzG3ZDjIhYPY9ACc3He35S0eTLPrCunu+lsp8FBPfa3pYiyaCVmuutU2kL+ZgDwAZ69C17xPNhZQvmMmPSjxrHxNjS2anXFLtOuoUJRnc8NvCdXA2/B5Wnx9ShNV7HMfbT2q4kZ/nNy5S/jz/ro5uvjBxUefRv8a6rCfFZX5BOAR2XMHC8lxz9TbhLZz9FG2xwA32xoE+9VSO2KwO2rNj4LbHQacCpy7zVm5wIMhS3MYzND2UCLhWn+BKbPv3Fj7ofsvM+ic/4MK15IpopiowCp1ElhMHGiWzKyLfEYqr9jMpPJ9Ey41W60em4ss3wDWqOopwrd24yddcbjae4ILQ5wBPja12H1FzF7veF+TJRK4kxVS1WmANBrNlAeSuDbAOSTZiCUIcAzTxJ3kerBXSTW+XseeEmwfZNE5u9Qt580JF6wwJRsOXrmboYJtxHk61XPx67ZI/xvAoXqsrAuAYLq53tnOIrBfaQbKjVwcqJOqJs0f65J8/c0S5zYLmnV/E5pWQkL27N9SPJPZbq2SVntD0zxH3j9wpzDqJsz2mgHGj351sybFcPb27wXTUCaBUy6UoEDR5bZSpgZkN8X8Bl45d7T/aw5M3Qt+OfD535tOcB97Z/9geE/3xRyafdbEyIwK6Bbc4sqNQ1e1Wnyk9AexVnwlJUb9hlDiE6ti90Vamk9WkeaIhOa0k68tUW2nP0F+Rej3T8v8Qfgplhd4x5q/W9TT2J1kbm1SELuUJhdtmAbWRWPwQT23HfPUv/pxSVlbjXIjYcaW0kxpM1j6bFE+ts8nWyZ6qX1UqZq/LfkTWdE7G9We8YeNH0hsmeDrg7eTCXhWw8Kzd1f9nfOdBzowbFPjkgZdgTSUwPlBOPqLlUC7gtAb0X6GVgbQ1ZGG220fQX4zmWZEMtvSFkVkFnVUnML7Dw2KwcmuhJjSdEjS+6tI7PbuZZh5mQUagBWKJ0eZSGDVA5m2MSAL0h6PJz3Mnjz3uXqvjO/KdlLmvAPUy3BaAnYVafOe5cGX0BqztrH5R60gukM7IQXwetuQJBvsf5JVVipOx4yenO58ULfSewke+2nDv4se+nl8D+/gfgNsO8Arq+mEwC/USRjsxmTNXCcVxRrHsixjlqP5cGKLVqUNjeHB+a2Oebk31QJ3xa7tJLK2r6JQL55y0ihaFZF3zm4Nb7RGpr35gbwAdY8nP0qZE8Avhz2JJlStqJXDKujOlPlbbkFJJS7la1nqOY0InMDDnB3frlmrFvOqc5V3IsFpaNPsweVr0DX/wKc9Q1GG2200fYd4CdAb5405jACbC+O8xjzoc+z8VD+bmLs7AtV2jD5NweQ4xxYCXoz0iPgRe9c3i489q9B/4QK6z1Jrqpt+qNSXUnyR8+tSAollSAlUAE8D771KnjZITfwg+qAXECZzXVfBt2G7c0lrL4WuZ8CfQIsJN8Gf1KVFxJ8tKRIcqvTnObO0TmHNS1QMsMR0LyEe5vhhA84BPOLZXWcC+bpBi4hvM8W5RsGj3PlgLPt72+WwTrBVt+JeLTw75OcoUTKeaZGfALYF0kIzzi2ZsPzqVNWsS83O1BosSsaQWfJJRWmVuFAF2P2ZFj7TXjZ1SNEGW200Q4AwNdCMEr9XXgYYMzY8AcVM7pE875sZTNl1UIWn2rasCXd6YQ5pC+g8tfRd34e1r6y3F341HWks8yV38HMhTB1p9lIbED/fYhpGoDZFph8YfwatvONcO4dDjxismb/zSarJXEOZvPHgmW8Qlm7+NINamf9GyoeC/wpZruRFbXDZIGmtSpvYiBguFf9cp4cai4ZPRPPrwOQvJvV2e/Wl7BZHkJLFXXTqeooVjHu1ahgQTJkoJ03Y1vsu539WbTzdIyXY3aNmQt8uA2Mtc7eITJa9gv4Wxr4Y+7gs+ckWwbUz/Hj2ktpAaYPIns0/pw3jtBktNFGO4AAP/ObaiPWcE3ubh2f3oJWKS1mwIQSFcbeN3RlMheN9c+E5nRDA6wC9vhYiEhRpsu+BfZbaO+Pwbmv2zyVAtc+KezFVmdpRSoYG0AU8WdTf5PqOmTWfmsa0dN8uhS4R2H628nkRf/rhgFOpuym2tupbSDHohMEa1J97m4QePPiRsJQZ30N+f+N05Mw+0/kDFlBDfGrZ/Pk1XasrbdpcwDivKJvkI1NttmIpHcNm6egYwlgMkWPrS9j1b4NA67YTDR+PVXqYnwaaPuB6qc5B64Z3sbNQyZznr30cnTWc3HFTwEfsvDMLjSvUwtkDyF2mFtVOLtPDZBVg0UZGZ5Dgwg/eCIW+/JqzL8CXfuTsPbxEZaMNtpoBxDgJ6mYMX4lxKln6mrLZq6n6YKWHn32FA5csyC2SKzOMaa631Qvou3ERLNOwcfWYmw9NGsW6F+pMLDrQG9E9kh09vPg/K9uup7033y94K2wYk23mpijiT6ogtJF8ZbpWxErSpmQPNg9pqXeCOc84QAi+yRQ2dUYx2YVw+mdGnXi3euiOjN2/hveKfOUq3+HJo8C9xzMvog5h5lrZGtd29GOkyUoJiqfmD5Q+Cs3L4ch6oxWSUWWKufZtR33bJ94kjpGyXP0B17NHDR9vKli8IOLlqa4WMYhGpJF7P65/od0M5XRmW3lGf8i7XqMsDVDl8bYfFelLbfJJRPWzskylyeA2suDsmtJa4fRzPzwzoRJOr0XFlddtDBMU9A7kXscWvtN+J3vjpBktNFGO+AAPykmp1CABkKF1CoJ1SlNZO2ybV2QqFpaPVUJMaF4PUWB9QatJK+K3/cKuoFeWHxRvUoMj9EkxooqydRkyfVUl8SuuMPCZCsG7ruY3oj8Y1H5i3D2f2/ernz9TrT+Qph+GUKoTnypFxvVre41T5KxGzeqBMtUGW2BjfWGTjSz18Pa2bC64wDgJbWZTVcFrWcT4YYTuDuZGi2PshIIKm8CBPXCy+HsV6C9j0B6EdKFgdVzrrq5VgCRxblIeIGTWZKQa2lirmuSrhP4rFz7VHNxKBFQswHrwpSymoRT60hdWcXMRi380FNesGtTIdtm7Wyc41nCBSjvWycTrqJ7MWS3DAY/tZd8B511juR+DPE6sMvBuVBRPG31asGohqWTqv3GEhbI0jmgFmPRrVmgblhba83szDNL51yyb6ZdaEXsP/9fyP0y0k/D6r+OUGS00Ua7YQB+3HQDcFI/QLtm9CoQ3VRX1Fx0Frdz+fCyikEnkQOPfyMUM7JUHc86DIt8gtEcDXj3rQ0zLOBl83vMYSsOcx7sk+DPQTwE+SfB2ntueD33G8PO+xziOaDr6uS0LoqYFxyakXFrjwGawrKtEx0B3iQvpIMxrWHujfDi0/YX4VcIMsmcoC2fo9b/xnHVcGdVEmtKI9cUsRMyZDc1PXr+l2H1bHAPQfZ/gP/AbBqAjHctZrENzJPzN5/MJR/yxAOSHp6vVrdRbLvUQaAXYhO0waNDrWo9aMfN9xGvqXJMarlDsxatAE6yXPlPprBzc1HX1iJLmlO0WvZr9udUQ81IcMjXJx5qETO3NDv7U3D2r6LiRxCvBrskhNWZNSFtPjqFkeGvx5QpHeatRreU/CCTE5E6z9UY7joN1me44sRQrEYbvbOPIZ4Ce34MzvkLWLtuhCGjjTbaDWFVJdvIN/mIC5MaPNnERkXIGCU/AjA3WbfCraxW/sAlC7Brin7IesfVqhIN23RJdAxiEmlNrVRkvUVg42iyJovEQ7DLDP8hGf8Pv/LP6Pk3U4WC1X9Aa68EzmqwrJPJGz3SyugJ+CW1BCoWa1ANJWyq1s53MFRHy7vHwPq9ceeehS//MkjhbZQOnYIV8R++dwre0G6eBjiqYe3T06ZUs66qmWUpp18sAZA86xvA76PVP8W5H0I8AewhMjsu0fqMT+g7oWgV7iur8gFNLL9oVHg6s9pawfA+8Q27UlMeVd8hwGqnWjE8uuMmWVhFnBrGtNY0V/xWS9eLFtANM1e+LiC1WUw0JEY1ZtXEfVhdvkpxXFqihlm1ga8HuSUVsyTdjCrZ7qud+XHg42j1NZj7ReCJUNy2cuItKOoG/1EpQ195/WUzzGLbqxPEWXFdbQmcsGYGokEobj1hG7TOd8kMM5lhaKdk/4n4E+HfDmvXjNBjtNFGu3EAvlW7dCNW0Ih4qFrWUO/cvg6wtRbwSuB5K0lQKfFUJnt6s6PXW19cOKuN0eIJp4zWIhqQWqRRzDAckfK6CvQV4FPI/xv4D4nzLqy+9uZt170cdtwTikdVCFxmcUdKgbpaoJ5kLxuqCtnC+5mWrMG1mdBUmLsN4o/NJg+Vzj0Pzv7SBulQ1btvJJjrELKWyH16dOMaL0XVmPCxiGw7wFy5h1wKW9tJyTuAd8CLT8P8I8x4KLJ7YhwLtlI7zapKN0S33IqGUTcDiqaYjvqpsOo6CWaYBcdNqr15GuK4rE4/wikZZpgP2DP1IKskEAXHwlzsHh8dyybcvrrxus5vzYiaU3MKt0nQfeV5hrZJ1jclLnW9Llo6fButRlfr3uuWGHu/2By5EHEmnPE63LbHmPQT4O4r4/BmFVN9IJJUZY74vF0UrwfsqzHe7V/fkE9qVwCvKq1U7MJFwD9L7m/g0P+AZ+wZ+2y00Ua7cQF+w9w766k8VHC9qollmWJYRQLwKmCRIicl5byrY+oieWNzJmpRHryhmVVRhC0YCa40c1N8uS6zq8G+i/imTF/A/AU4/ymm6xfCS6+85XXry6+F5z/LbPttRXE3ee/bwfPN0XUdkmUz2MjkZGao1Hu2QGxAL174QrgnYXoAds75eP0NrC1YNMy5kClXOuQiOLIEyHfcAUCdGwu3XNCuetXcYATKK7C+pKEPL/wini+K1d+DradAeT8o7olNv8ew28jsOLAdoC0iKVPQAjC5CkCtfrY0iUFypKHy7Xidogk1sfqUzGqk2vKh6vDxzlcEQsHooNvoqyuexhnlFrFz8zDWxiQQGXLNmKuWMqs5e5k6Fcc6vElNcFSNVpEdrDTr9mjBzv8GnteIP/1DJt+8p5v6h8p4sODuYMdUOUmKqr716V6lRZtK7IpezlHVBYpqEWHt9CHBKXHKDK4X9k2JT4H7V+A9cNaXx/4ZbbTRbjqAr+m7sZ3rBl5yHrOySuPyCsGHAVl5QnEjN4n0nWKgr4+ooohVcb2EN6xUJTSeSLtg+EDf1bmxwlsJzstk4B1Wvd+Ej4VdnClQhrYLyp3OuNzgu1Otfwe2fBu4Bp2z9+Za63Fj9tILVaw9DV/+HbhjQL4OiagAg7XBfSu8qqMCou7fO9rqoisjF8KzVLH5lAK7I+KPwD0Szn0ZnP2JBRDTxVB+PGyoKpJdtgGmtfkEiLqOr5GW+PUCV1Z8MvKG8V1wu5acsZwCF8bXX6OnbxVHHgXuJIriJLR+e7yOB3cs6GjQwcIXSfBBlINlEmj+hlOPVHxky81Iym2pXRnOh6R385GKr+QLwcwUgE/HRSw9cj4sC6YgURtjx+rvqnnvJIlgOgH3ubpS89LbNSU6/L9hOm3WuZg0FLyZCSFLXMh8Fd9hrUzoGB5pIZs8yQiNs8ztAn/tuHXl7Jd2M+XDHj6M/vQVcMmpUN7XjAcI7gvuFMwfAa6oC83VJ1uuTUxZTX/QnA5Xy43AnEdcDVxm2AVmfMybPkZpn4dPXxIK24022mij3ZR802g3czv757HJ6xDbgwpRAiRiXLqGKhylhyuJHKORr52U1aanAvhGjH01U2GYLjHjtd5P/gjOvHT4/ld3ANtgq8Eegy3JN+zNjN8qZ21L8lBbOu/bq/D3PfE9W+Muf8G1N4+N+SkrcNxW2LMS2u0a4FBC+22tnG0XJFXXY1vgQ9ttUdOuW6zdxlvUtN21wCGxD/YaHDzQH5VTsEWhnfdY+OzR8efuZ3t9U8I5V2dKTy+prR4MbAntuyc+M7HtzYXn3KL2+Oy2cdr21fv3Vk6ah61X3zxEAW5Me/FR4G8L/hSw20B5cmT3DzWzw4U/FNiOtLUurheIqL1g14JdJXQ56NvIXQK6CPxFUFwE699Z/ERytNFGG20E+KMdONDxLKx4maGJKu0SkfCG1i8f6pmvlZ4eAlimbDxpSE+lJuLj3uli3gT/I7lXwY6/gmfvGvtqtNFGu3Gd4dvtgKu3wZYVmBTUp81TAVNgNxy8C67btU9CAaONNtpoI8Af7YaxNxXw2ZdgxXMiwJalAvOxom+TINsZJGkF4u6lU1Y/o7jTel8y4qrYa8kZRgnuX5H9NnzyvePx9mijjTbaaKONNtq+WzE2wS3B3ix48H9gdmso7h2zxrpimG2XT428c0ta09qAf6Z1CyZ1w3fMMGcy886wUzGdjt3qPvCQq+B/fR0+MAL90UYbbbTRRhtttBHgj5a3D6zD3d+P7bg15u5F1Jcxk81E60Y/EbdbeTMtCT90jS7gJ/UtDAvp3BNwdzLzp5sV94MHXw93+SZ8dIxvHW200UYbbbTRRhsB/mh9+8geuO97sZVjzSb3olI8SYox1i/aYDynrpgC+7b8eR/gWyoJGBUpLFWNplKo8AImyE7DeAx20PfCg0p48DfgA7vHPhxttNFGG2200UYbAf5oLfvQXrjf+2DlKHD3ifKmsz8yxMBbB/zn/luB+0Rqp+0DJKWMLa38XgppgrnbYzwK7MHwQ9vg+78NH7x67MfRRhtttNFGG2202dBttFucPXM7HLqKFc8GTUIVGCBWMM2C+6GR01XM6SbUmg1LcTYXqWV8jCjJXovrm8UKa0Luyxh/j/hLOPt/xn4cbbTRRsNg9SCYHAPT20J5d4w7YOW/41/y/8bmGW20EeCPdouyVYezZyDOB9tOKGJkGx1AGvqLlB1hvc/UxbAq6yjRJVqbmLOgYqfLkL0Tyr8FPgRrV439Odpoo90y7BmHwyEnQXF3KO4K3AHTyUAocicdFOpY7vldOO+ZY3uNNtoI8Ee7Rdo5T8bcbwOHh4qzcwZNRzZfnb/VmFwLjrI6SMci0a8mBaD+VbxgHdHjXMwT3gv6HLi3gX8b8Omx6Mxoo41287DVHTA5CvwpOJ2Kd3cCf2esuA2Ut0EcCq4Ia6hQCIH0Mg8qHPhXorVnje042mgjwB/tlmrF6k/g3avBnRwLvDQ2EJ7Tit4R/XAeS5wBzR9tvWJZ9Weq36p/KiAMCoMCrLwa8WGwt4L+FbhwLEwz2mijLb89czscegRwCnAacCq424DdDnQKZkdibnuoLF6tgZ4QWqlIkkjt5booYPq76NyRwR9ttBHgj3bLtrPuZzb5A3D3hSmyWLtqUTY+fc+sole597PId1T7V3NDUdA/lOU1GXIOk5BdDPYfwLvAfQjKr8LadOzj0UYb7aaz1S2w5TAoT4TyDqDbQ3Ey5m+H7DY4HY84pCqmG5ZRYZjARaajU5BQvrNm+vjPiRP+d9E5I8AfbbQR4I822vNOMdv+crCfCnuIDzEyFsG1zHojJ5OUm42zHwLzQwm8Q6y/JCyw+WYVwFe8x7qilkMFmPfgvonsv4B/BftPuPpCeMX1Y1+PNtpoN4y9bgUuOQKK44Dbgj8Jc7dDuh1wEqaTgMPAVprFrlYpCLE2JMVFzKLUcPyzOgtpXZWwIj6EKByUI8AfbbQR4I82WmWrO3DuacheAHYklL49ZKLSTheIawD4JyDehspq5dj7bGGs9AclmptDSQEuBgq5ygv4NrILQB8GPgC7Pw4v+c7Y56ONNto+rJUHw+QImJ4EdnIIq+EUjBPAnwJ2a+BQpK1Yc7QZAHiD6q1aHNVdG+cda3YXXkucBTcC/NFGGwH+aKNlN68fxuy3oLiXMRVYlfBqc0fOjKJXg+9dxHqORYv1an+hGUFeM2L7QIOlF9gJfAX4ENL7YPpfMPnaGLc/2mijddbCbcDRwPFQnAL+DuDuiHE70HHAkaBDwBV9cC7RDXfsVBK0FvuhGvU38sOJullrvTTAqb24+hTgvwqt/d+x/0YbbQT4o43W3dhOxDgX7ElgBVIM8JT19jHLAPuhUJ4DMQJ7joFC+E66g7YBPQ3Yl+pzbxmYpshdBnwK3Aeh/A/gAli7YhwDo412i1nvDg1AfnIK+NPA3Q7Kk4HbAMdiOhrcdoIGZVzCYkRNtxKIXATtOQUCevJjMd7e0sUzRN20F9h+XZGavW+AvnmQczB9NTr3GWO/jjbaCPBHG21gnJzzFMytgR1rQUpTSkE+zE+a1X6MvIHPdqriJv+wJCa/+zcAH34RHACBHHIGRRXXvxN0EfBJjP/EFx+D674EL718HA6jjbZpAfwE2AEcDu7YCNxvG5JcXRUbfxzGYSbbSpXEX9XmCAo26uh9RSDuMttrxbw7VZxCf2HLrKMpQ1+r5nSJCktOMBuIL3OCEjRxsP4aeNFvjP0+2mgjwB9ttBl21n2geJEZDwdzwqsHvodGWC6pdp4ij80H+M3f1YD6Wiu/+/fuThhj+Ov3RLa/Pi93ltzwNaBvgH0K8VGYfAL2fgku/Q68fn0cG6ONtjQgfksA8dMjYPtRUJ5ASHS9HbiTwJ8EdjTG4cgfAs5Raweoot9D1n5XoUaAOeowmHqdKujHw3eKBsqUAnUzBcZei2zN/397Z/Zrx5XV4W/tOucOvrYTO3Yc46Qb9SCapEWjhoZ+IQHEJAYxSEGCZwgiEmpBI3U3LWHcSAiEGsTwxJ8A4gXxBH8BLzygJj0k7SaTndiOfe3r4Q6n9uJh17Br165Tda7tpDtZn3Rl33PPqVOnTlXt39p7rd/SKojIFT6l91jRYKE5Cyk6WIqOYZjAN4xRXjyKe+y30eIPwX2iGui0P8j09XRHp0t+Yj+Xj98ZxgaFfybZXzJRQ2dcVM1VCIdat+oxbf5W2Vi4ejs7wBsI3wT+B5Wvgf8m8AawbXn8hvEwx62/OgoHx2H3JKyfhvIs8CTiPgz6FMgZVB8DPQm6FdxqXHUjqpvmeW0m6EMue5P+ItFdR3udtTMTF22OTdP/Izev0Bl4JUnqyRoMSEfgh21Kd+IityggqjBz5oNvGCbwDWNFzn+/iHxOkd8V3BYs0i4ry4V6jgHXnDqNXlf11R98XlyT1o6QbQ5sN+un3V6w2g/zfE6CCmicMfZBbwKvoXoR5H+Bb4Quu3uvw1/eBEz0G8akcemrG3DrGBRnQL8P9HGQs4ieBTkHnAU9AzyKyiaiG+BmEs3BR43xtDvF4Hs3C6mcdavu2NUzNblkq9BfJM22fzCj6tT7V35qJA0MtMrB/xv4yuftlDIME/iGseL585VfQvwFKD4dGqz4tCZselrOhIFurGY3u52hF0bCvWNbMbShTkqPVILARw84qa3uwuf2ByDXgddRuYjoRcR9HV++AvNX4dg1+NyenUbGBwsV+LN14ChwHDhJyIX/COg5KM4ichr8GZDHwB9DdRNxrvWG11jKKyIqpMWnKX7ieCfasQMeGi2X/W1sNmMV57DBe6MM7ES9QjETKL8KF/7YzjnDMIFvGIfgS6dx8xfR4vfDjNuCwZFWBh6bMjt/GCvNyUGATowomk4yndfFC/rtK6vuM/WSenDU8KB3QK6ivBZm+PXbIC9D+SqUb8E71y2v3/je5fwsCPfFCVg7RSFnUH8ar2cROSdwVtHHBTmh6BYqx0C3EJl36mfUt3q7XXbT9lJytL8wYXlQ5YEJ7Mz9QeJVRonXBVNzr5FBecpqJUsEfrNYUQSBrybwDcMEvmHc38D+CXB/BPyWIEeVsjGhGFxUXtrMamAsWzIb3/hGy9jgrF1hMLQPS68Q7XabbHazzIy7dU5/LUjEtf8H0APU30PkCshl4BuI+xbeXwReA67A7k3YuAMXFnauGe8+zxfwoQ2Yb0BxFNZPgp6C8gmcnEHlbFXQehbkNHAK3BYwDz+pII2ja9VOkKxxMnv8moJ21cwlF2kbbNOzkpw2ey+RreXS1YBlYn2F4EGW5eEfNviQeruFgP8b9M8sRccwTOAbxoPgwnNI+WWUn4GZICVRYmzvzJPE3XJwoFRGUm3yc1rpJH32sXTWbCxIqPNcex+qbF/c2OJpUjRXd+iSVHhUnbjqvH5B8SVwO/zI2+Aug76N8jroq1C+Fh7jGtzagb/dZUL2kmG0ov3JNXh8A3Y3adJm/GkoTuHkDL48jcgZkFNB0PsTqBxHZBPYANYra9lqkryeeUerHk/V2lfXJrJ1hYlz5ukWl2r/piFVbLxsGGvz5LUn8NNrXZq0e6ni/lrg092vqaPn5NTAJfegXKfwwXucVLebdJVgJuj+38Kf/5Gd54ZhAt8wHhCf34LN5xH3eyCfgaIIS+5pigutSQ20DdwlPwA2Y+3U2S5d7UxvxtSVcmilzc6ReHbRR5kF0QY10fm5YVudJiO4a4R/43/tAb8Lso2yDVwFeQP8JdC3wb0FxTXgFhzcgvVbsLgDt+7A8T1z+nm/cd7ByTlc2oCNTeAIFFugx4Nw10dBT4QfdwrkMUQfDU4zcgJkHeRIcJ3xW+AKBNd3qtK2WVyngLV5nkR2kwO2tdU5L5q5piS5YF1n25mrVgcv8iaoqAT+svqeTsAhqwv7w7JEzOd2UpqYYySO15mACXzDMIFvGA+FLzwC67+MFL+D8KyouDDLjWotYpvxX6W3LF6Ns83En3RNcCYV2o7NxA/NkMUz/mNNvAaDijptIE7gJ5dkK62PP21/nWbnoh3VeNsIigRvbmk1kHjAeyj20dIjchfcLZCbqN9G5CpwHdUbwPXq51b4cTfhYBv8bZjfA/bDz4V9O58fNi/M4ew6sAb7c/DrMN8MQp2jFO4RykqwO44CR1GOovoIyAnEn6ieeww05LbDBkgBMgOdxStE7dx1JBfVa+UXo206jEYzzPXrqtdKlGc+dv0NNqyLUmLE9SKCbK+8ybnqA9fm4A1kiUvN0LUffc5mv3oOvLk3bLvT1mlJesgFuHTVQihEKf8ONR98wzCBbxgPjb86hrv3G6K8CPJjwY7Ga52qEvSsF10ykE4S9PfLsjz/sStIV7nKpk7btQWGgqBL3yyd9fTS/YNbsrNegQPUHSDsg98DuQ3uFspthLtVcfBNkHcQv413t0B2odwFtwN+B7gNxZ3w+GIvBAa7e7BxACzgpRL+pXx/3Dufd/B0badU/Wysw+4WrG/A3hoUc9A18JvAcSiPQrEGug6yiSuOoeUjKEdBTyDyCLgtVLeAKgXGbwJriJsFwS9zoGh6MgycCdorOlHAVaK9caPpTqPXM9idVDY/cHEM/T40Kz8WAMTpaVME7VCs/KCGxsMJ/PEJh2xR/oO7XzUUAou/R7/yORt/DMMEvmE8ZL54AtZ/HeEFcD8ehEYZVERIWZGlja3ix8YG1VW74U4b3g852D6I99C4W2U15xkX7tbHMrbw1Eg41Um6mhNhdRAgbRpQPbNKJ3e6teEL/cBEnGqoLN5HZQ/RfZB90BJ0gco9RO+A3AHugtwNwp8y/Og+sAt+H2UXJ3sge6B7eDkAV1bnCHULYpCy+oAeZBF+L337WBPoRAJcCihnIHMcrhW6ToJw1q1KeNcHQFCdg66Hf90asIEU60F46zqhgHSt+ncdZKttrEQhoSp0BjpXdC3839cdl0LztKq3ggwEX+Ew13Oz8TmQVdL1alibVx6OAY0c7uS4VedJM6Ptlpx/PpHWTWCYEfaF9M7uThwQL8NJf7M6PlIdagb/MNf0iqNkLvCQprNVXYuQuwNMvCNo7qnS7cqNF3QuyME/ol/5Axt3DMMEvmG8S3x+C478AjJ7EeFZtJihCxAfLW4vcbmRvk7NOu1odwzsFe+NBAq9YAPyhbqrXlVjKwVTCviCctCu+IqFfb0ZVyX0xAJ/6vvWD7jMhKMK4hBRaVISkipKSYRMT9t1OnW2MlbqosHgVKQ9F5Rag3ciFhXwGk1dV6lLnSRw6Z0Enf1Oij0kzjt31Z+157jYC366aixykhKFUiKnp+5xlSJ6bdlKdtE2haUT5MW/BsGu6qMZfCf58y5OvyHKVXcjgWiSA99RsfF1WbWckiU5bL0eGbJcWE9dIZvy2qF7hawQPKy0Ytd9o7pj7rJ9mD5xIdo9fnV9ztyBCXzD+KAys0NgvDd89Q7wr+gL/4ae/SnE/w6iP4sWj4bJzFLDwKUyOBgPCP+OIOkOhOMDcfI+OvB8HQo6xoIElgQgqwQIkn4op60wqQt8q/KF1iVEuvn98Q7qwP7Us40+zOeLCzag9fvhUVXt5A41QtEHIYvrfh2i1URmIo1VpJ5VV40CFtUokImPm9bSVzrfSsebXKXdFr0YoSM3NZkRT8Wp+HDwtFl1SoKq8G8bnNDGHbR1pe0Ubpx0Jcnp6dsvSjWjNqO89yYHnuAdL6A4sgWpnfOmkwTftXxVH+q6ofou4sBN8+eqSue70Z4ajf9f15tIvv5FVgh6hy44mbC93EQBK6wMrDBFFtf2aMZZS+qzeUToS7o2InUYGEeMzry0DMMEvmG8l/zTAfAfKP8JX/w0sv7rwC+DPI24eTUbrWPja95dOyPSUpGYpvxOTfmZ+hpGxlmZ8MJJIkJFRKsEDEc8ddpX8620bWbNW6/SdlpR+pFVVIjZlvmmbih9CdJOirfVBG1OQScCcvWqQ1cU07YTa3VRFZk0wYn2jpV2pklTcRwdoU4aUu1WROJyFIv5ODQoO2JSNSProiLK+oAIlTGSelRcIu7TTXROvDblSrKFm5I765osEZ0gG+v9FTLX0JIOqpq8USPgtX9y6/i10tmrCbPbaRZQL9CfsiqwYqrd5HKaJvBIHLVoI8wpef069Kl7ByuNxgzD+CBhKTrGdyHnTwLPIu43wf8cFI9VojKonyCKRHpT0fQ6zOYF3PggOmlQX3I1HTK7drV9muTVv2z/NaOKZOROEQlmjR1U6B9gyQmvuKhamtWBbg+Cbj5x3qUllvt+JPBLU3FSa6aMqOymImW+OodUs/iaquBYZDWiK5L+Uq2E1MbxvQZFtOJMVe7rjj2p6VuiE9M/dwJljSwuewq1f/B7aUTJS3QkZni3WeI53z8ujKfSyVKl36m70NxdopdiSOuypS4JSlW661IFsG9FtoZhAt8wvttQgT/9NMivIfwq6p5GXBEkVakaqbQ2y0ShkYtO7ktIH8ZrX/Pzp1MFvsTCZ0z06AqPdWblY+EWbTz27I/1d9aaVDSkcWgl10WGhUlm9yVXhNjd1bxQklxU0j9vMoXY3dndvMDPWqJmAo8gxP3yOpH0+ItkD/cUodiJE1a5e6c2r4c55zsrD9o13tE4x0WGIqzuznzX3F4yx3GogH8saJ5UPyNRHVDu32X75aKVtSplSqtgtTmPFx50O3TCLl8G/RqF/3cWf/lfNpYYhgl8w/gu5UunQZ6F+a8APwl8OMxQlXSaZ8V285LNL1jeaTKXpp8RVysL9imiYkiUR1vLJ1+MCZFahCc5z6p54RUJ0pUDotyHFjLFoNGM/IBGfFh3vDTrRgeWWqZ8x8PO5pEwHtK9D2nGetjFJXp8amF45hTRODBU6baizjSwGxXEI9Fgr8Nr5pzKb/cQo99YF9pkFweDwXqFqXOskpNw5Puv6mnCRauuqlivOwX7BehNlDdAvwX6NdCXYPcV2HsT9rar9EfDMEzgG8b3Cn9yDjd/DtXfAH0OnZ8iTtUQ1aaKrU7WVhk+6+9XWOqYqmc493epYE5m3ZPtZXOqs0IrFZqdlJVIqcg0YT81BWjia7KOJVPzoHXk+KaCbEy4pzPeU77zJeK4dyiTiVsZ0rtTWyQwPVid/Df6+xin0nfTnlrf/EGBLWlztqRLbPwGPWedQ16HU5yuck5aOhLFyf3dI7IBehsISZt65qI3PCiBS8ArIF9HeQn0ZShfheJtuLBtY4JhGCbwjfcZX/4oFL8I7tcQfgR1j4THPVBqK7Nk9OzPjvOriNsHIQKWpQdkBCtjdn6RsOr10+1N66Zqe/iw6dT0oLG/jwl0OeQxfpj53If4nnuzvSuK7EMHWmOBw5jN7FAglgaEQzPVIeDuueW0z2tz2oayekbPr8PY1A5dc1MFvk7bt65BbNzRS+raDQQXNxTbAb0K+jroK4h+Ha//DeU34C/eJldwYhiGYQLfeP/yzwX87zPATyPy88APA090fOG1TR+Pxl0Bp4jKJNG0THim/XqUUcu7UTE6lEOe2YaQT3nuCQzt2mYrMp4yEW8zlwIyJvqH0nBWON7ZdCcZOS6pOL3fIGzsXMg4SOoqHucD+94c7yUpXoPH4aGkPsVpX8tGlMT7KHYm0p770vRgaerfV+1CLat91yKqyaUVVxW3OXUiiFZCXryC3gW5BlwGeRX0ZYSvUeg3OeANYBsuLOy+bhiGCXzDaDjvgI8AzyLys6CfAf0wFLPgD+9BvK87traWiyJjmiArJHSC8pAJom2ZyM9r7e5k6eiORyIrEutdJ0l5YHeHsYBkcurJKnetqTUOUwO3gf1euQtx+mTXT8Mebbi2bAZ7SsO00UBzoFA292FVBwV9dz90OB0GWepaM+2708q/X2XZOSfZQvHc9183AGsi0547UtDslStSu9GQJ99cVI2Q3wG5AvI26i6CfwXHK3guAm/A8Xfg8/fsnm0Yhgl8w1iZLz8F8hngOZDPisjHgRNVM6CmZWowha+vCrdEOyvZJkKMFd5NFLoDeciDM5NjM5GaKNO0k+hQbsZQleaK8cSkG1HHTXLks4197sM2D1u2zyPiO3eoJwlxmRjgyYD4vZ/zLDmmg/nhTTCosYUn/ZyfbuuupsmXaC8VplvoLKsVMytN47Mg3KV/XU7tTt37nF0b2Oj4SlfwS1R/4BH8riLbqFwGea0S8t/C+5dg9n/Adbhw1+7FhmGYwDeMh8L54+A/DvKjiD4H8llwT4Kb07HZjHIH0g5Goqs58wyJtvT5qzTyWfX9mu2PJBnHzic67G5yWIEvskKX0GXHGLJuPYNCGqbbIcqwc8ukfKABN6bDiPPJx3nsrn6IwKUTsYgcJqqk74fa/3RNqk/aPnr5akNlWeqSL2VA4E9132mvD5Fm23UzuFJB7wWxLm+Cfx34DlpeBLkI8jos3oL1HUuvMQzDBL5hvKfn/xeego0fQvgsyI8CT4twTpm50FjLV93j62RblZ5NXup6t2p6CGm6T99eceXUk2VFgKqZnAVp03XiZIf6d5Xx7l3C4d1bVnn+SB66Tnn+/bz/WFC3dD+rBkeyJEgbUvhjRbojNQHT7/655ZAonSbTMTi3wZ5daNKpNbMypr3lL5GkoV2dQlOL+9wHTuppwsvqYvvQvwEFcVW7ZGkDBamXkPwOyFXQN1H+D/TbwLdCvvzuJbh7Hf5hhxUztQzDMEzgG8a7j4PzTwCfBP0JxH0W5RmB04ibtQ2StFXi/SyFafnQZHL4B8TgYHrG4PRuKyLpFJdmxHu8841wlIzWixvxDPhOjgn8KUJ0FfP5zOvriWbNFCZnFyZWDcAYqL0YCqpSr8wpQcuq/Qbit1GyjrDjdRATGi/lutdm/z50rLsz+d0C8DqnPbFz1XrZREBc9xvp2chqcoK2yTW1gG+tPj2I30fZAa6CexP8d4BvgF4EfSVYUF6+YX7yhmGYwDeM9xXPF/CpJ+DgB3HyI6L8uCLPAOdQthBXiQpfT+FrMzN7WBvNaAYz6zp4yILcpUJxsCNV7FUu+RhkQiCTCkqJtqeHzdeZGiQsOQ6DzjpTgoyxlYwpAUQuAyp1YIrf4jAWoUMfpJvUrk3zMUR6qzbpuSkTcn3yO195zKQ57l7bc6N71iuRwMdVp6RE0r19oP0M1StVS0V3QW8ixVtQvhUaQ8mrsHgVeB3KN2D9Grx0B/6ltHueYRgm8A3jA8mXTsPaR0GeQeQZ0KeBjwv+SaVYC5rDp7nu8dRlnWLQqmlVlSY6qBRUruJxmcDXXBesRKx3HPwY6DyrXYN9yYlLyeq7Xj+ujAge7MS7pAvuJFG/6p1Nh4V5zm50kuCXga8kO8M/MNWefduBGfXO4ktTO1K9jdb2p9LPe6/f32sn2mm2L+Ptd6Oq6K6wTmbmq91oZunTOhBNXXY6J4i07lYuWpZx9fZKYAHcAbkF/gaqlxH3Brp4DYpvw+I1WFwC/w789Y7dvwzDMIFvGMYEXpjD2bOEtJ5PhX/l46DnQB5DivVuF8+ySuypLz2vtfiOpv4jgS8D+SHxNlX7Aklk8HUM5NkHcUfeVH+JwM9ZHS5ZaZBqkrg3k3u/d6T7cc/hPt93lRUbfUD7mMaQiSd7wPUqVxtf+qZQNTlBOu5KzZeVT91iYLa/ObFSWyTtf0Ma5aE17jTNtktB7ipuB/W3EHcFuIQuLkLxcihqlbehuAHb23Bpx2biDcMwTOAbxkO4pr5wHNxZKD4C80/A4pMw/xDinwJOgT4CRREuv9pzu5lR1W7FqFSlf5Uwk4yaVB8JtaLaposVe0+MBhGuw3eBbE69JMJODtfVVTK1Bzn9N1azIOPlDkO2lTLgBhq/0cpe98uE/4D1aTrz37GI15EAorvT2m7D0+m0pJKcBRlxP3QuaGQrFK/RNMFnFJ/Wp5toVbJNdcbGH0CSA+0PUHZB7yK8A3IF1TdB34TiMuhlKC+HgtfyOuzfhq/ewzq6GoZhmMA3jPeeP9yEzUdhfgb4CMgPIHwC+Ci4c+BPg9sM+Qe1zg/5/XVSdK9QtK0Y1K6xuIC6SACr5OsCQhJH5ZwT2QvmqjUle+PQMcGaZnTkBPhAAAKrZeBoppvw0uZGSxS8MLK/qa97vBgi0deX7qxmcnnSoKVXM5ocv3TnIl/48HuZ7HQazmj2g8VtIPppXBJPs9etI9pK7p6AVwS84hegu1DcBLkhlFcULkHxJlq+DrM3obwC5Ttw7yYcvW1+8YZhGCbwDeN9IPyPnILZGSg+hDv4GFp8FNUPIzwJ8hiUJ2ny+zvCnlDgG5y6tXb1aURgazWYF+XaF3wdYS+VTWFd7JgR4yvfTSTxWPQyKK5VNKo6zWfv60C0MKXJ0VjjrJ6Al3BwDlMAG3LPQ/2Ftse251qU3eG6bsNVD9WxmE8832uR72kmuBVCmk69LKGaC6a07TYlcaQh0qbMdIPEEuAesAPsoBpm3+Ea+CvBJ969Bf4qLN6BtVuwdxOu3zZXGsMwDBP4hvEB5bwDjgCnYP8cFE/hio+h5cdgfhr1Z0AfR9wWyAb4zciaJlJwGuXqx7O82q3flLbtUCv0SXL64xChVnsuL7abKEH7VqLiMnLbJ4LZ1cFJopK1WZQIf/fxMoGM3daqlkraBkgun6zTsWT0nc8Zagi0+/eswHcQpct0C2XDtHzlCpMc2+gQ9mwsc8dOox3SZBJdI4FP5pgKolIdzlLB7YHsQrkPcgOVa+C2kYNbwBWUK+AuhVl3fy38cAPu3IV/2Md84Q3DMEzgG4Zx2Ov2xS149BjMTwCPh6JenkDc46j/EOj3QfEolCfAH0fcGugcmPVEYpr7IVXaNKpZ3/Oktau0VikMTotrHGxU768uCPSmuDPX4CsStZ27lY/2vZ+HI1WBaMeyX7023piS5oInnVTFBTckaCogNHsM0s1IEuxERaZ1Kk0zax/NwDc2RBLFYlLHJO2sfT3rX8dcdY1qHERoVdsRcuEXIAtF90F3QHZAbqN+G5G30OIq+G0ob4JcpSiuUfqrsLgJ5S1YvwMX9ky4G4ZhmMA3DOM95w/WYX4M1o5XQcBJ0MdAT4M+AbPHEU6Hgl+OobIJbgPYQtgEvwGFQ9UhuDbPWjuiVar/99Vf7ezjqid1RG6bV67J7HbqtSmu7YXUyHXXFfiS2i9KZSgTFYbWVcUSifmmLLXeoY6dULVKkLrDeGk3FwdF8S3VSZuEHxevNu8rjcmRRoGBpAFXs92yeuYBuF1w91DdQ+QAdE/gtuJuA+FH/Q3wV3ByA891cO/AYhsObsDiNmzdhmfuwW+a24xhGIYJfMMw3l+8MIetGfh1OLIOa1vAcfDHgvBfOw76aFgFkEcq958tYKt6zlFBj4AeUWQNKebAGurnCDOofjpFnn07RdHWnRNyXvmJB3zaoaqpSdCm/5F2HGKk89ZtR+I6WKAJVlrNPtCtqpOK4wbqdLWKBLTUsAywgCDGUV+C7oPcQ7gL7i6q98DdRdhFy71GqMNtKHaAW2H2/WAH5jehvAN6F2b3YH0P9vbCe1wu4Z8W2Gy7YRiGCXw7BIZhrMb5GbAGm3O4F0Q9i3VYPwIHR0M9gK4HVyCOQnEEdBNkPTxX5+HvuoXIBrAOcgRlE9wccTPwM5A5qkVQ0r4geIDOgAKVIqwsVFq6UvJNKo1Wyr5afRDEa9RyOMop8kCJyAKkEspUDZSkjJopleAXqNxB9B5IEOXIHsIuXu4Ae8A+lLugeyGffbEf/uUuFHeAXdjbBb8bhH65gMUCjh/AhdLEuWEYhmEC3zCM9wnPF3DCwdk66X4Gu0VV7epgX+CYg5sFzAWOONh1sFG9fiHtT80GMKsf87BQmGn7nIXARmU5s1/CkQXslbCnsFaC93CkhJMlXFe4rHC2hAvmwW4YhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEYhmEY7x7/DzEvJQ16qnh7AAAAAElFTkSuQmCC" alt="Dauto Tintas">
</div>
""", unsafe_allow_html=True)
aba_resumo, aba_vendas, aba_marcas, aba_clientes, aba_compras = st.tabs([
    "Visão geral", "Detalhamento de vendas", "Marcas e produtos", "Clientes", "Compras e conferência"
])
st.markdown(f"**Meses selecionados:** {', '.join(MES_NUM_TO_NOME[m] for m in mes_nums_sel)}")

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
data_ini_ts = pd.Timestamp(data_ini)
data_fim_exclusivo_ts = pd.Timestamp(data_fim) + pd.Timedelta(days=1)
df_f = df_f[(df_f["DATA"] >= data_ini_ts) & (df_f["DATA"] < data_fim_exclusivo_ts)]
df_f = df_f[df_f["DATA"].dt.month.isin(mes_nums_sel)].copy()

fat_total = float(df_f["FAT_LINHA"].sum()) if len(df_f) else 0.0

with aba_resumo:
    st.markdown("#### Painel executivo")
    st.caption("Resumo de faturamento, meta, previsão e vendas diárias por loja.")

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

    df_2026 = df.loc[df["DATA"].notna() & (df["DATA"].dt.year == ANO_ATUAL)]

    if lojas_sel_aplicadas:
        df_2026 = df_2026[df_2026["LOJA_N"].isin(lojas_sel_aplicadas)].copy()
    else:
        df_2026 = df_2026.iloc[0:0].copy()

    if vendedores_sel_aplicados:
        df_2026 = df_2026[df_2026["VENDEDOR_N"].isin(vendedores_sel_aplicados)].copy()
    else:
        df_2026 = df_2026.iloc[0:0].copy()

    df_2026 = df_2026[(df_2026["DATA"] >= data_ini_ts) & (df_2026["DATA"] < data_fim_exclusivo_ts)].copy()
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

    # Indicadores de ano anterior são exibidos junto aos KPIs executivos, mais abaixo.

    # ============================================================
    # KPI: META x REALIZADO (velocímetro)
    # ============================================================
    st.markdown("""<div class="bi-heading"><div class="bi-icon">↗</div><div><h2>Meta × Realizado</h2><p>Resultados e ritmo de fechamento</p></div></div>""", unsafe_allow_html=True)

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

    st.markdown("""<div class="bi-heading"><div class="bi-icon">▥</div><div><h2>Indicadores de desempenho</h2><p>Resumo executivo do período e das lojas selecionadas</p></div></div>""", unsafe_allow_html=True)
    kp1, kp2, kp3, kp4 = st.columns(4, gap="medium")
    kp1.metric(f"Ano-1 ({ANO_BASE})", "R$ " + format_brl(total_2025))
    kp2.metric("Realizado", "R$ " + format_brl(real_total_sel),
               (f"{var_pct_total:+.2f}% vs. Ano-1".replace(".", ",")) if var_pct_total is not None else None)
    kp3.metric("Média por dia de venda", "R$ " + format_brl(media_por_dia_venda) if media_por_dia_venda is not None else "—")
    kp4.metric("Clientes ativos", f"{clientes_ativos:,}".replace(",", "."))
    kp5, kp6, kp7, kp8 = st.columns(4, gap="medium")
    kp5.metric("Meta do período", "R$ " + format_brl(meta_total_sel))
    kp6.metric("Previsão de fechamento", "R$ " + format_brl(previsao_fechamento) if previsao_fechamento is not None else "—")
    kp7.metric("Dias comerciais realizados", format_decimal_pt(dias_venda_equiv, 1))
    kp8.metric("Dias comerciais restantes", format_decimal_pt(dias_restantes_equiv, 1))

    st.caption(
        "Dias de venda: segunda a sexta = 1, sábado = 0,5. "
        "Dias comerciais com venda zero na empresa inteira são tratados como feriado/dia sem expediente e saem da previsão."
        + (f" Dias ignorados no recorte: {', '.join(dias_ignorados_periodo)}." if dias_ignorados_periodo else "")
    )

    goal_col, gauge_col, gap_col = st.columns([1.65, 1.0, 1.0], gap="medium")
    with goal_col:
        _progresso = max(0.0, min(float(pct_meta or 0.0), 100.0))
        _largura = f"{_progresso:.2f}%"
        _rotulo = (format_decimal_pt(pct_meta, 1) + "%") if pct_meta is not None else "Sem meta"
        st.markdown(
            f'<div class="bi-goal"><div class="bi-goal-title">Evolução da meta</div>'
            f'<div class="bi-goal-values"><div>R$ {format_brl(real_total_sel)}<small>Realizado</small></div>'
            f'<div style="text-align:right">R$ {format_brl(meta_total_sel)}<small>Meta</small></div></div>'
            f'<div class="bi-progress"><span style="width:{_largura}">{_rotulo if _progresso >= 18 else ""}</span></div>'
            f'<div class="bi-scale"><span>0%</span><span>25%</span><span>50%</span><span>75%</span><span>100%</span></div></div>',
            unsafe_allow_html=True,
        )
    with gauge_col:
        if pct_meta is None:
            st.info("Meta zerada para o recorte selecionado.")
        else:
            fig_gauge = go.Figure(go.Indicator(
                mode="gauge+number", value=max(0, pct_meta),
                number={"suffix": "%", "valueformat": ".1f", "font": {"size": 35, "color": "#10234a"}},
                title={"text": "% da meta atingida", "font": {"size": 14, "color": "#183454"}},
                gauge={"axis": {"range": [0, 120], "tickcolor": "#91a1b6", "tickfont": {"size": 9}},
                       "bar": {"color": "#1b69e2", "thickness": 0.45},
                       "bgcolor": "white", "borderwidth": 0,
                       "steps": [{"range": [0, 80], "color": "#edf2f9"},
                                 {"range": [80, 100], "color": "#e2eaf5"},
                                 {"range": [100, 120], "color": "#d6e2f1"}],
                       "threshold": {"line": {"color": "#d73535", "width": 3}, "thickness": 0.8, "value": 100}},
            ))
            fig_gauge.update_layout(margin=dict(l=10, r=10, t=44, b=0), height=214,
                                    paper_bgcolor="white", font=dict(family="Arial", color="#183454"))
            bi_plotly_chart(fig_gauge, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)
    with gap_col:
        _dif_cor = "positive" if dif_meta_r >= 0 else ""
        st.markdown(
            f'<div class="bi-alert"><div class="bi-alert-label">Diferença (Realizado − Meta)</div>'
            f'<div class="bi-alert-value {_dif_cor}">R$ {format_brl(dif_meta_r)}</div></div>'
            f'<div class="bi-alert"><div class="bi-alert-label">Diferença percentual</div>'
            f'<div class="bi-alert-value {_dif_cor}">{format_decimal_pt(dif_meta_p, 2) if dif_meta_p is not None else "—"}%</div></div>',
            unsafe_allow_html=True,
        )

    # RELATÓRIO PRINCIPAL — faturamento diário das lojas
    st.markdown("""<div class="bi-heading"><div class="bi-icon">▦</div><div><h2>Faturamento diário por loja</h2><p>Tabela gerencial, totalização e detalhamento por unidade</p></div></div>""", unsafe_allow_html=True)
    # ------------------------------------------------------------
    # Vendas Diárias por Loja — visão geral + drill por loja
    # Baseado no mesmo recorte/filtros ativos do dashboard.
    # ------------------------------------------------------------


    vendas_dia_base = df_f.copy()
    vendas_dia_base = vendas_dia_base[vendas_dia_base["DATA"].notna()].copy()

    if vendas_dia_base.empty:
        st.info("Não há vendas no período/filtros selecionados para montar o relatório diário.")
    else:
        vendas_dia_base["DATA_DIA"] = vendas_dia_base["DATA"].dt.normalize()
        vendas_dia_base["VALOR_DIA"] = pd.to_numeric(vendas_dia_base[VAL_COL], errors="coerce").fillna(0.0)

        diario_loja = (
            vendas_dia_base.groupby(["DATA_DIA", "LOJA_KEY"], dropna=False)["VALOR_DIA"]
            .sum()
            .reset_index()
        )

        # Mantém somente dias em que houve movimento positivo na empresa, como no relatório do ERP.
        totais_dia_empresa = diario_loja.groupby("DATA_DIA")["VALOR_DIA"].sum()
        dias_com_movimento = totais_dia_empresa[totais_dia_empresa > 0].index
        diario_loja = diario_loja[diario_loja["DATA_DIA"].isin(dias_com_movimento)].copy()

        if diario_loja.empty:
            st.info("Não há dias com faturamento positivo no recorte selecionado.")
        else:
            lojas_diarias = sorted(
                diario_loja["LOJA_KEY"].dropna().astype(str).unique().tolist(),
                key=lambda k: (LOJA_KEY_RANK.get(k, DEFAULT_RANK), str(map_key_to_loja.get(k, k))),
            )

            nome_loja_diaria = {k: str(map_key_to_loja.get(k, k)) for k in lojas_diarias}

            aba_geral, aba_drill = st.tabs(["Visão geral", "Drill por loja"])

            with aba_geral:
                matriz = diario_loja.pivot_table(
                    index="DATA_DIA",
                    columns="LOJA_KEY",
                    values="VALOR_DIA",
                    aggfunc="sum",
                    fill_value=0.0,
                )
                matriz = matriz.reindex(columns=lojas_diarias, fill_value=0.0).sort_index()
                matriz = matriz.rename(columns=nome_loja_diaria)
                matriz["TOTAL"] = matriz.sum(axis=1)

                matriz_view = matriz.reset_index().rename(columns={"DATA_DIA": "DATA"})
                matriz_view["DATA"] = pd.to_datetime(matriz_view["DATA"]).dt.strftime("%d/%m/%Y")

                linha_total = {"DATA": "TOTAIS"}
                for col in matriz.columns:
                    linha_total[col] = float(matriz[col].sum())
                matriz_view = pd.concat([matriz_view, pd.DataFrame([linha_total])], ignore_index=True)

                cols_valor_diario = [c for c in matriz_view.columns if c != "DATA"]
                sty_diario = (
                    matriz_view.style
                    .apply(
                        lambda row: ["background-color: #e9f1ff; color: #10234a; font-weight: 900;"] * len(row)
                        if str(row.get("DATA", "")).upper() == "TOTAIS"
                        else ["" if row.name % 2 == 0 else "background-color: #f4f8fe;"] * len(row),
                        axis=1,
                    )
                    .format({c: lambda v: "R$ " + format_brl(v) for c in cols_valor_diario})
                )
                st.dataframe(sty_diario, use_container_width=True, hide_index=True, height=420)

                total_periodo_diario = float(matriz["TOTAL"].sum())
                qtd_dias_diario = int(len(matriz))
                media_empresa_diaria = (total_periodo_diario / qtd_dias_diario) if qtd_dias_diario else 0.0
                d1, d2, d3 = st.columns(3)
                d1.metric("Total do período", "R$ " + format_brl(total_periodo_diario))
                d2.metric("Dias com movimento", str(qtd_dias_diario))
                d3.metric("Média diária", "R$ " + format_brl(media_empresa_diaria))

            with aba_drill:
                opcoes_loja = [nome_loja_diaria[k] for k in lojas_diarias]
                nome_para_key = {nome_loja_diaria[k]: k for k in lojas_diarias}
                loja_escolhida_nome = st.selectbox(
                    "Selecione a loja",
                    options=opcoes_loja,
                    key="drill_vendas_diarias_loja",
                )
                loja_escolhida_key = nome_para_key[loja_escolhida_nome]

                drill = diario_loja[diario_loja["LOJA_KEY"] == loja_escolhida_key].copy()
                drill = (
                    drill.groupby("DATA_DIA", dropna=False)["VALOR_DIA"]
                    .sum()
                    .reset_index()
                    .sort_values("DATA_DIA")
                )

                # Completa os mesmos dias com movimento da empresa para enxergar dias zerados da loja.
                calendario_drill = pd.DataFrame({"DATA_DIA": sorted(dias_com_movimento)})
                drill = calendario_drill.merge(drill, on="DATA_DIA", how="left")
                drill["VALOR_DIA"] = drill["VALOR_DIA"].fillna(0.0)

                total_loja_diario = float(drill["VALOR_DIA"].sum())
                dias_empresa = int(len(drill))
                media_loja_diaria = (total_loja_diario / dias_empresa) if dias_empresa else 0.0
                melhor_dia_val = float(drill["VALOR_DIA"].max()) if len(drill) else 0.0
                melhor_dia_data = drill.loc[drill["VALOR_DIA"].idxmax(), "DATA_DIA"] if len(drill) else None

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Total faturado", "R$ " + format_brl(total_loja_diario))
                c2.metric("Média diária", "R$ " + format_brl(media_loja_diaria))
                c3.metric("Dias do relatório", str(dias_empresa))
                c4.metric(
                    "Melhor dia",
                    "R$ " + format_brl(melhor_dia_val),
                    pd.Timestamp(melhor_dia_data).strftime("%d/%m/%Y") if melhor_dia_data is not None else None,
                )

                drill_view = drill.rename(columns={"DATA_DIA": "DATA", "VALOR_DIA": "VENDAS_R$"}).copy()
                drill_view["DATA"] = pd.to_datetime(drill_view["DATA"]).dt.strftime("%d/%m/%Y")
                drill_view["% DO TOTAL"] = drill_view["VENDAS_R$"].apply(
                    lambda v: (v / total_loja_diario * 100) if total_loja_diario else 0.0
                )

                linha_total_drill = pd.DataFrame([{
                    "DATA": "TOTAIS",
                    "VENDAS_R$": total_loja_diario,
                    "% DO TOTAL": 100.0 if total_loja_diario else 0.0,
                }])
                drill_view = pd.concat([drill_view, linha_total_drill], ignore_index=True)

                sty_drill = (
                    drill_view.style
                    .apply(
                        lambda row: ["background-color: #e9f1ff; color: #10234a; font-weight: 900;"] * len(row)
                        if str(row.get("DATA", "")).upper() == "TOTAIS"
                        else ["" if row.name % 2 == 0 else "background-color: #f4f8fe;"] * len(row),
                        axis=1,
                    )
                    .format({
                        "VENDAS_R$": lambda v: "R$ " + format_brl(v),
                        "% DO TOTAL": lambda v: f"{float(v):.2f}%".replace(".", ","),
                    })
                )
                st.dataframe(sty_drill, use_container_width=True, hide_index=True, height=420)

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
                lambda row: ["background-color: #e9f1ff; color: #10234a; font-weight: 900;"] * len(row)
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
            .apply(lambda row: ["background-color: #e9f1ff; color: #10234a; font-weight: 900;"] * len(row) if str(row.get("LOJA", "")).upper() == "TOTAL" else [""] * len(row), axis=1)
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

        # A tabela detalhada é exibida no desktop; no celular os mesmos dados
        # aparecem automaticamente como cards sem rolagem horizontal.
        sty_prev = prev_view.style.apply(row_total_style, axis=1)
        for col in ['DIF_PREV_X_ANO1_R$', 'CRESC_%', 'DIF_PREV_X_META_R$']:
            sty_prev = styler_map_compat(sty_prev, color_pos_neg, subset=[col])
        sty_prev = sty_prev.bar(subset=['ATING_META_%'], color='#1565c0', vmin=0, vmax=120, width=80, height=52)
        sty_prev = sty_prev.format({
            'REALIZADO_R$': lambda v: 'R$ '+format_brl(v),
            'DIAS_VENDA': lambda v: format_decimal_pt(v, 1),
            'MEDIA_DIA_R$': lambda v: 'R$ '+format_brl(v),
            'PREVISAO_R$': lambda v: 'R$ '+format_brl(v),
            'ANO_1_R$': lambda v: 'R$ '+format_brl(v),
            'DIF_PREV_X_ANO1_R$': lambda v: 'R$ '+format_brl(v),
            'CRESC_%': lambda v: format_decimal_pt(v, 2)+'%',
            'META_R$': lambda v: 'R$ '+format_brl(v),
            'DIF_PREV_X_META_R$': lambda v: 'R$ '+format_brl(v),
            'ATING_META_%': lambda v: format_decimal_pt(v, 2)+'%',
        }, na_rep='—')
        with st.container(key="previsao_desktop"):
            st.dataframe(sty_prev, use_container_width=True, hide_index=True, height=420)
        tabela_previsao_mobile(prev_view)


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
    bi_plotly_chart(fig_comp, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

    st.divider()

with aba_vendas:
    st.markdown("#### Performance comercial")

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
        bi_plotly_chart(fig_loja, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

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


with aba_marcas:
    st.markdown("#### Exploração de portfólio")

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

    # O seletor de marcas fica em um formulário próprio. Assim, marcar/desmarcar
    # opções não provoca o recálculo de todo o dashboard a cada clique.
    # Somente o botão abaixo aplica a seleção e executa os agrupamentos pesados.
    marcas_aplicadas_anteriores = [
        m for m in st.session_state.get("marcas_x_loja_aplicadas", [])
        if m in marcas_mxl
    ]
    with st.form("form_marcas_x_loja", clear_on_submit=False):
        marcas_escolhidas_mxl = st.multiselect(
            "Selecione uma ou mais marcas",
            options=marcas_mxl,
            default=marcas_aplicadas_anteriores,
            key="marcas_x_loja_widget",
        )
        aplicar_marcas_mxl = st.form_submit_button(
            "Aplicar marcas",
            use_container_width=True,
            type="primary",
        )

    if aplicar_marcas_mxl:
        st.session_state["marcas_x_loja_aplicadas"] = list(marcas_escolhidas_mxl)

    marcas_sel_mxl = [
        m for m in st.session_state.get("marcas_x_loja_aplicadas", [])
        if m in marcas_mxl
    ]

    if len(marcas_sel_mxl) == 0:
        st.info("Selecione ao menos uma marca para ver o ranking das lojas, as linhas e os produtos.")
    else:
        df_mxl = df_f.loc[df_f["MARCA_N"].isin(marcas_sel_mxl)]

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
            bi_plotly_chart(fig_mxl, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

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

with aba_vendas:

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
        bi_plotly_chart(fig_vend, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

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
        bi_plotly_chart(fig_seg, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

        segmento_tbl = segmento.copy()
        segmento_tbl["Faturamento (R$)"] = segmento_tbl["Faturamento"].apply(lambda v: "R$ " + format_brl(v))
        st.dataframe(segmento_tbl[["SEGMENTO", "Faturamento (R$)"]], use_container_width=True, hide_index=True, height=260)

        # ------------------------------------------------------------
        # Drill: Faturamento por Loja x Segmento Gerencial
        # ------------------------------------------------------------
        with st.expander("Abrir drill: Faturamento por Loja x Segmento Gerencial"):
            st.markdown("#### Faturamento por Loja x Segmento Gerencial")

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
            bi_plotly_chart(fig_seg_loja, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

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
                .apply(lambda row: ["background-color: #e9f1ff; color: #10234a; font-weight: 900;"] * len(row) if str(row.get("LOJA", "")).upper() == "TOTAL" else [""] * len(row), axis=1)
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

with aba_marcas:

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
            bi_plotly_chart(fig_ml, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

    st.divider()

with aba_clientes:
    st.markdown("#### Inteligência de clientes")

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
        bi_plotly_chart(fig_cli, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

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
            bi_plotly_chart(fig_lc, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

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
            bi_plotly_chart(fig_mc, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)


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

with aba_compras:
    st.markdown("#### Compras, CMV e auditoria dos dados")

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
                                        lojas_sel_aplicadas, data_ini, data_fim, mes_nums_sel=None) -> pd.DataFrame:
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
                out = out[(out["DATA"] >= pd.Timestamp(data_ini)) & (out["DATA"] < (pd.Timestamp(data_fim) + pd.Timedelta(days=1)))]
                if mes_nums_sel:
                    out = out[out["DATA"].dt.month.isin(mes_nums_sel)]
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


    def _render_indicador_compras_drill(df_f: pd.DataFrame, lojas_sel_aplicadas, data_ini, data_fim, mes_nums_sel=None):
        # Drill (expander) + computação sob demanda para não pesar a abertura do dashboard
        with st.expander("Abrir Indicador de Compras", expanded=False):
            st.caption("Tabela por loja (com TOTAL ao final), usando o mesmo recorte de lojas e período do dashboard.")
            show_tbl = st.checkbox("Calcular/mostrar tabela", value=False, key="show_indicador_compras_tbl")
            if not show_tbl:
                st.info("Marque a opção acima para calcular e exibir a tabela. (Isso deixa o carregamento inicial mais leve.)")
                return

            df_compras, df_devol = carregar_movimentacoes_compras()
            tbl = _calc_indicador_compras_por_loja(df_f, df_compras, df_devol, lojas_sel_aplicadas, data_ini, data_fim, mes_nums_sel)

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
    _render_indicador_compras_drill(df_f, lojas_sel_aplicadas, data_ini, data_fim, mes_nums_sel)
