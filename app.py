import re
import unicodedata
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from datetime import date

st.set_page_config(page_title="Dashboard de Vendas", layout="wide")
st.title("Dashboard de Vendas Dauto Tintas")

top_card = st.empty()

ARQUIVO_EXCEL = "BASE .xlsx"


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
    for c in df.columns:
        ckey = canonical_key(c)
        for wanted in cand_norm:
            if wanted and wanted in ckey:
                return c
    return None


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
        return "color: #1f77b4; font-weight: 700;"
    if x < 0:
        return "color: #d62728; font-weight: 700;"
    return ""


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


@st.cache_data(ttl=10)
def carregar_dados():
    df = pd.read_excel(ARQUIVO_EXCEL, sheet_name=0)
    df.columns = df.columns.astype(str).str.strip()

    obj_cols = df.select_dtypes(include=["object"]).columns
    if len(obj_cols) > 0:
        df[obj_cols] = df[obj_cols].astype("string")

    if "DATA" in df.columns:
        df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce", dayfirst=True)
    else:
        df["DATA"] = pd.NaT

    if "QTD" in df.columns:
        df["QTD_NUM"] = to_float_series(df["QTD"])
    else:
        df["QTD_NUM"] = pd.Series([None] * len(df), dtype="float64")

    if "UNIT" in df.columns:
        df["UNIT_NUM"] = to_float_series(df["UNIT"])
    else:
        df["UNIT_NUM"] = pd.Series([None] * len(df), dtype="float64")

    df["FAT_LINHA"] = df["QTD_NUM"] * df["UNIT_NUM"]

    df["LOJA_N"] = normalize_dim(df["LOJA"], "SEM LOJA") if "LOJA" in df.columns else pd.Series(["SEM LOJA"] * len(df), dtype="string")
    df["VENDEDOR_N"] = normalize_dim(df["VENDEDOR"], "SEM VENDEDOR") if "VENDEDOR" in df.columns else pd.Series(["SEM VENDEDOR"] * len(df), dtype="string")
    df["MARCA_N"] = normalize_dim(df["MARCA"], "SEM MARCA") if "MARCA" in df.columns else pd.Series(["SEM MARCA"] * len(df), dtype="string")
    df["SEGMENTO_N"] = normalize_dim(df["SEGMENTO"], "SEM SEGMENTO") if "SEGMENTO" in df.columns else pd.Series(["SEM SEGMENTO"] * len(df), dtype="string")
    df["LINHA_N"] = normalize_dim(df["LINHA"], "SEM LINHA") if "LINHA" in df.columns else pd.Series(["SEM LINHA"] * len(df), dtype="string")

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

    col_cliente = pick_first_existing_col(df, ["CLIENTE", "CLIENTE_NOME", "NOMECLIENTE", "RAZAOSOCIAL", "RAZAO SOCIAL"])
    if col_cliente is not None:
        df["CLIENTE_N"] = normalize_dim(df[col_cliente], "SEM CLIENTE")
    else:
        df["CLIENTE_N"] = pd.Series(["SEM CLIENTE"] * len(df), dtype="string")

    col_vr = pick_first_existing_col(df, ["VR TOTAL", "VRTOTAL", "VR_TOTAL", "VALOR TOTAL", "VALOR_TOTAL", "VR TOTAL (NF)", "VALOR A FATURAR (NF)", "VALOR A FATURAR NF", "VR.TOTAL"])
    if col_vr is not None:
        df["VR_TOTAL_NUM"] = to_float_series(df[col_vr])
    else:
        df["VR_TOTAL_NUM"] = pd.Series([None] * len(df), dtype="float64")

    col_custo = pick_first_existing_col(
        df,
        [
            "CUSTO TT + ST", "CUSTO TT+ST", "CUSTOTT+ST", "CUSTO TOTAL + ST", "CUSTO TOTAL+ST",
            "CUSTO+ST", "CUSTO + ST", "CUSTO_ST", "CUSTO ST", "CUSTO COM ST", "CUSTO COM ST (NF)",
        ],
    )
    if col_custo is not None:
        df["CUSTO_ST_NUM"] = to_float_series(df[col_custo])
    else:
        df["CUSTO_ST_NUM"] = pd.Series([None] * len(df), dtype="float64")

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

    df["LOJA_KEY"] = df["LOJA_N"].astype("string").fillna("").apply(canonical_key)

    return df


@st.cache_data(ttl=10)
def carregar_movimentacoes_compras():
    def _try_read_sheet(candidates: list[str]) -> pd.DataFrame:
        for sh in candidates:
            try:
                dfx = pd.read_excel(ARQUIVO_EXCEL, sheet_name=sh)
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

        col_data = pick_first_existing_col(dfm, ["DATA", "DATA EMISSAO", "DATA EMISSÃO", "DATA EMISSÃO NF", "DATA EMISSAO NF", "DT EMISSAO", "DT EMISSÃO"])
        if col_data is not None:
            dfm["DATA"] = pd.to_datetime(dfm[col_data], errors="coerce", dayfirst=True)
        else:
            dfm["DATA"] = pd.NaT

        col_loja = pick_first_existing_col(dfm, ["LOJA", "UNIDADE", "FILIAL", "UNID", "UNID."])
        if col_loja is not None:
            dfm["LOJA_N"] = normalize_dim(dfm[col_loja], "SEM LOJA")
        else:
            dfm["LOJA_N"] = pd.Series(["SEM LOJA"] * len(dfm), dtype="string")

        dfm["LOJA_KEY"] = dfm["LOJA_N"].astype("string").fillna("").apply(canonical_key)

        col_tot = pick_first_existing_col(dfm, ["TOT. DOC", "TOT DOC", "TOT.DOC", "TOTAL DOC", "TOTAL DOCUMENTO", "VALOR DOC", "VALOR DOCUMENTO"])
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
def carregar_bases_planejamento():
    try:
        dias_raw = pd.read_excel(ARQUIVO_EXCEL, sheet_name="DIAS ÚTEIS")
        dias_raw.columns = dias_raw.columns.astype(str).str.strip()
    except Exception:
        dias_raw = pd.DataFrame()

    dias_rows = []
    if not dias_raw.empty:
        col_mes = pick_first_existing_col(dias_raw, ["MÊS", "MES"])
        col_dias = pick_first_existing_col(dias_raw, ["DIAS ÚTEIS EQUIVALENTES", "DIAS UTEIS EQUIVALENTES", "DIAS ÚTEIS", "DIAS UTEIS"])
        for _, r in dias_raw.iterrows():
            mes_txt = str(r[col_mes]).strip() if col_mes is not None and pd.notna(r[col_mes]) else ""
            dias_val = parse_number_any(r[col_dias]) if col_dias is not None else None
            mes_key = canonical_key(mes_txt)
            mapa_mes_ext = {
                "JANEIRO": 1, "FEVEREIRO": 2, "MARCO": 3, "MARÇO": 3, "ABRIL": 4,
                "MAIO": 5, "JUNHO": 6, "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9,
                "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12,
                "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
                "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12,
            }
            mes_num = mapa_mes_ext.get(mes_key)
            if mes_num is not None:
                dias_rows.append({"MES_NUM": mes_num, "DIAS_UTEIS": float(dias_val or 0.0)})
    dias_df = pd.DataFrame(dias_rows)
    if dias_df.empty:
        dias_df = pd.DataFrame({"MES_NUM": pd.Series(dtype="int64"), "DIAS_UTEIS": pd.Series(dtype="float64")})

    def _reshape_matriz(sheet_name: str, value_name: str) -> pd.DataFrame:
        try:
            raw = pd.read_excel(ARQUIVO_EXCEL, sheet_name=sheet_name)
            raw.columns = raw.columns.astype(str).str.strip()
        except Exception:
            raw = pd.DataFrame()

        if raw.empty:
            return pd.DataFrame({"LOJA_KEY": pd.Series(dtype="string"), "MES_NUM": pd.Series(dtype="int64"), value_name: pd.Series(dtype="float64")})

        col_mes = pick_first_existing_col(raw, ["MÊS", "MES", "LOJA"])
        col_ano = pick_first_existing_col(raw, ["ANO"])
        loja_cols = [c for c in raw.columns if c not in {col_mes, col_ano}]
        rows = []
        mapa_mes_ext = {
            "JANEIRO": 1, "FEVEREIRO": 2, "MARCO": 3, "MARÇO": 3, "ABRIL": 4,
            "MAIO": 5, "JUNHO": 6, "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9,
            "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12,
            "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
            "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12,
        }

        for _, r in raw.iterrows():
            mes_txt = str(r[col_mes]).strip() if col_mes is not None and pd.notna(r[col_mes]) else ""
            mes_num = mapa_mes_ext.get(canonical_key(mes_txt))
            if mes_num is None:
                continue
            for c in loja_cols:
                loja_key = canonical_key(c)
                val = parse_number_any(r[c])
                rows.append({"LOJA_KEY": loja_key, "MES_NUM": mes_num, value_name: float(val or 0.0)})

        out = pd.DataFrame(rows)
        if out.empty:
            out = pd.DataFrame({"LOJA_KEY": pd.Series(dtype="string"), "MES_NUM": pd.Series(dtype="int64"), value_name: pd.Series(dtype="float64")})
        return out

    meta_df = _reshape_matriz("META DAUTO TINTAS", "META")
    ano1_df = _reshape_matriz("ANO-1", "VENDAS_2025")

    return dias_df, meta_df, ano1_df


def obter_projecao_faturamento(df_base_ano: pd.DataFrame, dias_df: pd.DataFrame, meses_nums: list[int]):
    if df_base_ano is None or df_base_ano.empty or not meses_nums:
        return 0.0, 0.0, 0.0

    meses_nums = sorted(set(int(m) for m in meses_nums))
    dias_map = {}
    if dias_df is not None and not dias_df.empty:
        dias_map = dias_df.groupby("MES_NUM")["DIAS_UTEIS"].sum().to_dict()

    realizado_total = 0.0
    proj_total = 0.0
    dias_total = 0.0

    for mes_num in meses_nums:
        dfm = df_base_ano[df_base_ano["DATA"].dt.month == mes_num].copy()
        realizado_mes = float(dfm["FAT_LINHA"].sum()) if len(dfm) else 0.0
        dias_mes = float(dias_map.get(mes_num, 0.0) or 0.0)
        if dias_mes > 0:
            dias_decorridos = float(dfm["DATA"].dt.normalize().nunique()) if len(dfm) else 0.0
            dias_decorridos = min(dias_decorridos, dias_mes)
            media_dia = (realizado_mes / dias_decorridos) if dias_decorridos > 0 else 0.0
            proj_mes = media_dia * dias_mes
            dias_total += dias_mes
        else:
            proj_mes = realizado_mes
        realizado_total += realizado_mes
        proj_total += proj_mes

    media_total = (realizado_total / dias_total) if dias_total > 0 else 0.0
    return proj_total, media_total, dias_total


ANO_BASE = 2025
ANO_ATUAL = 2026


df = carregar_dados()
df = df[df["FAT_LINHA"].notna()].copy()

use_vr_total = df["VR_TOTAL_NUM"].notna().any()
VAL_COL = "VR_TOTAL_NUM" if use_vr_total else "FAT_LINHA"

dias_uteis_df, meta_plan_df, ano1_plan_df = carregar_bases_planejamento()

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

df_f = df.copy()
if lojas_sel_aplicadas:
    df_f = df_f[df_f["LOJA_N"].isin(lojas_sel_aplicadas)]
else:
    df_f = df_f.iloc[0:0]

df_f = df_f[df_f["DATA"].notna()]
df_f = df_f[(df_f["DATA"].dt.date >= data_ini) & (df_f["DATA"].dt.date <= data_fim)]
fat_total = float(df_f["FAT_LINHA"].sum()) if len(df_f) else 0.0

st.divider()
st.markdown("### Seleção de meses (para análise e tabelas)")
mes_opts = [nome for _, nome in MESES]
mes_sel_multi = st.multiselect("Selecione 1 ou mais meses", options=mes_opts, default=[mes_opts[0]])
mes_nome_to_num = {nome: num for num, nome in MESES}
mes_nums_sel = [mes_nome_to_num[m] for m in mes_sel_multi if m in mes_nome_to_num]
if not mes_nums_sel:
    mes_nums_sel = [num for num, _ in MESES]
mes_sel_label = ", ".join([m for m in mes_opts if mes_nome_to_num[m] in mes_nums_sel])
st.markdown(f"**Meses selecionados:** {mes_sel_label}")

st.subheader("Comparativo: 2025 (Ano-1) x 2026 (Ano Atual)")
lojas_keys_aplicadas = [canonical_key(x) for x in lojas_sel_aplicadas if canonical_key(x)]
df_2025 = ano1_plan_df.copy()
if lojas_keys_aplicadas:
    df_2025 = df_2025[df_2025["LOJA_KEY"].isin(lojas_keys_aplicadas)].copy()
else:
    df_2025 = df_2025.iloc[0:0].copy()

df_2026 = df.copy()
df_2026 = df_2026[df_2026["DATA"].notna()].copy()
df_2026 = df_2026[df_2026["DATA"].dt.year == ANO_ATUAL].copy()
if lojas_sel_aplicadas:
    df_2026 = df_2026[df_2026["LOJA_N"].isin(lojas_sel_aplicadas)].copy()
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
df_comp["VAR_%"] = df_comp.apply(lambda r: (r["VAR_R$"] / r["VENDAS_2025"] * 100) if r["VENDAS_2025"] not in (0, None) else None, axis=1)

map_key_to_loja = df[["LOJA_KEY", "LOJA_N"]].dropna().drop_duplicates().groupby("LOJA_KEY")["LOJA_N"].first().to_dict()
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

st.markdown("### Meta × Realizado")
df_meta = meta_plan_df.copy()
if lojas_keys_aplicadas:
    df_meta = df_meta[df_meta["LOJA_KEY"].isin(lojas_keys_aplicadas)].copy()
else:
    df_meta = df_meta.iloc[0:0].copy()
df_meta_sel = df_meta[df_meta["MES_NUM"].isin(mes_nums_sel)].copy()

meta_total_sel = float(df_meta_sel["META"].sum()) if len(df_meta_sel) else 0.0
real_total_sel = float(df_mes["VENDAS_2026"].sum()) if len(df_mes) else 0.0
projecao_total_sel, media_dia_util_sel, dias_uteis_total_sel = obter_projecao_faturamento(df_2026, dias_uteis_df, mes_nums_sel)
top_card.metric(
    "Faturamento Atual (R$)",
    "R$ " + format_brl(real_total_sel),
    delta="Proj. R$ " + format_brl(projecao_total_sel),
    help="Projeção calculada pela média faturada por dia útil × total de dias úteis do mês/meses selecionados.",
)

pct_meta = (real_total_sel / meta_total_sel * 100) if meta_total_sel != 0 else None
dif_meta_r = real_total_sel - meta_total_sel
dif_meta_p = (dif_meta_r / meta_total_sel * 100) if meta_total_sel != 0 else None

a1, a2, a3 = st.columns([1.0, 1.0, 1.6], gap="large")
with a1:
    st.metric("Meta (R$)", "R$ " + format_brl(meta_total_sel))
with a2:
    st.metric("Realizado (R$)", "R$ " + format_brl(real_total_sel))
with a3:
    if pct_meta is None:
        st.info("Meta zerada para o recorte selecionado (não é possível calcular %).")
    else:
        fig_gauge = go.Figure(go.Indicator(mode="gauge+number", value=max(0, pct_meta), number={"suffix": "%", "valueformat": ".1f"}, title={"text": "% da meta atingida"}, gauge={"axis": {"range": [0, 120]}, "bar": {"color": "#1f77b4"}, "steps": [{"range": [0, 80], "color": "#f2f2f2"}, {"range": [80, 100], "color": "#e6e6e6"}, {"range": [100, 120], "color": "#d9d9d9"}], "threshold": {"line": {"color": "#d62728", "width": 3}, "thickness": 0.75, "value": 100}}))
        fig_gauge.update_layout(margin=dict(l=20, r=20, t=60, b=10), height=260)
        st.plotly_chart(fig_gauge, use_container_width=True, config=PLOT_CONFIG_INTERACTIVE_NO_ZOOM)

b1, b2, b3, b4 = st.columns(4)
with b1:
    st.metric("Diferença (Realizado − Meta) (R$)", "R$ " + format_brl(dif_meta_r))
with b2:
    st.metric("Diferença (%)", (f"{dif_meta_p:.2f}%".replace(".", ",")) if dif_meta_p is not None else "—")
with b3:
    st.metric("Projeção de Faturamento (R$)", "R$ " + format_brl(projecao_total_sel))
with b4:
    st.metric("Média por Dia Útil (R$)", "R$ " + format_brl(media_dia_util_sel))

st.info("Colei a versão completa até aqui no canvas. Como o arquivo inteiro tem muitas linhas, continue copiando a partir do canvas lateral para não cortar nada no chat.")
