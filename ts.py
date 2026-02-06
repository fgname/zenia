import io
import unicodedata
import numpy as np
import pandas as pd
import streamlit as st

# =========================================================
# UI
# =========================================================
st.set_page_config(page_title="Automação ZEN — Relatórios", layout="wide")
st.title("Automação ZEN — Tarefas por Demanda + Estoque Mínimo")

st.write(
    "1) Envie os 3 arquivos (ZEN, saldo e estoque mínimo). "
    "2) Clique em **Gerar relatório**. "
    "3) Baixe o Excel final com as 2 abas prontas."
)

col1, col2, col3 = st.columns(3)
with col1:
    up_ped = st.file_uploader("ZEN.xlsx (base de pedidos)", type=["xlsx"])
with col2:
    up_saldo = st.file_uploader("dados_saldo.xlsx (saldo por endereço)", type=["xlsx"])
with col3:
    up_min = st.file_uploader("estoque minimo.xlsx (abas: estoque minimo + curva sku)", type=["xlsx"])

st.divider()

FILTRAR_SITUACAO = st.text_input("Filtrar por SITUACAO (opcional, ex: LBS). Deixe vazio para não filtrar:", "").strip()
if FILTRAR_SITUACAO == "":
    FILTRAR_SITUACAO = None


# =========================================================
# FUNÇÕES
# =========================================================
def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def require_cols(df: pd.DataFrame, cols, df_name="DataFrame"):
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise ValueError(
            f"{df_name}: faltando colunas {missing}. "
            f"Colunas disponíveis: {list(df.columns)}"
        )

def safe_int_series(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0).astype(int)

def strip_accents(s: str) -> str:
    s = s.replace("\n", " ").strip()
    s = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in s if not unicodedata.combining(ch))

def find_col_by_keywords(columns, keywords_any, keywords_all=()):
    cols = list(columns)
    norm = [strip_accents(c).lower() for c in cols]
    for orig, n in zip(cols, norm):
        any_ok = any(k.lower() in n for k in keywords_any)
        all_ok = all(k.lower() in n for k in keywords_all)
        if any_ok and all_ok:
            return orig
    return None

def classificar_endereco(endereco) -> str:
    """
    Regras:
    - Endereço == '01' => RUA
    - Senão pega 6º e 7º caractere (endereco[5:7])
      01-02 => PICKING
      >=03  => PULMAO
    """
    if pd.isna(endereco):
        return "DESCONHECIDO"
    e = str(endereco).strip()
    if e == "01":
        return "RUA"
    if len(e) < 7:
        return "DESCONHECIDO"
    nivel_txt = e[5:7]
    if not nivel_txt.isdigit():
        return "DESCONHECIDO"
    nivel = int(nivel_txt)
    if nivel in (1, 2):
        return "PICKING"
    if nivel >= 3:
        return "PULMAO"
    return "DESCONHECIDO"


def gerar_relatorio(ped: pd.DataFrame, saldo: pd.DataFrame, min_df: pd.DataFrame, curva_df: pd.DataFrame):
    # ---------------------------
    # validações
    # ---------------------------
    require_cols(ped, ["ITEM TECADI", "NUM PEDIDO"], "ZEN.xlsx")
    require_cols(saldo, ["Produto", "Endereço", "Saldo"], "dados_saldo.xlsx")
    require_cols(min_df, ["SKU (ZEN)"], "estoque minimo.xlsx -> aba 'estoque minimo'")
    require_cols(curva_df, ["SKU (ZEN)"], "estoque minimo.xlsx -> aba 'curva sku'")

    col_min = find_col_by_keywords(min_df.columns, keywords_any=["minimo"], keywords_all=["estoque"])
    if col_min is None:
        col_min = find_col_by_keywords(min_df.columns, keywords_any=["minimo"])
    if col_min is None:
        raise ValueError(f"Não encontrei a coluna de Estoque Mínimo na aba 'estoque minimo'. Colunas: {list(min_df.columns)}")

    col_curva = find_col_by_keywords(curva_df.columns, keywords_any=["curva"])
    if col_curva is None:
        raise ValueError(f"Não encontrei a coluna de Curva na aba 'curva sku'. Colunas: {list(curva_df.columns)}")

    # ---------------------------
    # limpeza / padronização
    # ---------------------------
    ped = ped.copy()
    saldo = saldo.copy()

    ped["ITEM TECADI"] = ped["ITEM TECADI"].astype(str).str.strip()
    ped["NUM PEDIDO"]  = ped["NUM PEDIDO"].astype(str).str.strip()

    if "QTDE ENVIADA" in ped.columns:
        ped["QTDE ENVIADA"] = pd.to_numeric(ped["QTDE ENVIADA"], errors="coerce").fillna(0)
    else:
        ped["QTDE ENVIADA"] = 0

    if "CANCELADO" in ped.columns:
        ped = ped[ped["CANCELADO"].fillna(0).astype(int) == 0]
    if "FATURADO" in ped.columns:
        ped = ped[ped["FATURADO"].fillna(0).astype(int) == 0]
    if FILTRAR_SITUACAO and "SITUACAO" in ped.columns:
        ped = ped[ped["SITUACAO"].astype(str).str.strip().eq(str(FILTRAR_SITUACAO))]

    saldo["Produto"]  = saldo["Produto"].astype(str).str.strip()
    saldo["Endereço"] = saldo["Endereço"].astype(str).str.strip()
    saldo["Saldo"]    = pd.to_numeric(saldo["Saldo"], errors="coerce").fillna(0)

    minimo = min_df[["SKU (ZEN)", col_min]].copy()
    minimo.columns = ["SKU", "Estoque mínimo"]
    minimo["SKU"] = minimo["SKU"].astype(str).str.strip()
    minimo["Estoque mínimo"] = pd.to_numeric(minimo["Estoque mínimo"], errors="coerce").fillna(0).astype(int)

    curva = curva_df[["SKU (ZEN)", col_curva]].copy()
    curva.columns = ["SKU", "Curva do SKU"]
    curva["SKU"] = curva["SKU"].astype(str).str.strip()
    curva["Curva do SKU"] = curva["Curva do SKU"].astype(str).str.strip()
    curva.loc[curva["Curva do SKU"].isin(["nan", "None"]), "Curva do SKU"] = ""

    # ---------------------------
    # classifica saldo
    # ---------------------------
    saldo2 = saldo.copy()
    saldo2["classe"] = saldo2["Endereço"].apply(classificar_endereco)

    saldo_picking = (
        saldo2[saldo2["classe"] == "PICKING"]
        .groupby("Produto", as_index=False)["Saldo"].sum()
        .rename(columns={"Produto": "SKU", "Saldo": "Saldo picking"})
    )

    saldo_pulmao = (
        saldo2[saldo2["classe"] == "PULMAO"]
        .groupby("Produto", as_index=False)["Saldo"].sum()
        .rename(columns={"Produto": "SKU", "Saldo": "Quantidade pulmão"})
    )

    saldo_rua = (
        saldo2[saldo2["classe"] == "RUA"]
        .groupby("Produto", as_index=False)["Saldo"].sum()
        .rename(columns={"Produto": "SKU", "Saldo": "Saldo rua"})
    )

    # =========================================================
    # RELATÓRIO 1) DEMANDA DE PEDIDOS
    # =========================================================
    skus_demanda = pd.Series(ped["ITEM TECADI"].dropna().unique(), name="SKU")

    demanda = (
        ped.groupby("ITEM TECADI", as_index=False)["QTDE ENVIADA"].sum()
        .rename(columns={"ITEM TECADI": "SKU", "QTDE ENVIADA": "Quantidade de peças demanda"})
    )

    sku_para_pedidos = (
        ped.groupby("ITEM TECADI")["NUM PEDIDO"]
        .apply(lambda s: set(s.dropna().astype(str)))
        .to_dict()
    )
    linhas_por_pedido = ped.groupby("NUM PEDIDO").size().to_dict()

    def linhas_dependem(sku: str) -> int:
        pedidos_do_sku = sku_para_pedidos.get(sku, set())
        return int(sum(linhas_por_pedido.get(p, 0) for p in pedidos_do_sku))

    linhas_dep = pd.DataFrame({
        "SKU": skus_demanda,
        "Linhas que dependem do abastecimento": [linhas_dependem(s) for s in skus_demanda]
    })

    out_demanda = (
        pd.DataFrame({"SKU": skus_demanda})
        .merge(demanda, on="SKU", how="left")
        .merge(linhas_dep, on="SKU", how="left")
        .merge(saldo_picking, on="SKU", how="left")
        .merge(saldo_pulmao.rename(columns={"Quantidade pulmão": "Saldo pulmão"}), on="SKU", how="left")
        .merge(saldo_rua, on="SKU", how="left")
    )

    out_demanda["Quantidade de peças demanda"] = safe_int_series(out_demanda["Quantidade de peças demanda"])
    out_demanda["Linhas que dependem do abastecimento"] = safe_int_series(out_demanda["Linhas que dependem do abastecimento"])
    out_demanda["Saldo picking"] = safe_int_series(out_demanda["Saldo picking"])
    out_demanda["Saldo pulmão"] = safe_int_series(out_demanda["Saldo pulmão"])
    out_demanda["Saldo rua"] = safe_int_series(out_demanda["Saldo rua"])

    out_demanda["Quantidade abastecimento"] = (out_demanda["Quantidade de peças demanda"] - out_demanda["Saldo picking"]).clip(lower=0).astype(int)

    def necessidade_demanda(row):
        dem = row["Quantidade de peças demanda"]
        pick = row["Saldo picking"]
        pul = row["Saldo pulmão"]
        if dem <= pick:
            return "OK (Picking cobre)"
        if dem <= pick + pul:
            return "Abastecer do pulmão"
        return "Sem Saldo Pulmão"

    out_demanda["Necessidade abastecimento?"] = out_demanda.apply(necessidade_demanda, axis=1)

    out_demanda = out_demanda[[
        "SKU",
        "Quantidade de peças demanda",
        "Linhas que dependem do abastecimento",
        "Saldo picking",
        "Saldo pulmão",
        "Saldo rua",
        "Necessidade abastecimento?",
        "Quantidade abastecimento"
    ]].sort_values(
        ["Necessidade abastecimento?", "Linhas que dependem do abastecimento", "Quantidade abastecimento"],
        ascending=[True, False, False]
    ).reset_index(drop=True)

    # =========================================================
    # RELATÓRIO 2) ESTOQUE MÍNIMO (com curva)
    # =========================================================
    out_min = (
        minimo[["SKU", "Estoque mínimo"]]
        .merge(curva, on="SKU", how="left")
        .merge(saldo_picking, on="SKU", how="left")
        .merge(saldo_pulmao, on="SKU", how="left")
        .merge(saldo_rua, on="SKU", how="left")
    )

    out_min["Saldo picking"] = safe_int_series(out_min["Saldo picking"])
    out_min["Quantidade pulmão"] = safe_int_series(out_min["Quantidade pulmão"])
    out_min["Saldo rua"] = safe_int_series(out_min["Saldo rua"])
    out_min["Estoque mínimo"] = safe_int_series(out_min["Estoque mínimo"])
    out_min["Curva do SKU"] = out_min["Curva do SKU"].fillna("")

    def necessidade_min(row):
        pul = row["Quantidade pulmão"]
        pick = row["Saldo picking"]
        minimo_val = row["Estoque mínimo"]
        if pul <= 0:
            return "Não"
        if (pick - minimo_val) <= 0:
            return "Sim"
        return "Não"

    out_min["Necessidade abastecimento"] = out_min.apply(necessidade_min, axis=1)

    falta = (out_min["Estoque mínimo"] - out_min["Saldo picking"]).clip(lower=0)
    out_min["Quantidade abastecimento"] = np.where(
        out_min["Necessidade abastecimento"].eq("Sim"),
        np.minimum(falta, out_min["Quantidade pulmão"]),
        0
    ).astype(int)

    out_min["% faltante estoque mínimo"] = np.clip(
        np.where(
            (out_min["Quantidade abastecimento"] == 0) | (out_min["Estoque mínimo"] <= 0),
            0.0,
            1 - (out_min["Saldo picking"] / out_min["Estoque mínimo"])
        ),
        0, 1
    )

    out_min = out_min[[
        "SKU",
        "Saldo picking",
        "Estoque mínimo",
        "Quantidade pulmão",
        "Necessidade abastecimento",
        "Quantidade abastecimento",
        "% faltante estoque mínimo",
        "Curva do SKU",
        "Saldo rua"
    ]].sort_values(
        ["Necessidade abastecimento", "Quantidade abastecimento", "% faltante estoque mínimo"],
        ascending=[True, False, False]
    ).reset_index(drop=True)

    return out_demanda, out_min


# =========================================================
# AÇÃO
# =========================================================
btn = st.button("🚀 Gerar relatório", type="primary", disabled=not (up_ped and up_saldo and up_min))

if btn:
    try:
        with st.spinner("Lendo arquivos e gerando relatório..."):
            ped = normalize_cols(pd.read_excel(up_ped, sheet_name=0))
            saldo = normalize_cols(pd.read_excel(up_saldo, sheet_name=0))
            min_df = normalize_cols(pd.read_excel(up_min, sheet_name="estoque minimo"))
            curva_df = normalize_cols(pd.read_excel(up_min, sheet_name="curva sku"))

            out_demanda, out_min = gerar_relatorio(ped, saldo, min_df, curva_df)

            # Exporta em memória (download)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                out_demanda.to_excel(writer, sheet_name="Tarefas por demanda de pedidos", index=False)
                out_min.to_excel(writer, sheet_name="Tarefas por estoque mínimo", index=False)

            buffer.seek(0)

        st.success("Relatório gerado com sucesso ✅")

        st.download_button(
            label="⬇️ Baixar Excel com as 2 abas",
            data=buffer,
            file_name="analise_tarefas_ZEN.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("Prévia — Demanda de pedidos"):
            st.dataframe(out_demanda, use_container_width=True)

        with st.expander("Prévia — Estoque mínimo"):
            st.dataframe(out_min, use_container_width=True)

    except Exception as e:
        st.error(f"Erro ao gerar relatório: {e}")
        st.exception(e)
