# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime

# ----- Ajustes globais de performance -----
pd.options.mode.copy_on_write = True

st.set_page_config(page_title="Controle de Manutenção", layout="wide")

# ===========================
# Helpers de limpeza
# ===========================

def _make_unique_cols(cols):
    """
    - Normaliza nomes (remove \n, trim)
    - Substitui vazios/Unnamed por BLANK_<idx>
    - Garante unicidade com sufixo __1, __2, ...
    """
    out = []
    seen = {}
    for i, c in enumerate(cols):
        name = "" if c is None else str(c)
        name = name.replace("\n", " ").strip()
        if name == "" or name.lower().startswith("unnamed"):
            name = f"BLANK_{i+1}"
        if name in seen:
            seen[name] += 1
            name = f"{name}__{seen[name]}"
        else:
            seen[name] = 0
        out.append(name)
    return out

def _normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    # tira brancos totais antes de renomear
    empty_cols = [c for c in df.columns if df[c].notna().sum() == 0]
    if empty_cols:
        df = df.drop(columns=empty_cols)
    # renomeia com unicidade
    df.columns = _make_unique_cols(df.columns)
    return df

# ===========================
# Carregamento OTIMIZADO
# ===========================

@st.cache_data(show_spinner=True)
def load_requisicoes_smart(xlsx_path: str, sheet_hint: str = "requis") -> pd.DataFrame:
    """
    Lê a planilha de Requisições em modo streaming (openpyxl read_only),
    sem puxar 1M de linhas vazias. Para quando encontra um bloco grande
    de linhas 100% vazias.
    """
    from openpyxl import load_workbook

    wb = load_workbook(xlsx_path, data_only=True, read_only=True)
    # Tenta localizar a aba pelo "hint"
    target_name = None
    for nm in wb.sheetnames:
        if sheet_hint.lower() in nm.lower():
            target_name = nm
            break
    ws = wb[target_name or wb.sheetnames[0]]

    rows = ws.iter_rows(values_only=True)
    header = None

    # Busca primeira linha de cabeçalho
    for r in rows:
        if r and any(r):
            header = [str(c).strip() if c is not None else "" for c in r]
            break
    if not header:
        return pd.DataFrame()

    data = []
    empty_streak = 0
    EMPTY_STOP = 400  # para após 400 linhas vazias seguidas (ajuste se necessário)

    for r in rows:
        # linha completamente vazia?
        if not r or not any(r):
            empty_streak += 1
            if empty_streak >= EMPTY_STOP:
                break
            continue
        empty_streak = 0
        data.append(list(r))

    df = pd.DataFrame(data, columns=header)
    df = df.dropna(how="all")
    df = _normalize_cols(df)

    # Tipagens importantes
    if "VALOR" in df.columns:
        df["VALOR"] = pd.to_numeric(df["VALOR"], errors="coerce").fillna(0)

    # Converte datas quando existentes
    for col in [
        "MÊS COMPETÊNCIA","DATA DE CRIAÇÃO","Data Aprovação","DATA RECEBIMENTO",
        "DATA DO DOC","DATA DE ENTRADA","DATA DE LANÇAMENTO"
    ]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Campo MÊS (YYYY-MM)
    if "MÊS COMPETÊNCIA" in df.columns and pd.api.types.is_datetime64_any_dtype(df["MÊS COMPETÊNCIA"]):
        df["MÊS"] = df["MÊS COMPETÊNCIA"].dt.to_period("M").astype(str)
    else:
        if "MÊS COMPETÊNCIA" in df.columns:
            try:
                _m = pd.to_datetime(df["MÊS COMPETÊNCIA"], errors="coerce")
                df["MÊS"] = _m.dt.to_period("M").astype(str)
            except Exception:
                df["MÊS"] = pd.NaT
        else:
            df["MÊS"] = pd.NaT

    # Ajuste de 'CD ' -> 'CD', se necessário
    if "CD" not in df.columns and "CD " in df.columns:
        df["CD"] = df["CD "]

    return df

@st.cache_data(show_spinner=True)
def load_budget(xlsx_path: str, sheet_hint: str = "budget") -> pd.DataFrame:
    # Lê só a aba de budget (formato largo com meses)
    xl = pd.ExcelFile(xlsx_path, engine="openpyxl")
    name = next((nm for nm in xl.sheet_names if sheet_hint.lower() in nm.lower()), xl.sheet_names[0])
    df = pd.read_excel(xl, sheet_name=name, dtype_backend="pyarrow")
    return _normalize_cols(df)

@st.cache_data(show_spinner=True)
def load_estornos(xlsx_path: str, sheet_hint: str = "estono") -> pd.DataFrame:
    try:
        xl = pd.ExcelFile(xlsx_path, engine="openpyxl")
        name = next((nm for nm in xl.sheet_names if sheet_hint.lower() in nm.lower()), None)
        if not name:
            return pd.DataFrame()
        df = pd.read_excel(xl, sheet_name=name, dtype_backend="pyarrow")
        return _normalize_cols(df)
    except Exception:
        return pd.DataFrame()

@st.cache_data(show_spinner=True)
def load_data():
    xlsx_path = "data/controle.xlsx"
    req_df = load_requisicoes_smart(xlsx_path, sheet_hint="requis")  # carga “inteligente”
    budget_df = load_budget(xlsx_path, sheet_hint="budget")
    estorno_df = load_estornos(xlsx_path, sheet_hint="estono")
    return {"req": req_df, "budget": budget_df, "estorno": estorno_df}

data = load_data()
req_df = data["req"]
budget_df = data["budget"]
estorno_df = data["estorno"]

# Abort early se não leu nada
if req_df.empty:
    st.error("Não foi possível carregar a planilha de Requisições (aba 'Base_Requisicoes'). Verifique o arquivo em data/controle.xlsx.")
    st.stop()

# ===========================
# Filtros / KPIs / Gráficos
# ===========================

# Sidebar filtros
st.sidebar.title('Filtros')

cd_opts = sorted([x for x in req_df.get('CD', pd.Series(dtype=str)).dropna().astype(str).unique()])
cd_sel = st.sidebar.multiselect('CD (Centro de Distribuição)', cd_opts, default=cd_opts)

ano_opts = sorted([int(x) for x in pd.to_numeric(req_df.get('ANO', pd.Series()), errors='coerce').dropna().unique()])
ano_sel = st.sidebar.multiselect('Ano', ano_opts, default=ano_opts)

mes_opts = sorted(req_df['MÊS'].dropna().unique())
mes_sel = st.sidebar.multiselect('Mês competência', mes_opts, default=mes_opts)

grupo_opts = sorted(req_df.get('Grupo', pd.Series(dtype=str)).dropna().astype(str).unique())
grupo_sel = st.sidebar.multiselect('Grupo', grupo_opts, default=grupo_opts)

subgrupo_opts = sorted(req_df.get('SubGrupo', pd.Series(dtype=str)).dropna().astype(str).unique())
subgrupo_sel = st.sidebar.multiselect('Subgrupo', subgrupo_opts, default=subgrupo_opts)

status_opts = sorted(req_df.get('STATUS', pd.Series(dtype=str)).dropna().astype(str).unique())
status_sel = st.sidebar.multiselect('Status Requisição', status_opts, default=status_opts)

# Aplicar filtros
f = (
    req_df
    .assign(
        CD=lambda d: d.get('CD', pd.Series(dtype=str)).astype(str),
        ANO=lambda d: pd.to_numeric(d.get('ANO', pd.Series()), errors='coerce'),
        Grupo=lambda d: d.get('Grupo', pd.Series(dtype=str)).astype(str),
        SubGrupo=lambda d: d.get('SubGrupo', pd.Series(dtype=str)).astype(str),
        STATUS=lambda d: d.get('STATUS', pd.Series(dtype=str)).astype(str),
    )
)

if cd_sel:
    f = f[f['CD'].isin(cd_sel)]
if ano_sel:
    f = f[f['ANO'].isin(ano_sel)]
if mes_sel:
    f = f[f['MÊS'].isin(mes_sel)]
if grupo_sel:
    f = f[f['Grupo'].isin(grupo_sel)]
if subgrupo_sel:
    f = f[f['SubGrupo'].isin(subgrupo_sel)]
if status_sel:
    f = f[f['STATUS'].isin(status_sel)]

# Top KPIs
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.metric('Requisições (itens)', f.shape[0])
with c2:
    st.metric('Valor total (R$)', float(f['VALOR'].sum()))
with c3:
    st.metric('Ticket médio (R$)', float((f['VALOR'].sum() / f.shape[0]) if f.shape[0] else 0))
with c4:
    aprovados = f[f.get('STATUS','')=='APROVADO']['VALOR'].sum() if 'STATUS' in f.columns else 0
    st.metric('Aprovado (R$)', float(aprovados))

st.markdown('---')

# Gráficos
g1, g2 = st.columns(2)
with g1:
    by_mes = f.groupby('MÊS', dropna=True)['VALOR'].sum().sort_index()
    fig, ax = plt.subplots(figsize=(6,3))
    by_mes.plot(kind='bar', ax=ax, color='#1f77b4')
    ax.set_title('Despesas por mês (R$)')
    ax.set_xlabel('Mês')
    ax.set_ylabel('Valor (R$)')
    st.pyplot(fig)
with g2:
    by_grupo = f.groupby('Grupo')['VALOR'].sum().sort_values(ascending=False).head(10)
    fig2, ax2 = plt.subplots(figsize=(6,3))
    by_grupo.plot(kind='barh', ax=ax2, color='#2ca02c')
    ax2.set_title('Top 10 grupos por valor (R$)')
    ax2.set_xlabel('Valor (R$)')
    ax2.set_ylabel('Grupo')
    st.pyplot(fig2)

st.markdown('---')

# Tabela detalhada
st.subheader('Requisições filtradas')
# .copy() só para garantir isolamento de view e evitar warnings
st.dataframe(f.copy(), use_container_width=True, height=400)

# Orç vs execução (BGT vs. REQ)
if budget_df is not None and not budget_df.empty:
    st.subheader('Orçado x Requisitado')

    # Transform budget (mensal) para formato longo
    budget = budget_df.copy()
    budget.columns = [str(c).strip().upper() for c in budget.columns]
    meses = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO']
    present = [m for m in meses if m in budget.columns]

    if present:
        bgt_long = budget.melt(
            id_vars=[c for c in budget.columns if c not in present],
            value_vars=present,
            var_name='MES_NOME', value_name='BGT_VALOR'
        )
        # Mapear nome mês para número/periodo
        mapa_mes = {mes:i+1 for i, mes in enumerate(meses)}
        bgt_long['MES_NUM'] = bgt_long['MES_NOME'].map(mapa_mes)
        bgt_long['ANO'] = datetime.now().year
        bgt_long['MÊS'] = bgt_long.apply(lambda r: f"{int(r['ANO'])}-{int(r['MES_NUM']):02d}", axis=1)

        # Vincular por Conta+Centro (quando existir)
        # Nas requisições, usar colunas 'Conta' e 'Centro de Custo'
        req_aux = req_df.copy()
        req_aux.columns = [str(c).strip().upper() for c in req_aux.columns]
        if 'CONTA' in req_aux.columns and 'CENTRO DE CUSTO' in req_aux.columns:
            req_aux['CHAVE'] = req_aux['CONTA'].astype(str)+'_'+req_aux['CENTRO DE CUSTO'].astype(str)
        elif 'CÓD. (CONTA+CENTRO)' in req_aux.columns:
            req_aux['CHAVE'] = req_aux['CÓD. (CONTA+CENTRO)'].astype(str)
        else:
            req_aux['CHAVE'] = None

        bgt_aux = bgt_long.copy()
        if 'CONTA' in bgt_aux.columns and 'CENTRO DE CUSTO' in bgt_aux.columns:
            bgt_aux['CHAVE'] = bgt_aux['CONTA'].astype(str)+'_'+bgt_aux['CENTRO DE CUSTO'].astype(str)
        elif 'CÓD.' in bgt_aux.columns:
            bgt_aux['CHAVE'] = bgt_aux['CÓD.'].astype(str)
        else:
            bgt_aux['CHAVE'] = None

        exec_mes = (
            req_aux
            .assign(MÊS=lambda d: pd.to_datetime(d.get('MÊS COMPETÊNCIA'), errors='coerce').dt.to_period('M').astype(str))
            .groupby(['CHAVE','MÊS'], dropna=False)['VALOR'].sum().reset_index(name='REQ_VALOR')
        )

        comp = (
            bgt_aux.merge(exec_mes, how='left', on=['CHAVE','MÊS'])
                  .groupby('MÊS', as_index=False)[['BGT_VALOR','REQ_VALOR']].sum()
                  .fillna(0)
                  .sort_values('MÊS')
        )

        fig3, ax3 = plt.subplots(figsize=(8,3))
        ax3.plot(comp['MÊS'], comp['BGT_VALOR'], marker='o', label='Orçado (BGT)')
        ax3.plot(comp['MÊS'], comp['REQ_VALOR'], marker='o', label='Requisitado (REQ)')
        ax3.set_title('BGT x REQ por mês (R$)')
        ax3.set_xlabel('Mês')
        ax3.set_ylabel('Valor (R$)')
        ax3.legend()
        plt.xticks(rotation=45)
        st.pyplot(fig3)

        comp['Disponível'] = comp['BGT_VALOR'] - comp['REQ_VALOR']
        st.dataframe(comp, use_container_width=True)

# Estornos Abertos
if estorno_df is not None and not estorno_df.empty:
    st.subheader('Estornos Abertos')
    st.dataframe(estorno_df, use_container_width=True, height=300)

# Download dos dados filtrados
st.markdown('---')

def to_excel_bytes(df_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, df in df_dict.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
    output.seek(0)
    return output

colA, colB = st.columns(2)
with colA:
    if st.button('Baixar Requisições Filtradas (Excel)'):
        bio = to_excel_bytes({'Requisicoes': f})
        st.download_button(
            'Download Requisicoes.xlsx',
            data=bio,
            file_name='Requisicoes_filtrado.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
with colB:
    if budget_df is not None and not budget_df.empty:
        if st.button('Baixar Budget (completo)'):
            bio2 = to_excel_bytes({'Budget': budget_df})
            st.download_button(
                'Download Budget.xlsx',
                data=bio2,
                file_name='Budget.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

st.caption('© 2026 - Painel criado para apoiar controle de requisições, NF e budget de Manutenção.')