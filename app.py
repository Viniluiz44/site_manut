
# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Controle de Manutenção", layout="wide")

@st.cache_data(show_spinner=False)
def load_data():
    xls = pd.read_excel('data/controle.xlsx', sheet_name=None, engine='openpyxl')
    def clean_cols(df):
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]
        return df
    sheets = {k: clean_cols(v) for k, v in xls.items()}
    return sheets

sheets = load_data()

# Infer sheet names by fuzzy matching
name_map = {k.lower().replace(' ', '_'): k for k in sheets.keys()}
get_sheet = lambda key: sheets[name_map.get(key, next((nm for nm in sheets if key in nm.lower().replace(' ','_')), list(sheets)[0]))]

# Try to access expected sheets
req_df = get_sheet('base_requis')
estorno_df = get_sheet('estono') if any('estono' in s.lower() for s in sheets) else None
budget_df = get_sheet('budget') if any('budget' in s.lower() for s in sheets) else None

# --- Padronização de campos comuns ---
if 'VALOR' in req_df.columns:
    req_df['VALOR'] = pd.to_numeric(req_df['VALOR'], errors='coerce').fillna(0)

# Converter datas
for col in ['MÊS COMPETÊNCIA','DATA DE CRIAÇÃO','Data Aprovação','DATA RECEBIMENTO','DATA DO DOC','DATA DE ENTRADA','DATA DE LANÇAMENTO']:
    if col in req_df.columns:
        req_df[col] = pd.to_datetime(req_df[col], errors='coerce')

# Criar campos auxiliares
if 'MÊS COMPETÊNCIA' in req_df.columns:
    req_df['MÊS'] = req_df['MÊS COMPETÊNCIA'].dt.to_period('M').astype(str)
else:
    req_df['MÊS'] = pd.NaT

if 'CD' not in req_df.columns and 'CD ' in req_df.columns:
    req_df['CD'] = req_df['CD ']

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
    st.metric('Valor total (R$)', f['VALOR'].sum())
with c3:
    st.metric('Ticket médio (R$)', (f['VALOR'].sum() / f.shape[0]) if f.shape[0] else 0)
with c4:
    aprovados = f[f.get('STATUS','')=='APROVADO']['VALOR'].sum() if 'STATUS' in f.columns else 0
    st.metric('Aprovado (R$)', aprovados)

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
st.dataframe(f, use_container_width=True, height=400)

# Orç vs execução (BGT vs. REQ)
if budget_df is not None:
    st.subheader('Orçado x Requisitado')
    # Transform budget (mensal) para formato longo
    budget = budget_df.copy()
    budget.columns = [str(c).strip().upper() for c in budget.columns]
    meses = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO']
    present = [m for m in meses if m in budget.columns]
    if present:
        bgt_long = budget.melt(id_vars=[c for c in budget.columns if c not in present], value_vars=present,
                               var_name='MES_NOME', value_name='BGT_VALOR')
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

        exec_mes = (req_aux
                    .assign(MÊS=lambda d: pd.to_datetime(d.get('MÊS COMPETÊNCIA'), errors='coerce').dt.to_period('M').astype(str))
                    .groupby(['CHAVE','MÊS'], dropna=False)['VALOR'].sum().reset_index(name='REQ_VALOR'))

        comp = (bgt_aux.merge(exec_mes, how='left', on=['CHAVE','MÊS'])
                        .groupby('MÊS', as_index=False)[['BGT_VALOR','REQ_VALOR']].sum().fillna(0))
        comp = comp.sort_values('MÊS')

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
        st.download_button('Download Requisicoes.xlsx', data=bio, file_name='Requisicoes_filtrado.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
with colB:
    if budget_df is not None:
        if st.button('Baixar Budget (completo)'):
            bio2 = to_excel_bytes({'Budget': budget_df})
            st.download_button('Download Budget.xlsx', data=bio2, file_name='Budget.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

st.caption('© 2026 - Painel criado para apoiar controle de requisições, NF e budget de Manutenção.')
