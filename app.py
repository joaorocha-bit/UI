import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# Configuração da página
st.set_page_config(
    page_title="Dashboard de Internação",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Otimização de estilos
st.markdown("""
<style>
    [data-testid="stMetricValue"] {
        font-size: 1.8rem !important;
        font-weight: 700 !important;
    }
    div[data-testid="stMetric"] {
        background-color: var(--background-color-secondary);
        padding: 12px 16px;
        border-radius: 10px;
        border: 1px solid rgba(128, 128, 128, 0.2);
    }
</style>
""", unsafe_allow_html=True)


def converter_valor(val):
    """Converte números da planilha mantendo a integridade monetária."""
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace('R$', '').replace(' ', '')
    if '(' in s and ')' in s:
        s = '-' + s.replace('(', '').replace(')', '')
    s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0


@st.cache_data
def carregar_dados_excel(file_path_or_buffer):
    """Carrega e trata os dados das abas com precisão."""
    xls = pd.ExcelFile(file_path_or_buffer)
    setores = ['5B', 'A1', 'B1', 'C1', 'C2', 'D1', 'D2', 'D3', 'E1', 'E3', 'F3']
    
    # 1. Carregar Ticket Médio
    df_tm = pd.read_excel(xls, sheet_name='Ticket Médio')
    
    sector_rows = []
    for idx, row in df_tm.iterrows():
        setor_nome = str(row.iloc[0]).strip()
        if setor_nome in setores:
            sector_rows.append({
                'Setor': setor_nome,
                'Atend_2024': pd.to_numeric(row.iloc[1], errors='coerce'),
                'Atend_2025': pd.to_numeric(row.iloc[2], errors='coerce'),
                'Atend_2026_JS': pd.to_numeric(row.iloc[3], errors='coerce'),
                'Rec_2024': converter_valor(row.iloc[6]),
                'Rec_2025': converter_valor(row.iloc[7]),
                'Rec_2026_JS': converter_valor(row.iloc[8]),
                'TM_2024': pd.to_numeric(row.iloc[11], errors='coerce'),
                'TM_2025': pd.to_numeric(row.iloc[12], errors='coerce'),
                'TM_2026_JS': pd.to_numeric(row.iloc[13], errors='coerce'),
            })
            
    df_ticket_medio = pd.DataFrame(sector_rows)

    # 2. Carregar DRE
    dre_dict = {}
    for s in setores + ['Resumo Anual']:
        sheet_target = s if s in xls.sheet_names else 'Resumo Anual'
        df_s = pd.read_excel(xls, sheet_name=sheet_target)
        
        header_row_idx = None
        for i, row in df_s.iterrows():
            if str(row.iloc[0]).strip().lower() in ['indicador', 'indicadores']:
                header_row_idx = i
                break
        
        if header_row_idx is not None:
            df_clean = df_s.iloc[header_row_idx+1:].copy()
            df_clean = df_clean.dropna(how='all')
            df_clean.columns = ['Indicador', '2024 (Ano)', '2025 (Ano)', '2026 (Jan–Set)', '2025 (Jan–Set)', '2024 (Jan–Set)', 'Var_25_24_Ano', 'Var_26_25_JS', 'Var_25_24_JS']
            df_clean = df_clean.dropna(subset=['Indicador'])
            
            cols_num = ['2024 (Ano)', '2025 (Ano)', '2026 (Jan–Set)', '2025 (Jan–Set)', '2024 (Jan–Set)']
            for col in cols_num:
                df_clean[col] = df_clean[col].apply(converter_valor)
                
            dre_dict[s] = df_clean.reset_index(drop=True)

    return setores, df_ticket_medio, dre_dict


EXCEL_FILE = 'Analise_Internacao_2024-2026 (5).xlsx'

st.sidebar.image("https://img.icons8.com/color/96/hospital-2.png", width=60)
st.sidebar.title("Hospital Dashboard")
st.sidebar.markdown("---")

uploaded_file = st.sidebar.file_uploader("Carregar planilha (opcional)", type=["xlsx"])

if uploaded_file:
    setores, df_tm, dre_dict = carregar_dados_excel(uploaded_file)
elif os.path.exists(EXCEL_FILE):
    setores, df_tm, dre_dict = carregar_dados_excel(EXCEL_FILE)
else:
    st.error("⚠️ Envie a planilha `.xlsx` no menu lateral para visualizar os dados.")
    st.stop()

# Filtros
setor_selecionado = st.sidebar.selectbox("SELECIONE O SETOR", ["Todos os Setores (Geral)"] + setores)
visao_temporal = st.sidebar.radio("PERÍODO DE ANÁLISE DRE", ["Jan–Set (2024 vs 2025 vs 2026)", "Ano Fechado (2024 vs 2025)"])

st.sidebar.markdown("---")

# Título
st.title(f"📊 Dashboard de Internação — {setor_selecionado}")
st.markdown("Acompanhamento consolidado da **Produção (Atendimentos)**, **Ticket Médio** e **DRE Hospitalar**.")

# CÁLCULOS
if setor_selecionado == "Todos os Setores (Geral)":
    df_dre_curr = dre_dict.get('Resumo Anual', pd.DataFrame())
    
    rec_2024 = df_dre_curr.loc[df_dre_curr['Indicador'].str.contains('Receita bruta', case=False, na=False), '2024 (Ano)'].values[0]
    rec_2025 = df_dre_curr.loc[df_dre_curr['Indicador'].str.contains('Receita bruta', case=False, na=False), '2025 (Ano)'].values[0]
    rec_2026_js = df_dre_curr.loc[df_dre_curr['Indicador'].str.contains('Receita bruta', case=False, na=False), '2026 (Jan–Set)'].values[0]
    
    atend_2024 = df_tm['Atend_2024'].sum()
    atend_2025 = df_tm['Atend_2025'].sum()
    atend_2026_js = df_tm['Atend_2026_JS'].sum()
    
    tm_2024 = rec_2024 / atend_2024
    tm_2025 = rec_2025 / atend_2025
    tm_2026_js = rec_2026_js / atend_2026_js
else:
    row_s = df_tm[df_tm['Setor'] == setor_selecionado].iloc[0]
    rec_2024 = row_s['Rec_2024']
    rec_2025 = row_s['Rec_2025']
    rec_2026_js = row_s['Rec_2026_JS']
    
    atend_2024 = row_s['Atend_2024']
    atend_2025 = row_s['Atend_2025']
    atend_2026_js = row_s['Atend_2026_JS']
    
    tm_2024 = row_s['TM_2024']
    tm_2025 = row_s['TM_2025']
    tm_2026_js = row_s['TM_2026_JS']
    
    df_dre_curr = dre_dict.get(setor_selecionado, pd.DataFrame())

# Percentuais
if "Jan–Set" in visao_temporal:
    rec_2025_prop = df_dre_curr.loc[df_dre_curr['Indicador'].str.contains('Receita bruta', case=False, na=False), '2025 (Jan–Set)'].values[0] if not df_dre_curr.empty else rec_2025*0.75
    var_rec = ((rec_2026_js - rec_2025_prop) / rec_2025_prop) * 100 if rec_2025_prop else 0
    var_atend = ((atend_2026_js - atend_2025*0.75) / (atend_2025*0.75)) * 100
    var_tm = ((tm_2026_js - tm_2025) / tm_2025) * 100 if tm_2025 else 0
    
    periodo_label = "2026 (Jan–Set) vs 2025 (Jan–Set)"
    rec_destaque = rec_2026_js
    atend_destaque = atend_2026_js
    tm_destaque = tm_2026_js
else:
    var_rec = ((rec_2025 - rec_2024) / rec_2024) * 100 if rec_2024 else 0
    var_atend = ((atend_2025 - atend_2024) / atend_2024) * 100 if atend_2024 else 0
    var_tm = ((tm_2025 - tm_2024) / tm_2024) * 100 if tm_2024 else 0
    
    periodo_label = "2025 vs 2024 (Ano)"
    rec_destaque = rec_2025
    atend_destaque = atend_2025
    tm_destaque = tm_2025

margem_contrib_val = 0
if not df_dre_curr.empty:
    row_mc = df_dre_curr[df_dre_curr['Indicador'].str.contains('Margem de contribuição', case=False, na=False)]
    if not row_mc.empty:
        col_target = '2026 (Jan–Set)' if "Jan–Set" in visao_temporal else '2025 (Ano)'
        margem_contrib_val = row_mc[col_target].values[0]

# KPIs Nativos
c1, c2, c3, c4 = st.columns(4)

with c1:
    st.metric(
        label=f"Receita Bruta ({'2026 JS' if 'Jan–Set' in visao_temporal else '2025'})",
        value=f"R$ {rec_destaque/1e6:.1f}M",
        delta=f"{var_rec:+.1f}% ({periodo_label})"
    )

with c2:
    st.metric(
        label="Produção (Atendimentos)",
        value=f"{int(atend_destaque):,}".replace(",", "."),
        delta=f"{var_atend:+.1f}% vs Período Ant."
    )

with c3:
    st.metric(
        label="Ticket Médio",
        value=f"R$ {tm_destaque:,.0f}".replace(",", "."),
        delta=f"{var_tm:+.1f}% Variação"
    )

with c4:
    st.metric(
        label="Margem Contribuição",
        value=f"R$ {margem_contrib_val/1e6:.1f}M",
        delta="Resultado DRE",
        delta_color="normal"
    )

st.markdown("<br>", unsafe_allow_html=True)

# GRÁFICOS PRINCIPAIS
g1, g2 = st.columns([1.25, 1])

with g1:
    st.subheader("📈 Evolução Anual: Receita Bruta vs Atendimentos")
    
    df_chart_evo = pd.DataFrame({
        'Período': ['2024 (Ano)', '2025 (Ano)', '2026 (Jan-Set)'],
        'Receita Bruta (R$)': [rec_2024, rec_2025, rec_2026_js],
        'Atendimentos': [atend_2024, atend_2025, atend_2026_js]
    })
    
    fig_combo = go.Figure()
    
    fig_combo.add_trace(go.Bar(
        x=df_chart_evo['Período'],
        y=df_chart_evo['Receita Bruta (R$)'],
        name='Receita Bruta (R$)',
        marker_color='#28a745',
        text=[f"R$ {v/1e6:.1f}M" for v in df_chart_evo['Receita Bruta (R$)']],
        textposition='inside',
        insidetextanchor='middle'
    ))
    
    fig_combo.add_trace(go.Scatter(
        x=df_chart_evo['Período'],
        y=df_chart_evo['Atendimentos'],
        name='Atendimentos',
        yaxis='y2',
        mode='lines+markers+text',
        line=dict(color='#17a2b8', width=3),
        marker=dict(size=8),
        text=[f"{int(v):,}".replace(",", ".") for v in df_chart_evo['Atendimentos']],
        textposition='top center'
    ))
    
    fig_combo.update_layout(
        template='plotly_dark' if st.get_option("theme.base") == "dark" else 'plotly_white',
        height=380,
        margin=dict(l=20, r=20, t=30, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        yaxis=dict(title="Receita Bruta (R$)", showgrid=True, gridcolor="rgba(128,128,128,0.2)"),
        yaxis2=dict(title="Atendimentos", overlaying='y', side='right', showgrid=False)
    )
    st.plotly_chart(fig_combo, use_container_width=True)

with g2:
    st.subheader("🎯 Evolução do Ticket Médio (R$)")
    
    df_chart_tm = pd.DataFrame({
        'Período': ['2024 (Ano)', '2025 (Ano)', '2026 (Jan-Set)'],
        'Ticket Médio': [tm_2024, tm_2025, tm_2026_js]
    })
    
    fig_tm = px.line(
        df_chart_tm,
        x='Período',
        y='Ticket Médio',
        markers=True,
        text=[f"R$ {v:,.0f}".replace(",", ".") for v in df_chart_tm['Ticket Médio']],
        color_discrete_sequence=['#00b4d8']
    )
    fig_tm.update_traces(textposition="top center", line=dict(width=3), marker=dict(size=8))
    fig_tm.update_layout(
        template='plotly_dark' if st.get_option("theme.base") == "dark" else 'plotly_white',
        height=380,
        margin=dict(l=20, r=20, t=30, b=20),
        yaxis=dict(title="Ticket Médio (R$)", showgrid=True, gridcolor="rgba(128,128,128,0.2)")
    )
    st.plotly_chart(fig_tm, use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

# TABELA DRE E COMPARATIVO DE SETORES
sec_col1, sec_col2 = st.columns([1.1, 0.9])

with sec_col1:
    st.subheader(f"📑 Demonstração do Resultado (DRE) — {setor_selecionado}")
    
    if not df_dre_curr.empty:
        df_display = df_dre_curr.copy()
        
        indicadores_filtro = [
            'Receita bruta', 'Receita de Medicamentos', 'Receita de Mat. Méd. Hospitalar',
            'Deduções da receita', 'Receita líquida', 'Custo de Medicamentos',
            'Custo de Mat. Méd. Hospitalar', 'Custos variáveis (total)', 'Margem de contribuição'
        ]
        
        df_display_filtered = df_display[df_display['Indicador'].str.strip().isin(indicadores_filtro)].copy()
        
        # Formatação das colunas para evitar o bug de carregamento
        for col in ['2024 (Ano)', '2025 (Ano)', '2026 (Jan–Set)']:
            if col in df_display_filtered.columns:
                df_display_filtered[col] = df_display_filtered[col].apply(
                    lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                )
        
        st.dataframe(
            df_display_filtered[['Indicador', '2024 (Ano)', '2025 (Ano)', '2026 (Jan–Set)']],
            use_container_width=True,
            hide_index=True,
            height=360
        )
    else:
        st.info("Informações da DRE indisponíveis para a seleção atual.")

with sec_col2:
    st.subheader("🏢 Comparativo do Ticket Médio por Setor")
    
    col_target_tm = 'TM_2026_JS' if "Jan–Set" in visao_temporal else 'TM_2025'
    df_tm_sorted = df_tm.sort_values(by=col_target_tm, ascending=True)
    
    fig_sector = px.bar(
        df_tm_sorted,
        x=col_target_tm,
        y='Setor',
        orientation='h',
        text_auto='.0f',
        title=f"Ticket Médio por Setor ({'2026 Jan-Set' if 'Jan–Set' in visao_temporal else '2025'})",
        color_discrete_sequence=['#0b5cab']
    )
    fig_sector.update_layout(
        template='plotly_dark' if st.get_option("theme.base") == "dark" else 'plotly_white',
        height=360,
        margin=dict(l=10, r=10, t=40, b=20),
        xaxis=dict(title="Ticket Médio (R$)")
    )
    st.plotly_chart(fig_sector, use_container_width=True)

st.markdown("---")
st.caption("Dashboard de Internação | Dados oficiais 2024–2026.")
