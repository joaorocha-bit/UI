import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# Configuração da página
st.set_page_config(
    page_title="Dashboard de Internação | DRE & Performance",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização CSS inspirada no Kondado / Dashboard Executivo Clean
st.markdown("""
<style>
    /* Fundo e Container Principal */
    .main {
        background-color: #f4f7f9;
    }
    
    /* Top Bar / Headers */
    .stAppHeader {
        background-color: #f4f7f9;
    }
    
    /* Sidebar Personalizada */
    [data-testid="stSidebar"] {
        background-color: #0b5cab;
        color: white;
    }
    [data-testid="stSidebar"] * {
        color: white !important;
    }
    [data-testid="stSidebar"] .stSelectbox div {
        color: #333333 !important;
    }

    /* Cards de Indicadores (KPIs) */
    .kpi-card {
        background-color: #ffffff;
        border-radius: 10px;
        padding: 18px 20px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        border: 1px solid #e2e8f0;
        margin-bottom: 15px;
    }
    .kpi-title {
        font-size: 0.85rem;
        color: #64748b;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .kpi-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1e293b;
        margin: 5px 0;
    }
    .kpi-sub {
        font-size: 0.82rem;
        font-weight: 600;
    }
    .kpi-positive {
        color: #10b981;
    }
    .kpi-negative {
        color: #ef4444;
    }

    /* Tabelas */
    div[data-testid="stTable"] table {
        border-radius: 8px;
        overflow: hidden;
    }
</style>
""", unsafe_allow_html=True)


@st.cache_data
def carregar_dados_excel(file_path_or_buffer):
    """Carrega e estrutura as abas da planilha Excel."""
    xls = pd.ExcelFile(file_path_or_buffer)
    
    # Lista de setores identificados nas abas
    setores = ['5B', 'A1', 'B1', 'C1', 'C2', 'D1', 'D2', 'D3', 'E1', 'E3', 'F3']
    
    # 1. Ticket Médio e Produção
    df_tm = pd.read_excel(xls, sheet_name='Ticket Médio')
    
    # Limpeza e extração da tabela de Ticket Médio
    # Localiza a linha dos setores
    sector_rows = []
    for idx, row in df_tm.iterrows():
        setor_nome = str(row.iloc[0]).strip()
        if setor_nome in setores:
            sector_rows.append({
                'Setor': setor_nome,
                'Atend_2024': pd.to_numeric(row.iloc[1], errors='coerce'),
                'Atend_2025': pd.to_numeric(row.iloc[2], errors='coerce'),
                'Atend_2026_JS': pd.to_numeric(row.iloc[3], errors='coerce'),
                'Rec_2024': str(row.iloc[6]).replace('.', '').replace(',', '.'),
                'Rec_2025': str(row.iloc[7]).replace('.', '').replace(',', '.'),
                'Rec_2026_JS': str(row.iloc[8]).replace('.', '').replace(',', '.'),
                'TM_2024': pd.to_numeric(row.iloc[11], errors='coerce'),
                'TM_2025': pd.to_numeric(row.iloc[12], errors='coerce'),
                'TM_2026_JS': pd.to_numeric(row.iloc[13], errors='coerce'),
                'Var_TM_25_24': pd.to_numeric(row.iloc[16], errors='coerce'),
                'Var_TM_26_25': pd.to_numeric(row.iloc[17], errors='coerce')
            })
            
    df_ticket_medio = pd.DataFrame(sector_rows)
    for col in ['Rec_2024', 'Rec_2025', 'Rec_2026_JS']:
        df_ticket_medio[col] = pd.to_numeric(df_ticket_medio[col], errors='coerce')

    # 2. Resumo da DRE por Setor
    dre_dict = {}
    for s in setores + ['Resumo Anual']:
        sheet_name = s if s in xls.sheet_names else 'Resumo Anual'
        df_s = pd.read_excel(xls, sheet_name=s if s in xls.sheet_names else 'Resumo Anual')
        
        # Encontrar linha onde começam os indicadores
        header_row_idx = None
        for i, row in df_s.iterrows():
            if str(row.iloc[0]).strip().lower() in ['indicador', 'indicadores']:
                header_row_idx = i
                break
        
        if header_row_idx is not None:
            df_clean = df_s.iloc[header_row_idx+1:].copy()
            df_clean.columns = ['Indicador', '2024 (Ano)', '2025 (Ano)', '2026 (Jan–Set)', '2025 (Jan–Set)', '2024 (Jan–Set)', 'Var_25_24_Ano', 'Var_26_25_JS', 'Var_25_24_JS']
            df_clean = df_clean.dropna(subset=['Indicador'])
            
            # Limpeza de valores numéricos
            cols_num = ['2024 (Ano)', '2025 (Ano)', '2026 (Jan–Set)', '2025 (Jan–Set)', '2024 (Jan–Set)']
            for col in cols_num:
                df_clean[col] = df_clean[col].astype(str).str.replace('.', '', regex=False).str.replace('(', '-', regex=False).str.replace(')', '', regex=False).str.replace(',', '.', regex=False)
                df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0)
                
            dre_dict[s] = df_clean.reset_index(drop=True)

    # 3. Evolução Mensal Geral
    df_dados = pd.read_excel(xls, sheet_name='Dados')
    df_dados = df_dados.dropna(subset=['Ano', 'Mês']).copy()
    
    return setores, df_ticket_medio, dre_dict, df_dados

# Carregamento do arquivo
EXCEL_FILE = 'Analise_Internacao_2024-2026 (5).xlsx'

st.sidebar.image("https://img.icons8.com/color/96/hospital-2.png", width=60)
st.sidebar.title("Hospital Dashboard")
st.sidebar.markdown("---")

uploaded_file = st.sidebar.file_uploader("Upload da Planilha Excel (opcional)", type=["xlsx"])

if uploaded_file:
    setores, df_tm, dre_dict, df_dados = carregar_dados_excel(uploaded_file)
elif os.path.exists(EXCEL_FILE):
    setores, df_tm, dre_dict, df_dados = carregar_dados_excel(EXCEL_FILE)
else:
    st.error("⚠️ Por favor, envie a planilha `Analise_Internacao_2024-2026 (5).xlsx` no menu lateral.")
    st.stop()

# Filtros na Sidebar
setor_selecionado = st.sidebar.selectbox("SELECIONE O SETOR", ["Todos os Setores (Geral)"] + setores)
visao_temporal = st.sidebar.radio("PERÍODO DE COMPARAÇÃO DRE", ["Jan–Set (2024 vs 2025 vs 2026)", "Ano Fechado (2024 vs 2025)"])

st.sidebar.markdown("---")
st.sidebar.caption("📊 Analítico de Internação (2024–2026)")

# --- TITULO PRINCIPAL ---
st.title(f"📈 Dashboard de Internação — {setor_selecionado}")
st.markdown("Acompanhamento da **Produção (Atendimentos)**, **Ticket Médio** e **DRE Hospitalar**.")

# --- COMPUTAÇÃO DOS KPIS ---
if setor_selecionado == "Todos os Setores (Geral)":
    # Métrica Geral
    rec_2024 = df_tm['Rec_2024'].sum()
    rec_2025 = df_tm['Rec_2025'].sum()
    rec_2026_js = df_tm['Rec_2026_JS'].sum()
    
    atend_2024 = df_tm['Atend_2024'].sum()
    atend_2025 = df_tm['Atend_2025'].sum()
    atend_2026_js = df_tm['Atend_2026_JS'].sum()
    
    tm_2024 = rec_2024 / atend_2024 if atend_2024 else 0
    tm_2025 = rec_2025 / atend_2025 if atend_2025 else 0
    tm_2026_js = rec_2026_js / atend_2026_js if atend_2026_js else 0
    
    df_dre_curr = dre_dict.get('Resumo Anual', pd.DataFrame())
else:
    # Métrica do Setor Escolhido
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

# Cálculo de Variação %
if "Jan–Set" in visao_temporal:
    var_rec = ((rec_2026_js - df_tm['Rec_2025'].sum()*0.75) / (df_tm['Rec_2025'].sum()*0.75))*100 if setor_selecionado == "Todos os Setores (Geral)" else ((rec_2026_js - rec_2025*0.75)/(rec_2025*0.75))*100
    var_atend = ((atend_2026_js - atend_2025*0.75) / (atend_2025*0.75))*100
    var_tm = ((tm_2026_js - tm_2025) / tm_2025)*100 if tm_2025 else 0
    periodo_label = "2026 (Jan-Set) vs 2025 (Proporcional)"
    rec_destaque = rec_2026_js
    atend_destaque = atend_2026_js
    tm_destaque = tm_2026_js
else:
    var_rec = ((rec_2025 - rec_2024) / rec_2024)*100 if rec_2024 else 0
    var_atend = ((atend_2025 - atend_2024) / atend_2024)*100 if atend_2024 else 0
    var_tm = ((tm_2025 - tm_2024) / tm_2024)*100 if tm_2024 else 0
    periodo_label = "2025 vs 2024 (Ano)"
    rec_destaque = rec_2025
    atend_destaque = atend_2025
    tm_destaque = tm_2025

# --- CARDS KPIS (ESTILO KONDADO) ---
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Receita Bruta ({periodo_label.split()[0]})</div>
        <div class="kpi-value">R$ {rec_destaque:,.0f}</div>
        <div class="kpi-sub {'kpi-positive' if var_rec>=0 else 'kpi-negative'}">
            {'▲' if var_rec>=0 else '▼'} {var_rec:.1f}% ({periodo_label})
        </div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Produção (Atendimentos)</div>
        <div class="kpi-value">{int(atend_destaque):,}</div>
        <div class="kpi-sub {'kpi-positive' if var_atend>=0 else 'kpi-negative'}">
            {'▲' if var_atend>=0 else '▼'} {var_atend:.1f}% vs Ano Anterior
        </div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Ticket Médio</div>
        <div class="kpi-value">R$ {tm_destaque:,.0f}</div>
        <div class="kpi-sub {'kpi-positive' if var_tm>=0 else 'kpi-negative'}">
            {'▲' if var_tm>=0 else '▼'} {var_tm:.1f}% de Variação
        </div>
    </div>
    """, unsafe_allow_html=True)

# Extrair Margem de Contribuição da DRE
margem_contrib_val = 0
margem_contrib_pct = 0
if not df_dre_curr.empty:
    row_mc = df_dre_curr[df_dre_curr['Indicador'].str.contains('Margem de contribuição', case=False, na=False)]
    if not row_mc.empty:
        col_target = '2026 (Jan–Set)' if "Jan–Set" in visao_temporal else '2025 (Ano)'
        margem_contrib_val = row_mc[col_target].values[0]

with col4:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Margem de Contribuição</div>
        <div class="kpi-value">R$ {margem_contrib_val:,.0f}</div>
        <div class="kpi-sub kpi-positive">
            🟢 Resultado DRE ({col_target if 'col_target' in locals() else 'Período'})
        </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# --- GRÁFICOS PRINCIPAIS ---
g_col1, g_col2 = st.columns([1.2, 1])

with g_col1:
    st.subheader("📊 Evolução Anual: Receita vs Atendimentos")
    
    # Preparar DataFrame para o Gráfico da Evolução
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
        textposition='auto'
    ))
    fig_combo.add_trace(go.Scatter(
        x=df_chart_evo['Período'],
        y=df_chart_evo['Atendimentos'],
        name='Atendimentos',
        yaxis='y2',
        mode='lines+markers+text',
        line=dict(color='#0b5cab', width=3),
        marker=dict(size=10),
        text=[f"{int(v)}" for v in df_chart_evo['Atendimentos']],
        textposition='top center'
    ))
    
    fig_combo.update_layout(
        template='plotly_white',
        height=380,
        margin=dict(l=20, r=20, t=30, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        yaxis=dict(title="Receita (R$)", showgrid=True, gridcolor="#f0f0f0"),
        yaxis2=dict(title="Atendimentos", overlaying='y', side='right', showgrid=False)
    )
    st.plotly_chart(fig_combo, use_container_width=True)

with g_col2:
    st.subheader("📈 Evolução do Ticket Médio (R$)")
    
    df_chart_tm = pd.DataFrame({
        'Período': ['2024', '2025', '2026 (Jan-Set)'],
        'Ticket Médio': [tm_2024, tm_2025, tm_2026_js]
    })
    
    fig_tm = px.line(
        df_chart_tm,
        x='Período',
        y='Ticket Médio',
        markers=True,
        text=[f"R$ {v:,.0f}" for v in df_chart_tm['Ticket Médio']],
        color_discrete_sequence=['#17a2b8']
    )
    fig_tm.update_traces(textposition="top center", line=dict(width=3), marker=dict(size=9))
    fig_tm.update_layout(
        template='plotly_white',
        height=380,
        margin=dict(l=20, r=20, t=30, b=20),
        yaxis=dict(title="Ticket Médio (R$)", showgrid=True, gridcolor="#f0f0f0")
    )
    st.plotly_chart(fig_tm, use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

# --- SEÇÃO SEGUNDÁRIA: DRE E COMPARATIVO DE SETORES ---
sec_col1, sec_col2 = st.columns([1.1, 0.9])

with sec_col1:
    st.subheader(f"📑 Demonstração do Resultado (DRE) — {setor_selecionado}")
    
    if not df_dre_curr.empty:
        # Formatação de Exibição
        df_display = df_dre_curr.copy()
        
        # Filtrar principais linhas do DRE para tabela limpa
        indicadores_filtro = [
            'Receita bruta', 'Receita de Medicamentos', 'Receita de Mat. Méd. Hospitalar',
            'Deduções da receita', 'Receita líquida', 'Custo de Medicamentos',
            'Custo de Mat. Méd. Hospitalar', 'Custos variáveis (total)', 'Margem de contribuição'
        ]
        
        df_display_filtered = df_display[df_display['Indicador'].str.strip().isin(indicadores_filtro)].copy()
        
        # Formatar colunas monetárias
        for col in ['2024 (Ano)', '2025 (Ano)', '2026 (Jan–Set)']:
            if col in df_display_filtered.columns:
                df_display_filtered[col] = df_display_filtered[col].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        
        st.dataframe(
            df_display_filtered[['Indicador', '2024 (Ano)', '2025 (Ano)', '2026 (Jan–Set)']],
            use_container_width=True,
            hide_index=True,
            height=360
        )
    else:
        st.info("Informações detalhadas de DRE indisponíveis para este filtro.")

with sec_col2:
    st.subheader("🏢 Comparativo do Ticket Médio por Setor")
    
    # Gráfico comparativo entre todos os setores
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
        template='plotly_white',
        height=360,
        margin=dict(l=10, r=10, t=40, b=20),
        xaxis=dict(title="Ticket Médio (R$)")
    )
    st.plotly_chart(fig_sector, use_container_width=True)

# Rodapé
st.markdown("---")
st.caption("Dashboard desenvolvido em Python / Streamlit | Dados da planilha hospitalar 2024-2026.")