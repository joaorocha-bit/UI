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

st.markdown("""
<style>
    [data-testid="stMetricValue"] { font-size: 1.8rem !important; font-weight: 700 !important; }
    div[data-testid="stMetric"] {
        background-color: var(--background-color-secondary);
        padding: 12px 16px; border-radius: 10px; border: 1px solid rgba(128, 128, 128, 0.2);
    }
</style>
""", unsafe_allow_html=True)

def clean_money(val):
    """Limpeza segura de valores financeiros da planilha"""
    if pd.isna(val) or val == '' or val == '-':
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    
    s = str(val).strip().replace('R$', '').replace(' ', '')
    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1]
        
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    else:
        if s.count('.') > 0:
            parts = s.split('.')
            if len(parts[-1]) == 3 and len(parts) > 1:
                s = s.replace('.', '')
    try:
        return float(s)
    except Exception:
        return 0.0

@st.cache_data
def carregar_dados_excel(file_path_or_buffer):
    xls = pd.ExcelFile(file_path_or_buffer)
    setores = ['5B', 'A1', 'B1', 'C1', 'C2', 'D1', 'D2', 'D3', 'E1', 'E3', 'F3']
    
    # 1. Carregar Ticket Médio (Mapeando TODAS as colunas Ano e Jan-Set)
    df_tm = pd.read_excel(xls, sheet_name='Ticket Médio')
    sector_rows = []
    
    for idx, row in df_tm.iterrows():
        setor_nome = str(row.iloc[0]).strip()
        if setor_nome in setores:
            sector_rows.append({
                'Setor': setor_nome,
                # Atendimentos
                'Atend_2024_Ano': pd.to_numeric(row.iloc[1], errors='coerce'),
                'Atend_2025_Ano': pd.to_numeric(row.iloc[2], errors='coerce'),
                'Atend_2026_JS': pd.to_numeric(row.iloc[3], errors='coerce'),
                'Atend_2025_JS': pd.to_numeric(row.iloc[4], errors='coerce'),
                'Atend_2024_JS': pd.to_numeric(row.iloc[5], errors='coerce'),
                # Receitas
                'Rec_2024_Ano': clean_money(row.iloc[6]),
                'Rec_2025_Ano': clean_money(row.iloc[7]),
                'Rec_2026_JS': clean_money(row.iloc[8]),
                'Rec_2025_JS': clean_money(row.iloc[9]),
                'Rec_2024_JS': clean_money(row.iloc[10]),
                # Ticket Médio
                'TM_2024_Ano': pd.to_numeric(row.iloc[11], errors='coerce'),
                'TM_2025_Ano': pd.to_numeric(row.iloc[12], errors='coerce'),
                'TM_2026_JS': pd.to_numeric(row.iloc[13], errors='coerce'),
                'TM_2025_JS': pd.to_numeric(row.iloc[14], errors='coerce'),
                'TM_2024_JS': pd.to_numeric(row.iloc[15], errors='coerce'),
            })
            
    df_ticket_medio = pd.DataFrame(sector_rows)

    # 2. Carregar DRE
    dre_dict = {}
    for s in setores + ['Resumo Anual']:
        sheet_target = s if s in xls.sheet_names else 'Resumo Anual'
        df_s = pd.read_excel(xls, sheet_name=sheet_target)
        
        header_idx = None
        for i, row in df_s.iterrows():
            if str(row.iloc[0]).strip().lower() in ['indicador', 'indicadores']:
                header_idx = i
                break
        
        if header_idx is not None:
            df_clean = df_s.iloc[header_idx+1:].copy()
            df_clean = df_clean.dropna(how='all').dropna(subset=[df_clean.columns[0]])
            
            cols = ['Indicador', '2024 (Ano)', '2025 (Ano)', '2026 (Jan–Set)', '2025 (Jan–Set)', '2024 (Jan–Set)', 'Var_25_24_Ano', 'Var_26_25_JS', 'Var_25_24_JS']
            df_clean.columns = cols[:len(df_clean.columns)]
            
            cols_num = ['2024 (Ano)', '2025 (Ano)', '2026 (Jan–Set)', '2025 (Jan–Set)', '2024 (Jan–Set)']
            for col in cols_num:
                if col in df_clean.columns:
                    df_clean[col] = df_clean[col].apply(clean_money)
                
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
visao_temporal = st.sidebar.radio("PERÍODO DE ANÁLISE", ["Ano Fechado (2024 vs 2025)", "Jan–Set (2024 vs 2025 vs 2026)"])

st.sidebar.markdown("---")
st.title(f"📊 Dashboard de Internação — {setor_selecionado}")
st.markdown("Acompanhamento consolidado da **Produção (Atendimentos)**, **Ticket Médio** e **DRE Hospitalar**.")

# ---------------------------------------------------------
# LÓGICA DE DADOS SEPARADA POR VISÃO TEMPORAL ESTrita
# ---------------------------------------------------------

df_dre_curr = dre_dict.get('Resumo Anual' if setor_selecionado == "Todos os Setores (Geral)" else setor_selecionado, pd.DataFrame())

if setor_selecionado == "Todos os Setores (Geral)":
    # Agrupamentos Gerais
    if "Ano Fechado" in visao_temporal:
        rec_atual = df_dre_curr.loc[df_dre_curr['Indicador'].str.contains('Receita bruta', case=False, na=False), '2025 (Ano)'].values[0]
        rec_ant = df_dre_curr.loc[df_dre_curr['Indicador'].str.contains('Receita bruta', case=False, na=False), '2024 (Ano)'].values[0]
        
        atend_atual = df_tm['Atend_2025_Ano'].sum()
        atend_ant = df_tm['Atend_2024_Ano'].sum()
    else:
        rec_atual = df_dre_curr.loc[df_dre_curr['Indicador'].str.contains('Receita bruta', case=False, na=False), '2026 (Jan–Set)'].values[0]
        rec_ant = df_dre_curr.loc[df_dre_curr['Indicador'].str.contains('Receita bruta', case=False, na=False), '2025 (Jan–Set)'].values[0]
        rec_ant2 = df_dre_curr.loc[df_dre_curr['Indicador'].str.contains('Receita bruta', case=False, na=False), '2024 (Jan–Set)'].values[0]
        
        atend_atual = df_tm['Atend_2026_JS'].sum()
        atend_ant = df_tm['Atend_2025_JS'].sum()
        atend_ant2 = df_tm['Atend_2024_JS'].sum()

    tm_atual = rec_atual / atend_atual if atend_atual else 0
    tm_ant = rec_ant / atend_ant if atend_ant else 0
    if "Jan–Set" in visao_temporal:
        tm_ant2 = rec_ant2 / atend_ant2 if atend_ant2 else 0

else:
    # Setor Específico
    row_s = df_tm[df_tm['Setor'] == setor_selecionado].iloc[0]
    
    if "Ano Fechado" in visao_temporal:
        rec_atual, rec_ant = row_s['Rec_2025_Ano'], row_s['Rec_2024_Ano']
        atend_atual, atend_ant = row_s['Atend_2025_Ano'], row_s['Atend_2024_Ano']
        tm_atual, tm_ant = row_s['TM_2025_Ano'], row_s['TM_2024_Ano']
    else:
        rec_atual, rec_ant, rec_ant2 = row_s['Rec_2026_JS'], row_s['Rec_2025_JS'], row_s['Rec_2024_JS']
        atend_atual, atend_ant, atend_ant2 = row_s['Atend_2026_JS'], row_s['Atend_2025_JS'], row_s['Atend_2024_JS']
        tm_atual, tm_ant, tm_ant2 = row_s['TM_2026_JS'], row_s['TM_2025_JS'], row_s['TM_2024_JS']

# Variações e Labels
var_rec = ((rec_atual - rec_ant) / rec_ant) * 100 if rec_ant else 0
var_atend = ((atend_atual - atend_ant) / atend_ant) * 100 if atend_ant else 0
var_tm = ((tm_atual - tm_ant) / tm_ant) * 100 if tm_ant else 0

periodo_label = "2025 vs 2024" if "Ano Fechado" in visao_temporal else "2026 vs 2025 (Jan-Set)"
ano_label = "2025 (Ano)" if "Ano Fechado" in visao_temporal else "2026 (Jan-Set)"

# Margem DRE
margem_contrib_val = 0
if not df_dre_curr.empty:
    row_mc = df_dre_curr[df_dre_curr['Indicador'].str.contains('Margem de contribuição', case=False, na=False)]
    if not row_mc.empty:
        col_target = '2025 (Ano)' if "Ano Fechado" in visao_temporal else '2026 (Jan–Set)'
        if col_target in row_mc.columns:
            margem_contrib_val = row_mc[col_target].values[0]

# KPIs Dinâmicos
c1, c2, c3, c4 = st.columns(4)
with c1: st.metric(label=f"Receita Bruta ({ano_label})", value=f"R$ {rec_atual/1e6:.1f}M", delta=f"{var_rec:+.1f}% ({periodo_label})")
with c2: st.metric(label=f"Produção ({ano_label})", value=f"{int(atend_atual):,}".replace(",", "."), delta=f"{var_atend:+.1f}% vs Período Ant.")
with c3: st.metric(label=f"Ticket Médio ({ano_label})", value=f"R$ {tm_atual:,.0f}".replace(",", "."), delta=f"{var_tm:+.1f}% Variação")
with c4: st.metric(label="Margem Contribuição", value=f"R$ {margem_contrib_val/1e6:.1f}M", delta="Resultado DRE", delta_color="normal")

st.markdown("<br>", unsafe_allow_html=True)

# ---------------------------------------------------------
# GRÁFICOS DINÂMICOS CONFORME A VISÃO TEMPORAL
# ---------------------------------------------------------
g1, g2 = st.columns([1.25, 1])

if "Ano Fechado" in visao_temporal:
    periodos_grafico = ['2024 (Ano)', '2025 (Ano)']
    rec_grafico = [rec_ant, rec_atual]
    atend_grafico = [atend_ant, atend_atual]
    tm_grafico = [tm_ant, tm_atual]
else:
    periodos_grafico = ['2024 (Jan-Set)', '2025 (Jan-Set)', '2026 (Jan-Set)']
    rec_grafico = [rec_ant2, rec_ant, rec_atual]
    atend_grafico = [atend_ant2, atend_ant, atend_atual]
    tm_grafico = [tm_ant2, tm_ant, tm_atual]

with g1:
    st.subheader(f"📈 Evolução: Receita vs Atendimentos")
    df_chart_evo = pd.DataFrame({'Período': periodos_grafico, 'Receita Bruta (R$)': rec_grafico, 'Atendimentos': atend_grafico})
    
    fig_combo = go.Figure()
    fig_combo.add_trace(go.Bar(
        x=df_chart_evo['Período'], y=df_chart_evo['Receita Bruta (R$)'], name='Receita Bruta (R$)',
        marker_color='#28a745', text=[f"R$ {v/1e6:.1f}M" for v in df_chart_evo['Receita Bruta (R$)']],
        textposition='inside', insidetextanchor='middle'
    ))
    fig_combo.add_trace(go.Scatter(
        x=df_chart_evo['Período'], y=df_chart_evo['Atendimentos'], name='Atendimentos', yaxis='y2',
        mode='lines+markers+text', line=dict(color='#17a2b8', width=3), marker=dict(size=8),
        text=[f"{int(v):,}".replace(",", ".") for v in df_chart_evo['Atendimentos']], textposition='top center'
    ))
    fig_combo.update_layout(
        template='plotly_dark' if st.get_option("theme.base") == "dark" else 'plotly_white',
        height=380, margin=dict(l=20, r=20, t=30, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        yaxis=dict(title="Receita Bruta (R$)", showgrid=True, gridcolor="rgba(128,128,128,0.2)"),
        yaxis2=dict(title="Atendimentos", overlaying='y', side='right', showgrid=False)
    )
    st.plotly_chart(fig_combo, use_container_width=True)

with g2:
    st.subheader("🎯 Evolução do Ticket Médio (R$)")
    df_chart_tm = pd.DataFrame({'Período': periodos_grafico, 'Ticket Médio': tm_grafico})
    fig_tm = px.line(
        df_chart_tm, x='Período', y='Ticket Médio', markers=True,
        text=[f"R$ {v:,.0f}".replace(",", ".") for v in df_chart_tm['Ticket Médio']], color_discrete_sequence=['#00b4d8']
    )
    fig_tm.update_traces(textposition="top center", line=dict(width=3), marker=dict(size=8))
    fig_tm.update_layout(
        template='plotly_dark' if st.get_option("theme.base") == "dark" else 'plotly_white',
        height=380, margin=dict(l=20, r=20, t=30, b=20),
        yaxis=dict(title="Ticket Médio (R$)", showgrid=True, gridcolor="rgba(128,128,128,0.2)")
    )
    st.plotly_chart(fig_tm, use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

# ---------------------------------------------------------
# TABELA DRE E RANKING CONFORME A VISÃO TEMPORAL
# ---------------------------------------------------------
sec_col1, sec_col2 = st.columns([1.1, 0.9])

with sec_col1:
    st.subheader(f"📑 DRE — {setor_selecionado}")
    if not df_dre_curr.empty:
        indicadores_filtro = [
            'Receita bruta', 'Receita de Medicamentos', 'Receita de Mat. Méd. Hospitalar',
            'Deduções da receita', 'Receita líquida', 'Custo de Medicamentos',
            'Custo de Mat. Méd. Hospitalar', 'Custos variáveis (total)', 'Margem de contribuição'
        ]
        df_display_filtered = df_dre_curr[df_dre_curr['Indicador'].str.strip().isin(indicadores_filtro)].copy()
        
        # Só exibe as colunas relevantes baseadas na seleção atual
        if "Ano Fechado" in visao_temporal:
            col_view = ['2024 (Ano)', '2025 (Ano)']
        else:
            col_view = ['2024 (Jan–Set)', '2025 (Jan–Set)', '2026 (Jan–Set)']
            
        for col in col_view:
            if col in df_display_filtered.columns:
                df_display_filtered[col] = df_display_filtered[col].apply(
                    lambda x: f"R$ {float(x):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if pd.notnull(x) else "-"
                )
        
        cols_final = ['Indicador'] + [c for c in col_view if c in df_display_filtered.columns]
        st.dataframe(df_display_filtered[cols_final], use_container_width=True, hide_index=True, height=360)
    else:
        st.info("Informações da DRE indisponíveis para a seleção atual.")

with sec_col2:
    st.subheader("🏢 Ranking de Ticket Médio por Setor")
    
    col_target_tm = 'TM_2025_Ano' if "Ano Fechado" in visao_temporal else 'TM_2026_JS'
    df_tm_sorted = df_tm.sort_values(by=col_target_tm, ascending=True)
    
    fig_sector = px.bar(
        df_tm_sorted, x=col_target_tm, y='Setor', orientation='h', text_auto='.0f',
        title=f"Ticket Médio por Setor ({ano_label})",
        color_discrete_sequence=['#0b5cab']
    )
    fig_sector.update_layout(
        template='plotly_dark' if st.get_option("theme.base") == "dark" else 'plotly_white',
        height=360, margin=dict(l=10, r=10, t=40, b=20), xaxis=dict(title="Ticket Médio (R$)")
    )
    st.plotly_chart(fig_sector, use_container_width=True)

st.markdown("---")
st.caption("Dashboard de Internação | Dados oficiais 2024–2026 consolidados.")
