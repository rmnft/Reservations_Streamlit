# -------------------------------------------------
#  📊 President's Inn – Streamlit BI Dashboard
# -------------------------------------------------
#  • Carrega dados de reservas a partir de Reservations.xlsx
#  • Se o arquivo não estiver na pasta, permite upload interativo
#  • Exibe métricas-chave, gráficos e Balanced Scorecard
# -------------------------------------------------

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pathlib import Path
import numpy as np
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

# ---------- CONFIGURAÇÃO DA PÁGINA ---------------
st.set_page_config(
    page_title="President's Inn Dashboard",
    page_icon="🏨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ---------- CSS PERSONALIZADO --------------------
st.markdown(
    """
    <style>
        .main-header {
            background: linear-gradient(90deg, #1e3c72 0%, #2a5298 100%);
            padding: 2rem;
            border-radius: 10px;
            color: white;
            text-align: center;
            margin-bottom: 2rem;
        }
        .metric-card {
            background: white;
            border-radius: 15px;
            padding: 1.5rem;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            text-align: center;
            border-left: 4px solid #2a5298;
            transition: transform 0.2s;
        }
        .metric-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(0,0,0,0.15);
        }
        .metric-value {
            font-size: 2rem;
            font-weight: bold;
            color: #1e3c72;
            margin: 0;
        }
        .metric-label {
            color: #666;
            font-size: 0.9rem;
            margin: 0;
        }
        .insight-box {
            background: #f8f9fa;
            border-left: 4px solid #28a745;
            padding: 1rem;
            border-radius: 5px;
            margin: 1rem 0;
        }
        .warning-box {
            background: #fff3cd;
            border-left: 4px solid #ffc107;
            padding: 1rem;
            border-radius: 5px;
            margin: 1rem 0;
        }
        .stTabs [data-baseweb="tab-list"] {
            gap: 2px;
        }
        .stTabs [data-baseweb="tab"] {
            height: 50px;
            padding-left: 20px;
            padding-right: 20px;
            background-color: #f0f2f6;
            border-radius: 10px 10px 0 0;
        }
        .stTabs [aria-selected="true"] {
            background-color: #2a5298;
            color: white;
        }
        footer {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------- FUNÇÃO DE CARREGAMENTO --------------
@st.cache_data(show_spinner=False)
def load_dataframe(f) -> pd.DataFrame:
    """
    Recebe Path ou BytesIO de um Excel e devolve dataframe enriquecido.
    Aceita vários aliases para Arrival/Departure/Daily Rate.
    """
    try:
        df = pd.read_excel(f)
    except Exception as e:
        st.error(f"❌ Erro ao ler arquivo Excel: {str(e)}")
        st.stop()

    if df.empty:
        st.error("❌ Arquivo Excel está vazio!")
        st.stop()

    # Normaliza cabeçalhos: remove espaços e usa Title Case
    df.columns = df.columns.str.strip().str.title()

    # Mapeia aliases → nome canônico
    aliases = {
        "Arrival Date":   ["Arrival Date", "Arrival", "Check-In", "Data De Chegada", "Checkin"],
        "Departure Date": ["Departure Date", "Departure", "Check-Out", "Data De Saída", "Checkout"],
        "Daily Rate":     ["Daily Rate", "Adr", "Tarifa Diária", "Rate", "Price"],
        "Room Type":      ["Room Type", "Type", "Tipo De Quarto", "Category"],
        "No Of Guests":   ["No Of Guests", "Guests", "Hóspedes", "Pax"],
        "Room":           ["Room", "Room Number", "Quarto", "Número Do Quarto"]
    }

    def first_present(possible):
        for c in possible:
            if c in df.columns:
                return c
        return None

    # Mapeia colunas essenciais
    column_mapping = {}
    missing_columns = []
    
    for canonical_name, possible_names in aliases.items():
        found_column = first_present(possible_names)
        if found_column:
            column_mapping[found_column] = canonical_name
        else:
            missing_columns.append(canonical_name)

    if missing_columns:
        st.error(f"❌ Colunas ausentes: {', '.join(missing_columns)}. Colunas disponíveis: {', '.join(df.columns.tolist())}")
        st.info("💡 Certifique-se de que o Excel contém pelo menos: Arrival Date, Departure Date, Daily Rate, Room Type, No Of Guests, Room")
        st.stop()

    # Renomeia colunas
    df.rename(columns=column_mapping, inplace=True)

    # Conversões e validações
    try:
        df["Arrival Date"] = pd.to_datetime(df["Arrival Date"])
        df["Departure Date"] = pd.to_datetime(df["Departure Date"])
    except Exception as e:
        st.error(f"❌ Erro ao converter datas: {str(e)}")
        st.stop()

    # Remove linhas com datas inválidas
    df = df.dropna(subset=["Arrival Date", "Departure Date"])
    
    # Valida se departure é depois de arrival
    invalid_dates = df["Departure Date"] <= df["Arrival Date"]
    if invalid_dates.any():
        st.warning(f"⚠️ {invalid_dates.sum()} reservas com datas inválidas foram removidas")
        df = df[~invalid_dates]

    # Calcula métricas derivadas
    df["Nights"] = (df["Departure Date"] - df["Arrival Date"]).dt.days
    df["Revenue"] = df["Daily Rate"] * df["Nights"]
    df["Month"] = df["Arrival Date"].dt.to_period('M')
    df["Year"] = df["Arrival Date"].dt.year
    df["Weekday"] = df["Arrival Date"].dt.day_name()

    # Remove outliers extremos
    q1_rate = df["Daily Rate"].quantile(0.01)
    q99_rate = df["Daily Rate"].quantile(0.99)
    df = df[(df["Daily Rate"] >= q1_rate) & (df["Daily Rate"] <= q99_rate)]

    return df

# ---------- BUSCA AUTOMÁTICA DO ARQUIVO ---------
def load_data():
    DATA_PATH = Path(__file__).parent / "Reservations.xlsx" if "__file__" in globals() else Path("Reservations.xlsx")
    
    if DATA_PATH.exists():
        return load_dataframe(DATA_PATH)
    else:
        st.markdown(
            """
            <div class="warning-box">
                <h4>🔍 Arquivo não encontrado</h4>
                <p>O arquivo <code>Reservations.xlsx</code> não foi encontrado. Por favor, faça upload do arquivo abaixo.</p>
            </div>
            """, 
            unsafe_allow_html=True
        )
        
        uploaded = st.file_uploader(
            "📂 Envie o arquivo Reservations.xlsx", 
            type="xlsx",
            help="Arquivo deve conter colunas: Arrival Date, Departure Date, Daily Rate, Room Type, No Of Guests, Room"
        )
        
        if uploaded is not None:
            return load_dataframe(uploaded)
        else:
            st.info("👆 Aguardando upload do arquivo...")
            st.stop()

# ---------- CARREGA DADOS ------------------------
with st.spinner("📊 Carregando dados..."):
    df = load_data()

if df is None or df.empty:
    st.error("❌ Não foi possível carregar os dados")
    st.stop()

# ---------- HEADER PRINCIPAL ---------------------
st.markdown(
    """
    <div class="main-header">
        <h1>🏨 President's Inn</h1>
        <h3>Dashboard de Desempenho Operacional</h3>
        <p>Análise inteligente de reservas • Atualizado em tempo real</p>
    </div>
    """,
    unsafe_allow_html=True
)

# ---------- FILTROS LATERAIS --------------------
with st.sidebar:
    st.header("🎛️ Filtros")
    
    # Filtro de período
    min_date = df["Arrival Date"].min().date()
    max_date = df["Arrival Date"].max().date()
    
    date_range = st.date_input(
        "📅 Período de análise",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
        key="date_filter"
    )
    
    # Filtro de tipo de quarto
    room_types = ["Todos"] + sorted(df["Room Type"].unique().tolist())
    selected_room_types = st.multiselect(
        "🏠 Tipos de Quarto",
        options=room_types,
        default=["Todos"]
    )
    
    # Filtro de número de hóspedes
    guest_counts = ["Todos"] + sorted(df["No Of Guests"].unique().tolist())
    selected_guests = st.multiselect(
        "👥 Número de Hóspedes",
        options=guest_counts,
        default=["Todos"]
    )

# Aplica filtros
df_filtered = df.copy()

if len(date_range) == 2:
    start_date, end_date = date_range
    df_filtered = df_filtered[
        (df_filtered["Arrival Date"].dt.date >= start_date) & 
        (df_filtered["Arrival Date"].dt.date <= end_date)
    ]

if "Todos" not in selected_room_types and selected_room_types:
    df_filtered = df_filtered[df_filtered["Room Type"].isin(selected_room_types)]

if "Todos" not in selected_guests and selected_guests:
    df_filtered = df_filtered[df_filtered["No Of Guests"].isin(selected_guests)]

# ---------- CÁLCULO DE KPIS ----------------------
total_reservations = len(df_filtered)
total_revenue = df_filtered["Revenue"].sum()
avg_daily_rate = df_filtered["Daily Rate"].mean()
avg_length_stay = df_filtered["Nights"].mean()
occupancy_rate = len(df_filtered) / len(df) * 100 if len(df) > 0 else 0
revpar = total_revenue / df_filtered["Room"].nunique() if df_filtered["Room"].nunique() > 0 else 0

# ---------- DASHBOARD PRINCIPAL ------------------
# KPIs principais
col1, col2, col3, col4, col5, col6 = st.columns(6)

kpis = [
    ("💰", "Receita Total", f"R$ {total_revenue:,.0f}"),
    ("🏨", "Total de Reservas", f"{total_reservations:,}"),
    ("📊", "ADR Médio", f"R$ {avg_daily_rate:.0f}"),
    ("📅", "Estadia Média", f"{avg_length_stay:.1f} noites"),
    ("📈", "Taxa de Ocupação", f"{occupancy_rate:.1f}%"),
    ("💎", "RevPAR", f"R$ {revpar:.0f}")
]

for col, (icon, label, value) in zip([col1, col2, col3, col4, col5, col6], kpis):
    col.markdown(
        f"""
        <div class="metric-card">
            <div style="font-size: 2rem; margin-bottom: 0.5rem;">{icon}</div>
            <p class="metric-value">{value}</p>
            <p class="metric-label">{label}</p>
        </div>
        """,
        unsafe_allow_html=True
    )

st.markdown("---")

# ---------- TABS DE ANÁLISE ----------------------
tab1, tab2, tab3, tab4 = st.tabs(["📊 Visão Geral", "💰 Análise Financeira", "🏠 Quartos & Hóspedes", "📈 Tendências"])

with tab1:
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Receita por Tipo de Quarto")
        revenue_by_room = df_filtered.groupby("Room Type")["Revenue"].sum().reset_index()
        fig1 = px.bar(
            revenue_by_room,
            x="Room Type",
            y="Revenue",
            title="",
            color="Revenue",
            color_continuous_scale="Blues"
        )
        fig1.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig1, use_container_width=True)
    
    with col2:
        st.subheader("Distribuição por Número de Hóspedes")
        guest_distribution = df_filtered["No Of Guests"].value_counts().reset_index()
        guest_distribution.columns = ["No Of Guests", "Count"]
        fig2 = px.pie(
            guest_distribution,
            values="Count",
            names="No Of Guests",
            title=""
        )
        fig2.update_layout(height=400)
        st.plotly_chart(fig2, use_container_width=True)

with tab2:
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Receita Mensal")
        monthly_revenue = df_filtered.groupby("Month")["Revenue"].sum().reset_index()
        monthly_revenue["Month_str"] = monthly_revenue["Month"].astype(str)
        fig3 = px.line(
            monthly_revenue,
            x="Month_str",
            y="Revenue",
            title="",
            markers=True
        )
        fig3.update_layout(height=400)
        st.plotly_chart(fig3, use_container_width=True)
    
    with col2:
        st.subheader("ADR por Tipo de Quarto")
        adr_by_room = df_filtered.groupby("Room Type")["Daily Rate"].mean().reset_index()
        fig4 = px.bar(
            adr_by_room,
            x="Room Type",
            y="Daily Rate",
            title="",
            color="Daily Rate",
            color_continuous_scale="Greens"
        )
        fig4.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig4, use_container_width=True)

with tab3:
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Ocupação por Dia da Semana")
        weekday_counts = df_filtered["Weekday"].value_counts().reset_index()
        weekday_counts.columns = ["Weekday", "Count"]
        # Ordenar dias da semana
        weekday_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        weekday_counts["Weekday"] = pd.Categorical(weekday_counts["Weekday"], categories=weekday_order, ordered=True)
        weekday_counts = weekday_counts.sort_values("Weekday")
        
        fig5 = px.bar(
            weekday_counts,
            x="Weekday",
            y="Count",
            title=""
        )
        fig5.update_layout(height=400)
        st.plotly_chart(fig5, use_container_width=True)
    
    with col2:
        st.subheader("Duração da Estadia")
        nights_distribution = df_filtered["Nights"].value_counts().head(10).reset_index()
        nights_distribution.columns = ["Nights", "Count"]
        fig6 = px.bar(
            nights_distribution,
            x="Nights",
            y="Count",
            title=""
        )
        fig6.update_layout(height=400)
        st.plotly_chart(fig6, use_container_width=True)

with tab4:
    st.subheader("Análise de Tendências Temporal")
    
    # Métricas ao longo do tempo
    monthly_metrics = df_filtered.groupby("Month").agg({
        "Revenue": "sum",
        "Daily Rate": "mean",
        "Nights": "mean",
        "Room": "nunique"
    }).reset_index()
    monthly_metrics["Month_str"] = monthly_metrics["Month"].astype(str)
    
    fig7 = make_subplots(
        rows=2, cols=2,
        subplot_titles=("Receita Mensal", "ADR Médio", "Estadia Média", "Quartos Únicos"),
        specs=[[{"secondary_y": False}, {"secondary_y": False}],
               [{"secondary_y": False}, {"secondary_y": False}]]
    )
    
    fig7.add_trace(
        go.Scatter(x=monthly_metrics["Month_str"], y=monthly_metrics["Revenue"], name="Receita"),
        row=1, col=1
    )
    fig7.add_trace(
        go.Scatter(x=monthly_metrics["Month_str"], y=monthly_metrics["Daily Rate"], name="ADR"),
        row=1, col=2
    )
    fig7.add_trace(
        go.Scatter(x=monthly_metrics["Month_str"], y=monthly_metrics["Nights"], name="Estadia"),
        row=2, col=1
    )
    fig7.add_trace(
        go.Scatter(x=monthly_metrics["Month_str"], y=monthly_metrics["Room"], name="Quartos"),
        row=2, col=2
    )
    
    fig7.update_layout(height=600, showlegend=False)
    st.plotly_chart(fig7, use_container_width=True)

# ---------- INSIGHTS AUTOMÁTICOS ----------------
st.markdown("---")
st.header("🔍 Insights Automáticos")

col1, col2 = st.columns(2)

with col1:
    # Top performers
    top_room_type = df_filtered.groupby("Room Type")["Revenue"].sum().idxmax()
    top_room_revenue = df_filtered.groupby("Room Type")["Revenue"].sum().max()
    
    st.markdown(
        f"""
        <div class="insight-box">
            <h4>🏆 Melhor Performance</h4>
            <p><strong>{top_room_type}</strong> é o tipo de quarto mais lucrativo, gerando R$ {top_room_revenue:,.0f} em receita.</p>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # Oportunidades
    lowest_adr_room = df_filtered.groupby("Room Type")["Daily Rate"].mean().idxmin()
    lowest_adr = df_filtered.groupby("Room Type")["Daily Rate"].mean().min()
    
    st.markdown(
        f"""
        <div class="warning-box">
            <h4>💡 Oportunidade</h4>
            <p>Quartos <strong>{lowest_adr_room}</strong> têm ADR mais baixo (R$ {lowest_adr:.0f}). Considere estratégias de upselling.</p>
        </div>
        """,
        unsafe_allow_html=True
    )

with col2:
    # Padrões sazonais
    best_month = monthly_revenue.loc[monthly_revenue["Revenue"].idxmax(), "Month"] if not monthly_revenue.empty else "N/A"
    
    st.markdown(
        f"""
        <div class="insight-box">
            <h4>📅 Sazonalidade</h4>
            <p>Melhor mês: <strong>{best_month}</strong>. Use estes dados para otimizar pricing e marketing.</p>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # Guest behavior
    avg_guests = df_filtered["No Of Guests"].mean()
    st.markdown(
        f"""
        <div class="insight-box">
            <h4>👥 Perfil do Cliente</h4>
            <p>Média de <strong>{avg_guests:.1f} hóspedes</strong> por reserva. Foque em experiências para este segmento.</p>
        </div>
        """,
        unsafe_allow_html=True
    )

# ---------- BALANCED SCORECARD -------------------
st.markdown("---")
st.header("🎯 Balanced Scorecard")

scorecard_data = {
    "Perspectiva": [
        "💰 Financeira",
        "🎯 Clientes",
        "⚙️ Processos Internos",
        "📚 Aprendizado & Crescimento"
    ],
    "Objetivos Estratégicos": [
        f"Aumentar RevPAR para R$ {revpar * 1.1:.0f} (+10%)",
        f"Manter ADR médio acima de R$ {avg_daily_rate * 0.95:.0f}",
        "Reduzir tempo de check-in/out em 15%",
        "Implementar analytics preditivos"
    ],
    "Indicadores-Chave": [
        f"RevPAR atual: R$ {revpar:.0f}",
        f"ADR atual: R$ {avg_daily_rate:.0f}",
        f"Estadia média: {avg_length_stay:.1f} noites",
        "Dashboard implementado ✅"
    ],
    "Status": [
        "🟡 Em andamento",
        "🟢 Atingido",
        "🔴 Atenção necessária",
        "🟢 Concluído"
    ]
}

scorecard_df = pd.DataFrame(scorecard_data)
st.dataframe(scorecard_df, use_container_width=True, hide_index=True)

# ---------- FOOTER -------------------------------
st.markdown("---")
st.markdown(
    f"""
    <div style="text-align: center; color: #666; padding: 2rem;">
        <p>📊 Dashboard atualizado automaticamente • {len(df_filtered)} reservas analisadas</p>
        <p>Período: {df_filtered['Arrival Date'].min().strftime('%d/%m/%Y')} - {df_filtered['Arrival Date'].max().strftime('%d/%m/%Y')}</p>
    </div>
    """,
    unsafe_allow_html=True
)

# ---------- BOTÃO DE REFRESH --------------------
if st.button("🔄 Atualizar Dashboard", type="primary"):
    st.cache_data.clear()
    st.rerun()