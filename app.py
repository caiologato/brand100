import streamlit as st
import pandas as pd
import io
import plotly.express as px

# 1. Configuração da Página e Look and Feel (Cencosud Media)
st.set_page_config(page_title="Simulador DOOH - Cencosud Media", page_icon="🟢", layout="wide")

# Forçando cores claras para o fundo e textos escuros para resolver o bug de leitura
st.markdown("""
    <style>
    .stApp {
        background-color: #F8F9FA;
        color: #333333;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 20px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #ffffff;
        border-radius: 4px 4px 0px 0px;
        gap: 1px;
        padding-top: 10px;
        padding-bottom: 10px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #13C18E !important;
        color: white !important;
        font-weight: bold;
    }
    .stButton>button, .stDownloadButton>button {
        background-color: #13C18E !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
        font-weight: bold !important;
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        background-color: #0A7051 !important;
    }
    /* Estilizando as métricas */
    div[data-testid="metric-container"] {
        background-color: white;
        border: 1px solid #e0e0e0;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0px 4px 6px rgba(0,0,0,0.05);
    }
    </style>
""", unsafe_allow_html=True)

# Cabeçalho
st.markdown("### 🟢 cencosud **media**")
st.title("Simulador de Propostas DOOH - Brand100")

# 2. Carregamento e Tratamento dos Dados
@st.cache_data
def load_data():
    file_path = "Proposta Comercial Interna - 220726 (2).xlsx"
    df = pd.read_excel(file_path, sheet_name='Proposta DOOH - Simulador', header=5)
    df = df[df['Status'] == 'INSTALADA'].copy()
    
    # Trazendo mais dados para o Dashboard e FORÇANDO que sejam lidos como números
    df_ui = pd.DataFrame({
        'Selecionar': False,
        'Bandeira': df['Bandeira'],
        'Loja': df['Nome da Loja'],
        'Cidade': df['Cidade / Municipio'],
        'UF': df['UF'],
        'Classe': df['Público'],
        # O pd.to_numeric garante que o sistema converta texto em número para não quebrar os gráficos
        '% Fem': pd.to_numeric(df['% Público Feminino'], errors='coerce'),
        '% Masc': pd.to_numeric(df['% Público Masculino'], errors='coerce'),
        'Valor Diária Base (R$)': pd.to_numeric(df['Valor diária / cota'], errors='coerce').fillna(0),
        'Diárias': 30,
        'Cotas': 1,
        'Impactos/Dia': (pd.to_numeric(df['Impactos IAB / Impressões OTS'], errors='coerce') / pd.to_numeric(df['Período da campanha (em dias)'], errors='coerce')).fillna(0),
        'Alcance/Dia': pd.to_numeric(df['Tráfego por dia'], errors='coerce').fillna(0)
    })
    return df_ui
    
try:
    df_ui = load_data()
except Exception as e:
    st.error(f"Erro ao carregar a planilha. Detalhe: {e}")
    st.stop()

# 3. Sidebar
st.sidebar.image("logo.jpg", use_container_width=True)
st.sidebar.header("Configurações Globais")
desconto = st.sidebar.slider("Desconto Negociado (%)", 0, 100, 0)

# Criando as Abas (Tabs) para separar as áreas
tab_plan, tab_admin = st.tabs(["📊 Planejamento do Plano", "⚙️ Área Admin"])

# ================= ÁREA ADMIN =================
with tab_admin:
    st.header("Configurações Administrativas")
    st.write("Ajuste o valor base cobrado pela diária em todas as lojas.")
    col_a, col_b = st.columns([1, 3])
    with col_a:
        novo_preco = st.number_input("Novo Preço Padrão (R$):", min_value=0.0, value=349.0, step=10.0)
    with col_b:
        st.write("")
        st.write("")
        if st.button("Aplicar Preço Global"):
            st.session_state['preco_global'] = novo_preco
            st.success("Preço global atualizado com sucesso!")

# Atualiza o preço se foi alterado no admin
if 'preco_global' in st.session_state:
    df_ui['Valor Diária Base (R$)'] = st.session_state['preco_global']

# ================= ÁREA DE PLANEJAMENTO =================
with tab_plan:
    st.write("### 1. Seleção de Praças e Lojas")
    st.write("Selecione as lojas e ajuste a quantidade de diárias e cotas diretamente na tabela.")
    
    # Tabela Editável
    edited_df = st.data_editor(
        df_ui,
        column_config={
            "Selecionar": st.column_config.CheckboxColumn("Selecionar", default=False),
            "Valor Diária Base (R$)": st.column_config.NumberColumn("Valor Diária (R$)", format="R$ %.2f"),
            "% Fem": None, "% Masc": None, "Impactos/Dia": None, "Alcance/Dia": None # Ocultando dados sensíveis da tabela
        },
        disabled=["Bandeira", "Loja", "Cidade", "UF", "Classe"],
        hide_index=True,
        use_container_width=True
    )

    selected_stores = edited_df[edited_df['Selecionar']].copy()

    if not selected_stores.empty:
        # Cálculos de Negócio
        selected_stores['Investimento Bruto'] = selected_stores['Valor Diária Base (R$)'] * selected_stores['Diárias'] * selected_stores['Cotas']
        selected_stores['Investimento Líquido'] = selected_stores['Investimento Bruto'] * (1 - desconto/100)
        selected_stores['Impactos Totais'] = selected_stores['Impactos/Dia'] * selected_stores['Diárias']
        selected_stores['Alcance Total'] = selected_stores['Alcance/Dia'] * selected_stores['Diárias']

        total_liquido = selected_stores['Investimento Líquido'].sum()
        total_impactos = selected_stores['Impactos Totais'].sum()
        total_alcance = selected_stores['Alcance Total'].sum()
        cpm = (total_liquido / total_impactos) * 1000 if total_impactos > 0 else 0

        st.write("---")
        st.write("### 2. Resumo da Audiência e Resultados")
        
        # KPIs Principais
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Lojas Ativas", len(selected_stores))
        col2.metric("Alcance Total", f"{total_alcance:,.0f}".replace(',','.'))
        col3.metric("Impactos (IAB)", f"{total_impactos:,.0f}".replace(',','.'))
        col4.metric("Investimento Líquido", f"R$ {total_liquido:,.2f}".replace(',','x').replace('.',',').replace('x','.'))
        col5.metric("CPM Médio", f"R$ {cpm:,.2f}".replace(',','x').replace('.',',').replace('x','.'))

        st.write("#### 👁️ Perfil da Audiência (Dashboard)")
        
        # Dashboard Visual
        dash_col1, dash_col2, dash_col3 = st.columns(3)
        
        with dash_col1:
            # Gráfico de Gênero
            avg_fem = selected_stores['% Fem'].mean() * 100
            avg_masc = selected_stores['% Masc'].mean() * 100
            df_gender = pd.DataFrame({'Gênero': ['Feminino', 'Masculino'], 'Porcentagem': [avg_fem, avg_masc]})
            fig_gender = px.pie(df_gender, values='Porcentagem', names='Gênero', hole=0.6, 
                                color='Gênero', color_discrete_map={'Feminino':'#13C18E', 'Masculino':'#0A7051'})
            fig_gender.update_layout(title_text='Perfil de Gênero', margin=dict(t=40, b=0, l=0, r=0), height=300)
            st.plotly_chart(fig_gender, use_container_width=True)

        with dash_col2:
            # Gráfico de Classe Social
            df_class = selected_stores.groupby('Classe')['Alcance Total'].sum().reset_index()
            fig_class = px.bar(df_class, x='Classe', y='Alcance Total', text_auto='.2s', 
                               color_discrete_sequence=['#13C18E'])
            fig_class.update_layout(title_text='Alcance por Classe Social', margin=dict(t=40, b=0, l=0, r=0), height=300)
            st.plotly_chart(fig_class, use_container_width=True)

        with dash_col3:
            # Gráfico de Distribuição por UF
            df_uf = selected_stores.groupby('UF')['Alcance Total'].sum().reset_index().sort_values('Alcance Total', ascending=True)
            fig_uf = px.bar(df_uf, x='Alcance Total', y='UF', orientation='h', text_auto='.2s', 
                            color_discrete_sequence=['#0A7051'])
            fig_uf.update_layout(title_text='Alcance por Estado (UF)', margin=dict(t=40, b=0, l=0, r=0), height=300)
            st.plotly_chart(fig_uf, use_container_width=True)

        # 3. Download do Excel
        st.write("---")
        output = io.BytesIO()
        df_export = selected_stores.drop(columns=['Selecionar', '% Fem', '% Masc', 'Impactos/Dia', 'Alcance/Dia']).rename(columns={'Valor Diária Base (R$)': 'Valor Diária (R$)'})
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_export.to_excel(writer, index=False, sheet_name='Resumo_Proposta')
        
        st.download_button(
            label="📥 Baixar Plano Completo em Excel",
            data=output.getvalue(),
            file_name="Plano_DOOH_Cencosud.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("💡 Navegue pela tabela acima e selecione pelo menos uma loja para liberar o dashboard de audiência e o download do plano.")
