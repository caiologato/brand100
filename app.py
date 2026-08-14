import streamlit as st
import pandas as pd
import io

# 1. Configuração da Página e Look and Feel (Cencosud Media)
st.set_page_config(page_title="Simulador DOOH - Cencosud Media", page_icon="🟢", layout="wide")

st.markdown("""
    <style>
    .stApp {
        background-color: #f8f9fa;
    }
    /* Estilização dos botões com o Verde Cencosud Media */
    .stButton>button, .stDownloadButton>button {
        background-color: #13C18E !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
        font-weight: bold !important;
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        background-color: #0e9c72 !important;
    }
    /* Cards de Métricas */
    div[data-testid="metric-container"] {
        background-color: white;
        border: 1px solid #e0e0e0;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 2px 2px 10px rgba(0,0,0,0.05);
    }
    </style>
""", unsafe_allow_html=True)

# Cabeçalho
col1, col2 = st.columns([1, 4])
with col1:
    # Caso queira subir a imagem do logo, coloque o arquivo na mesma pasta e descomente a linha abaixo
    # st.image("Captura de Tela 2026-08-14 às 08.43.38.jpg", use_column_width=True)
    st.markdown("### 🟢 cencosud **media**")
with col2:
    st.title("Simulador de Propostas DOOH - Brand100")

# 2. Carregamento e Tratamento dos Dados
@st.cache_data
def load_data():
    file_path = "Proposta Comercial Interna - 220726 (2).xlsx"
    # Lendo a aba correta e pulando as linhas em branco do cabeçalho
    df = pd.read_excel(file_path, sheet_name='Proposta DOOH - Simulador', header=5)
    
    # Filtrando apenas as lojas ativas
    df = df[df['Status'] == 'INSTALADA'].copy()
    
    # Preparando o dataframe para a interface
    df_ui = pd.DataFrame({
        'Selecionar': False,
        'Bandeira': df['Bandeira'],
        'Loja': df['Nome da Loja'],
        'UF': df['UF'],
        'Valor Diária Base (R$)': df['Valor diária / cota'].astype(float),
        'Diárias': 30,
        'Cotas': 1,
        'Impactos/Dia': (df['Impactos IAB / Impressões OTS'] / df['Período da campanha (em dias)']).astype(float),
        'Alcance/Dia': df['Tráfego por dia'].astype(float)
    })
    return df_ui

try:
    df_ui = load_data()
except Exception as e:
    st.error(f"Erro ao carregar a planilha. Certifique-se de que o arquivo 'Proposta Comercial Interna - 220726 (2).xlsx' está na mesma pasta. Detalhe: {e}")
    st.stop()

# 3. Área Admin (Global)
with st.expander("⚙️ Área Admin - Ajuste Global de Preço da Diária"):
    col_a, col_b = st.columns([1, 3])
    with col_a:
        novo_preco = st.number_input("Novo Preço Padrão (R$):", min_value=0.0, value=349.0, step=10.0)
    with col_b:
        st.write("")
        st.write("")
        if st.button("Aplicar Preço Global"):
            st.session_state['preco_global'] = novo_preco
            st.success("Preço global atualizado nas lojas abaixo!")

if 'preco_global' in st.session_state:
    df_ui['Valor Diária Base (R$)'] = st.session_state['preco_global']

# 4. Configurações da Proposta (Sidebar)
st.sidebar.image("logo.jpg", use_container_width=True)
st.sidebar.header("Configurações do Plano")
desconto = st.sidebar.slider("Desconto Negociado (%)", 0, 100, 0)

# 5. Interface de Seleção da Equipe
st.write("---")
st.write("### Seleção de Lojas")
st.write("Marque as lojas desejadas. Você pode ajustar a quantidade de **Diárias**, **Cotas** e editar o **Valor da Diária** de cada loja diretamente na tabela clicando sobre o número.")

# Tabela editável
edited_df = st.data_editor(
    df_ui,
    column_config={
        "Selecionar": st.column_config.CheckboxColumn("Selecionar", default=False),
        "Valor Diária Base (R$)": st.column_config.NumberColumn("Valor Diária (R$)", format="R$ %.2f"),
        "Impactos/Dia": None, # Ocultamos da UI para deixar limpo, mas usamos no cálculo
        "Alcance/Dia": None   # Ocultamos da UI para deixar limpo, mas usamos no cálculo
    },
    disabled=["Bandeira", "Loja", "UF"],
    hide_index=True,
    use_container_width=True
)

# 6. Resumo e Cálculos Finais
selected_stores = edited_df[edited_df['Selecionar']].copy()

if not selected_stores.empty:
    # Matemática do negócio
    selected_stores['Investimento Bruto'] = selected_stores['Valor Diária Base (R$)'] * selected_stores['Diárias'] * selected_stores['Cotas']
    selected_stores['Investimento Líquido'] = selected_stores['Investimento Bruto'] * (1 - desconto/100)
    selected_stores['Impactos Totais'] = selected_stores['Impactos/Dia'] * selected_stores['Diárias']
    selected_stores['Alcance Total'] = selected_stores['Alcance/Dia'] * selected_stores['Diárias']

    total_bruto = selected_stores['Investimento Bruto'].sum()
    total_liquido = selected_stores['Investimento Líquido'].sum()
    total_impactos = selected_stores['Impactos Totais'].sum()
    total_alcance = selected_stores['Alcance Total'].sum()
    cpm = (total_liquido / total_impactos) * 1000 if total_impactos > 0 else 0

    st.write("---")
    st.write("### 📊 Resumo Executivo da Proposta")
    
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Lojas Selecionadas", len(selected_stores))
    col2.metric("Alcance de Pessoas", f"{total_alcance:,.0f}".replace(',','.'))
    col3.metric("Impactos IAB", f"{total_impactos:,.0f}".replace(',','.'))
    col4.metric("Investimento Final", f"R$ {total_liquido:,.2f}".replace(',','x').replace('.',',').replace('x','.'))
    col5.metric("CPM Médio", f"R$ {cpm:,.2f}".replace(',','x').replace('.',',').replace('x','.'))

    # 7. Download do Excel de Resumo
    output = io.BytesIO()
    
    # Prepara a planilha de saída bonitinha
    df_export = selected_stores.drop(columns=['Selecionar', 'Impactos/Dia', 'Alcance/Dia']).rename(columns={
        'Valor Diária Base (R$)': 'Valor Diária (R$)'
    })
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Resumo_Proposta')
    
    st.write("")
    st.download_button(
        label="📥 Baixar Resumo da Proposta em Excel",
        data=output.getvalue(),
        file_name="Resumo_Proposta_Brand100.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("💡 Selecione pelo menos uma loja na tabela acima para gerar o resumo e o Excel da proposta.")
