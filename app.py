import streamlit as st
import pandas as pd
import io
import plotly.express as px
import sqlite3
from datetime import datetime, timedelta

# 1. Configuração da Página e Cores
st.set_page_config(page_title="Simulador DOOH - Cencosud Media", page_icon="🟢", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #F8F9FA; color: #333333; }
    .stTabs [data-baseweb="tab-list"] { gap: 20px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; background-color: #ffffff; border-radius: 4px 4px 0px 0px; padding: 10px; }
    .stTabs [aria-selected="true"] { background-color: #13C18E !important; color: white !important; font-weight: bold; }
    .stButton>button, .stDownloadButton>button { background-color: #13C18E !important; color: white !important; border-radius: 8px !important; border: none !important; font-weight: bold !important; }
    .stButton>button:hover, .stDownloadButton>button:hover { background-color: #0A7051 !important; }
    div[data-testid="metric-container"] { background-color: white; border: 1px solid #e0e0e0; padding: 15px; border-radius: 10px; box-shadow: 0px 4px 6px rgba(0,0,0,0.05); }
    </style>
""", unsafe_allow_html=True)

# 2. Inicialização do Banco de Dados Local (SQLite)
def init_db():
    conn = sqlite3.connect('cencosud_dooh.db')
    c = conn.cursor()
    # Tabela de propostas salvas para Follow-up
    c.execute('''CREATE TABLE IF NOT EXISTS propostas
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, data_criacao TEXT, nome_plano TEXT, cliente TEXT, 
                  email TEXT, telefone TEXT, inicio TEXT, fim TEXT, lojas_ativas INTEGER, investimento REAL, impactos REAL, cpm REAL)''')
    # Tabela para salvar preços customizados individualmente
    c.execute('''CREATE TABLE IF NOT EXISTS precos_lojas (loja TEXT PRIMARY KEY, preco REAL)''')
    conn.commit()
    conn.close()

init_db()

# 3. Carregamento dos Dados
@st.cache_data(ttl=60) # Recarrega a cada 60s para atualizar preços do BD
def load_data():
    file_path = "Proposta Comercial Interna - 220726 (2).xlsx"
    df = pd.read_excel(file_path, sheet_name='Proposta DOOH - Simulador', header=5)
    df = df[df['Status'] == 'INSTALADA'].copy()
    
    # Extração de Coordenadas para o Mapa
    coords = df['Coordenadas'].astype(str).str.split(',', expand=True)
    df['Lat'] = pd.to_numeric(coords[0], errors='coerce')
    df['Lon'] = pd.to_numeric(coords[1], errors='coerce')
    
    # Busca preços customizados no Banco de Dados
    conn = sqlite3.connect('cencosud_dooh.db')
    df_precos = pd.read_sql_query("SELECT loja, preco as preco_customizado FROM precos_lojas", conn)
    conn.close()
    
    # Mesclando preços e montando UI
    df_ui = pd.DataFrame({
        'Selecionar': False,
        'Bandeira': df['Bandeira'],
        'Loja': df['Nome da Loja'],
        'Cidade': df['Cidade / Municipio'],
        'UF': df['UF'],
        'Classe': df['Público'],
        'Lat': df['Lat'],
        'Lon': df['Lon'],
        '% Fem': pd.to_numeric(df['% Público Feminino'], errors='coerce'),
        '% Masc': pd.to_numeric(df['% Público Masculino'], errors='coerce'),
        'Valor Diária Base (R$)': pd.to_numeric(df['Valor diária / cota'], errors='coerce').fillna(349.0),
        'Diárias': 30,
        'Cotas': 1,
        'Impactos/Dia': (pd.to_numeric(df['Impactos IAB / Impressões OTS'], errors='coerce') / pd.to_numeric(df['Período da campanha (em dias)'], errors='coerce')).fillna(0),
        'Alcance/Dia': pd.to_numeric(df['Tráfego por dia'], errors='coerce').fillna(0)
    })
    
    # Aplica o preço customizado se existir no BD
    df_ui = pd.merge(df_ui, df_precos, how='left', left_on='Loja', right_on='loja')
    df_ui['Valor Diária Base (R$)'] = df_ui['preco_customizado'].combine_first(df_ui['Valor Diária Base (R$)'])
    df_ui.drop(columns=['loja', 'preco_customizado'], inplace=True)
    
    return df_ui

try:
    df_ui = load_data()
except Exception as e:
    st.error(f"Erro ao carregar a planilha. Detalhe: {e}")
    st.stop()

# Cabeçalho
st.markdown("### 🟢 cencosud **media**")
st.title("Simulador de Propostas DOOH - Brand100")

# Abas
tab_plan, tab_admin = st.tabs(["📊 Planejamento da Campanha", "⚙️ Área Admin & Comerciais"])

# ================= ÁREA DE PLANEJAMENTO =================
with tab_plan:
    st.write("### 1. Dados da Proposta")
    
    # Linha 1: Metadados
    col_a, col_b, col_c, col_d = st.columns(4)
    nome_plano = col_a.text_input("Nome do Plano (Ex: Q3 Lançamento)")
    nome_cliente = col_b.text_input("Nome do Cliente / Agência")
    email_cliente = col_c.text_input("E-mail Contato")
    tel_cliente = col_d.text_input("Telefone Contato")
    
    # Linha 2: Datas (Ajuste automático de diárias)
    col_e, col_f, col_g = st.columns([1, 1, 2])
    data_inicio = col_e.date_input("Início da Campanha", datetime.today())
    data_fim = col_f.date_input("Fim da Campanha", datetime.today() + timedelta(days=14))
    
    # Cálculo dinâmico das diárias pelo período
    dias_campanha = max((data_fim - data_inicio).days + 1, 1)
    df_ui['Diárias'] = dias_campanha
    
    desconto = col_g.slider("Desconto Negociado (%)", 0, 100, 0)
    
    st.write("---")
    st.write("### 2. Seleção de Praças e Lojas")
    st.info(f"💡 Dica: O período selecionado definiu as diárias padrão para **{dias_campanha} dias**. Você ainda pode ajustar individualmente na tabela.")
    
    # Tabela Editável
    edited_df = st.data_editor(
        df_ui,
        column_config={
            "Selecionar": st.column_config.CheckboxColumn("Selecionar", default=False),
            "Valor Diária Base (R$)": st.column_config.NumberColumn("Valor Diária (R$)", format="R$ %.2f"),
            "Lat": None, "Lon": None, "% Fem": None, "% Masc": None, "Impactos/Dia": None, "Alcance/Dia": None
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
        st.write("### 3. Resumo da Audiência e Resultados")
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Lojas Ativas", len(selected_stores))
        col2.metric("Alcance Total", f"{total_alcance:,.0f}".replace(',','.'))
        col3.metric("Impactos (IAB)", f"{total_impactos:,.0f}".replace(',','.'))
        col4.metric("Investimento Líquido", f"R$ {total_liquido:,.2f}".replace(',','x').replace('.',',').replace('x','.'))
        col5.metric("CPM Médio", f"R$ {cpm:,.2f}".replace(',','x').replace('.',',').replace('x','.'))

        st.write("#### 👁️ Perfil da Audiência (Dashboard)")
        
        dash_col1, dash_col2 = st.columns([1, 2])
        with dash_col1:
            df_uf = selected_stores.groupby('UF')['Alcance Total'].sum().reset_index().sort_values('Alcance Total', ascending=True)
            fig_uf = px.bar(df_uf, x='Alcance Total', y='UF', orientation='h', text_auto='.2s', color_discrete_sequence=['#13C18E'])
            fig_uf.update_layout(title_text='Alcance Consolidado por Estado (UF)', margin=dict(t=40, b=0, l=0, r=0), height=350)
            st.plotly_chart(fig_uf, use_container_width=True)

        with dash_col2:
            # Novo: Mapa Interativo de Alcance
            fig_map = px.scatter_mapbox(
                selected_stores.dropna(subset=['Lat', 'Lon']),
                lat="Lat", lon="Lon", size="Alcance Total", color="UF", hover_name="Loja",
                zoom=3, mapbox_style="carto-positron", color_discrete_sequence=['#13C18E', '#0A7051', '#32CD32']
            )
            fig_map.update_layout(title_text='Distribuição Geográfica do Impacto', margin=dict(t=40, b=0, l=0, r=0), height=350)
            st.plotly_chart(fig_map, use_container_width=True)

        # Ações Finais: Salvar BD e Exportar Excel Customizado
        st.write("---")
        action_col1, action_col2 = st.columns([1, 1])
        
        # Rotina de criação do Excel com XlsxWriter
        output = io.BytesIO()
        workbook = pd.ExcelWriter(output, engine='xlsxwriter')
        df_export = selected_stores.drop(columns=['Selecionar', '% Fem', '% Masc', 'Lat', 'Lon', 'Impactos/Dia', 'Alcance/Dia']).rename(columns={'Valor Diária Base (R$)': 'Valor Diária (R$)'})
        df_export.to_excel(workbook, index=False, sheet_name='Proposta_DOOH', startrow=10)
        
        wb = workbook.book
        ws = workbook.sheets['Proposta_DOOH']
        
        # Estilos Customizados
        header_fmt = wb.add_format({'bold': True, 'bg_color': '#13C18E', 'font_color': 'white', 'border': 1})
        title_fmt = wb.add_format({'bold': True, 'font_size': 14})
        
        try: ws.insert_image('A1', 'logo.jpg', {'x_scale': 0.4, 'y_scale': 0.4})
        except: pass
        
        data_validade = (datetime.today() + timedelta(days=15)).strftime("%d/%m/%Y")
        ws.write('D2', 'RESUMO DA PROPOSTA COMERCIAL - DOOH', title_fmt)
        ws.write('A6', 'Nome do Plano:', header_fmt); ws.write('B6', nome_plano)
        ws.write('A7', 'Cliente/Agência:', header_fmt); ws.write('B7', nome_cliente)
        ws.write('A8', 'Período:', header_fmt); ws.write('B8', f"{data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}")
        ws.write('D6', 'Investimento Total:', header_fmt); ws.write('E6', f"R$ {total_liquido:,.2f}")
        ws.write('D7', 'Impactos Totais:', header_fmt); ws.write('E7', f"{total_impactos:,.0f}")
        ws.write('D8', 'Validade Proposta:', header_fmt); ws.write('E8', f"{data_validade} (15 dias)")
        
        workbook.close()

        with action_col1:
            if st.button("💾 Salvar Proposta no Sistema (Follow-up)"):
                if not nome_plano or not nome_cliente:
                    st.warning("Preencha o Nome do Plano e o Cliente para salvar!")
                else:
                    conn = sqlite3.connect('cencosud_dooh.db')
                    c = conn.cursor()
                    c.execute("INSERT INTO propostas (data_criacao, nome_plano, cliente, email, telefone, inicio, fim, lojas_ativas, investimento, impactos, cpm) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                              (datetime.today().strftime("%Y-%m-%d %H:%M"), nome_plano, nome_cliente, email_cliente, tel_cliente, str(data_inicio), str(data_fim), len(selected_stores), total_liquido, total_impactos, cpm))
                    conn.commit()
                    conn.close()
                    st.success("✅ Proposta salva com sucesso! O comercial já pode visualizar na Área Admin.")

        with action_col2:
            st.download_button(
                label="📥 Baixar Proposta Executiva em Excel",
                data=output.getvalue(),
                file_name=f"Proposta_{nome_plano.replace(' ','_')}_{datetime.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ================= ÁREA ADMIN & FOLLOW UP =================
with tab_admin:
    st.header("Área Restrita: Configurações e Follow-up")
    
    # Bloqueio por Senha
    user = st.text_input("Usuário", key="user_admin")
    senha = st.text_input("Senha", type="password", key="pass_admin")
    
    if user == "cencomedia" and senha == "brand100":
        st.success("Autenticado com sucesso.")
        admin_tab1, admin_tab2 = st.tabs(["💰 Ajuste de Preços (Lojas)", "📈 Follow-up de Propostas"])
        
        with admin_tab1:
            st.write("### Tabela Mestra de Preços")
            st.write("Altere o preço da diária base de qualquer loja aqui. A alteração ficará salva para as próximas vezes.")
            
            # Tabela admin para alterar o preço do banco
            df_admin_precos = df_ui[['Loja', 'Valor Diária Base (R$)']].copy()
            edited_precos = st.data_editor(df_admin_precos, use_container_width=True, hide_index=True)
            
            # Botão para aplicar preço global em lote para simplificar a vida
            novo_global = st.number_input("Atribuir Preço Global para TODAS as lojas", min_value=0.0, value=349.0)
            if st.button("Forçar Preço Global"):
                edited_precos['Valor Diária Base (R$)'] = novo_global
                st.rerun() # Atualiza a tela com os novos preços
                
            if st.button("💾 Salvar Alterações de Preço"):
                conn = sqlite3.connect('cencosud_dooh.db')
                c = conn.cursor()
                for index, row in edited_precos.iterrows():
                    c.execute("INSERT OR REPLACE INTO precos_lojas (loja, preco) VALUES (?, ?)", (row['Loja'], row['Valor Diária Base (R$)']))
                conn.commit()
                conn.close()
                st.cache_data.clear() # Força o app a reler os dados novos
                st.success("Preços atualizados com sucesso no Banco de Dados!")

        with admin_tab2:
            st.write("### Histórico Comercial")
            st.write("Acompanhe as simulações geradas pela equipe.")
            conn = sqlite3.connect('cencosud_dooh.db')
            df_propostas = pd.read_sql_query("SELECT * FROM propostas ORDER BY id DESC", conn)
            conn.close()
            
            if not df_propostas.empty:
                st.dataframe(df_propostas, use_container_width=True, hide_index=True)
                # Exportar o CSV do pipeline de vendas
                csv_pipeline = df_propostas.to_csv(index=False).encode('utf-8')
                st.download_button("📥 Exportar Histórico do Funil", csv_pipeline, "pipeline_vendas_brand100.csv", "text/csv")
            else:
                st.info("Nenhuma proposta salva no sistema ainda.")
    else:
        if user or senha:
            st.error("Usuário ou senha incorretos.")