import streamlit as st
import pandas as pd
import io
import plotly.express as px
import sqlite3
from datetime import datetime, timedelta

# Banco de dados v4
DB_NAME = 'cencosud_dooh_v4.db'

# 1. Configuração da Página e Cores
st.set_page_config(page_title="Simulador de Campanhas DOOH", page_icon="🟢", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #F8F9FA; color: #333333; }
    .stTabs [data-baseweb="tab-list"] { gap: 20px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; background-color: #ffffff; border-radius: 4px 4px 0px 0px; padding: 10px; font-weight: 500;}
    .stTabs [aria-selected="true"] { background-color: #13C18E !important; color: white !important; font-weight: bold; }
    .stButton>button, .stDownloadButton>button { background-color: #13C18E !important; color: white !important; border-radius: 8px !important; border: none !important; font-weight: bold !important; }
    .stButton>button:hover, .stDownloadButton>button:hover { background-color: #0A7051 !important; }
    div[data-testid="metric-container"] { background-color: white; border: 1px solid #e0e0e0; padding: 15px; border-radius: 10px; box-shadow: 0px 4px 6px rgba(0,0,0,0.05); }
    </style>
""", unsafe_allow_html=True)

# 2. Inicialização do Banco de Dados Local
def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS propostas
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, data_criacao TEXT, comercial TEXT, nome_plano TEXT, 
                  cliente TEXT, contato_nome TEXT, contato_email TEXT, contato_telefone TEXT, 
                  inicio TEXT, fim TEXT, investimento REAL, impactos REAL, status TEXT, data_followup TEXT)''')
    
    c.execute('''CREATE TABLE IF NOT EXISTS precos_lojas (loja TEXT PRIMARY KEY, preco REAL)''')
    c.execute('''CREATE TABLE IF NOT EXISTS comerciais (nome TEXT PRIMARY KEY, email TEXT, telefone TEXT)''')
    
    c.execute("SELECT COUNT(*) FROM comerciais")
    if c.fetchone()[0] == 0:
        default_comerciais = [
            ("Victoria Osti", "victoria.osti@cencosud.com.br", "+55 11 98919-4893"),
            ("Amanda Miguel", "amanda.galvao@cencosud.com.br", "+55 11 99883-2734"),
            ("Vivian Tostes", "vivian.tostes@cencosud.com.br", "+55 21 99922-1919"),
            ("Caio Logato", "caio.logato@cencosud.com.br", "+55 11 94167-9472")
        ]
        c.executemany("INSERT INTO comerciais VALUES (?,?,?)", default_comerciais)
        
    conn.commit()
    conn.close()

init_db()

# 3. Carregamento dos Dados
@st.cache_data
def load_data():
    file_path = "Proposta Comercial Interna - 220726 (2).xlsx"
    df = pd.read_excel(file_path, sheet_name='Proposta DOOH - Simulador', header=5)
    
    def extract_coord(val, idx):
        try: return float(str(val).split(',')[idx].strip())
        except: return None
        
    df['Lat'] = pd.to_numeric(df['Coordenadas'].apply(lambda x: extract_coord(x, 0)), errors='coerce')
    df['Lon'] = pd.to_numeric(df['Coordenadas'].apply(lambda x: extract_coord(x, 1)), errors='coerce')
    
    def clean_status(val):
        if isinstance(val, pd.Timestamp) or isinstance(val, datetime):
            return "EM IMPLANTAÇÃO"
        return str(val).strip()
    df['Status'] = df['Status'].apply(clean_status)
    
    conn = sqlite3.connect(DB_NAME)
    df_precos = pd.read_sql_query("SELECT loja, preco as preco_customizado FROM precos_lojas", conn)
    conn.close()
    
    df_ui = pd.DataFrame({
        'Selecionar': False,
        'Status': df['Status'],
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
    
    df_ui = pd.merge(df_ui, df_precos, how='left', left_on='Loja', right_on='loja')
    df_ui['Valor Diária Base (R$)'] = df_ui['preco_customizado'].combine_first(df_ui['Valor Diária Base (R$)'])
    df_ui.drop(columns=['loja', 'preco_customizado'], inplace=True)
    return df_ui

def get_comerciais():
    conn = sqlite3.connect(DB_NAME)
    df_comerciais = pd.read_sql_query("SELECT * FROM comerciais", conn)
    conn.close()
    return df_comerciais

try:
    df_ui = load_data()
    df_comerciais = get_comerciais()
except Exception as e:
    st.error(f"Erro ao carregar a planilha. Detalhe: {e}")
    st.stop()

# Inicialização de Variáveis de Estado e Chaves Dinâmicas
if 'selecionadas' not in st.session_state:
    st.session_state['selecionadas'] = set()
if 'store_edits' not in st.session_state:
    st.session_state['store_edits'] = {}
if 'plan_key' not in st.session_state:
    st.session_state['plan_key'] = 0
if 'admin_key' not in st.session_state:
    st.session_state['admin_key'] = 0

st.image("logo.jpg", width=250)
st.title("Simulador de Campanhas DOOH")

tab_plan, tab_admin = st.tabs(["📊 Planejamento da Campanha", "⚙️ Área Admin & Comerciais"])

# ================= ÁREA DE PLANEJAMENTO =================
with tab_plan:
    st.write("### 1. Dados da Proposta")
    col_a, col_b, col_c = st.columns([2, 2, 1.5])
    nome_plano = col_a.text_input("Nome do Plano (Ex: Q3 Lançamento)")
    nome_cliente = col_b.text_input("Cliente / Agência")
    
    lista_nomes = df_comerciais['nome'].tolist()
    comercial_selecionado = col_c.selectbox("Comercial (Cencosud)", options=lista_nomes)
    dados_comercial = df_comerciais[df_comerciais['nome'] == comercial_selecionado].iloc[0]
    
    col_cont1, col_cont2, col_cont3 = st.columns(3)
    contato_nome = col_cont1.text_input("Nome do Contato (Agência/Cliente)")
    contato_email = col_cont2.text_input("E-mail do Contato")
    contato_telefone = col_cont3.text_input("Telefone do Contato")
    
    col_d, col_e, col_f = st.columns([1, 1, 2])
    data_inicio = col_d.date_input("Início da Campanha", datetime.today())
    data_fim = col_e.date_input("Fim da Campanha", datetime.today() + timedelta(days=14))
    
    dias_campanha = max((data_fim - data_inicio).days + 1, 1)
    df_ui['Diárias'] = dias_campanha
    desconto = col_f.slider("Desconto Negociado (%)", 0, 100, 0)
    
    st.write("---")
    st.write("### 2. Seleção de Praças e Lojas")
    
    df_ui['Selecionar'] = df_ui['Loja'].isin(st.session_state['selecionadas'])
    for idx, row in df_ui.iterrows():
        loja = row['Loja']
        if loja in st.session_state['store_edits']:
            edits = st.session_state['store_edits'][loja]
            df_ui.at[idx, 'Diárias'] = edits.get('Diárias', dias_campanha)
            df_ui.at[idx, 'Cotas'] = edits.get('Cotas', 1)
            df_ui.at[idx, 'Valor Diária Base (R$)'] = edits.get('Valor Diária Base (R$)', row['Valor Diária Base (R$)'])

    col_filt1, col_filt2 = st.columns([2, 2])
    status_disponiveis = sorted(df_ui['Status'].unique().tolist())
    status_padrao = ['INSTALADA'] if 'INSTALADA' in status_disponiveis else status_disponiveis
    
    status_selecionados = col_filt1.multiselect("Filtrar por Status da Loja", options=status_disponiveis, default=status_padrao)
    busca_loja = col_filt2.text_input("🔍 Buscar por Loja, Cidade ou UF:", "")
    
    mask = df_ui['Status'].isin(status_selecionados)
    if busca_loja:
        mask = mask & (df_ui['Loja'].str.contains(busca_loja, case=False, na=False) | \
               df_ui['Cidade'].str.contains(busca_loja, case=False, na=False) | \
               df_ui['UF'].str.contains(busca_loja, case=False, na=False))
               
    df_display = df_ui[mask].copy()

    col_btn1, col_btn2, _ = st.columns([2, 2, 6])
    
    # CORREÇÃO: Chave dinâmica força a atualização imediata da tabela ao clicar
    if col_btn1.button("✅ Selecionar Lojas Visíveis"):
        for loja in df_display['Loja']: st.session_state['selecionadas'].add(loja)
        st.session_state['plan_key'] += 1 # Força recarregar o widget
        st.rerun()
        
    if col_btn2.button("🟩 Desmarcar Lojas Visíveis"):
        for loja in df_display['Loja']: st.session_state['selecionadas'].discard(loja)
        st.session_state['plan_key'] += 1 # Força recarregar o widget
        st.rerun()

    # Tabela com chave dinâmica associada
    edited_df = st.data_editor(
        df_display,
        key=f"plan_table_{st.session_state['plan_key']}",
        column_config={
            "Selecionar": st.column_config.CheckboxColumn("Selecionar", default=False),
            "Valor Diária Base (R$)": st.column_config.NumberColumn("Valor Diária (R$)", format="R$ %.2f"),
            "Lat": None, "Lon": None, "% Fem": None, "% Masc": None, "Impactos/Dia": None, "Alcance/Dia": None
        },
        disabled=["Status", "Bandeira", "Loja", "Cidade", "UF", "Classe"],
        hide_index=True,
        use_container_width=True
    )

    for _, row in edited_df.iterrows():
        loja = row['Loja']
        if row['Selecionar']: st.session_state['selecionadas'].add(loja)
        else: st.session_state['selecionadas'].discard(loja)
        
        if loja not in st.session_state['store_edits']: st.session_state['store_edits'][loja] = {}
        st.session_state['store_edits'][loja]['Diárias'] = row['Diárias']
        st.session_state['store_edits'][loja]['Cotas'] = row['Cotas']
        st.session_state['store_edits'][loja]['Valor Diária Base (R$)'] = row['Valor Diária Base (R$)']
        
        mask_ui = df_ui['Loja'] == loja
        df_ui.loc[mask_ui, 'Selecionar'] = row['Selecionar']
        df_ui.loc[mask_ui, 'Diárias'] = row['Diárias']
        df_ui.loc[mask_ui, 'Cotas'] = row['Cotas']
        df_ui.loc[mask_ui, 'Valor Diária Base (R$)'] = row['Valor Diária Base (R$)']

    selected_stores = df_ui[df_ui['Selecionar']].copy()

    if not selected_stores.empty:
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
        
        row1_col1, row1_col2 = st.columns(2)
        row2_col1, row2_col2 = st.columns(2)
        
        with row1_col1:
            avg_fem = selected_stores['% Fem'].mean() * 100
            avg_masc = selected_stores['% Masc'].mean() * 100
            df_gender = pd.DataFrame({'Gênero': ['Feminino', 'Masculino'], 'Porcentagem': [avg_fem, avg_masc]})
            fig_gender = px.pie(df_gender, values='Porcentagem', names='Gênero', hole=0.6, color='Gênero', color_discrete_map={'Feminino':'#13C18E', 'Masculino':'#0A7051'})
            fig_gender.update_layout(title_text='Perfil de Gênero', margin=dict(t=40, b=0, l=0, r=0), height=300)
            st.plotly_chart(fig_gender, use_container_width=True)

        with row1_col2:
            df_class = selected_stores.groupby('Classe')['Alcance Total'].sum().reset_index()
            fig_class = px.bar(df_class, x='Classe', y='Alcance Total', text_auto='.2s', color_discrete_sequence=['#13C18E'])
            fig_class.update_layout(title_text='Alcance por Classe Social', margin=dict(t=40, b=0, l=0, r=0), height=300)
            st.plotly_chart(fig_class, use_container_width=True)
            
        with row2_col1:
            df_uf = selected_stores.groupby('UF')['Alcance Total'].sum().reset_index().sort_values('Alcance Total', ascending=True)
            fig_uf = px.bar(df_uf, x='Alcance Total', y='UF', orientation='h', text_auto='.2s', color_discrete_sequence=['#13C18E'])
            fig_uf.update_layout(title_text='Alcance Consolidado por Estado (UF)', margin=dict(t=40, b=0, l=0, r=0), height=350)
            st.plotly_chart(fig_uf, use_container_width=True)

        with row2_col2:
            st.markdown("**Distribuição Geográfica do Impacto**")
            df_map = selected_stores.dropna(subset=['Lat', 'Lon'])
            if not df_map.empty:
                # CORREÇÃO DO MAPA: Agora utilizando explicitamente o 'open-street-map' 
                # que não bloqueia requisições e tamanho fixo para não dar erro com valores zerados
                fig_map = px.scatter_mapbox(
                    df_map, lat="Lat", lon="Lon", hover_name="Loja", 
                    zoom=3, mapbox_style="open-street-map", color_discrete_sequence=['#13C18E']
                )
                fig_map.update_traces(marker=dict(size=10, opacity=0.8))
                fig_map.update_layout(margin=dict(t=0, b=0, l=0, r=0), height=350)
                st.plotly_chart(fig_map, use_container_width=True)
            else:
                st.info("As lojas selecionadas não possuem coordenadas válidas para plotagem.")

        st.write("---")
        
        st.write("#### 💾 Gestão da Proposta")
        save_col1, save_col2, _ = st.columns([1.5, 1.5, 3])
        data_followup_input = save_col1.date_input("📅 Data limite para Follow-up", datetime.today() + timedelta(days=3))
        
        with save_col2:
            st.write("") 
            st.write("")
            if st.button("Salvar no Sistema de Vendas", use_container_width=True):
                if not nome_plano or not nome_cliente:
                    st.warning("Preencha o Nome do Plano e o Cliente no topo da página para salvar!")
                else:
                    conn = sqlite3.connect(DB_NAME)
                    c = conn.cursor()
                    c.execute("""INSERT INTO propostas 
                                 (data_criacao, comercial, nome_plano, cliente, contato_nome, contato_email, contato_telefone, 
                                  inicio, fim, investimento, impactos, status, data_followup) 
                                 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                              (datetime.today().strftime("%Y-%m-%d %H:%M"), dados_comercial['nome'], nome_plano, 
                               nome_cliente, contato_nome, contato_email, contato_telefone, 
                               str(data_inicio), str(data_fim), total_liquido, total_impactos, "Enviada", str(data_followup_input)))
                    conn.commit()
                    conn.close()
                    st.success("✅ Proposta enviada para o funil!")

        output = io.BytesIO()
        workbook = pd.ExcelWriter(output, engine='xlsxwriter')
        df_export = selected_stores.drop(columns=['Selecionar', 'Status', '% Fem', '% Masc', 'Lat', 'Lon', 'Impactos/Dia', 'Alcance/Dia', 'Investimento Bruto']).rename(columns={'Valor Diária Base (R$)': 'Valor Diária (R$)', 'Investimento Líquido': 'Investimento Final (R$)'})
        df_export.to_excel(workbook, index=False, sheet_name='Proposta_DOOH', startrow=12)
        
        wb = workbook.book
        ws = workbook.sheets['Proposta_DOOH']
        
        ws.set_column('A:A', 15)
        ws.set_column('B:B', 30)
        ws.set_column('C:C', 20)
        ws.set_column('D:E', 10)
        ws.set_column('F:H', 12)
        ws.set_column('I:J', 20)
        ws.set_column('K:K', 15)
        
        header_fmt = wb.add_format({'bold': True, 'bg_color': '#13C18E', 'font_color': 'white', 'border': 1})
        title_fmt = wb.add_format({'bold': True, 'font_size': 16, 'font_color': '#13C18E'})
        money_fmt = wb.add_format({'num_format': 'R$ #,##0.00'})
        num_fmt = wb.add_format({'num_format': '#,##0'})
        
        try: ws.insert_image('A1', 'logo.jpg', {'x_scale': 0.15, 'y_scale': 0.15})
        except: pass
        
        data_validade = (datetime.today() + timedelta(days=15)).strftime("%d/%m/%Y")
        
        ws.merge_range('D2:I2', 'RESUMO DA PROPOSTA COMERCIAL - DOOH', title_fmt)
        ws.write('A7', 'Plano:', header_fmt); ws.write('B7', nome_plano)
        ws.write('A8', 'Agência/Anunciante:', header_fmt); ws.write('B8', nome_cliente)
        ws.write('A9', 'Contato (Cliente):', header_fmt); ws.write('B9', f"{contato_nome} | {contato_telefone}")
        ws.write('A10', 'Período:', header_fmt); ws.write('B10', f"{data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}")
        
        ws.write('D7', 'Investimento Total:', header_fmt); ws.write('E7', total_liquido, money_fmt)
        ws.write('D8', 'Impactos Totais:', header_fmt); ws.write('E8', total_impactos, num_fmt)
        ws.write('D9', 'Validade:', header_fmt); ws.write('E9', f"{data_validade} (15 dias)")
        
        ws.write('G7', 'Comercial (Cencosud):', header_fmt); ws.write('H7', dados_comercial['nome'])
        ws.write('G8', 'E-mail:', header_fmt); ws.write('H8', dados_comercial['email'])
        ws.write('G9', 'Telefone:', header_fmt); ws.write('H9', dados_comercial['telefone'])
        
        ws.set_column('I:I', 20, money_fmt)
        ws.set_column('J:K', 15, num_fmt)
        workbook.close()
        
        st.write("")
        st.download_button(
            label="📥 Baixar Proposta Executiva em Excel",
            data=output.getvalue(),
            file_name=f"Proposta_{nome_plano.replace(' ','_')}_{datetime.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ================= ÁREA ADMIN & FOLLOW UP =================
with tab_admin:
    st.header("Área Restrita")
    user = st.text_input("Usuário", key="user_admin")
    senha = st.text_input("Senha", type="password", key="pass_admin")
    
    if user == "cencomedia" and senha == "brand100":
        st.success("Autenticado com sucesso.")
        admin_tab1, admin_tab2, admin_tab3 = st.tabs(["📈 Funil e Follow-up", "👥 Gerenciar Comerciais", "💰 Ajuste de Preços Globais"])
        
        with admin_tab1:
            st.write("### Pipeline de Vendas")
            st.write("Dê dois cliques no **Status** ou na **Data de Follow-up** para atualizá-los e clique em Salvar.")
            
            conn = sqlite3.connect(DB_NAME)
            df_propostas = pd.read_sql_query("SELECT * FROM propostas ORDER BY id DESC", conn)
            conn.close()
            
            if not df_propostas.empty:
                col_busca, col_filtro = st.columns([3, 1])
                busca_hist = col_busca.text_input("🔍 Buscar no histórico (Plano, Cliente, Contato...)", "")
                
                lista_filtro_comerciais = ["Todos"] + df_propostas['comercial'].unique().tolist()
                filtro_comercial = col_filtro.selectbox("Filtrar por Comercial", options=lista_filtro_comerciais)
                
                if busca_hist:
                    mask = df_propostas.astype(str).apply(lambda x: x.str.contains(busca_hist, case=False)).any(axis=1)
                    df_propostas = df_propostas[mask]
                    
                if filtro_comercial != "Todos":
                    df_propostas = df_propostas[df_propostas['comercial'] == filtro_comercial]
                
                df_propostas['data_followup'] = pd.to_datetime(df_propostas['data_followup'], errors='coerce').dt.date
                
                edited_propostas = st.data_editor(
                    df_propostas,
                    column_config={
                        "id": st.column_config.Column("ID", disabled=True),
                        "data_criacao": st.column_config.Column("Criação", disabled=True),
                        "comercial": st.column_config.Column("Comercial (CS)", disabled=True),
                        "nome_plano": st.column_config.Column("Plano", disabled=True),
                        "cliente": st.column_config.Column("Agência/Cliente", disabled=True),
                        "contato_nome": st.column_config.Column("Contato Agência", disabled=True),
                        "contato_email": st.column_config.Column("E-mail Contato", disabled=True),
                        "contato_telefone": st.column_config.Column("Tel. Contato", disabled=True),
                        "inicio": st.column_config.Column("Início", disabled=True),
                        "fim": st.column_config.Column("Fim", disabled=True),
                        "investimento": st.column_config.NumberColumn("Investimento", format="R$ %.2f", disabled=True),
                        "impactos": st.column_config.NumberColumn("Impactos", disabled=True),
                        "status": st.column_config.SelectboxColumn("Status do Pipeline", options=["Enviada", "Sem Retorno", "Em revisão", "Aprovada", "Reprovada", "Cancelada"], required=True),
                        "data_followup": st.column_config.DateColumn("Data Follow-up", format="DD/MM/YYYY")
                    },
                    hide_index=True,
                    use_container_width=True
                )
                
                if st.button("💾 Salvar Alterações no Funil"):
                    conn = sqlite3.connect(DB_NAME)
                    c = conn.cursor()
                    for index, row in edited_propostas.iterrows():
                        c.execute("UPDATE propostas SET status = ?, data_followup = ? WHERE id = ?", (row['status'], str(row['data_followup']), row['id']))
                    conn.commit()
                    conn.close()
                    st.success("Tabela de CRM atualizada com sucesso!")

                st.write("---")
                output_funil = io.BytesIO()
                with pd.ExcelWriter(output_funil, engine='xlsxwriter') as writer:
                    df_export_funil = edited_propostas.copy()
                    df_export_funil.to_excel(writer, index=False, sheet_name='Pipeline_Vendas')
                
                st.download_button(
                    label="📥 Baixar Pipeline Completo em Excel",
                    data=output_funil.getvalue(),
                    file_name=f"Funil_DOOH_Brand100_{datetime.today().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.info("Nenhuma proposta salva no funil de vendas ainda.")

        with admin_tab2:
            st.write("### Equipe Comercial")
            edited_comerciais = st.data_editor(df_comerciais, num_rows="dynamic", use_container_width=True, hide_index=True)
            if st.button("💾 Salvar Contatos"):
                conn = sqlite3.connect(DB_NAME)
                c = conn.cursor()
                c.execute("DELETE FROM comerciais")
                for index, row in edited_comerciais.iterrows():
                    if pd.notna(row['nome']) and str(row['nome']).strip() != "":
                        c.execute("INSERT INTO comerciais (nome, email, telefone) VALUES (?, ?, ?)", (row['nome'], row['email'], row['telefone']))
                conn.commit()
                conn.close()
                st.cache_data.clear()
                st.success("Lista atualizada com sucesso! A página será recarregada.")
                st.rerun()

        with admin_tab3:
            st.write("### Tabela Mestra de Preços")
            st.write("Selecione as lojas que deseja alterar, digite o novo preço e clique em Salvar.")
            
            df_admin_precos = df_ui[['Loja', 'Valor Diária Base (R$)']].copy()
            df_admin_precos.insert(0, 'Selecionar Lojas', False)
            
            col_tabela, col_acao = st.columns([2, 1])
            
            with col_tabela:
                # CORREÇÃO: Chave dinâmica força a atualização imediata da tabela de Admin
                edited_precos = st.data_editor(
                    df_admin_precos, 
                    key=f"admin_table_{st.session_state['admin_key']}",
                    column_config={
                        "Selecionar Lojas": st.column_config.CheckboxColumn(required=True),
                        "Loja": st.column_config.Column(disabled=True),
                        "Valor Diária Base (R$)": st.column_config.NumberColumn(format="R$ %.2f", disabled=True)
                    },
                    use_container_width=True, 
                    hide_index=True
                )
                
            with col_acao:
                st.write("**Ação em Lote**")
                lojas_selecionadas = edited_precos[edited_precos['Selecionar Lojas']]['Loja'].tolist()
                
                st.info(f"{len(lojas_selecionadas)} loja(s) selecionada(s).")
                
                novo_preco_lote = st.number_input("Digite o Novo Preço (R$)", min_value=0.0, value=349.0, step=10.0)
                
                if st.button("💾 Aplicar Preço nas Lojas Selecionadas", use_container_width=True):
                    if len(lojas_selecionadas) > 0:
                        conn = sqlite3.connect(DB_NAME)
                        c = conn.cursor()
                        for loja in lojas_selecionadas:
                            c.execute("INSERT OR REPLACE INTO precos_lojas (loja, preco) VALUES (?, ?)", (loja, novo_preco_lote))
                        conn.commit()
                        conn.close()
                        
                        # Limpa o cache e atualiza a chave da tabela para forçar o recarregamento visual
                        load_data.clear() 
                        st.session_state['admin_key'] += 1 
                        
                        st.success("Preços atualizados com sucesso!")
                        st.rerun()
                    else:
                        st.warning("Selecione pelo menos uma loja na tabela ao lado.")
