import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import hashlib

# Configuração da página
st.set_page_config(
    page_title="Análise Financeira - 2025",
    page_icon="📊",
    layout="wide"
)

# Função para criar hash da senha
def make_hash(password):
    return hashlib.sha256(str.encode(password)).hexdigest()

# Função para verificar o login
def check_login(username, password):
    if username == "cintia.ferreira" and make_hash(password) == make_hash("Cf2025"):
        return True
    return False

# Função para a página principal
def main_page():
    # Lista de todos os indicadores financeiros
    INDICADORES = [
        'ATIVO',
        'PASSIVO',
        'PATRIMONIO LIQUIDO',
        'ESTOQUES',
        'COMPENSACOES',
        'RECEITA OPERACIONAL BRUTA',
        'DEDUCOES/CUSTOS/DESPESAS',
        'APURACAO DO RESULTADO',
        'CONTAS DEVEDORAS',
        'CONTAS CREDORAS',
        'RESULTADO DO MES',
        'RESULTADO DO EXERCÍCIO'
    ]

    # Função para carregar e preparar os dados
    @st.cache_data
    def load_data():
        try:
            # Lê a planilha consolidada
            df = pd.read_excel('Consolidadas_1tri_2025.xlsx')
            
            # Renomeia as colunas para o formato esperado
            df = df.rename(columns={
                'emopresa': 'empresa',
                datetime(2025, 1, 31): '01/2025',
                datetime(2025, 2, 28): '02/2025',
                datetime(2025, 3, 31): '03/2025'
            })
            
            # Converte valores monetários para float
            for col in ['01/2025', '02/2025', '03/2025', 'Saldo acumulado']:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace('$', '').str.replace(',', ''), errors='coerce')
            
            return df
        except Exception as e:
            st.error(f"Erro ao carregar os dados: {str(e)}")
            return None

    # Carregar os dados
    df = load_data()

    if df is not None:
        # Título principal
        st.title("📊 Dashboard Financeiro - 1º Trimestre 2025")

        # Sidebar para filtros
        st.sidebar.header("Filtros")
        empresas_selecionadas = st.sidebar.multiselect(
            "Selecione as Empresas",
            options=df['empresa'].unique(),
            default=df['empresa'].unique()
        )

        indicadores_selecionados = st.sidebar.multiselect(
            "Selecione os Indicadores",
            options=INDICADORES,
            default=['ATIVO', 'PASSIVO', 'PATRIMONIO LIQUIDO']
        )

        # Filtrar dados
        df_filtrado = df[
            (df['empresa'].isin(empresas_selecionadas)) &
            (df['info'].isin(indicadores_selecionados))
        ]

        # Layout em duas colunas
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("Evolução Mensal por Indicador")
            
            # Criar gráfico de linha para evolução mensal
            for indicador in indicadores_selecionados:
                dados_indicador = df_filtrado[df_filtrado['info'] == indicador]
                
                if not dados_indicador.empty:
                    fig = go.Figure()
                    
                    for empresa in empresas_selecionadas:
                        dados_empresa = dados_indicador[dados_indicador['empresa'] == empresa]
                        if not dados_empresa.empty:
                            fig.add_trace(go.Scatter(
                                x=['Jan/25', 'Fev/25', 'Mar/25'],
                                y=[dados_empresa['01/2025'].iloc[0], 
                                   dados_empresa['02/2025'].iloc[0], 
                                   dados_empresa['03/2025'].iloc[0]],
                                name=empresa,
                                mode='lines+markers'
                            ))
                    
                    fig.update_layout(
                        title=f"{indicador}",
                        xaxis_title="Mês",
                        yaxis_title="Valor (R$)",
                        height=400
                    )
                    st.plotly_chart(fig, use_container_width=True)

        with col2:
            st.subheader("Análise Comparativa")
            
            # Tabela com valores mensais
            for empresa in empresas_selecionadas:
                st.write(f"### {empresa}")
                
                dados_empresa = df_filtrado[df_filtrado['empresa'] == empresa]
                if not dados_empresa.empty:
                    tabela_dados = []
                    for indicador in indicadores_selecionados:
                        dados_indicador = dados_empresa[dados_empresa['info'] == indicador]
                        if not dados_indicador.empty:
                            tabela_dados.append({
                                'Indicador': indicador,
                                'Janeiro': f"R$ {dados_indicador['01/2025'].iloc[0]:,.2f}",
                                'Fevereiro': f"R$ {dados_indicador['02/2025'].iloc[0]:,.2f}",
                                'Março': f"R$ {dados_indicador['03/2025'].iloc[0]:,.2f}",
                                'Saldo Acumulado': f"R$ {dados_indicador['Saldo acumulado'].iloc[0]:,.2f}"
                            })
                    
                    if tabela_dados:
                        df_tabela = pd.DataFrame(tabela_dados)
                        st.table(df_tabela)

        # Métricas importantes
        st.subheader("Métricas Consolidadas (Saldo Acumulado)")
        
        # Criar um DataFrame com os saldos acumulados
        saldos_consolidados = []
        for indicador in INDICADORES:
            saldo = df[df['info'] == indicador]['Saldo acumulado'].sum()
            saldos_consolidados.append({
                'Indicador': indicador,
                'Saldo Acumulado': f"R$ {saldo:,.2f}"
            })
        
        # Dividir os indicadores em três colunas
        df_saldos = pd.DataFrame(saldos_consolidados)
        n_indicadores = len(INDICADORES)
        n_por_coluna = (n_indicadores + 2) // 3  # Arredonda para cima
        
        col_saldos1, col_saldos2, col_saldos3 = st.columns(3)
        
        with col_saldos1:
            st.table(df_saldos.iloc[0:n_por_coluna])
        
        with col_saldos2:
            st.table(df_saldos.iloc[n_por_coluna:2*n_por_coluna])
        
        with col_saldos3:
            st.table(df_saldos.iloc[2*n_por_coluna:])

        # Análise de Tendências
        st.subheader("Análise de Tendências")
        col_tendencias1, col_tendencias2 = st.columns(2)

        with col_tendencias1:
            # Gráfico de barras para comparação entre empresas
            dados_resultado = df[df['info'] == 'RESULTADO DO MES']
            fig_barras = go.Figure()
            
            for empresa in empresas_selecionadas:
                dados_empresa = dados_resultado[dados_resultado['empresa'] == empresa]
                if not dados_empresa.empty:
                    fig_barras.add_trace(go.Bar(
                        name=empresa,
                        x=['Janeiro', 'Fevereiro', 'Março'],
                        y=[dados_empresa['01/2025'].iloc[0],
                           dados_empresa['02/2025'].iloc[0],
                           dados_empresa['03/2025'].iloc[0]]
                    ))
            
            fig_barras.update_layout(
                title="Comparativo do Resultado Mensal por Empresa",
                barmode='group',
                height=400
            )
            st.plotly_chart(fig_barras, use_container_width=True)

        with col_tendencias2:
            # Gráfico de área para patrimônio líquido
            dados_pl = df[df['info'] == 'PATRIMONIO LIQUIDO']
            fig_area = go.Figure()
            
            for empresa in empresas_selecionadas:
                dados_empresa = dados_pl[dados_pl['empresa'] == empresa]
                if not dados_empresa.empty:
                    fig_area.add_trace(go.Scatter(
                        name=empresa,
                        x=['Janeiro', 'Fevereiro', 'Março'],
                        y=[dados_empresa['01/2025'].iloc[0],
                           dados_empresa['02/2025'].iloc[0],
                           dados_empresa['03/2025'].iloc[0]],
                        fill='tonexty'
                    ))
            
            fig_area.update_layout(
                title="Evolução do Patrimônio Líquido",
                height=400
            )
            st.plotly_chart(fig_area, use_container_width=True)

        # Notas e Observações
        st.markdown("""
        ### Observações Importantes
        - Os dados apresentados correspondem ao primeiro trimestre de 2025
        - Todos os indicadores podem ser selecionados para análise detalhada
        - Os gráficos e tabelas são atualizados automaticamente com base nos filtros
        - O Resultado do Exercício é calculado de forma acumulativa
        - Os valores negativos em indicadores como Deduções/Custos/Despesas indicam saídas
        """)

# Página de login
def login_page():
    st.title("Login")
    
    # Criar colunas para centralizar o formulário
    col1, col2, col3 = st.columns([1,2,1])
    
    with col2:
        st.markdown("### Acesso ao Dashboard Financeiro")
        username = st.text_input("Usuário")
        password = st.text_input("Senha", type="password")
        
        if st.button("Entrar"):
            if check_login(username, password):
                st.session_state['logged_in'] = True
                st.rerun()
            else:
                st.error("Usuário ou senha incorretos!")

# Inicialização da sessão
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

# Roteamento principal
if st.session_state['logged_in']:
    main_page()
else:
    login_page() 