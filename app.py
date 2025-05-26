import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import hashlib
import os
import traceback
from pathlib import Path

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

# Função para carregar o arquivo Excel com tratamento de erros
def load_excel_file(file_path):
    try:
        # Verifica se o arquivo existe
        if not os.path.exists(file_path):
            st.error(f"❌ Arquivo não encontrado: {file_path}")
            st.info("📁 Diretório atual: " + os.getcwd())
            st.info("📄 Arquivos disponíveis: " + ", ".join(os.listdir()))
            return None

        # Verifica o tamanho do arquivo
        file_size = os.path.getsize(file_path)
        if file_size == 0:
            st.error("❌ O arquivo está vazio!")
            return None

        # Tenta ler o arquivo
        st.info(f"📊 Tentando ler o arquivo: {file_path}")
        df = pd.read_excel(file_path)
        
        if df.empty:
            st.error("❌ O arquivo foi lido mas está vazio!")
            return None
            
        st.success(f"✅ Arquivo carregado com sucesso! Dimensões: {df.shape}")
        return df

    except FileNotFoundError as e:
        st.error(f"❌ Erro ao encontrar o arquivo: {str(e)}")
        return None
    except PermissionError as e:
        st.error(f"❌ Erro de permissão ao ler o arquivo: {str(e)}")
        return None
    except pd.errors.EmptyDataError as e:
        st.error(f"❌ O arquivo está vazio ou mal formatado: {str(e)}")
        return None
    except Exception as e:
        st.error(f"❌ Erro inesperado ao ler o arquivo: {str(e)}")
        st.error("Detalhes do erro:")
        st.code(traceback.format_exc())
        return None

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
            # Define o caminho do arquivo
            file_path = 'Consolidadas_1tri_2025.xlsx'
            
            # Carrega o arquivo Excel
            df = load_excel_file(file_path)
            if df is None:
                return None
            
            # Renomeia as colunas para o formato esperado
            try:
                df = df.rename(columns={
                    'emopresa': 'empresa',
                    datetime(2025, 1, 31): '01/2025',
                    datetime(2025, 2, 28): '02/2025',
                    datetime(2025, 3, 31): '03/2025'
                })
            except Exception as e:
                st.error("❌ Erro ao renomear colunas:")
                st.error(str(e))
                st.info("📊 Colunas disponíveis: " + ", ".join(df.columns))
                return None
            
            # Converte valores monetários para float
            try:
                for col in ['01/2025', '02/2025', '03/2025', 'Saldo acumulado']:
                    if col not in df.columns:
                        st.error(f"❌ Coluna não encontrada: {col}")
                        continue
                    df[col] = pd.to_numeric(df[col].astype(str).str.replace('$', '').str.replace(',', ''), errors='coerce')
            except Exception as e:
                st.error("❌ Erro ao converter valores monetários:")
                st.error(str(e))
                return None
            
            return df

        except Exception as e:
            st.error("❌ Erro inesperado ao preparar os dados:")
            st.error(str(e))
            st.code(traceback.format_exc())
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
        
        # Criar um DataFrame com os saldos acumulados por empresa
        saldos_consolidados = []
        for empresa in empresas_selecionadas:
            for indicador in INDICADORES:
                dados = df[(df['empresa'] == empresa) & (df['info'] == indicador)]
                if not dados.empty:
                    saldo = dados['Saldo acumulado'].iloc[0]  # Pega o saldo acumulado para esta empresa e indicador
                    saldos_consolidados.append({
                        'Empresa': empresa,
                        'Indicador': indicador,
                        'Saldo Acumulado': f"R$ {saldo:,.2f}"
                    })
        
        # Criar DataFrame com os saldos
        if saldos_consolidados:
            df_saldos = pd.DataFrame(saldos_consolidados)
            
            # Dividir os dados por empresa
            for empresa in empresas_selecionadas:
                st.write(f"### {empresa}")
                df_empresa = df_saldos[df_saldos['Empresa'] == empresa]
                if not df_empresa.empty:
                    st.table(df_empresa[['Indicador', 'Saldo Acumulado']])

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
                title="Resultado Mensal por Empresa",
                xaxis_title="Mês",
                yaxis_title="Valor (R$)",
                height=400
            )
            st.plotly_chart(fig_barras, use_container_width=True)

        with col_tendencias2:
            # Gráfico de pizza para distribuição do resultado
            dados_exercicio = df[df['info'] == 'RESULTADO DO EXERCÍCIO']
            if not dados_exercicio.empty:
                valores = dados_exercicio['Saldo acumulado'].values
                empresas = dados_exercicio['empresa'].values
                
                fig_pizza = go.Figure(data=[go.Pie(
                    labels=empresas,
                    values=valores,
                    hole=.3
                )])
                
                fig_pizza.update_layout(
                    title="Distribuição do Resultado do Exercício",
                    height=400
                )
                st.plotly_chart(fig_pizza, use_container_width=True)

# Página de login
def login_page():
    st.title("🔐 Login")
    
    username = st.text_input("Usuário")
    password = st.text_input("Senha", type="password")
    
    if st.button("Entrar"):
        if check_login(username, password):
            st.session_state['logged_in'] = True
            st.rerun()
        else:
            st.error("Usuário ou senha incorretos!")

# Controle de sessão
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

if st.session_state['logged_in']:
    main_page()
else:
    login_page() 