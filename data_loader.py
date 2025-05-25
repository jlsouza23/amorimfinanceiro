import pandas as pd
import glob
import os
from typing import Dict, List, Tuple
import logging
import warnings
from pathlib import Path
import subprocess
import struct
import io
import xlwings as xw

# Configuração de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class FinancialDataLoader:
    def __init__(self, data_dir: str = "."):
        """Inicializa o carregador de dados financeiros."""
        self.data_dir = data_dir
        self.companies = ["Marina", "PortoMarina", "PortoSantaMarina"]
        
    def get_excel_files(self) -> List[str]:
        """Retorna lista de arquivos Excel no diretório."""
        excel_files = glob.glob(os.path.join(self.data_dir, "*.xls*"))
        logger.info(f"Arquivos Excel encontrados: {excel_files}")
        return sorted(excel_files)

    def read_excel_with_xlwings(self, file_path: str) -> pd.DataFrame:
        """Lê arquivo Excel usando xlwings."""
        try:
            # Abre o arquivo Excel
            app = xw.App(visible=False)
            wb = app.books.open(file_path)
            
            try:
                # Pega a primeira planilha
                sheet = wb.sheets[0]
                
                # Pega a área usada
                used_range = sheet.used_range
                
                # Converte para DataFrame
                df = pd.DataFrame(used_range.value)
                
                logger.info(f"Dados lidos com xlwings:\n{df.head()}")
                return df
                
            finally:
                # Fecha o arquivo e a aplicação
                wb.close()
                app.quit()
                
        except Exception as e:
            logger.error(f"Erro ao ler com xlwings: {str(e)}")
            raise

    def try_read_excel(self, file_path: str) -> pd.DataFrame:
        """Tenta diferentes métodos para ler o arquivo Excel."""
        errors = []
        
        # Lista de tentativas de leitura com diferentes engines e parâmetros
        attempts = [
            # Tentativa 1: xlwings
            lambda: self.read_excel_with_xlwings(file_path),
            
            # Tentativa 2: pandas com engine auto e sem cabeçalho
            lambda: pd.read_excel(file_path, header=None),
            
            # Tentativa 3: xlrd específico para xls e sem cabeçalho
            lambda: pd.read_excel(file_path, engine='xlrd', header=None),
            
            # Tentativa 4: openpyxl específico para xlsx e sem cabeçalho
            lambda: pd.read_excel(file_path, engine='openpyxl', header=None),
            
            # Tentativa 5: pyxlsb para arquivos binários
            lambda: pd.read_excel(file_path, engine='pyxlsb', header=None)
        ]
        
        for attempt_func in attempts:
            try:
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    df = attempt_func()
                    if not df.empty:
                        logger.info(f"\nEstrutura do arquivo {file_path}:")
                        logger.info(f"Dimensões: {df.shape}")
                        logger.info(f"Primeiras 10 linhas:\n{df.head(10)}")
                        return df
            except Exception as e:
                errors.append(str(e))
                continue
        
        # Se todas as tentativas falharem, loga os erros e levanta exceção
        error_msg = "\n".join(errors)
        logger.error(f"Todas as tentativas de leitura falharam para {file_path}:\n{error_msg}")
        raise ValueError(f"Não foi possível ler o arquivo {file_path}")

    def find_header_row(self, df: pd.DataFrame) -> int:
        """Encontra a linha que contém o cabeçalho do Comparativo de Movimento."""
        # Procura por "COMPARATIVO DE MOVIMENTO" nas primeiras 10 linhas
        for idx in range(min(10, len(df))):
            row = df.iloc[idx].astype(str)
            if any('COMPARATIVO DE MOVIMENTO' in str(val).upper() for val in row):
                # O cabeçalho geralmente está 2 linhas abaixo
                return idx + 2
        return 6  # Retorna o padrão se não encontrar

    def extract_company_name_from_excel(self, df: pd.DataFrame, file_path: str) -> str:
        """Extrai o nome da empresa do DataFrame ou nome do arquivo."""
        try:
            # Procura por "Empresa:" nas primeiras 3 linhas
            for idx in range(min(3, len(df))):
                row = df.iloc[idx].astype(str)
                for valor in row:
                    if 'Empresa:' in str(valor):
                        empresa = valor.split('Empresa:')[-1].strip()
                        logger.info(f"Nome da empresa encontrado na linha {idx}: {empresa}")
                        return empresa
            
            # Se não encontrar, tenta extrair do nome do arquivo
            file_name = os.path.basename(file_path)
            for company in self.companies:
                if company in file_name:
                    logger.info(f"Nome da empresa extraído do arquivo: {company}")
                    return company
                    
        except Exception as e:
            logger.error(f"Erro ao extrair nome da empresa: {str(e)}")
        return None

    def clean_column_name(self, col: str) -> str:
        """Limpa e padroniza o nome da coluna."""
        if not isinstance(col, str):
            return str(col)
        
        # Remove caracteres não imprimíveis
        col = ''.join(char for char in col if char.isprintable())
        
        # Remove espaços extras
        col = col.strip()
        
        return col

    def process_movimento_block(self, df: pd.DataFrame) -> pd.DataFrame:
        """Processa o bloco de Comparativo de Movimento."""
        try:
            # Encontra a linha do cabeçalho
            header_row_idx = self.find_header_row(df)
            logger.info(f"Linha do cabeçalho encontrada: {header_row_idx}")
            
            # Pega o cabeçalho e limpa os nomes das colunas
            header_row = df.iloc[header_row_idx].apply(self.clean_column_name)
            logger.info(f"Cabeçalho encontrado: {header_row.tolist()}")
            
            # Cria novo DataFrame com os dados após o cabeçalho
            movimento_df = df.iloc[header_row_idx + 1:].copy()
            movimento_df.columns = header_row
            
            # Identifica colunas numéricas (meses)
            meses_cols = []
            for col in movimento_df.columns:
                col_clean = self.clean_column_name(col)
                if '/' in col_clean:
                    meses_cols.append(col)
                elif 'acumulado' in col_clean.lower():
                    meses_cols.append(col)
            
            logger.info(f"Colunas de meses encontradas: {meses_cols}")
            
            # Procura por colunas que podem ter nomes diferentes
            col_mapping = {
                'Código': ['Código', 'Codigo', 'CÓDIGO', 'CODIGO', '0', 0],
                'Classificação': ['Classificação', 'Classificacao', 'CLASSIFICAÇÃO', 'CLASSIFICACAO', '1', 1],
                'Descrição': ['Descrição', 'Descricao', 'DESCRIÇÃO', 'DESCRICAO', '2', 2]
            }
            
            # Mapeia as colunas encontradas
            colunas_encontradas = {}
            for col_name, alternatives in col_mapping.items():
                for alt in alternatives:
                    if str(alt) in [str(col) for col in movimento_df.columns]:
                        colunas_encontradas[col_name] = alt
                        break
            
            logger.info(f"Mapeamento de colunas: {colunas_encontradas}")
            
            # Renomeia as colunas encontradas
            movimento_df = movimento_df.rename(columns=dict(zip(colunas_encontradas.values(), colunas_encontradas.keys())))
            
            # Seleciona as colunas na ordem correta
            colunas_finais = list(colunas_encontradas.keys()) + meses_cols
            movimento_df = movimento_df[colunas_finais]
            
            # Remove linhas vazias
            movimento_df = movimento_df.dropna(how='all')
            
            # Adiciona tipo do bloco
            movimento_df['Bloco'] = 'Comparativo de Movimento'
            
            return movimento_df
        
        except Exception as e:
            logger.error(f"Erro ao processar bloco de movimento: {str(e)}")
            logger.error(f"Colunas disponíveis: {df.columns.tolist()}")
            return pd.DataFrame()

    def process_balancete_block(self, df: pd.DataFrame) -> pd.DataFrame:
        """Processa o bloco de Resumo do Balancete."""
        try:
            # Procura o início do bloco de balancete
            balancete_start = None
            for idx, row in df.iterrows():
                row_str = row.astype(str)
                if any('ATIVO' in str(val).upper() for val in row_str):
                    balancete_start = idx
                    logger.info(f"Início do balancete encontrado na linha {idx}")
                    break
                    
            if balancete_start is None:
                logger.warning("Bloco de balancete não encontrado")
                return pd.DataFrame()
            
            # Lista de categorias do balancete (case insensitive)
            categorias = [
                'ATIVO', 'PASSIVO', 'PATRIMONIO LIQUIDO', 'ESTOQUES', 'COMPENSACOES',
                'RECEITA OPERACIONAL BRUTA', 'DEDUCOES/CUSTOS/DESPESAS', 'APURACAO DO RESULTADO',
                'CONTAS DEVEDORAS', 'CONTAS CREDORAS', 'RESULTADO DO MES', 'RESULTADO DO EXERCÍCIO'
            ]
            
            # Cria DataFrame do balancete
            balancete_df = df.iloc[balancete_start:].copy()
            
            # Usa o mesmo cabeçalho do bloco de movimento
            header_row_idx = self.find_header_row(df)
            header_row = df.iloc[header_row_idx].apply(self.clean_column_name)
            balancete_df.columns = header_row
            
            # Identifica colunas numéricas (meses)
            meses_cols = []
            for col in balancete_df.columns:
                col_clean = self.clean_column_name(col)
                if '/' in col_clean:
                    meses_cols.append(col)
                elif 'acumulado' in col_clean.lower():
                    meses_cols.append(col)
            
            # Pega apenas as linhas que são categorias (case insensitive)
            balancete_df['temp_upper'] = balancete_df.iloc[:, 0].astype(str).str.upper()
            balancete_df = balancete_df[balancete_df['temp_upper'].isin([cat.upper() for cat in categorias])]
            balancete_df = balancete_df.drop('temp_upper', axis=1)
            
            # Renomeia a primeira coluna para 'Descrição'
            balancete_df = balancete_df.rename(columns={balancete_df.columns[0]: 'Descrição'})
            
            # Seleciona apenas as colunas relevantes
            colunas_finais = ['Descrição'] + meses_cols
            balancete_df = balancete_df[colunas_finais]
            
            # Adiciona tipo do bloco
            balancete_df['Bloco'] = 'Resumo do Balancete'
            
            return balancete_df
            
        except Exception as e:
            logger.error(f"Erro ao processar bloco de balancete: {str(e)}")
            logger.error(f"Colunas disponíveis: {df.columns.tolist()}")
            return pd.DataFrame()

    def load_excel_file(self, file_path: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Carrega um arquivo Excel e retorna os dois blocos de dados processados."""
        try:
            logger.info(f"\nCarregando arquivo: {file_path}")
            
            # Lê todo o arquivo
            df = self.try_read_excel(file_path)
            
            # Extrai nome da empresa
            empresa = self.extract_company_name_from_excel(df, file_path)
            logger.info(f"Empresa detectada: {empresa}")
            
            # Processa os dois blocos
            movimento_df = self.process_movimento_block(df)
            balancete_df = self.process_balancete_block(df)
            
            # Adiciona informações comuns
            for data in [movimento_df, balancete_df]:
                if not data.empty:
                    data['Empresa'] = empresa
                    data['Arquivo_Origem'] = os.path.basename(file_path)
                    # Extrai o ano do nome do arquivo
                    if "2024" in file_path:
                        data['Ano'] = "2024"
                    elif "2025" in file_path:
                        data['Ano'] = "2025"
            
            return movimento_df, balancete_df
            
        except Exception as e:
            logger.error(f"Erro ao carregar {file_path}: {str(e)}")
            return pd.DataFrame(), pd.DataFrame()
    
    def load_all_data(self) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Carrega todos os arquivos Excel e retorna dois DataFrames consolidados."""
        all_movimento = []
        all_balancete = []
        
        for excel_file in self.get_excel_files():
            try:
                movimento_df, balancete_df = self.load_excel_file(excel_file)
                if not movimento_df.empty:
                    all_movimento.append(movimento_df)
                if not balancete_df.empty:
                    all_balancete.append(balancete_df)
            except Exception as e:
                logger.error(f"Erro ao processar {excel_file}: {str(e)}")
                continue
                
        if not all_movimento and not all_balancete:
            raise ValueError("Nenhum dado foi carregado!")
            
        # Consolida os dados
        final_movimento = pd.concat(all_movimento, ignore_index=True) if all_movimento else pd.DataFrame()
        final_balancete = pd.concat(all_balancete, ignore_index=True) if all_balancete else pd.DataFrame()
        
        logger.info(f"\nDados consolidados:")
        logger.info(f"Movimento - Total de linhas: {len(final_movimento)}")
        logger.info(f"Balancete - Total de linhas: {len(final_balancete)}")
        
        return final_movimento, final_balancete 