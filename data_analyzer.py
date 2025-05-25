import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from typing import Dict, List

class FinancialAnalyzer:
    def __init__(self, data: pd.DataFrame):
        """Inicializa o analisador com os dados financeiros.
        
        Args:
            data (pd.DataFrame): DataFrame com os dados consolidados
        """
        self.data = data
        self.process_data()
        
    def process_data(self):
        """Processa e limpa os dados iniciais."""
        # Remove linhas totalmente vazias
        self.data = self.data.dropna(how='all')
        
        # Converte colunas numéricas
        numeric_columns = self.data.select_dtypes(include=['object']).columns
        for col in numeric_columns:
            try:
                self.data[col] = pd.to_numeric(self.data[col].astype(str).str.replace('R$', '').str.replace('.', '').str.replace(',', '.'), errors='ignore')
            except:
                continue
                
    def get_companies(self) -> List[str]:
        """Retorna lista de empresas disponíveis."""
        return sorted(self.data['Company'].unique())
    
    def get_years(self) -> List[str]:
        """Retorna lista de anos disponíveis."""
        return sorted(self.data['Year'].unique())
    
    def get_periods(self) -> List[str]:
        """Retorna lista de períodos disponíveis."""
        return sorted(self.data['Period'].unique())
    
    def filter_data(self, company: str = None, year: str = None, period: str = None) -> pd.DataFrame:
        """Filtra os dados com base nos parâmetros fornecidos."""
        filtered_data = self.data.copy()
        
        if company:
            filtered_data = filtered_data[filtered_data['Company'] == company]
        if year:
            filtered_data = filtered_data[filtered_data['Year'] == year]
        if period:
            filtered_data = filtered_data[filtered_data['Period'] == period]
            
        return filtered_data
    
    def create_comparison_chart(self, metric_column: str, companies: List[str] = None) -> go.Figure:
        """Cria um gráfico comparativo entre empresas para uma métrica específica."""
        if companies is None:
            companies = self.get_companies()
            
        fig = go.Figure()
        
        for company in companies:
            company_data = self.data[self.data['Company'] == company]
            
            fig.add_trace(go.Bar(
                name=company,
                x=company_data['Year'].astype(str) + ' ' + company_data['Period'],
                y=company_data[metric_column],
                text=company_data[metric_column].round(2),
                textposition='auto',
            ))
            
        fig.update_layout(
            title=f'Comparação de {metric_column} por Empresa',
            xaxis_title='Período',
            yaxis_title=metric_column,
            barmode='group'
        )
        
        return fig
    
    def create_pie_chart(self, metric_column: str, year: str = None, period: str = None) -> go.Figure:
        """Cria um gráfico de pizza para comparação entre empresas."""
        filtered_data = self.filter_data(year=year, period=period)
        
        fig = px.pie(
            filtered_data,
            values=metric_column,
            names='Company',
            title=f'Distribuição de {metric_column} por Empresa'
        )
        
        return fig
    
    def create_time_series(self, metric_column: str, companies: List[str] = None) -> go.Figure:
        """Cria um gráfico de série temporal para uma métrica específica."""
        if companies is None:
            companies = self.get_companies()
            
        fig = go.Figure()
        
        for company in companies:
            company_data = self.data[self.data['Company'] == company]
            
            fig.add_trace(go.Scatter(
                name=company,
                x=company_data['Year'].astype(str) + ' ' + company_data['Period'],
                y=company_data[metric_column],
                mode='lines+markers'
            ))
            
        fig.update_layout(
            title=f'Evolução de {metric_column} ao Longo do Tempo',
            xaxis_title='Período',
            yaxis_title=metric_column
        )
        
        return fig 