# Dashboard Financeiro - 1º Trimestre 2025

Este é um dashboard interativo desenvolvido com Streamlit para análise financeira do primeiro trimestre de 2025.

## Funcionalidades

- Visualização de indicadores financeiros por empresa
- Gráficos comparativos de evolução mensal
- Análise de tendências
- Métricas consolidadas
- Sistema de login seguro

## Requisitos

- Python 3.13.3 ou superior
- Dependências listadas em `requirements.txt`

## Instalação

1. Clone o repositório
2. Instale as dependências:
```bash
pip install -r requirements.txt
```

## Uso

1. Coloque o arquivo `Consolidadas_1tri_2025.xlsx` na mesma pasta do `app.py`
2. Execute o aplicativo:
```bash
streamlit run app.py
```

3. Acesse o dashboard usando as seguintes credenciais:
   - Usuário: cintia.ferreira
   - Senha: Cf2025

## Estrutura de Dados

O arquivo Excel deve conter as seguintes colunas:
- empresa
- info
- 01/2025
- 02/2025
- 03/2025
- Saldo acumulado 