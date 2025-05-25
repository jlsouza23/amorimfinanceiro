import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime

# Dados das empresas
data = {
    'empresa': [
        'PORTO SANTA MARIA EMPREENDIMENTOS TURIST', 'PORTO SANTA MARIA EMPREENDIMENTOS TURIST',
        'PORTO SANTA MARIA EMPREENDIMENTOS TURIST', 'PORTO SANTA MARIA EMPREENDIMENTOS TURIST',
        'PORTO SANTA MARIA EMPREENDIMENTOS TURIST', 'PORTO SANTA MARIA EMPREENDIMENTOS TURIST',
        'PORTO SANTA MARIA EMPREENDIMENTOS TURIST', 'PORTO SANTA MARIA EMPREENDIMENTOS TURIST',
        'PORTO SANTA MARIA EMPREENDIMENTOS TURIST', 'PORTO SANTA MARIA EMPREENDIMENTOS TURIST',
        'PORTO SANTA MARIA EMPREENDIMENTOS TURIST', 'PORTO SANTA MARIA EMPREENDIMENTOS TURIST',
        'MARINA ASTURIAS - SERVICOS NAVAIS LTDA', 'MARINA ASTURIAS - SERVICOS NAVAIS LTDA'
    ],
    'info': [
        'ATIVO', 'PASSIVO', 'PATRIMONIO LIQUIDO', 'ESTOQUES', 'COMPENSACOES',
        'RECEITA OPERACIONAL BRUTA', 'DEDUCOES/CUSTOS/DESPESAS', 'APURACAO DO RESULTADO',
        'CONTAS DEVEDORAS', 'CONTAS CREDORAS', 'RESULTADO DO MES', 'RESULTADO DO EXERCÍCIO',
        'ATIVO', 'PASSIVO'
    ],
    '01/2025': [
        0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 14369.96, 0.00, 14369.96, 0.00, 14369.96, 14369.96,
        0.00, 0.00
    ],
    '02/2025': [
        0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 14369.84, 0.00, 14369.84, 0.00, 14369.84, 28739.80,
        0.00, 0.00
    ],
    '03/2025': [
        0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 17317.07, 0.00, 17317.07, 0.00, 17317.07, 46056.87,
        0.00, 0.00
    ],
    'Saldo acumulado': [
        0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 46056.87, 0.00, 46056.87, 0.00, 46056.87, 46056.87,
        0.00, 0.00
    ]
}

# Criar DataFrame
df = pd.DataFrame(data)

def analisar_por_empresa():
    """Análise separada por empresa"""
    print("\n=== Análise por Empresa ===")
    
    for empresa in df['empresa'].unique():
        print(f"\nEmpresa: {empresa}")
        df_empresa = df[df['empresa'] == empresa]
        
        # Análise dos principais indicadores
        for info in ['DEDUCOES/CUSTOS/DESPESAS', 'RESULTADO DO MES', 'RESULTADO DO EXERCÍCIO']:
            dados = df_empresa[df_empresa['info'] == info]
            if not dados.empty:
                print(f"\n{info}:")
                print(f"Janeiro 2025: R$ {dados['01/2025'].iloc[0]:,.2f}")
                print(f"Fevereiro 2025: R$ {dados['02/2025'].iloc[0]:,.2f}")
                print(f"Março 2025: R$ {dados['03/2025'].iloc[0]:,.2f}")
                print(f"Saldo Acumulado: R$ {dados['Saldo acumulado'].iloc[0]:,.2f}")

def analisar_consolidado():
    """Análise consolidada de todas as empresas"""
    print("\n=== Análise Consolidada ===")
    
    # Soma dos valores por informação
    consolidado = df.groupby('info').agg({
        '01/2025': 'sum',
        '02/2025': 'sum',
        '03/2025': 'sum',
        'Saldo acumulado': 'sum'
    })
    
    # Mostrar principais indicadores consolidados
    for info in ['DEDUCOES/CUSTOS/DESPESAS', 'RESULTADO DO MES', 'RESULTADO DO EXERCÍCIO']:
        if info in consolidado.index:
            print(f"\n{info} (Consolidado):")
            print(f"Janeiro 2025: R$ {consolidado.loc[info, '01/2025']:,.2f}")
            print(f"Fevereiro 2025: R$ {consolidado.loc[info, '02/2025']:,.2f}")
            print(f"Março 2025: R$ {consolidado.loc[info, '03/2025']:,.2f}")
            print(f"Saldo Acumulado: R$ {consolidado.loc[info, 'Saldo acumulado']:,.2f}")

def gerar_graficos():
    """Gera gráficos para visualização dos dados"""
    # Configurar estilo dos gráficos
    plt.style.use('default')  # Usando estilo padrão
    
    # Gráfico 1: Evolução do Resultado do Exercício por Empresa
    plt.figure(figsize=(12, 6))
    for empresa in df['empresa'].unique():
        dados = df[(df['empresa'] == empresa) & (df['info'] == 'RESULTADO DO EXERCÍCIO')]
        if not dados.empty:
            plt.plot(['Jan/25', 'Fev/25', 'Mar/25'], 
                    [dados['01/2025'].iloc[0], dados['02/2025'].iloc[0], dados['03/2025'].iloc[0]],
                    marker='o', label=empresa.split(' - ')[0])
    
    plt.title('Evolução do Resultado do Exercício por Empresa')
    plt.xlabel('Mês')
    plt.ylabel('Valor (R$)')
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()
    plt.savefig('resultado_exercicio.png')
    plt.close()
    
    # Gráfico 2: Comparativo de Deduções/Custos/Despesas
    plt.figure(figsize=(12, 6))
    dados_deducoes = df[df['info'] == 'DEDUCOES/CUSTOS/DESPESAS']
    empresas = dados_deducoes['empresa'].unique()
    
    x = range(len(empresas))
    width = 0.25
    
    plt.bar([i - width for i in x], dados_deducoes['01/2025'], width, label='Janeiro')
    plt.bar(x, dados_deducoes['02/2025'], width, label='Fevereiro')
    plt.bar([i + width for i in x], dados_deducoes['03/2025'], width, label='Março')
    
    plt.title('Comparativo de Deduções/Custos/Despesas por Empresa')
    plt.xlabel('Empresa')
    plt.ylabel('Valor (R$)')
    plt.xticks(x, [emp.split(' - ')[0] for emp in empresas], rotation=45)
    plt.legend()
    plt.tight_layout()
    plt.savefig('deducoes_custos.png')
    plt.close()

def main():
    print("=== Análise Financeira - 1º Trimestre 2025 ===")
    
    # Análise individual por empresa
    analisar_por_empresa()
    
    # Análise consolidada
    analisar_consolidado()
    
    # Gerar gráficos
    gerar_graficos()
    
    print("\nGráficos gerados:")
    print("1. resultado_exercicio.png - Evolução do Resultado do Exercício por Empresa")
    print("2. deducoes_custos.png - Comparativo de Deduções/Custos/Despesas por Empresa")

if __name__ == "__main__":
    main() 