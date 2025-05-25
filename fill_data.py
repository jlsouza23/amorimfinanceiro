from create_simplified_excel import (
    create_balancete_structure,
    create_consolidated_sheet,
    create_company_sheet,
    save_to_excel,
    create_sheet_2024,
    create_sheet_2025
)
import pandas as pd

def format_valor(valor, is_debito=False):
    """Formata o valor no padrão desejado."""
    if valor == 0:
        return '0,00'
    
    # Formata o número com duas casas decimais
    valor_str = f"{abs(valor):,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')
    
    # Adiciona o 'd' se for débito
    if is_debito:
        valor_str += 'd'
    
    return valor_str

def fill_marina_2024_data(df):
    """Preenche os dados de 2024 da Marina."""
    # Exemplo de preenchimento (você deve substituir pelos dados reais)
    data = {
        'ATIVO': [1000000.00, 1050000.00, 1100000.00, 1150000.00, 1200000.00, 1250000.00,
                 1300000.00, 1350000.00, 1400000.00, 1450000.00, 1500000.00, 1550000.00],
        'PASSIVO': [400000.00, 420000.00, 440000.00, 460000.00, 480000.00, 500000.00,
                   520000.00, 540000.00, 560000.00, 580000.00, 600000.00, 620000.00],
        'PATRIMONIO LIQUIDO': [600000.00, 630000.00, 660000.00, 690000.00, 720000.00, 750000.00,
                             780000.00, 810000.00, 840000.00, 870000.00, 900000.00, 930000.00],
        'ESTOQUES': [200000.00, 210000.00, 220000.00, 230000.00, 240000.00, 250000.00,
                    260000.00, 270000.00, 280000.00, 290000.00, 300000.00, 310000.00],
        'COMPENSACOES': [50000.00, 52500.00, 55000.00, 57500.00, 60000.00, 62500.00,
                       65000.00, 67500.00, 70000.00, 72500.00, 75000.00, 77500.00],
        'RECEITA OPERACIONAL BRUTA': [300000.00, 315000.00, 330000.00, 345000.00, 360000.00, 375000.00,
                                    390000.00, 405000.00, 420000.00, 435000.00, 450000.00, 465000.00],
        'DEDUCOES/CUSTOS/DESPESAS': [-180000.00, -189000.00, -198000.00, -207000.00, -216000.00, -225000.00,
                                   -234000.00, -243000.00, -252000.00, -261000.00, -270000.00, -279000.00],
        'APURACAO DO RESULTADO': [120000.00, 126000.00, 132000.00, 138000.00, 144000.00, 150000.00,
                                156000.00, 162000.00, 168000.00, 174000.00, 180000.00, 186000.00],
        'CONTAS DEVEDORAS': [80000.00, 84000.00, 88000.00, 92000.00, 96000.00, 100000.00,
                           104000.00, 108000.00, 112000.00, 116000.00, 120000.00, 124000.00],
        'CONTAS CREDORAS': [-60000.00, -63000.00, -66000.00, -69000.00, -72000.00, -75000.00,
                          -78000.00, -81000.00, -84000.00, -87000.00, -90000.00, -93000.00],
        'RESULTADO DO MES': [20000.00, 21000.00, 22000.00, 23000.00, 24000.00, 25000.00,
                           26000.00, 27000.00, 28000.00, 29000.00, 30000.00, 31000.00],
        'RESULTADO DO EXERCÍCIO': [20000.00, 41000.00, 63000.00, 86000.00, 110000.00, 135000.00,
                                161000.00, 188000.00, 216000.00, 245000.00, 275000.00, 306000.00]
    }
    
    # Preenche o DataFrame
    for info, valores in data.items():
        idx = df[df['Info'] == info].index[0]
        for mes, valor in enumerate(valores):
            col = f'Jan/2024' if mes == 0 else f'{["Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"][mes-1]}/2024'
            df.at[idx, col] = valor
        df.at[idx, 'Saldo Consolidado'] = valores[-1]  # Último valor como saldo consolidado

def fill_marina_2025_data(df):
    """Preenche os dados de 2025 da Marina."""
    # Exemplo de preenchimento (você deve substituir pelos dados reais)
    data = {
        'ATIVO': [1600000.00, 1650000.00, 1700000.00],
        'PASSIVO': [640000.00, 660000.00, 680000.00],
        'PATRIMONIO LIQUIDO': [960000.00, 990000.00, 1020000.00],
        'ESTOQUES': [320000.00, 330000.00, 340000.00],
        'COMPENSACOES': [80000.00, 82500.00, 85000.00],
        'RECEITA OPERACIONAL BRUTA': [480000.00, 495000.00, 510000.00],
        'DEDUCOES/CUSTOS/DESPESAS': [-288000.00, -297000.00, -306000.00],
        'APURACAO DO RESULTADO': [192000.00, 198000.00, 204000.00],
        'CONTAS DEVEDORAS': [128000.00, 132000.00, 136000.00],
        'CONTAS CREDORAS': [-96000.00, -99000.00, -102000.00],
        'RESULTADO DO MES': [32000.00, 33000.00, 34000.00],
        'RESULTADO DO EXERCÍCIO': [32000.00, 65000.00, 99000.00]
    }
    
    # Preenche o DataFrame
    for info, valores in data.items():
        idx = df[df['Info'] == info].index[0]
        for mes, valor in enumerate(valores):
            col = f'{["Jan", "Fev", "Mar"][mes]}/2025'
            df.at[idx, col] = valor
        df.at[idx, 'Saldo Consolidado'] = valores[-1]  # Último valor como saldo consolidado

def fill_porto_santa_maria_data(df):
    """Preenche os dados do Porto Santa Maria."""
    # Dados do exemplo
    data = {
        'DEDUCOES/CUSTOS/DESPESAS': [-14369.96, -14369.84, -17317.07],
        'CONTAS DEVEDORAS': [-14369.96, -14369.84, -17317.07],
        'RESULTADO DO MES': [-14369.96, -14369.84, -17317.07],
        'RESULTADO DO EXERCÍCIO': [-14369.96, -28739.80, -46056.87]
    }
    
    # Preenche o DataFrame
    for info in df['Info'].unique():
        if info == '':  # Pula linhas em branco
            continue
            
        idx = df[df['Info'] == info].index[0]
        
        if info in data:
            valores = data[info]
            # Preenche os meses
            df.at[idx, '01/2025'] = format_valor(valores[0], True)
            df.at[idx, '02/2025'] = format_valor(valores[1], True)
            df.at[idx, '03/2025'] = format_valor(valores[2], True)
            
            # Preenche o saldo acumulado
            if info == 'RESULTADO DO EXERCÍCIO':
                df.at[idx, 'Saldo Acumulado'] = format_valor(valores[2], True)
            else:
                soma = sum(valores)
                df.at[idx, 'Saldo Acumulado'] = format_valor(soma, True)
        else:
            # Zera os valores para as outras informações
            df.at[idx, '01/2025'] = '0,00'
            df.at[idx, '02/2025'] = '0,00'
            df.at[idx, '03/2025'] = '0,00'
            df.at[idx, 'Saldo Acumulado'] = '0,00'

def fill_marina_data(df):
    """Preenche os dados da Marina."""
    # Por enquanto, mantém os valores zerados
    pass

def fill_porto_marina_data(df):
    """Preenche os dados do Porto Marina."""
    # Por enquanto, mantém os valores zerados
    pass

def fill_porto_santa_maria_2025():
    """Preenche os dados de 2025 do Porto Santa Maria."""
    # Cria um novo DataFrame
    df = create_sheet_2025()
    df['Info'] = create_balancete_structure()
    
    # Dados reais do exemplo
    data = {
        'DEDUCOES/CUSTOS/DESPESAS': [-14369.96, -14369.84, -17317.07],
        'CONTAS DEVEDORAS': [-14369.96, -14369.84, -17317.07],
        'RESULTADO DO MES': [-14369.96, -14369.84, -17317.07],
        'RESULTADO DO EXERCÍCIO': [-14369.96, -28739.80, -46056.87]
    }
    
    # Preenche o DataFrame
    for info in df['Info'].unique():
        if info == '':  # Pula linhas em branco
            continue
            
        idx = df[df['Info'] == info].index[0]
        
        if info in data:
            valores = data[info]
            # Preenche os meses
            df.at[idx, '01/2025'] = format_valor(valores[0], True)
            df.at[idx, '02/2025'] = format_valor(valores[1], True)
            df.at[idx, '03/2025'] = format_valor(valores[2], True)
            
            # Preenche o saldo acumulado
            if info == 'RESULTADO DO EXERCÍCIO':
                df.at[idx, 'Saldo Acumulado'] = format_valor(valores[2], True)
            else:
                soma = sum(valores)
                df.at[idx, 'Saldo Acumulado'] = format_valor(soma, True)
        else:
            # Zera os valores para as outras informações
            df.at[idx, '01/2025'] = '0,00'
            df.at[idx, '02/2025'] = '0,00'
            df.at[idx, '03/2025'] = '0,00'
            df.at[idx, 'Saldo Acumulado'] = '0,00'
    
    return df

def fill_porto_santa_maria_2024():
    """Preenche os dados de 2024 do Porto Santa Maria."""
    # Cria um novo DataFrame
    df = create_sheet_2024()
    df['Info'] = create_balancete_structure()
    
    # Dados de exemplo para 2024
    data = {
        'ATIVO': [1000000] * 12,  # Valor constante para exemplo
        'PASSIVO': [400000] * 12,
        'PATRIMONIO LIQUIDO': [600000] * 12,
        'ESTOQUES': [list(range(200000, 260000, 5000))] * 12,  # Crescimento mensal
        'COMPENSACOES': [50000] * 12,
        'RECEITA OPERACIONAL BRUTA': [list(range(300000, 420000, 10000))] * 12,
        'DEDUCOES/CUSTOS/DESPESAS': [-180000] * 12,
        'APURACAO DO RESULTADO': [120000] * 12,
        'CONTAS DEVEDORAS': [80000] * 12,
        'CONTAS CREDORAS': [-60000] * 12,
        'RESULTADO DO MES': [20000] * 12,
        'RESULTADO DO EXERCÍCIO': [list(range(20000, 260000, 20000))] * 12
    }
    
    # Preenche o DataFrame
    for info in df['Info'].unique():
        if info == '':  # Pula linhas em branco
            continue
            
        idx = df[df['Info'] == info].index[0]
        
        if info in data:
            valores = data[info]
            # Preenche os meses
            for mes in range(12):
                col = f'{mes+1:02d}/2024'
                valor = valores[mes] if isinstance(valores[mes], (int, float)) else valores[mes][mes]
                df.at[idx, col] = format_valor(valor, valor < 0)
            
            # Preenche o saldo acumulado
            if info == 'RESULTADO DO EXERCÍCIO':
                df.at[idx, 'Saldo Acumulado'] = format_valor(valores[-1][-1], valores[-1][-1] < 0)
            else:
                soma = sum(valores) if isinstance(valores[0], (int, float)) else sum(v[-1] for v in valores)
                df.at[idx, 'Saldo Acumulado'] = format_valor(soma, soma < 0)
        else:
            # Zera os valores para as outras informações
            for mes in range(12):
                col = f'{mes+1:02d}/2024'
                df.at[idx, col] = '0,00'
            df.at[idx, 'Saldo Acumulado'] = '0,00'
    
    return df

def fill_marina_2024():
    """Preenche os dados de 2024 da Marina."""
    df = create_sheet_2024()
    df['Info'] = create_balancete_structure()
    
    # Dados de exemplo para Marina (valores maiores que Porto Santa Maria)
    data = {
        'ATIVO': [2000000] * 12,
        'PASSIVO': [800000] * 12,
        'PATRIMONIO LIQUIDO': [1200000] * 12,
        'ESTOQUES': [list(range(400000, 520000, 10000))] * 12,
        'COMPENSACOES': [100000] * 12,
        'RECEITA OPERACIONAL BRUTA': [list(range(600000, 840000, 20000))] * 12,
        'DEDUCOES/CUSTOS/DESPESAS': [-360000] * 12,
        'APURACAO DO RESULTADO': [240000] * 12,
        'CONTAS DEVEDORAS': [160000] * 12,
        'CONTAS CREDORAS': [-120000] * 12,
        'RESULTADO DO MES': [40000] * 12,
        'RESULTADO DO EXERCÍCIO': [list(range(40000, 520000, 40000))] * 12
    }
    
    # Preenche o DataFrame (mesmo código do Porto Santa Maria 2024)
    for info in df['Info'].unique():
        if info == '':  # Pula linhas em branco
            continue
            
        idx = df[df['Info'] == info].index[0]
        
        if info in data:
            valores = data[info]
            # Preenche os meses
            for mes in range(12):
                col = f'{mes+1:02d}/2024'
                valor = valores[mes] if isinstance(valores[mes], (int, float)) else valores[mes][mes]
                df.at[idx, col] = format_valor(valor, valor < 0)
            
            # Preenche o saldo acumulado
            if info == 'RESULTADO DO EXERCÍCIO':
                df.at[idx, 'Saldo Acumulado'] = format_valor(valores[-1][-1], valores[-1][-1] < 0)
            else:
                soma = sum(valores) if isinstance(valores[0], (int, float)) else sum(v[-1] for v in valores)
                df.at[idx, 'Saldo Acumulado'] = format_valor(soma, soma < 0)
        else:
            # Zera os valores para as outras informações
            for mes in range(12):
                col = f'{mes+1:02d}/2024'
                df.at[idx, col] = '0,00'
            df.at[idx, 'Saldo Acumulado'] = '0,00'
    
    return df

def fill_marina_2025():
    """Preenche os dados de 2025 da Marina."""
    df = create_sheet_2025()
    df['Info'] = create_balancete_structure()
    
    # Dados de exemplo para Marina 2025 (crescimento em relação a 2024)
    data = {
        'ATIVO': [2200000, 2300000, 2400000],
        'PASSIVO': [880000, 920000, 960000],
        'PATRIMONIO LIQUIDO': [1320000, 1380000, 1440000],
        'ESTOQUES': [520000, 540000, 560000],
        'COMPENSACOES': [110000, 115000, 120000],
        'RECEITA OPERACIONAL BRUTA': [840000, 880000, 920000],
        'DEDUCOES/CUSTOS/DESPESAS': [-396000, -414000, -432000],
        'APURACAO DO RESULTADO': [264000, 276000, 288000],
        'CONTAS DEVEDORAS': [176000, 184000, 192000],
        'CONTAS CREDORAS': [-132000, -138000, -144000],
        'RESULTADO DO MES': [44000, 46000, 48000],
        'RESULTADO DO EXERCÍCIO': [44000, 90000, 138000]
    }
    
    # Preenche o DataFrame
    for info in df['Info'].unique():
        if info == '':  # Pula linhas em branco
            continue
            
        idx = df[df['Info'] == info].index[0]
        
        if info in data:
            valores = data[info]
            # Preenche os meses
            df.at[idx, '01/2025'] = format_valor(valores[0], valores[0] < 0)
            df.at[idx, '02/2025'] = format_valor(valores[1], valores[1] < 0)
            df.at[idx, '03/2025'] = format_valor(valores[2], valores[2] < 0)
            
            # Preenche o saldo acumulado
            if info == 'RESULTADO DO EXERCÍCIO':
                df.at[idx, 'Saldo Acumulado'] = format_valor(valores[2], valores[2] < 0)
            else:
                soma = sum(valores)
                df.at[idx, 'Saldo Acumulado'] = format_valor(soma, soma < 0)
        else:
            # Zera os valores para as outras informações
            df.at[idx, '01/2025'] = '0,00'
            df.at[idx, '02/2025'] = '0,00'
            df.at[idx, '03/2025'] = '0,00'
            df.at[idx, 'Saldo Acumulado'] = '0,00'
    
    return df

def fill_porto_marina_2024():
    """Preenche os dados de 2024 do Porto Marina."""
    df = create_sheet_2024()
    df['Info'] = create_balancete_structure()
    
    # Dados de exemplo para Porto Marina (valores entre Marina e Porto Santa Maria)
    data = {
        'ATIVO': [1500000] * 12,
        'PASSIVO': [600000] * 12,
        'PATRIMONIO LIQUIDO': [900000] * 12,
        'ESTOQUES': [list(range(300000, 390000, 7500))] * 12,
        'COMPENSACOES': [75000] * 12,
        'RECEITA OPERACIONAL BRUTA': [list(range(450000, 630000, 15000))] * 12,
        'DEDUCOES/CUSTOS/DESPESAS': [-270000] * 12,
        'APURACAO DO RESULTADO': [180000] * 12,
        'CONTAS DEVEDORAS': [120000] * 12,
        'CONTAS CREDORAS': [-90000] * 12,
        'RESULTADO DO MES': [30000] * 12,
        'RESULTADO DO EXERCÍCIO': [list(range(30000, 390000, 30000))] * 12
    }
    
    # Preenche o DataFrame (mesmo código do Porto Santa Maria 2024)
    for info in df['Info'].unique():
        if info == '':  # Pula linhas em branco
            continue
            
        idx = df[df['Info'] == info].index[0]
        
        if info in data:
            valores = data[info]
            # Preenche os meses
            for mes in range(12):
                col = f'{mes+1:02d}/2024'
                valor = valores[mes] if isinstance(valores[mes], (int, float)) else valores[mes][mes]
                df.at[idx, col] = format_valor(valor, valor < 0)
            
            # Preenche o saldo acumulado
            if info == 'RESULTADO DO EXERCÍCIO':
                df.at[idx, 'Saldo Acumulado'] = format_valor(valores[-1][-1], valores[-1][-1] < 0)
            else:
                soma = sum(valores) if isinstance(valores[0], (int, float)) else sum(v[-1] for v in valores)
                df.at[idx, 'Saldo Acumulado'] = format_valor(soma, soma < 0)
        else:
            # Zera os valores para as outras informações
            for mes in range(12):
                col = f'{mes+1:02d}/2024'
                df.at[idx, col] = '0,00'
            df.at[idx, 'Saldo Acumulado'] = '0,00'
    
    return df

def fill_porto_marina_2025():
    """Preenche os dados de 2025 do Porto Marina."""
    df = create_sheet_2025()
    df['Info'] = create_balancete_structure()
    
    # Dados de exemplo para Porto Marina 2025 (crescimento em relação a 2024)
    data = {
        'ATIVO': [1650000, 1725000, 1800000],
        'PASSIVO': [660000, 690000, 720000],
        'PATRIMONIO LIQUIDO': [990000, 1035000, 1080000],
        'ESTOQUES': [390000, 405000, 420000],
        'COMPENSACOES': [82500, 86250, 90000],
        'RECEITA OPERACIONAL BRUTA': [630000, 660000, 690000],
        'DEDUCOES/CUSTOS/DESPESAS': [-297000, -310500, -324000],
        'APURACAO DO RESULTADO': [198000, 207000, 216000],
        'CONTAS DEVEDORAS': [132000, 138000, 144000],
        'CONTAS CREDORAS': [-99000, -103500, -108000],
        'RESULTADO DO MES': [33000, 34500, 36000],
        'RESULTADO DO EXERCÍCIO': [33000, 67500, 103500]
    }
    
    # Preenche o DataFrame
    for info in df['Info'].unique():
        if info == '':  # Pula linhas em branco
            continue
            
        idx = df[df['Info'] == info].index[0]
        
        if info in data:
            valores = data[info]
            # Preenche os meses
            df.at[idx, '01/2025'] = format_valor(valores[0], valores[0] < 0)
            df.at[idx, '02/2025'] = format_valor(valores[1], valores[1] < 0)
            df.at[idx, '03/2025'] = format_valor(valores[2], valores[2] < 0)
            
            # Preenche o saldo acumulado
            if info == 'RESULTADO DO EXERCÍCIO':
                df.at[idx, 'Saldo Acumulado'] = format_valor(valores[2], valores[2] < 0)
            else:
                soma = sum(valores)
                df.at[idx, 'Saldo Acumulado'] = format_valor(soma, soma < 0)
        else:
            # Zera os valores para as outras informações
            df.at[idx, '01/2025'] = '0,00'
            df.at[idx, '02/2025'] = '0,00'
            df.at[idx, '03/2025'] = '0,00'
            df.at[idx, 'Saldo Acumulado'] = '0,00'
    
    return df

def main():
    # Lista de empresas e seus arquivos
    empresas = {
        'porto_santa_maria.xlsx': {
            '2024': fill_porto_santa_maria_2024,
            '2025': fill_porto_santa_maria_2025
        },
        'marina.xlsx': {
            '2024': fill_marina_2024,
            '2025': fill_marina_2025
        },
        'porto_marina.xlsx': {
            '2024': fill_porto_marina_2024,
            '2025': fill_porto_marina_2025
        }
    }
    
    # Para cada empresa
    for arquivo, anos in empresas.items():
        print(f"Processando {arquivo}...")
        
        # Para cada ano
        for ano, func_preencher in anos.items():
            # Cria e preenche o DataFrame
            df = func_preencher()
            
            # Salva no arquivo
            save_to_excel(df, arquivo, ano)
        
        print(f"Arquivo {arquivo} atualizado com sucesso!")
    
    print("Todos os arquivos foram atualizados com sucesso!")

if __name__ == "__main__":
    main() 