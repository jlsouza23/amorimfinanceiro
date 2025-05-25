from data_loader import FinancialDataLoader
import logging

# Configuração de logging mais detalhado
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

def test_file_loading():
    loader = FinancialDataLoader(".")
    files = loader.get_excel_files()
    
    print("\nArquivos encontrados:")
    for file in files:
        print(f"- {file}")
        
    print("\nTentando carregar cada arquivo:")
    for file in files:
        try:
            movimento_df, balancete_df = loader.load_excel_file(file)
            print(f"\nArquivo {file} carregado com sucesso!")
            if not movimento_df.empty:
                print(f"Movimento shape: {movimento_df.shape}")
            if not balancete_df.empty:
                print(f"Balancete shape: {balancete_df.shape}")
        except Exception as e:
            print(f"\nErro ao carregar {file}:")
            print(str(e))

if __name__ == "__main__":
    test_file_loading() 