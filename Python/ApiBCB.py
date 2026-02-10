import requests
import pandas as pd

def get_dados_bcb(codigo_serie, data_inicio, data_fim):
    print(f"Buscando dados da série {codigo_serie}...")
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{codigo_serie}/dados?formato=json&dataInicial={data_inicio}&dataFinal={data_fim}"
    try:
        r = requests.get(url)
        r.raise_for_status()
        dados = r.json()
        if not dados:
            print("Nenhum dado encontrado.")
            return None
        df = pd.DataFrame(dados)
        df['data'] = pd.to_datetime(df['data'], format='%d/%m/%Y')
        df['valor'] = pd.to_numeric(df['valor'])
        df.set_index('data', inplace=True)
        print("Dados obtidos com sucesso.")
        return df
    except Exception as e:
        print(f"Erro: {e}")
        return None
