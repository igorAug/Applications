from api_bcb import get_dados_bcb
from datetime import datetime

def analisar_selic(data_inicio, data_fim):
    codigo_selic = 11
    df = get_dados_bcb(codigo_selic, data_inicio, data_fim)
    if df is not None and not df.empty:
        print("\n--- Análise da Taxa SELIC ---")
        print(f"Período: {data_inicio} a {data_fim}\n")
        print("Estatísticas:")
        print(df['valor'].describe())
        valor_ultimo = df.iloc[-1]
        print(f"\nÚltimo valor ({valor_ultimo.name.strftime('%d/%m/%Y')}): {valor_ultimo['valor']}%")
        df['media_movel_60d'] = df['valor'].rolling(60).mean()
        print("\nÚltimos 5 registros:")
        print(df.tail(5))
    else:
        print("Erro ao analisar SELIC.")

if __name__ == "__main__":
    hoje = datetime.now().strftime('%d/%m/%Y')
    analisar_selic("01/01/2023", hoje)
