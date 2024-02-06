from decimal import Decimal
import pandas as pd
import os

# Verificar o diretório atual
print("Diretório Atual:", os.getcwd())

# Mude o diretório de trabalho para o diretório do script
os.chdir(os.path.dirname(os.path.abspath(__file__)))

def calcular_apostas(valor_total, odd_aposta1, odd_aposta2):
    valor_total = Decimal(str(valor_total))  # Converta para Decimal
    odd_aposta1 = Decimal(str(odd_aposta1))  # Converta para Decimal
    odd_aposta2 = Decimal(str(odd_aposta2))  # Converta para Decimal

    melhor_retorno = Decimal('0.00')
    melhor_aposta1 = Decimal('0.00')
    melhor_aposta2 = Decimal('0.00')

    for centavos_aposta1 in range(1, int(valor_total * 100)):
        aposta1 = Decimal(centavos_aposta1) / Decimal(100)
        aposta2 = valor_total - aposta1

        retorno_aposta1 = aposta1 * odd_aposta1
        retorno_aposta2 = aposta2 * odd_aposta2

        retorno_total = min(retorno_aposta1, retorno_aposta2)
        if retorno_total > melhor_retorno:
            melhor_retorno = retorno_total
            melhor_aposta1 = aposta1
            melhor_aposta2 = aposta2

    if melhor_retorno > valor_total:
        return melhor_aposta1, melhor_aposta2
    else:
        return None

def processar_excel(entrada_excel, saida_excel):
    try:
        df = pd.read_excel(entrada_excel)
    except FileNotFoundError as e:
        print(f"Erro: {e}")
        return

    resultados = []
    for index, row in df.iterrows():
        resultado = calcular_apostas(row.iloc[0], row.iloc[1], row.iloc[2])
        resultados.append(resultado)

    df['aposta1'] = [f"Aposta 1: Apostei {res[0]:.2f}" if res else f"Não é possível realizar as apostas para obter um retorno acima de {row.iloc[0]}" for res in resultados]
    df['aposta2'] = [f"Aposta 2: Apostei {res[1]:.2f}" if res else None for res in resultados]

    df.to_excel(saida_excel, index=False)
    print(f"Resultados escritos em {saida_excel}.")

# Substitua 'ODDS.xlsx' e 'saida.xlsx' pelos nomes reais dos seus arquivos Excel
caminho_arquivo_odds = r'C:\Users\Matheus A Aleixo\Documents\WORKSPACE_PYTHON\script_aposta_excel\ODDS2.xlsx'
caminho_arquivo_saida = r'C:\Users\Matheus A Aleixo\Documents\WORKSPACE_PYTHON\script_aposta_excel\Resultado_ODDS.xlsx'
processar_excel(caminho_arquivo_odds, caminho_arquivo_saida)
