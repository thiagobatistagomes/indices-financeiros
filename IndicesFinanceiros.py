import pandas as pd
import matplotlib.pyplot as plt

# Carregar o arquivo Excel Para os Indices
df = pd.read_excel(r'C:\Users\fabio\OneDrive\Documentos\Faculdade\Contabilidade em Informática\Trabalho\teste.xlsx')

def calcularEstruturaDeCapital(df):
    # Calcular a Participação de Capital de Terceiros (PCT) em %
    df['PCT'] = df['Passivo'] / df['Patrimônio Líquido (2023)'] * 100
    
    # Calcular o Endividamento em %
    df['Endividamento'] = df['Passivo Circulante'] / df['Passivo'] * 100
    
    # Calcular a Imobilização (IPL) em %
    df['IPL'] = df['Imobilizado'] / df['Patrimônio Líquido (2023)'] * 100
    
    return df[['Empresa', 'PCT', 'Endividamento', 'IPL']]

def calcularLiquidez(df):
    # Calcular a Liquidez Geral
    df['Liquidez Geral'] = (df['Ativo Circulante'] + df['ARLP']) / df['Passivo']

    # Calcular a Liquidez Corrente
    df['Liquidez Corrente'] = df['Ativo Circulante'] / df['Passivo Circulante']

    # Calcular a Liquidez Seca
    df['Liquidez Seca'] = df['Disponibilidades'] / df['Passivo Circulante']

    return df[['Empresa', 'Liquidez Geral', 'Liquidez Corrente', 'Liquidez Seca']]

def calcularRentabilidade(df):
    # Calcular o Giro do Ativo
    df['Giro Ativo'] = df['Vendas Líquidas'] / df['Total Ativo'] * 100
    
    # Calcular a Margem Líquida em %
    df['Margem Líquida'] = df['Lucro Líquido'] / df['Vendas Líquidas'] * 100

    # Calcular a Rentabilidade do Ativo em %
    df['Rentabilidade do Ativo'] = df['Lucro Líquido'] / df['Total Ativo'] * 100

    # Calcular a Rentabilidade do PL em %
    df['Rentabilidade do PL'] = df['Lucro Líquido'] / df['PL Médio'] * 100

    return df[['Empresa', 'Giro Ativo', 'Margem Líquida', 'Rentabilidade do Ativo', 'Rentabilidade do PL']]


def main():
    # Calcular métricas
    estruturaDeCapital = calcularEstruturaDeCapital(df)
    liquidez = calcularLiquidez(df)
    rentabilidade = calcularRentabilidade(df)

    # Exibir os resultados
    print("Resultados para Vivo (2023):")
    print("=============================")
    print("\nEstrutura De Capital:")
    print(estruturaDeCapital.to_string(index=False))

    print("\nLiquidez:")
    print(liquidez.to_string(index=False))

    print("\nRentabilidade:")
    print(rentabilidade.to_string(index=False))

    # Gráficos de barras
    plt.figure(figsize=(12, 8))

    # Gráfico para Estrutura de Capital
    plt.subplot(2, 2, 1)
    ax1 = estruturaDeCapital.set_index('Empresa').plot(kind='bar', ax=plt.gca())
    plt.title('Estrutura de Capital (%)')
    for p in ax1.patches:
        ax1.annotate(f"{p.get_height():.2f}%", (p.get_x() + p.get_width() / 2., p.get_height()),
                    ha='center', va='center', xytext=(0, 10), textcoords='offset points')

    # Gráfico para Liquidez
    plt.subplot(2, 2, 2)
    ax2 = liquidez.set_index('Empresa').plot(kind='bar', ax=plt.gca())
    plt.title('Liquidez')
    for p in ax2.patches:
        ax2.annotate(f"{p.get_height():.2f}", (p.get_x() + p.get_width() / 2., p.get_height()),
                    ha='center', va='center', xytext=(0, 10), textcoords='offset points')

    # Gráfico para Rentabilidade
    plt.subplot(2, 2, 3)
    ax3 = rentabilidade.set_index('Empresa').plot(kind='bar', ax=plt.gca())
    plt.title('Rentabilidade (%)')
    for p in ax3.patches:
        ax3.annotate(f"{p.get_height():.2f}%", (p.get_x() + p.get_width() / 2., p.get_height()),
                    ha='center', va='center', xytext=(0, 10), textcoords='offset points')

    plt.tight_layout()
    plt.show()

if __name__ == "__main__":
    main()
