import pandas as pd


def formatar_numero(valor):
    # Verificar se o valor é numérico
    if isinstance(valor, (int, float)):
        # Converter números para o formato brasileiro
        return '{:,.2f}'.format(valor).replace('.', '*').replace(',', '.').replace('*', ',')
    else:
        return valor

def transformar_planilha(input_file, output_file):
    # Carregar a planilha Excel em um DataFrame usando pandas
    df = pd.read_excel(input_file)
    
    # Extrair apenas a primeira palavra da coluna 1 (Coluna A)
    df['Produto'] = df['Produto'].str.split().str[0]
    df['Instituição'] = df['Instituição'].str.split().str[0]

    df.loc[df['Tipo de Evento'].str.lower() == 'juros sobre capital próprio', 'Tipo de Evento'] = 'JSCP'
    df.loc[df['Tipo de Evento'].str.lower() == 'Reembolso - dividendo', 'Tipo de Evento'] = 'REEMBOLSO(DIV)'
    df.loc[df['Tipo de Evento'].str.lower() == 'Reembolso - juros sobre capital próprio', 'Tipo de Evento'] = 'REEMBOLSO(JSCP)'    
    df.loc[df['Tipo de Evento'].str.lower() == 'rendimento', 'Tipo de Evento'] = 'RENDIMENTO'    
    df.loc[df['Tipo de Evento'].str.lower() == 'dividendo', 'Tipo de Evento'] = 'DIVIDENDO' 

    # Excluir as colunas 5 e 6 (Coluna E e Coluna F)
    df = df.drop(columns=['Quantidade', 'Preço unitário'])

    df['Instituição'], df['Valor líquido'] = df['Valor líquido'], df['Instituição']

    df.insert(4, 'Coluna 5', '')
    df['Coluna 6'] = 'BRL'

    df['Valor líquido'], df['Coluna 6'] = df['Coluna 6'], df['Valor líquido']


    df['Instituição'] = df['Instituição'].apply(formatar_numero)

    df.rename(columns={'Produto': 'ativo'}, inplace=True)
    df.rename(columns={'Pagamento': 'date'}, inplace=True)
    df.rename(columns={'Tipo de Evento': 'evento'}, inplace=True)
    df.rename(columns={'Instituição': 'valor'}, inplace=True)
    df.rename(columns={'Coluna 5': 'irrf'}, inplace=True)
    df.rename(columns={'Valor líquido': 'moeda'}, inplace=True)
    df.rename(columns={'Coluna 6': 'corretora'}, inplace=True)
    


    df = df.drop(df.tail(3).index)

    # Salvar a planilha transformada em um novo arquivo
    #df.to_excel(output_file, index=False)
    df.to_csv(output_file, index=False, sep=';', encoding='utf-8')


if __name__ == "__main__":
    # Defina o nome do arquivo de entrada e saída
    arquivo_entrada = './proventos.xlsx'
    arquivo_saida = './prov.csv'
    
    # Chame a função para realizar a transformação
    transformar_planilha(arquivo_entrada, arquivo_saida)
