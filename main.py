import argparse
import os
import pathlib
import pandas as pd

def processar_planilha(input_path):
    # Carregar a planilha Excel
    df = pd.read_excel(input_path)

    # Iterar sobre todas as linhas e colunas 4
    for indice, linha in df.iterrows():
        # Realizar o processamento desejado, aqui você pode fazer o que for necessário com o valor da célula
        valor_coluna_4 = linha.iloc[3]
        print(valor_coluna_4)
        # Adicionar uma coluna indicando que a célula foi processada
        df.at[indice, 'Processado'] = 'Jacare'

    # Salvar a planilha com a nova coluna (sobrescrevendo a planilha original)
    df.to_excel(input_path, index=False)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description = ("Le a planilha de chamados"))
    parser.add_argument("-p", "--path", help = "Informe o path/to/directory contendo a planilha", required = True)

    args = parser.parse_args()

    if not pathlib.Path(args.path).is_dir():
        print(f"{args.path} does not exist.")
        exit(1)

    files = [f for f in pathlib.Path(args.path).glob("*.xlsx")]
    if len(files) == 0:
        print("Nao encontrei arquivos para processar [", args.path, "].")
        exit(1)

    latest_file = max(files, key=os.path.getmtime)
    print('Caminho..: ', args.path)
    print("Mais novo: ",latest_file)

    # Substitua 'caminho_entrada.xlsx' pelo caminho do seu arquivo de entrada
    processar_planilha(latest_file)

