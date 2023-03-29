import pandas as pd
import pathlib
import datetime
from PIL import Image

# Verificar se o arquivo de texto existe
file_path = 'Primeira quinzena.txt'
if not pathlib.Path(file_path).is_file():
    print(f'O arquivo {file_path} não existe!')
    exit()

# Ler os caminhos de arquivo do arquivo de texto
with open(file_path, 'r') as f:
    file_paths = f.read().splitlines()

data_frames = [] # criar uma lista vazia para armazenar os DataFrames

for file_path in file_paths:
    # Verificar se o caminho de arquivo é válido
    if not pathlib.Path(file_path).is_file():
        print(f'O arquivo {file_path} não existe!')
        continue

    # Ler o arquivo Excel e adicionar o DataFrame à lista
    df = pd.read_excel(file_path, sheet_name=None, usecols='A:D', skiprows=1, nrows=12)
    data_frames.append(df)

data = {}       
for df in data_frames:
    for sheet_name in df.keys():
        if sheet_name in data:
            data[sheet_name].append(df[sheet_name])
        else:
            data[sheet_name] = [df[sheet_name]]

# Comparar os dados das abas com o mesmo nome nos arquivos Excel
for sheet_name, data_list in data.items():
    for i, data1 in enumerate(data_list):
        for data2 in data_list[i+1:]:
            if not data1.equals(data2):
                print(f"As abas '{sheet_name}' dos arquivos são diferentes.")
            else:
                # Comparar os dados das colunas 'DATA', 'CAPITAL', 'GRANDE SP' e 'NOVAS'
                dados_comparados = {
                    'DATA': data1['DATA'].tolist(),
                    'CAPITAL': data1['CAPITAL'].tolist(),
                    'GRANDE SP': data1['GRANDE SP'].tolist(),
                    'NOVAS': data1['NOVAS'].tolist()
                }

                # abrir a pasta de trabalho e a aba "Demonstrativo"
                workbook = openpyxl.load_workbook('Fatura.xlsx')
                sheet = workbook['Demonstrativo']

                # gerar identificador único de fatura
                fatura_id = uuid.uuid4().hex[:7]
                sheet['C9'] = fatura_id
                            
                # nome do cliente
                sheet['C11'] = sheet_name

                # colar os dados
                for col, col_dados in zip(['B', 'C', 'D', 'E'], dados_comparados.values()):
                    for row, dados in enumerate(col_dados, start=15):
                        sheet[f'{col}{row}'] = dados
                
                # data atual
                data_atual = datetime.date.today()
                dtf = data_atual.strftime('%d-%m-%Y')
                sheet['C10'] = dtf

                # carrega a imagem
                img = Image.open('Color.jpg')

                # adiciona a imagem à planilha, ancorada na célula G2
                width, height = img.size
                sheet.column_dimensions['G'].width = width/6.6
                sheet.row_dimensions[2].height = height/6.6
                sheet.add_picture('Color.jpg', 'G2')

                # criar a pasta se ela não existir
                pasta = f'Faturas-{dtf}'
                if not os.path.exists(pasta):
                    os.makedirs(pasta)

                # criar o caminho completo do arquivo
                arquivo = f'{pasta}/Fatura {sheet_name}-{dtf}.xlsx'

                # salvar o arquivo
                try:
                    workbook.save(arquivo)
                    print(f'Fatura_{sheet_name}.xlsx salva!')
                except Exception as e:
                    print(f'Erro ao salvar o arquivo Fatura_{sheet_name}.xlsx: {e}')

