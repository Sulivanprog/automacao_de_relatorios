import pandas as pd
import json
import re
from unidecode import unidecode
from pathlib import Path

def clean_header(header):
    header = unidecode(header)  # Remove acentos
    header = re.sub(r'[^a-zA-Z0-9_ ]', '', header)  # Remove caracteres especiais
    header = header.lower().replace(" ", "_")  # Converte para minúsculas e troca espaço por _
    return header

def try_convert_number(value):
    """Tenta converter strings para int ou float quando possível"""
    try:
        return int(value)
    except ValueError:
        try:
            return float(value)
        except ValueError:
            return value

def convert_timestamps(value):
    """Converte objetos Timestamp para strings no formato desejado"""
    if isinstance(value, pd.Timestamp):
        return value.strftime('%Y-%m-%d %H:%M:%S')  # Formato de data e hora
    return value

def excel_to_json(file_path, file_name=None):
    if file_name is None:
        # Lendo o arquivo Excel mantendo os tipos originais
        df = pd.read_excel(file_path, header=0)
    else:    
        # Lendo o arquivo Excel mantendo os tipos originais
        df = pd.read_excel((file_path / file_name), header=0)
    
    # Removendo colunas e linhas totalmente vazias
    df = df.dropna(how='all').dropna(axis=1, how='all')
    
    # Ajustando os nomes das colunas
    df.columns = [clean_header(str(col)) for col in df.columns]
    
    # Substituindo valores NaN por ""
    df = df.where(pd.notna(df), "")
    
    # Convertendo valores numéricos
    df = df.map(lambda x: try_convert_number(x) if isinstance(x, str) else x)
    
    # Aplicando a conversão de Timestamp para string
    df = df.applymap(convert_timestamps)
    
    # Convertendo para dicionário e depois para lista de registros
    data = df.to_dict(orient='records')
    
    # Obtém o caminho da pasta onde o script está localizado
    main_folder_path = Path(__file__).parent.parent
    data_path = main_folder_path / "data"

    # Criando a pasta de "data", se não existir
    data_directory = data_path
    data_directory.mkdir(parents=True, exist_ok=True)

    # Salvando como JSON
    order_name = data_directory / "pedidos.json"

    # Obtém o caminho da pasta onde o script está localizado
    main_folder_path = Path(__file__).parent
    final_path = main_folder_path / order_name
    
    with open(final_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
    
    #print("Arquivo pedidos.json criado com sucesso!")
    
    return data
