import json
import re
from collections import defaultdict
from pathlib import Path

def formatar_pedidos_por_cliente(pedidos):
    """
    Organiza os pedidos por cliente e calcula a dívida total de cada cliente.
    Remove qualquer conteúdo dentro de parênteses no campo 'paciente'.
    
    :param pedidos: Lista de dicionários contendo os pedidos.
    :return: Dicionário estruturado por cliente, contendo seus pedidos e dívida total.
    """
    
    # Dicionário aninhado para armazenar pedidos organizados por cliente
    pedidos_por_cliente = defaultdict(lambda: defaultdict(lambda: {
        "no_pedido": "", "paciente": "", "data_entrada": "", "data_final": "", "items": [], "total": 0.00
    }))
    
    # Itera sobre todos os pedidos
    for pedido in pedidos:
        cliente = pedido["cliente"]  # Nome do cliente
        no_pedido = pedido["no_pedido"]  # Número do pedido

        # Remove qualquer conteúdo entre parênteses no campo 'paciente'
        paciente_limpo = re.sub(r"\s*\(.*?\)", "", pedido["paciente"])

        # Se o pedido ainda não foi registrado, inicializa as informações
        if not pedidos_por_cliente[cliente][no_pedido]["paciente"]:
            pedidos_por_cliente[cliente][no_pedido]["no_pedido"] = pedido["no_pedido"]
            pedidos_por_cliente[cliente][no_pedido]["paciente"] = paciente_limpo
            pedidos_por_cliente[cliente][no_pedido]["data_entrada"] = pedido.get("data_entrada", "")
            pedidos_por_cliente[cliente][no_pedido]["data_final"] = pedido.get("data_final", "")
        
        # Adiciona os detalhes do item ao pedido
        item = (
            pedido["quantidade"], 
            #pedido["qtd"], 
            
            pedido["servicoproduto"], 
            
            pedido["valor_bruto"],
            #pedido["valor_unitario"],
            
            pedido["valor_liquido"], 
            #pedido["valor_total"], 
            
            pedido["tipo"], 
            pedido["dentes"] 

        )
        pedidos_por_cliente[cliente][no_pedido]["items"].append(item)
        pedidos_por_cliente[cliente][no_pedido]["total"] += pedido["valor_liquido"]
        #pedidos_por_cliente[cliente][no_pedido]["total"] += pedido["valor_total"]
    
    # Constrói a estrutura final do dicionário, incluindo a dívida total
    pedidos_final = {}
    for cliente, pedidos in pedidos_por_cliente.items():
        divida_total = round(sum(pedido["total"] for pedido in pedidos.values()), 2)
        pedidos_final[cliente] = {"pedidos": list(pedidos.values()), "divida_total": divida_total}


    # Obtém o caminho da pasta onde o script está localizado
    main_folder_path = Path(__file__).parent.parent
    data_path = main_folder_path / "data"

    # Criando a pasta de "data", se não existir
    data_directory = data_path
    data_directory.mkdir(parents=True, exist_ok=True)

    # Salvando os pedidos em um arquivo JSON
    json_orders_file_path = data_directory / "orders.json"
    with open(json_orders_file_path, "w", encoding="utf-8") as json_file:
        json.dump(pedidos_final, json_file, indent=4, ensure_ascii=False)
    return pedidos_final

#print("Arquivo 'pedidos_por_cliente.json' gerado com sucesso!")
