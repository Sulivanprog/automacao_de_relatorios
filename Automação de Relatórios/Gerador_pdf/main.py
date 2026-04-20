from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Image
from reportlab.lib.units import cm
from pathlib import Path
import sys
import json
import random
import string

# Funções de utilitário
from utils.leitor_de_excel import excel_to_json
from utils.formatar_pedidos import formatar_pedidos_por_cliente as formatacao


# Função para formatar valores monetários
def format_currency(value):
    """
    Formata valores monetários no padrão brasileiro.
    Exemplo: 1000.50 -> R$ 1.000,50
    """
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# Função para gerar o PDF para cada cliente
def generate_pdf_for_client(client, orders, filename, path_file):
    """
    Gera um relatório PDF para o cliente com os detalhes dos pedidos.
    """
    doc = SimpleDocTemplate(filename, pagesize=A4, topMargin=2*cm)
    styles = {
        "header": ParagraphStyle(name="Header", fontSize=10, fontName="Helvetica-Bold", leading=14, spaceAfter=20),
        "title": ParagraphStyle(name="Bold", fontSize=8, fontName="Helvetica-Bold"),
        "bold": ParagraphStyle(name="Bold", fontSize=7, fontName="Helvetica-Bold"),
        "normal": ParagraphStyle(name="Normal", fontSize=7, wordWrap='CJK')
    }

    content = []

    # Adiciona a imagem no topo esquerdo
    logo_path = path_file / "image" / "logo.png"
    logo = Image(str(logo_path))  # Certifique-se de que logo_path seja convertido para string
    logo.drawHeight = 1*cm
    logo.drawWidth = 3*cm
    logo.x = 2*cm  # Distância da margem esquerda
    logo.y = 29.7*cm  # Distância da margem superior
    content.append(logo)  # Adiciona a imagem no topo do conteúdo

    # Cabeçalho =====================================================================
    header = Paragraph(f"<b>Data de finalização | Cliente:</b> {client}", styles["normal"])
    header_table = Table([
        [header, ""],
        ["", ""],
    ], colWidths=[16*cm, 4*cm])

    # Estilos da tabela de cabeçalho
    header_style = TableStyle([
        ('ALIGN', (0,0), (0,-1), 'LEFT'),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ])
    header_table.setStyle(header_style)
    content.append(header_table)

    # Tabela principal =====================================================================
    table_data = [
        [Paragraph("Nº pedido", styles["bold"]), Paragraph("Cliente", styles["bold"]), 
         Paragraph("Paciente", styles["bold"]), Paragraph("Data de finalização", styles["bold"]),
         Paragraph("Valor bruto", styles["bold"]), Paragraph("Valor liquido", styles["bold"])]
    ]

    table_style = TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
    ])

    # Preenche a tabela com os dados dos pedidos
    for order in orders["pedidos"]:
        if len(table_data) > 1:
            table_style.add('LINEBELOW', (0, len(table_data)-1), (-1, len(table_data)-1), 1, colors.silver)

        main_row = [
            Paragraph(str(order["no_pedido"]), styles["normal"]),
            Paragraph(client, styles["normal"]),
            Paragraph(order["paciente"], styles["normal"]),
            Paragraph(order["data_entrada"], styles["normal"]),
            Paragraph(order["data_final"], styles["normal"]),
            Paragraph(format_currency(order["total"]), styles["bold"])
        ]
        table_data.append(main_row)

        for item in order["items"]:
            qty, desc, val_bruto, val_liquido, tipo, dentes = map(str, item)
            produto = "{}    {}".format(str(qty), desc)
            if tipo == "Serviço":
                if dentes != "":
                    table_data.append([
                        "", 
                        Paragraph(f'{produto} ({dentes})', styles["normal"]), "", "", 
                        Paragraph(format_currency(float(val_bruto)), styles["normal"]), 
                        Paragraph(format_currency(float(val_liquido)), styles["normal"]),
                    ])
                else:
                    table_data.append([
                        "", 
                        Paragraph(produto, styles["normal"]), "", "", 
                        Paragraph(format_currency(float(val_bruto)), styles["normal"]), 
                        Paragraph(format_currency(float(val_liquido)), styles["normal"]),
                    ])
            else:
                table_data.append([
                        "", 
                        Paragraph(produto, styles["normal"]), "", "", 
                        Paragraph(format_currency(float(val_bruto)), styles["normal"]), 
                        Paragraph(format_currency(float(val_liquido)), styles["normal"]),
                    ])
            table_style.add('SPAN', (1, len(table_data)-1), (3, len(table_data)-1))
            table_style.add('TEXTCOLOR', (1, len(table_data)-1), (1, len(table_data)-1), colors.silver)

    table = Table(table_data, colWidths=[1.4*cm, 4*cm, 5.6*cm, 3*cm, 3*cm, 3*cm])
    table.setStyle(table_style)
    content.append(table)

    # Tabela de totais =====================================================================
    totals_table = Table([
        ["", Paragraph("Total", styles["header"]), format_currency(orders["divida_total"])]
    ], colWidths=[16*cm, 1*cm, 3*cm])

    totals_style = TableStyle([
        ('ALIGN', (1,0), (1,-1), 'LEFT'),
        ('ALIGN', (2,0), (2,-1), 'RIGHT'),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ])
    totals_table.setStyle(totals_style)
    content.append(Spacer(1, 15))
    content.append(totals_table)

    doc.build(content)


# Função para gerar string aleatória
def generate_random_string(length=8):
    chars = string.ascii_letters + string.digits
    return ''.join(random.choices(chars, k=length))


# Função para carregar ou criar o arquivo JSON de sufixos aleatórios
def load_or_create_random_suffixes(file_path):
    if file_path.exists():
        with open(file_path, "r", encoding="utf-8") as file:
            return json.load(file)
    else:
        return {}

def save_random_suffixes(file_path, suffixes):
    with open(file_path, "w", encoding="utf-8") as file:
        json.dump(suffixes, file, indent=4, ensure_ascii=False)


# EXECUÇÃO =====================================================================

# Obtém o caminho da pasta onde o script está localizado
main_folder_path = Path(__file__).parent

# Define o diretório base do executável (para uso em pacotes executáveis)
BASE_DIR = Path(sys._MEIPASS) if hasattr(sys, '_MEIPASS') else Path(__file__).parent

# Caminho do arquivo Excel
excel_file_path = BASE_DIR / "relatorio_pedidos_item_detalhado.xlsx"

# Lê os dados do Excel
dict_orders = excel_to_json(excel_file_path)

# Formata os pedidos por cliente
orders_by_client = formatacao(dict_orders)

# Definindo a pasta de destino para os relatórios
relatorio_path = main_folder_path / "Relatorios de clientes"


# Criando a pasta de relatórios, se não existir
output_directory = relatorio_path
output_directory.mkdir(parents=True, exist_ok=True)

data_path = main_folder_path / "data"

# Criando a pasta de "data", se não existir
data_directory = data_path
data_directory.mkdir(parents=True, exist_ok=True)

# Carrega os sufixos aleatórios previamente gerados
suffixes_file_path = data_path / "random_suffixes.json"
random_suffixes = load_or_create_random_suffixes(suffixes_file_path)

# Lista para armazenar os nomes dos arquivos gerados
generated_files = []

# Gerando o PDF para cada cliente
for client, orders in orders_by_client.items():
    # Verifica se o cliente já tem um sufixo aleatório
    if client not in random_suffixes:
        random_suffix = generate_random_string()
        random_suffixes[client] = random_suffix
    else:
        random_suffix = random_suffixes[client]

    filename = relatorio_path / f"{client.replace(' ', '_')}-{random_suffix}-RELATORIO.pdf"
    fullfile = f'{client} | {filename}'

    
    #git_path = "data/files/MARCELO_LIEVORE_DE_BRANDÃO-LYpPoPCK-RELATORIO.pdf"
    git_path = f"{"data/files/"}{client.replace(' ', '_')}-{random_suffix}-RELATORIO.pdf"
    git_file = {"name": "Relatório InoviLab" ,
               "path": git_path }
    generate_pdf_for_client(client, orders, str(filename), main_folder_path)
    #generated_files.append(str(fullfile))

    generated_files.append(git_file)
    print(f"PDF gerado para o cliente {client}")


# Salva os nomes dos arquivos em um JSON
json_file_path = data_path / "arquivos_gerados.json"
with open(json_file_path, "w", encoding="utf-8") as json_file:
    json.dump(generated_files, json_file, indent=4, ensure_ascii=False)

# Salva os sufixos aleatórios no arquivo JSON
save_random_suffixes(suffixes_file_path, random_suffixes)

print(f"Lista de arquivos salva em: {json_file_path}")
