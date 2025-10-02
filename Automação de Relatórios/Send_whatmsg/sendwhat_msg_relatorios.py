import pyautogui
import pandas as pd
import time
import webbrowser
import urllib.parse
from datetime import datetime

# Carregar a planilha Excel
df = pd.read_excel("clientes.xlsx")  # Substitua pelo caminho correto do seu arquivo Excel

# Lista de meses em português
meses_em_portugues = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho", 
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
]

# Função para obter o mês em português
def obter_mes_por_extenso(data):
    mes = data.month - 1  # O mês no objeto datetime é 1 baseado, então subtraímos 1
    return meses_em_portugues[mes]

# Iterar sobre os contatos e enviar as mensagens
for index, row in df.iterrows():
    numero = str(row['Telefone']).strip()  # Recupera o número de telefone e converte para string
    arquivo = row['Arquivo']  # Recupera o arquivo a ser enviado / (Situacional)
    data_relatorio = pd.to_datetime(row['Data'])  # Conversão de 'Data' para datetime
    vencimento = pd.to_datetime(row['Vencimento']).strftime('%d/%m/%Y')  # Formato da data de vencimento

    # Concatenar a mensagem com as informações do Excel
    mensagem = f"""Boa tarde, tudo bem?
Espero que você e sua família estejam bem.
Segue abaixo o relatório dos pedidos finalizados pelo laboratório no mês de {obter_mes_por_extenso(data_relatorio)} de {data_relatorio.year}. O boleto será enviado para o seu e-mail cadastrado, com vencimento para o dia {vencimento}, mas também oferecemos as opções de pagamento via PIX ou cartão de crédito.

- *Pagamento via PIX:* De forma rápida e sem custos adicionais.
        Chave pix: 46696392000176
- *Pagamento via cartão de crédito:* Com a cobrança da taxa do cartão já inclusa, podendo ser parcelado em várias vezes conforme sua preferência.

📌 Acesse agora o relatório: {arquivo}
_(Este link ficará disponível por 15 dias)_

Fico à disposição para qualquer dúvida. 🤝
Atenciosamente,
Sulivan Santos"""

    # Remover a parte ".0" do número, caso exista (numérico como float)
    numero = numero.split(".")[0]  # Isso remove qualquer parte decimal (como o ".0")

    # Ajustar o formato do número de telefone
    if not numero.startswith("+"):
        numero = "+55" + numero.strip()  # Adiciona o código do país (+55 para Brasil)

    # Remover qualquer caractere não numérico, deixando apenas os números
    numero_formatado = ''.join([c for c in numero if c.isdigit()])

    # Garantir que o número esteja no formato internacional completo
    if len(numero_formatado) != 13:  # Considerando o código do país +55 e o número de 9 dígitos
        print(f"Erro: número de telefone {numero} inválido. Verifique o formato.")
        continue

    # Codificar a mensagem para que ela seja válida na URL
    mensagem_codificada = urllib.parse.quote(mensagem)

    # Criar a URL de envio no WhatsApp Web
    url = f"https://web.whatsapp.com/send?phone={numero_formatado}&text={mensagem_codificada}"

    # Abrir a URL de envio
    webbrowser.open(url)
    time.sleep(10)  # Espera o WhatsApp carregar a conversa

    # Simula o pressionamento de 'Enter' para enviar a mensagem
    pyautogui.press('enter')  # Pressiona Enter para enviar a mensagem

    print(f"Mensagem enviada para {numero_formatado}")
    time.sleep(3)  # Pausa de 3 segundos entre os envios
    pyautogui.hotkey('ctrl', 'w')  # Fechar a aba
