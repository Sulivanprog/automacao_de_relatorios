import pyautogui
import pandas as pd
import time
import webbrowser
import urllib.parse

# Carregar a planilha Excel
df = pd.read_excel("todos_clientes.xlsx")  # Substitua pelo caminho correto do seu arquivo Excel

# Abra o WhatsApp Web no navegador
#webbrowser.open("https://web.whatsapp.com/")
#time.sleep(10)  # Aguardar o WhatsApp Web carregar (ajuste se necessário)

# Iterar sobre os contatos e enviar as mensagens
for index, row in df.iterrows():
    numero = str(row['Telefone']).strip()  # Recupera o número de telefone e converte para string
    mensagem = row['Mensagem']  # Recupera a mensagem a ser enviada

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
    time.sleep(3)  # Pausa de 5 segundos entre os envios
    pyautogui.hotkey('ctrl', 'w')  # Fechar a aba
    
