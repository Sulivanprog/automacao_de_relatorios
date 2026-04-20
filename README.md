
```markdown
# 📊 Automação de Relatórios em PDF com Python

Este projeto realiza a automação da geração de relatórios em PDF a partir de dados extraídos de uma planilha Excel, organizando as informações por cliente e gerando documentos personalizados automaticamente.

---

🚀 Objetivo

O objetivo do sistema é reduzir o trabalho manual na criação de relatórios, automatizando:

- Leitura de dados de planilhas Excel
- Processamento e organização de pedidos por cliente
- Geração de relatórios individuais em PDF
- Controle e rastreamento dos arquivos gerados

---

⚙️ Tecnologias utilizadas

- Python 3
- ReportLab (geração de PDF)
- JSON (armazenamento de dados auxiliares)
- Pathlib (manipulação de arquivos e diretórios)
- Estruturação de dados com listas e dicionários

---

🧠 Funcionalidades

- Conversão de dados de Excel para estrutura JSON
- Agrupamento de pedidos por cliente
- Geração automática de relatórios personalizados em PDF
- Organização de dados de pedidos (itens, valores e serviços)
- Criação de nomes únicos de arquivos para evitar duplicidade
- Persistência de controle de arquivos gerados via JSON

---

📂 Estrutura do projeto

```

automacao_de_relatorios/
│
├── utils/
│   ├── leitor_de_excel.py
│   ├── formatar_pedidos.py
│
├── data/
│   ├── arquivos_gerados.json
│   ├── random_suffixes.json
│
├── image/
│   └── logo.png
│
├── Relatorios de clientes/
│
├── main.py
└── README.md

````

---

▶️ Como executar o projeto

1. Instale as dependências necessárias:

```bash
pip install reportlab
````

2. Execute o script principal:

```bash
python main.py
```

---

📌 Como funciona

1. O sistema lê uma planilha Excel com pedidos.
2. Os dados são organizados por cliente.
3. Para cada cliente, um relatório em PDF é gerado automaticamente.
4. Os arquivos são salvos com nomes únicos.
5. Um arquivo JSON registra todos os relatórios gerados.

---

💡 Exemplo de uso

Ao executar o sistema, será gerado automaticamente:

* Um PDF individual para cada cliente
* Relatórios contendo:

  * Dados do cliente
  * Lista de pedidos
  * Valores brutos e líquidos
  * Resumo final

---

📈 Aprendizados

Este projeto reforça habilidades em:

* Automação de processos
* Manipulação de dados estruturados
* Geração de documentos dinâmicos
* Organização de lógica de negócio em Python
* Trabalho com arquivos e persistência de dados

---

👨‍💻 Autor

Sulivan Batista dos Santos

```
