# Spreadsheet Tab Extractor

Uma ferramenta em Python para gerar automaticamente arquivos Excel individuais a partir das abas de uma planilha consolidada.

## Sobre o Projeto

O Spreadsheet Tab Extractor foi desenvolvido para automatizar a separação de planilhas que contêm múltiplas abas em arquivos independentes.

A aplicação permite que usuários selecionem um arquivo Excel contendo diversas abas e, com apenas alguns cliques, gerem um arquivo separado para cada aba encontrada.

A solução preserva a estrutura, formatação, fórmulas, validações e demais elementos existentes nas abas originais.

---

## Funcionalidades

- Interface gráfica simples e intuitiva.
- Seleção de um arquivo Excel de origem.
- Geração automática de um arquivo para cada aba.
- Preservação da estrutura original da planilha.
- Exclusão automática das demais abas em cada arquivo gerado.
- Possibilidade de ignorar abas específicas.
- Processamento rápido de dezenas de abas simultaneamente.

---

## Tecnologias Utilizadas

- Python
- OpenPyXL
- Tkinter

---

## Requisitos

- Python 3.10 ou superior

---

## Instalação

Clone o repositório:

```bash
git clone https://github.com/SEU-USUARIO/spreadsheet-tab-extractor.git
```

Entre na pasta:

```bash
cd spreadsheet-tab-extractor
```

Instale as dependências:

```bash
pip install -r requirements.txt
```

---

## Arquivo requirements.txt

```txt
openpyxl
```

---

## Como Utilizar

### 1. Executar a aplicação

```bash
python create_files_for_each_spreadsheet_tab.py
```

### 2. Selecionar a planilha consolidada

Clique em:

```text
📎 Anexar
```

e escolha o arquivo Excel.

### 3. Escolher a pasta de saída

Selecione o diretório onde os arquivos gerados serão armazenados.

### 4. Iniciar o processamento

Clique em:

```text
▶️ Processar
```

### 5. Aguardar a conclusão

A aplicação criará automaticamente um arquivo Excel para cada aba da planilha selecionada.

---

## Exemplo

### Entrada

```text
Workbook.xlsx
│
├── Dashboard
├── Report A
├── Report B
├── Report C
└── Template
```

### Saída

```text
Report A.xlsx
Report B.xlsx
Report C.xlsx
```

---

## Benefícios

- Redução de trabalho manual repetitivo.
- Maior produtividade.
- Menor risco de erro humano.
- Padronização do processo.
- Facilidade de utilização por usuários não técnicos.

---

## Casos de Uso

- Consolidação de relatórios.
- Distribuição de arquivos por equipe.
- Separação de dados por área.
- Processamento de planilhas administrativas.
- Organização de planilhas com múltiplas abas.

---

## Estrutura do Projeto

```text
excel-spreadsheet-tab-extractor-python/
│
├── create_files_for_each_spreadsheet_tab.py
├── requirements.txt
└── README.md
```

---

## Melhorias Futuras

- Barra de progresso.
- Logs detalhados de execução.
- Exportação para outros formatos.
- Configuração de abas ignoradas via interface.
- Empacotamento multiplataforma.

---

## Licença

Este projeto está disponível sob a licença MIT.

---

## Autor

Felipe Alves Bueno
