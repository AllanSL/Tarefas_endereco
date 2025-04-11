# 📊 Locais - Tarefas

Interface gráfica para processar relatórios de tarefas e gerar rankings por local de origem com base na classificação de giro dos produtos.

---

## 🚀 Funcionalidades

- Seleciona o relatório de tarefas via interface gráfica.
- Faz a junção com o arquivo `GIRO.xlsx`, que classifica os produtos por giro (`a1`, `a2`, etc.).
- Gera um ranking por local de origem com base na quantidade de tarefas.
- Exporta um arquivo Excel com o resultado já formatado.

---

## 📄 Arquivo `GIRO.xlsx`

Este projeto depende de um arquivo chamado `GIRO.xlsx`, que contém informações sobre os produtos e suas respectivas classificações de giro. Por motivos de confidencialidade, o arquivo incluído neste repositório contém **apenas o nome das colunas**, sem dados.

### 🛠️ Como preparar o arquivo:

1. **Nome do arquivo:** `GIRO.xlsx`  
2. **Formato:** Excel (`.xlsx`)
3. **Colunas obrigatórias:**
   - `Produto`: Código ou nome do produto.
   - `GIRO`: Classificação do giro (ex: `a1`, `a2`, `c2`, etc.)

### 🧹 Exemplo:

| Produto | GIRO |
|---------|------|
| 12345   | a1   |
| 67890   | c2   |

> 📌 **Importante:** Certifique-se de que a coluna `Produto` contenha apenas strings (sem espaços extras), pois o código trata os dados como `str` e aplica `.strip()` antes de fazer os cruzamentos.

---

## 🖥️ Como usar

1. **Instale os pacotes necessários manualmente:**

```bash
pip install pandas openpyxl
```

> 💡 O `tkinter` já vem incluído por padrão com o Python (a partir da versão 3.8).

2. **Execute o programa:**

```bash
python main.py
```

3. **Selecione o arquivo de relatório** (arquivo Excel com as tarefas).
4. O arquivo de saída será salvo com o nome `locais_tarefas_DD_MM_AA.xlsx` e será aberto automaticamente.

---

## 🧱 Estrutura do Projeto

```
locais-tarefas/
├── main.py              # Interface gráfica (Tkinter)
├── processamento.py     # Lógica de processamento dos arquivos Excel
├── utils.py             # Função utilitária para acesso ao arquivo empacotado
├── GIRO.xlsx            # Arquivo auxiliar com classificação de produtos
└── README.md
```

---

## 📦 Gerar executável (opcional)

Se desejar empacotar este projeto como executável, siga os passos abaixo:

1. **Instale o pyinstaller:**

```bash
pip install pyinstaller
```

2. **Execute o comando abaixo para gerar o `.exe`:**

```bash
pyinstaller --noconfirm --onefile --windowed --add-data "GIRO.xlsx;." main.py
```

- O parâmetro `--add-data` garante que o arquivo `GIRO.xlsx` seja incluído no executável.
- O executável final estará na pasta `dist/`.

> 💡 Dica: o projeto já está preparado para rodar corretamente mesmo após empacotado (graças à função `resource_path()` em `utils.py`).

---

## 🧪 Requisitos

- Python 3.8+
- Pandas
- openpyxl
- tkinter (incluído no Python padrão)

---

