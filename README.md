# 📄 DocField Mapper & Comparator

> **Automação Inteligente para Mapeamento e Comparação de Documentos Jurídicos (DOCX/ODT)**

Este projeto é uma ferramenta avançada desenvolvida em Python para analisar documentos estruturados (como Termos de Referência e Contratos). Ele resolve dois problemas complexos que bibliotecas padrões falham em resolver:
1.  **Extração de Índices Reais (1.1, 1.2.1):** Utiliza o motor do LibreOffice para calcular a numeração automática de tópicos.
2.  **Identificação de Campos Opcionais:** Realiza uma varredura profunda no XML do documento para identificar campos marcados com cores (Vermelho, Destaque, Estilos de Ênfase), classificando-os como "Escolha/Opcional".

## 🚀 Funcionalidades

* **Híbrido e Robusto:** Combina a precisão visual do LibreOffice com a análise de dados do Python.
* **Scanner XML Profundo:** Detecta cores e estilos de destaque mesmo quando ocultos em *Smart Tags*, *Links* ou estilos customizados do Word.
* **Classificação Automática:**
    * **Preenchimento:** Identifica padrões `[...]`, `XX`, `<...>`, `(...)`.
    * **Alternativa:** Identifica termos como `OU`.
    * **Escolha:** Classifica baseado na cor da fonte (Preto = Obrigatório / Colorido = Opcional).
* **Extração de Notas:** Captura comentários inseridos no Word e os vincula ao texto.
* **Modo Comparação:** Gera um relatório "De/Para" cruzando dois documentos e apontando o que foi mantido, adicionado ou removido.
* **Saída Organizada:** Gera planilhas Excel prontas para análise.

---

## 🛠️ Pré-requisitos do Sistema

Para que o script funcione, você precisa de dois softwares instalados no seu computador:

1.  **Python 3.8+**
2.  **LibreOffice** (Obrigatório para o cálculo dos índices `1.1`, `1.2`).

---

## 💻 Guia de Instalação e Configuração

### 🐧 No Linux (Ubuntu, Kali, Debian)

O Linux é o ambiente nativo recomendado para este script.

1.  **Atualize o sistema e instale o Python/Pip:**
    ```bash
    sudo apt update
    sudo apt install python3 python3-pip -y
    ```

2.  **Instale o LibreOffice:**
    O script usa o comando `soffice` no terminal.
    ```bash
    sudo apt install libreoffice -y
    ```

3.  **Instale as bibliotecas Python necessárias:**
    Navegue até a pasta do projeto e execute:
    ```bash
    pip3 install pandas python-docx lxml odfpy openpyxl
    ```
    *(Ou, se tiver o arquivo requirements.txt: `pip3 install -r requirements.txt`)*

---

### 🪟 No Windows

O Windows requer um passo extra importante: adicionar o LibreOffice às Variáveis de Ambiente (PATH).

1.  **Instale o Python:**
    * Baixe em [python.org](https://www.python.org/downloads/).
    * ⚠️ **Importante:** Na tela de instalação, marque a caixinha **"Add Python to PATH"**.

2.  **Instale o LibreOffice:**
    * Baixe e instale a versão mais recente em [libreoffice.org](https://www.libreoffice.org/).

3.  **Configurar o PATH (Passo Crítico):**
    Para que o Python consiga "chamar" o LibreOffice, o Windows precisa saber onde ele está.
    * Abra o menu Iniciar e digite **"Editar as variáveis de ambiente do sistema"**.
    * Clique em **Variáveis de Ambiente**.
    * Em "Variáveis do sistema" (parte de baixo), encontre a linha **Path** e clique em **Editar**.
    * Clique em **Novo** e cole o caminho onde o LibreOffice foi instalado. Geralmente é:
        `C:\Program Files\LibreOffice\program`
    * Clique em OK em tudo e reinicie o seu terminal (CMD ou PowerShell).

4.  **Instale as bibliotecas Python:**
    Abra o CMD ou PowerShell na pasta do projeto e rode:
    ```powershell
    pip install pandas python-docx lxml odfpy openpyxl
    ```

---

## 📂 Estrutura de Arquivos

Para o script funcionar corretamente, mantenha a seguinte organização:

```text
📁 /pasta-do-projeto
│
├── 📜 mapeador.py          # O script principal
├── 📜 requirements.txt     # Lista de dependências
├── 📜 README.md            # Este arquivo
│
├── 📄 contrato_base.docx   # Seu documento (Coloque aqui!)
└── 📄 contrato_novo.docx   # Outro documento (Coloque aqui!)
```

## ▶️ Como Usar

1.  **Abra o terminal** na pasta do projeto.

2.  **Execute o script:**
    * **Linux/Mac:**
        ```bash
        python3 mapeador.py
        ```
    * **Windows:**
        ```bash
        python mapeador.py
        ```

3.  **Siga o Menu Interativo:**
    * O script listará os arquivos encontrados. Digite o número do **Documento Principal**.
    * Ele perguntará: `Comparar com outro arquivo? (S/N)`.
        * Digite **S** para selecionar um segundo arquivo e gerar um comparativo cruzado.
        * Digite **N** para apenas mapear os campos do arquivo principal.

4.  **Verifique o Resultado:**
    * Um arquivo Excel será gerado na mesma pasta, nomeado como `Mapeamento_NomeDoArquivo.xlsx`.
    * O script cria e apaga automaticamente arquivos `.txt` temporários durante o processo.

---

## ❓ Solução de Problemas Comuns

| Problema | Causa Provável | Solução |
| :--- | :--- | :--- |
| **Erro: "LibreOffice falhou"** | O LibreOffice não está instalado ou não está no PATH. | **Windows:** Verifique se `C:\Program Files\LibreOffice\program` está no PATH.<br>**Linux:** Rode `sudo apt install libreoffice`. |
| **Erro: "Permission denied" ao salvar Excel** | O arquivo Excel gerado anteriormente está aberto. | Feche o arquivo Excel no seu computador e tente rodar o script novamente. |
| **Índices aparecem vazios** | O arquivo pode estar corrompido ou protegido. | Abra o arquivo no Word, clique em "Salvar Como" e salve uma nova cópia limpa. |
| **Cores não detectadas** | O texto usa um estilo complexo não mapeado. | O script atual usa uma varredura XML profunda ("Qualquer coisa que não seja preto é cor"). Verifique se o texto não está realmente preto (Automático). |

---

## 🧠 Entendendo a Lógica (Para Desenvolvedores)

Se você deseja modificar o código, aqui está como ele "pensa":

1.  **Normalização:** O script primeiro converte o `.docx` para `.txt` usando o LibreOffice em modo *headless*. Isso força a renderização dos números de lista (ex: transforma a lista automática do Word em texto puro "1.1 Objeto").
2.  **Mapeamento:** Ele lê esse TXT e cria um mapa: `{'Texto do Parágrafo': '1.1'}`.
3.  **Análise de Metadados:** Em seguida, ele usa a biblioteca `lxml` para ler a estrutura profunda do `.docx` original. Ele procura tags `<w:color>`, `<w:highlight>` ou `<w:shd>` dentro dos parágrafos.
4.  **Fusão:** Por fim, ele cruza os dados: pega o índice descoberto no passo 1 e combina com as cores/comentários descobertos no passo 3.
