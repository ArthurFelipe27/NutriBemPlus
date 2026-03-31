# 🏥 NutriBem+ – Gestão de Dietas e Etiquetas Hospitalares

![GitHub repo size](https://img.shields.io/github/repo-size/ArthurFelipe27/NutriBemPlus?style=for-the-badge)
![GitHub language count](https://img.shields.io/github/languages/count/ArthurFelipe27/NutriBemPlus?style=for-the-badge)
![GitHub last commit](https://img.shields.io/github/last-commit/ArthurFelipe27/NutriBemPlus?style=for-the-badge)
![License](https://img.shields.io/github/license/ArthurFelipe27/NutriBemPlus?style=for-the-badge)

> **NutriBem+** é um sistema desktop moderno integrado com tecnologias web, focado em otimizar o fluxo de trabalho de nutricionistas no ambiente hospitalar. Desenvolvido com **Python e Pandas** no processamento de dados e **HTML, CSS e JavaScript** na interface, o software automatiza a gestão de planilhas, gera painéis analíticos em tempo real e emite etiquetas em PDF prontas para impressão.

---

## ✨ Funcionalidades Principais

O sistema foi arquitetado para lidar com múltiplos setores de um hospital (Enfermarias, UTI, UPA) e atua em três frentes principais:

* 📊 **Dashboard Analítico** Visualização em tempo real da distribuição de dietas através de gráficos inteligentes, com categorização automática e higienização de dados (*Data Cleansing*).

* 🖨️ **Fila de Impressão e Relatórios** Geração automatizada de etiquetas de dieta em matriz (PDF) e emissão de relatórios de ocupação ou "Mapas Gerais" (Auditoria) separados por alas hospitalares.

* 📝 **Gestor de Dados Integrado** Interface estilo *spreadsheet* dentro do próprio aplicativo para edição rápida do banco de dados (Excel) com sugestões de preenchimento para padronizar prescrições.

---

## 🛡️ Segurança e Prevenção de Erros Clínicos

* **Higienização de Dados:** O sistema normaliza os textos de dietas digitadas, removendo acentos e variações textuais para garantir um agrupamento analítico preciso.
* **Sistema de Alertas Visuais:** Algoritmo que lê o campo de "Observações" e destaca pacientes com termos críticos (ex: *alergia*, *broncoaspiração*, *restrição severa*), prevenindo erros médicos.
* **Backups Automáticos:** Antes de qualquer alteração na base de dados, o backend cria um snapshot temporal (`.backups/`) para evitar perda acidental de informações.

---

## 📊 Dashboard Analítico

* Visão geral do setor selecionado (Enfermarias, UTI ou UPA)
* Gráficos dinâmicos de distribuição de dietas (Chart.js)
* Fila de impressão interativa com adição/remoção de pacientes
* Geração de PDFs de Etiquetas e Relatórios de Ocupação

---

## 📝 Gestor de Dados (Editor)

* Leitura e escrita bidirecional em planilhas Excel (Pandas/OpenPyXL)
* Inserção padronizada de dados com `datalist` embutido
* Destaque em tempo real para leitos com alertas clínicos críticos
* Interface de exclusão e adição de novas linhas sem a necessidade de abrir o Microsoft Excel

---

## 💻 Pré-requisitos

Antes de executar o projeto, certifique-se de ter:

* 🐍 **Python 3.8 ou superior**
* 📦 **Pip** (Gerenciador de pacotes do Python)
* 💻 Sistema operacional **Windows, Linux ou macOS** (Windows recomendado para abertura nativa de PDFs)

---

## 🚀 Tecnologias Utilizadas

### ⚙️ Backend (Python)

* 🐍 **Python** — Core da aplicação
* 🌐 **pywebview** — Renderização do frontend web como um aplicativo Desktop
* 🐼 **Pandas & openpyxl** — Processamento de dados e motor ETL (Excel)
* 📄 **ReportLab** — Criação programática de arquivos PDF vetoriais

### 🖥️ Frontend (Web)

* 🧱 **HTML5 & CSS3** — Interface responsiva com suporte a Dark/Light Mode
* ⚡ **JavaScript (ES6+)** — Lógica de interface, higienização e integração com a API Python
* 📊 **Chart.js** — Renderização de gráficos de rosca (*Doughnut*)
* 🔔 **Toastify-JS** — Notificações dinâmicas estilo *toast*

---

## ⚙️ Como Executar o Projeto

### 1️⃣ Clone o repositório

```bash
git clone [https://github.com/ArthurFelipe27/NutriBemPlus.git](https://github.com/ArthurFelipe27/NutriBemPlus.git)
cd NutriBemPlus

### 2️⃣ Instale as dependências
No terminal, execute o comando para instalar as bibliotecas necessárias:

Bash
''pip install -r requirements.txt''

### 3️⃣ Execute a Aplicação
Inicie o controlador principal:

Bash
python main.py
A interface desktop abrirá automaticamente. A base de dados (pacientes.xlsx) e a pasta oculta de backups (.backups/) serão geradas de forma autônoma pelo sistema.

### 📂 Estrutura de Pastas

nutribemplus/
├── .backups/              # Snapshots de segurança do Excel (gerado via código)
├── web/
│   └── index.html         # Frontend unificado (HTML, CSS, JS, Chart.js, Toastify)
├── main.py                # Controller principal da API (Comunicação JS -> Python)
├── excel_service.py       # Camada de serviço responsável pelo motor do Pandas e Backups
├── pdf_service.py         # Camada de serviço dedicada ao motor do ReportLab (PDFs)
├── logger.py              # Utilitário de registro de logs de erro
├── pacientes.xlsx         # Banco de dados simulado (gerado via código)
├── requirements.txt       # Dependências do Python
└── README.md              # Documentação oficial

### 📸 Demonstração
Dashboard Analítico e Alertas Clínicos
<img width="1920" height="1080" alt="Dashboard NutriBem" src="COLOQUE_O_LINK_DA_IMAGEM_AQUI" />

Gestor de Dados (Edição em Tempo Real)
<img width="1920" height="1080" alt="Editor de Planilha" src="COLOQUE_O_LINK_DA_IMAGEM_AQUI" />

Emissão de Etiquetas (PDF)
<img width="1920" height="1080" alt="Exemplo de Etiquetas PDF" src="COLOQUE_O_LINK_DA_IMAGEM_AQUI" />

📌 Status do Projeto
Projeto finalizado e refatorado, aplicando conceitos sólidos de Engenharia de Software como Separation of Concerns (SoC), Data Cleansing e Resiliência de Dados.

🧑‍💻 Autor
Arthur Felipe da Silva Matos 
🔗 LinkedIn: https://www.linkedin.com/in/arthurfelipedasilvamatos

🌐 GitHub: https://github.com/ArthurFelipe27

📝 Licença
Este projeto está licenciado sob a Licença MIT.

Consulte o arquivo LICENSE para mais detalhes.