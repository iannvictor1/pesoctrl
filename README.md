#### **📦 PesoCtrl**



Sistema de automação para controle de pesagens industriais com suporte a múltiplas impressoras, processamento em tempo real e interface web.



#### **🚀 Sobre o projeto**



O PesoCtrl é um sistema desenvolvido para automatizar o processo de pesagem em ambientes logísticos/industriais, eliminando tarefas manuais repetitivas e reduzindo erros operacionais.



O sistema monitora arquivos gerados por impressoras de etiquetas, processa esses dados automaticamente e registra todas as informações em planilhas estruturadas, permitindo rastreabilidade e controle em tempo real.



#### **⚙️ Principais funcionalidades**

* 🖨️ Suporte a múltiplas impressoras independentes
* 🚚 Controle de recebimento global (um fluxo para várias impressoras)
* 📦 Controle de pallets por impressora
* 🔄 Função "novo pallet" (fluxo contínuo sem interrupção)
* 👀 Monitoramento automático de arquivos (watchdog)
* 📥 Sistema de fila para processamento seguro
* 🧠 Extração automática de dados das etiquetas:
* Peso (KG)
* Quantidade de etiquetas
* Peso total
* Validade
* Descrição do produto
* 📊 Geração automática de planilhas Excel em tempo real
* 🗂️ Histórico de pesagens e etiquetas
* 🧹 Limpeza automática de arquivos antigos
* 📡 Interface web interativa com Streamlit
* 🟢 Indicadores visuais de status (impressoras, pallets e monitor)



#### 🧠 Arquitetura do sistema



O sistema foi dividido em camadas para garantir organização e escalabilidade:

PesoCtrl/

├── config.py        → Configuração central do sistema

├── monitor.py       → Motor de processamento e automação

├── tela\_controle.py           → Interface web (Streamlit)



##### **🔧 config.py**



Responsável por:



* Configuração das impressoras
* Definição dos caminhos do sistema
* Padronização das estruturas de diretório



##### **⚙️ monitor.py**



Core do sistema:



* Monitoramento de arquivos
* Controle de fila
* Processamento de etiquetas
* Extração de dados via regex
* Escrita em Excel
* Controle de sessões (pallets)
* Controle de recebimento global
* Logs e limpeza automática



##### **🖥️ app.py**



Interface web com:



* Painel operacional em tempo real
* Controle de recebimento e pallets
* Controle do monitor
* Visualização de métricas
* Sidebar com status das impressoras



##### **📂 Estrutura de diretórios**



Cada impressora possui seu próprio ambiente isolado:



Controle\_Pesagem/

├── recebimento\_global.json

├── impressora\_1/

│   ├── entrada\_etiquetas/

│   ├── fila\_processamento/

│   ├── historico\_etiquetas/

│   ├── historico\_pesagens/

│   ├── fila\_descartada/

│   ├── pesagens\_em\_andamento.xlsx

│   ├── monitor\_log.txt

│   ├── sessao\_status.json

│   └── monitor\_pid.json



##### **🔄 Fluxo de funcionamento**



**Recebimento iniciado**

&#x20;       **↓**

**Início do pallet**

&#x20;       **↓**

**Arquivo chega na pasta de entrada**

&#x20;       **↓**

**Sistema captura automaticamente**

&#x20;       **↓**

**Arquivo vai para fila**

&#x20;       **↓**

**Processamento (extração de dados)**

&#x20;       **↓**

**Registro no Excel**

&#x20;       **↓**

**Encerramento do pallet**

&#x20;       **↓**

**Arquivo movido para histórico**



##### **🛠️ Tecnologias utilizadas**

* **Python**
* **Streamlit**
* **Pandas**
* **OpenPyXL**
* **Watchdog**
* **JSON (persistência de estado)**
* **Pathlib / OS / Shutil (manipulação de arquivos)**
* **Regex (extração de dados)**



##### **▶️ Como executar**

**1. Clonar o repositório**



**git clone https://github.com/seuusuario/pesoctrl.git**

**cd pesoctrl**



##### **2. Instalar dependências**



**pip install -r requirements.txt**





##### **3. Executar aplicação**



**streamlit run tela\_controle.py**



##### **📌 Possíveis melhorias futuras**

**🔐 Sistema de autenticação**

**📊 Dashboard com histórico consolidado**

**📤 Exportação de relatórios**

**📡 Integração com banco de dados**

**📲 Alertas automáticos (falhas, erros, paradas)**

**🧠 Machine Learning para análise de padrões de pesagem**



##### **🎯 Objetivo**



**O objetivo do projeto é aumentar a eficiência operacional, reduzir erros humanos e garantir rastreabilidade no processo de pesagem em ambientes com múltiplas impressoras.**



##### **⭐ Destaque**



**Este projeto foi desenvolvido durante estágio, aplicando conceitos reais de:**



**Automação de processos**

**Manipulação de arquivos em tempo real**

**Concorrência e filas**

**Estruturação de sistemas escaláveis**

**Desenvolvimento de interfaces operacionais**



##### **👨‍💻 Autor**

##### 

##### **Desenvolvido por Iann Victor**

