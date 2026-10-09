# PesoCtrl no Docker

Uma aplicacao Compose chamada `pesoctrl` inicia PostgreSQL 16, API FastAPI,
frontend React servido pelo Nginx e Streamlit (controle das impressoras).
Cada servico tem seu container, verificacao de saude e reinicio automatico.

## Instalar e iniciar no Windows

Requer Docker Desktop em modo Linux, com Docker Compose v2. Python e Node
nao precisam estar instalados no Windows. Instale tambem o Git para clonar
e atualizar o sistema. Em uma pasta destinada a aplicacao:

```powershell
git clone https://github.com/iannvictor1/pesoctrl.git
cd pesoctrl
Copy-Item .env.example .env
```

Na primeira instalacao, configure `POSTGRES_PASSWORD` no `.env` antes de
inicializar o banco. O `.env` fica somente na maquina, fora do Git.
Depois, na pasta do projeto:

```powershell
.\scripts\preparar_producao.ps1
```

O script constroi as imagens, prepara arquivos de estado ausentes e aguarda
os servicos ficarem saudaveis. Para iniciar novamente sem reconstruir:

```powershell
.\scripts\iniciar_docker_producao.ps1
```

Para instalar a tarefa de login e liberar as portas na rede local, execute
em PowerShell como Administrador, usando o usuario que executa Docker Desktop:

```powershell
.\scripts\criar_tarefas_producao.ps1
```

A tarefa `PesoCtrl-Docker` abre Docker Desktop se necessario, espera o motor
Linux e executa Compose. Ela roda **apos login**, nao antes do login. A tarefa
diaria `PesoCtrl-BackupPostgres` permanece as 23:30; requer usuario conectado.
As tarefas antigas PesoCtrl-API, PesoCtrl-Frontend, PesoCtrl-Postgres,
PesoCtrl-Streamlit e PesoCtrlStreamlit sao desativadas, sem exclusao.
Processos iniciados manualmente nas mesmas portas devem ser encerrados antes.

## Acesso pela rede

* Sistema: `http://IP-DO-SERVIDOR:5174`
* API e documentacao: `http://IP-DO-SERVIDOR:5174/docs`
* Saude (inclui conexao com banco): `http://IP-DO-SERVIDOR:5174/api/health`
* Controle das impressoras: `http://IP-DO-SERVIDOR:8501`

O frontend chama `/api` no mesmo endereco. O Nginx encaminha o IP do cliente
para os endpoints que identificam coletores pelo IP. A API fica na rede interna
do Compose; PostgreSQL continua acessivel no servidor em `127.0.0.1:15432`.
O script de tarefas libera somente as portas web para a sub-rede local,
inclusive quando a rede do Windows esta classificada como Publica.

Portas podem ser definidas em `.env` com `FRONTEND_PORT` e `STREAMLIT_PORT`
(veja `.env.example`). Reexecute o script de tarefas se mudar as portas.
A URL DATABASE_URL de desenvolvimento nao substitui a URL interna do Compose.

## Dados existentes

O volume `pesoctrl_pesoctrl_pgdata` e reutilizado. **Nao execute
`docker compose down -v`**: isso remove o banco. A senha padrao existente foi
mantida; mudar POSTGRES_PASSWORD no `.env` nao muda a senha de um banco ja
inicializado. Para alterar, ajuste tambem o usuario dentro do PostgreSQL.

As pastas `impressora_1`, `impressora_2`, `historico_somador_entregas` e os
arquivos `recebimento_global.json` e `somador_entrega.json` ficam no Windows,
montados em `/data`. Assim as impressoras continuam gravando nas mesmas pastas.
O monitor usa polling no Docker para detectar arquivos nas pastas Windows.
Os botoes do Streamlit continuam iniciando/parando os monitores.

As chaves ADB ficam no volume `pesoctrl_pesoctrl_adb`. Coletores via rede usam
ADB instalado na imagem; pode ser necessario autorizar a nova chave no coletor.
USB conectado ao Windows nao e repassado automaticamente ao container Linux:
o fluxo de restauracao USB exige encaminhamento USB ou ADB no host. Valide os
coletores reais e a deteccao de IP pelo Docker Desktop antes da operacao.

## Manutencao

```powershell
docker compose ps
docker compose logs --tail 100
.\scripts\backup_postgres.ps1
# Reconstruir apos atualizar o codigo:
git pull --ff-only origin main
.\scripts\preparar_producao.ps1
# Parar sem apagar dados:
docker compose down
```

Backups do banco ficam em `backups/postgres`. Copie tambem os arquivos e pastas
de `/data` acima para proteger planilhas, etiquetas e estados. Os scripts antigos
`iniciar_api.ps1` e `iniciar_frontend.ps1` ficam disponiveis para desenvolvimento;
nao sao usados pela nova tarefa. `atualizar_producao.ps1` e o atualizador legado;
para Docker, atualize o codigo e execute `preparar_producao.ps1`.

Referencia: [ordem de inicializacao do Docker Compose](https://docs.docker.com/compose/how-tos/startup-order/).
