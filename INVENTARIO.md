# Inventário pelo coletor

Abra a aba **Inventário** no painel ou acesse `/inventario` no mesmo endereço do frontend para uma tela dedicada ao coletor.

Para criar um inventário, preencha **Nome do operador do novo inventário** e clique em **Criar inventário**. O nome é obrigatório, fica salvo no banco e aparece no inventário, na consulta e na aba Resumo da planilha exportada. Inventários antigos mostram “Não informado”. Reinicie a API após esta atualização para adicionar automaticamente a coluna `operador_nome` ao banco e gere o novo build do frontend.

## Vários coletores

Cada aparelho acessa o mesmo endereço e informa seu nome no campo **Nome deste coletor**. O nome fica salvo no navegador e acompanha as novas caixas registradas. Registros antigos permanecem sem identificação de coletor.

Em cada tela, selecione explicitamente o inventário e o pallet. Para contar o mesmo inventário, os coletores selecionam o mesmo inventário e cada um cria ou seleciona seu pallet. Para contagens separadas, selecionam inventários diferentes. A seleção fica salva por aba do navegador e é restaurada ao recarregar; criar inventários ou pallets em outro aparelho não altera essa seleção.

O resumo dos pallets e as caixas são atualizados a cada três segundos, sem tirar o foco do campo de leitura. A coluna de coletores mostra quem já registrou caixas no pallet, não presença online. Os pallets não são reservados exclusivamente: os operadores devem escolher os pallets combinados para sua contagem. Cada leitura leva seu destino explícito e o banco coordena gravações e encerramentos simultâneos. Um inventário só pode ser encerrado depois de todos os seus pallets.

Teste de concorrência no PostgreSQL: defina `INVENTARIO_TEST_POSTGRES=1` e execute os testes abaixo. Eles usam schemas temporários isolados, sem gravar caixas nos inventários operacionais.

1. Clique em **Criar inventário**. Bipe o **Unitizador do novo pallet**, com Enter ou Tab: o número é salvo como texto, preservando zeros à esquerda, e cada leitura cria um novo pallet, inclusive se o unitizador se repetir.
2. Bipe o **Código de barras do endereço**. O sistema identifica e mostra o endereço abreviado; códigos desconhecidos bloqueiam a leitura de caixas até um endereço válido ser informado.
3. Toque no campo **Leitura da caixa** e bipe as etiquetas. Configure o coletor como teclado, com Enter ou Tab ao final da leitura.
4. Cada leitura é enviada ao banco com código, nome do produto, validade e peso. O campo permanece disponível para novas leituras enquanto a fila é salva. Confira a confirmação na tela.
5. Se os dados não forem identificados, use **Revisar leitura**, abra **Completar dados de uma etiqueta não identificada**, preencha os dados e clique em **Adicionar caixa**.
6. Use **Encerrar pallet** ao terminar. Toque em **Bipar novo unitizador** para iniciar outro pallet, ou encerre o inventário após encerrar todos os seus pallets. A criação de outro pallet não encerra o anterior automaticamente.

O cadastro contém os 400 endereços de `Export (25).xlsx`, armazenados em `backend/data/enderecos_inventario.json`, sem depender da pasta Downloads. O código é comparado como texto, removendo apenas espaços e terminadores nas extremidades. Unitizador, endereço abreviado e código do endereço ficam salvos no pallet e acompanham a exportação nas abas Pallets e Caixas. Um endereço confirmado não pode ser substituído nesse pallet.

Ao atualizar uma instalação existente, reinicie a API para adicionar as três colunas ao banco e publique o novo build do frontend. Pallets antigos permanecem consultáveis, com os campos ausentes indicados como não informados; para novas caixas, crie um pallet pelo fluxo de unitizador e endereço.

Os seletores permitem consultar inventários anteriores e suas caixas por pallet. Os números de pallet são sequenciais entre inventários. Registros encerrados ficam disponíveis para consulta e não recebem novas caixas.

A leitura usa o código cadastrado da caixa ou o formato de etiqueta já interpretado pelo sistema. As informações ausentes precisam ser preenchidas; nenhum nome, peso ou validade é inventado. O mesmo código é bloqueado quando já está no inventário, inclusive em outro pallet; pode ser contado novamente em outro inventário. Esse controle pressupõe códigos individuais por caixa.

O inventário não depende do logcat/ADB e não dá baixa no estoque. As tabelas `inventarios`, `inventario_pallets` e `inventario_caixas` são criadas na inicialização da API. Após atualizar uma API já em execução, reinicie-a para disponibilizar as rotas e criar as tabelas.

É necessária conexão com a API. Leituras pendentes ficam apenas na memória da página até a confirmação do banco. Em caso de falha, confira a mensagem e revise a leitura; não feche a página com leituras pendentes ou não confirmadas.

Validação: ` .\.venv\Scripts\python.exe -m unittest backend.test_inventario -v` (dependências em `requirements-dev.txt`). No frontend: `npx tsc --noEmit` e `npm run build`.
