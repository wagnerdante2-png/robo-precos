# RoboPrecos

RPA local para coleta dos totalizadores de Auditoria de Precos no sistema PDA.

## Estado atual - v0.2

O nucleo de leitura direta do PDA foi validado em ambiente real.

O robo consegue:

- abrir o Chrome controlado localmente;
- autenticar no PDA;
- abrir diretamente a tela de Auditoria de Preco;
- selecionar loja e periodo;
- pesquisar;
- ler diretamente da pagina:
  - OK;
  - Divergente;
  - Sem etiqueta;
  - Total auditado;
- validar que OK + Divergente + Sem etiqueta = Total;
- descobrir automaticamente as lojas existentes no dropdown Centro;
- percorrer a rede inteira;
- salvar checkpoint apos cada loja;
- retomar uma coleta interrompida;
- refazer login automaticamente se a sessao expirar;
- gerar CSV consolidado da rede.

Nesta etapa o robo ainda nao altera a planilha de controle. A escrita nas abas ETIQUETAS, DIVERGENCIAS e SEM PRECO sera conectada depois da validacao da coleta em massa.

## Premissas

- Windows 10/11
- Google Chrome
- PowerShell nativo do Windows
- sem Python
- sem Selenium
- sem instalacao de bibliotecas
- sem GitHub Actions
- execucao local

## Execucao

Execute:

RoboPrecos.cmd

O menu oferece:

1. Testar uma unica loja
2. Coletar todas as lojas retornadas pelo PDA

No modo rede inteira o robo nao usa uma lista fixa de lojas. Ele le diretamente as opcoes atuais do campo Centro do sistema PDA.

## Checkpoint e retomada

A cada loja concluida o resultado e salvo em:

output/checkpoints/

Se ocorrer queda, timeout, fechamento ou erro em alguma loja, execute novamente o mesmo periodo.

As lojas ja registradas como OK e matematicamente validadas sao preservadas. O robo tenta novamente somente as lojas pendentes ou com erro.

O consolidado fica em:

output/auditoria_rede_YYYYMMDD_YYYYMMDD.csv

## Sessao expirada

Se o PDA retornar para o login durante a coleta, o robo usa a credencial local protegida pelo Windows, refaz a autenticacao, reabre a Auditoria de Preco e repete a loja que estava em processamento.

## Seguranca local

Os arquivos abaixo ficam fora do Git:

- config.precos.json
- data/pda_credential.json
- output/

A credencial PDA e protegida localmente pelo Windows via DPAPI e nao deve ser adicionada manualmente ao repositorio.

## Arquitetura

RoboPrecos.cmd
  -> RoboPrecos.ps1
      -> src/bootstrap.ps1
      -> src/cdp.ps1
      -> src/pda.ps1
      -> src/network.ps1

A base herdada do robo-horas permanece no repositorio para reaproveitamento, mas o fluxo do RoboPrecos usa somente os modulos necessarios.

## Proximas etapas

1. validar a coleta de varias lojas / rede inteira;
2. mapear de forma exata as linhas e colunas da planilha Controle de Auditoria de Precos;
3. escrever Total na aba ETIQUETAS;
4. escrever Divergente na aba DIVERGENCIAS;
5. escrever Sem etiqueta na aba SEM PRECO;
6. preservar formulas, formatacao e historico existentes;
7. integrar posteriormente os dados de descontos vindos do BI.
