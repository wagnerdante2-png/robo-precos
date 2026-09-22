# RoboPrecos

RPA local para coleta dos totalizadores de Auditoria de Precos no sistema PDA e gravacao segura na planilha corporativa de controle.

## Estado atual - v0.3

O fluxo operacional cobre coleta e escrita da planilha:

- abre o Chrome controlado localmente;
- autentica no PDA;
- abre diretamente a tela de Auditoria de Preco;
- seleciona loja e periodo;
- coleta OK, Divergente, Sem etiqueta e Total auditado;
- valida que OK + Divergente + Sem etiqueta = Total;
- descobre automaticamente as lojas existentes no campo Centro;
- percorre a rede inteira;
- salva checkpoint apos cada loja;
- retoma uma coleta interrompida;
- refaz login automaticamente se a sessao expirar;
- grava a planilha Controle de Auditoria de Precos;
- preserva as colunas existentes das abas operacionais;
- localiza o mes pela linha 2, sem depender de coluna fixa;
- localiza as lojas pela coluna B, sem depender de linha fixa;
- recalcula a planilha antes da whitelist, permitindo reconhecer novas lojas cadastradas pelas formulas;
- cria backup antes de qualquer gravacao.

## Local da planilha de controle

A planilha operacional nao e mais gravada em Downloads.

Quando o RoboPrecos e executado sozinho, a raiz operacional e a propria pasta do RoboPrecos.

Quando ele e executado pela Plataforma RPA em:

`plataforma-rpa/robots/robo-precos/`

a raiz operacional passa a ser automaticamente:

`plataforma-rpa/`

Portanto, a planilha final permanece ao lado de `central.ps1` e `Iniciar Central.cmd`.

O robo aceita:

`Controle de Auditoria de Precos.xlsx`

ou uma unica variante compativel:

`Controle de Auditoria de Precos*.xlsx`

Isso contempla versoes nomeadas, por exemplo, com sufixos de revisao.

### Migracao automatica

Se nenhuma planilha for encontrada na raiz operacional, o robo procura uma unica copia compativel:

1. na pasta do proprio RoboPrecos, quando ele estiver acoplado a Plataforma;
2. na pasta Downloads do usuario, apenas para compatibilidade com a versao antiga.

Quando encontra uma unica copia antiga, ela e copiada para a raiz operacional e toda gravacao subsequente passa a ocorrer somente na copia da raiz.

Os backups sao criados em:

`RoboPrecos_Backups/`

dentro da mesma raiz onde esta a planilha operacional.

## Compatibilidade com a planilha 2026-2028

O robô nao depende de posicoes fixas para encontrar o mes ou a loja:

- ETIQUETAS: loja na coluna B, meses na linha 2;
- DIVERGENCIAS: loja na coluna B, meses na linha 2;
- SEM PRECO: loja na coluna B, meses na linha 2.

Por isso a expansao ate janeiro de 2028 permanece compativel.

As linhas futuras de lojas podem ser preenchidas por formula. Antes de validar as whitelists, o robo forca um recalculo completo do Excel para materializar uma nova loja cadastrada.

## Premissas

- Windows 10/11;
- Google Chrome;
- Microsoft Excel instalado;
- PowerShell nativo do Windows;
- sem Python;
- sem Selenium;
- sem instalacao de bibliotecas;
- sem GitHub Actions;
- execucao local.

## Execucao

Execute:

`RoboPrecos.cmd`

O menu oferece:

1. testar uma unica loja;
2. coletar a rede e preencher a planilha de controle.

No modo rede inteira, o robo le diretamente as opcoes atuais do campo Centro do PDA.

## Checkpoint e retomada

A cada loja concluida, o resultado de coleta e salvo em:

`output/checkpoints/`

Se ocorrer queda, timeout, fechamento ou erro em alguma loja, execute novamente o mesmo periodo.

As lojas ja registradas como OK e matematicamente validadas sao preservadas. O robo tenta novamente somente as lojas pendentes ou com erro.

O consolidado da coleta fica em:

`output/auditoria_rede_YYYYMMDD_YYYYMMDD.csv`

## Seguranca local

Permanecem fora do Git:

- `config.precos.json`;
- `data/pda_credential.json`;
- `output/`;
- a planilha operacional;
- `RoboPrecos_Backups/`.

A credencial PDA e protegida localmente pelo Windows via DPAPI.

## Arquitetura

`RoboPrecos.cmd`
  -> `RoboPrecos.ps1`
      -> `src/bootstrap.ps1`
      -> `src/cdp.ps1`
      -> `src/pda.ps1`
      -> `src/network.ps1`
      -> `src/control_workbook.ps1`

A base herdada do robo-horas permanece no repositorio para reaproveitamento, mas o fluxo do RoboPrecos usa somente os modulos necessarios.
