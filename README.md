# RoboPrecos

RPA local para coleta da Auditoria de Precos no PDA e preenchimento automatizado dos descontos de **PRECO ERRADO** pelo Power BI.

## Estado atual - v0.4

Fluxo principal:

1. recebe o periodo no CMD;
2. coleta e valida a Auditoria de Precos no PDA;
3. grava ETIQUETAS, DIVERGENCIAS e SEM PRECO;
4. abre o Power BI em perfil dedicado;
5. decide automaticamente qual fonte de descontos usar;
6. coleta somente PRECO ERRADO;
7. grava Quantidade de Cupons e Desconto na aba DESCONTOS;
8. salva a planilha na raiz operacional.

Nao existe exportacao intermediaria do Power BI para Excel no novo fluxo.

## Regra de descontos

O mes do periodo informado define a fonte:

- **mes atual**: pagina Resumo, visual DESCONTO POR MOTIVO;
- **mes anterior/fechado**: pagina Descontos Mes Anterior;
- **mes futuro**: bloqueado.

No historico, o campo **Valor Total** e deliberadamente ignorado.

Os campos gravados sao exclusivamente:

- Quantidade de Cupons;
- Desconto;
- Motivo = PRECO ERRADO.

### Ausencia nao vira zero

O Power BI pode nao retornar todas as lojas.

Por isso:

- loja com Quantidade e Desconto validos -> grava;
- zero explicito nos dois campos -> grava zero;
- loja ausente -> nao toca na celula;
- campo vazio/nulo -> nao toca na celula;
- loja do BI inexistente na planilha -> ignora e registra no log.

O robo nunca converte ausencia de informacao em zero.

## Filtro historico

A pagina Descontos Mes Anterior possui slicer Empresa.

Antes da leitura, o robo precisa confirmar que o slicer esta em **Todos**. Se nao conseguir limpar/confirmar esse filtro, o modulo interrompe a coleta de descontos e nao grava dados parciais.

## Credencial Power BI

Na primeira execucao do modulo de descontos, o CMD solicita:

- usuario/e-mail Power BI;
- senha Power BI.

A senha e armazenada localmente usando a protecao do Windows, da mesma forma que a credencial PDA.

Arquivo local:

\`data/bi_credential.json\`

Ele e ignorado pelo Git.

O Chrome do BI usa um perfil persistente separado:

\`output/chrome_bi/\`

Se a sessao ainda estiver autenticada, o relatorio abre diretamente. Se a Microsoft solicitar login, o robo tenta preencher a credencial protegida. MFA/aprovacoes adicionais, quando existirem, precisam ser concluidas na janela aberta; o robo aguarda e retoma automaticamente.

## Modos

Execute:

\`RoboPrecos.cmd\`

Menu:

1. testar uma unica loja no PDA;
2. **fluxo completo PDA + Power BI + planilha**;
3. **testar somente a leitura do Power BI sem gravar a planilha**.

O modo 3 foi criado para validar o novo modulo com seguranca antes da primeira gravacao real.

## Planilha operacional

Quando executado pela Plataforma RPA em:

\`plataforma-rpa/robots/robo-precos/\`

a planilha operacional fica na raiz:

\`plataforma-rpa/\`

Quando executado isoladamente, fica na raiz do RoboPrecos.

O robo reconhece uma unica planilha:

\`Controle de Auditoria de Precos*.xlsx\`

Se houver duas copias compativeis na raiz, ele interrompe para evitar gravacao no arquivo errado.

Uma copia antiga encontrada em Downloads pode ser migrada automaticamente para a raiz, mas Downloads nao e mais o destino operacional.

## Backups

Antes da gravacao da Auditoria e antes da gravacao de Descontos sao criados backups em:

\`RoboPrecos_Backups/\`

## Arquivos de controle

Coleta PDA:

\`output/checkpoints/\`

Snapshots do Power BI:

\`output/descontos/\`

Os snapshots permitem auditar exatamente quais lojas/valores foram lidos antes da gravacao.

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

## Arquitetura

\`RoboPrecos.cmd\`
  -> \`RoboPrecos.ps1\`
      -> \`src/bootstrap.ps1\`
      -> \`src/cdp.ps1\`
      -> \`src/pda.ps1\`
      -> \`src/network.ps1\`
      -> \`src/control_workbook.ps1\`
      -> \`src/bi.ps1\`
      -> \`src/discount_workbook.ps1\`


## Login Power BI por estados

O RoboPrecos nao abre mais o link direto do relatorio antes de autenticar.

Sequencia esperada:

1. pagina inicial do Power BI: informa o e-mail corporativo e envia;
2. pagina Microsoft de conta/e-mail, quando apresentada;
3. pagina Microsoft de senha;
4. pergunta para permanecer conectado;
5. confirmacao real de sessao autenticada;
6. somente entao abre o link do relatorio de descontos.

Paginas `singleSignOn` nao sao consideradas sessao autenticada.

Se houver MFA ou aprovacao externa, o robo aguarda a intervencao humana e retoma depois.

A classificacao de mes corrente/historico usa apenas ano e mes, evitando diferencas de milissegundos entre objetos DateTime.
