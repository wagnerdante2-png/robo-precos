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


## Login por acessibilidade CDP

A partir da v0.4.3, o login do Power BI nao depende de localizar inputs por CSS/DOM comum.

O robo usa a arvore de acessibilidade do Chrome para:

- localizar o campo de e-mail visivel;
- focar o controle via CDP;
- digitar o usuario fornecido no CMD;
- acionar Enviar;
- localizar e preencher a senha;
- acionar Entrar;
- confirmar a tela de permanecer conectado.

Isso cobre telas renderizadas em shadow DOM/controles que nao aparecem para querySelector, mas estao visiveis ao usuario.


## Login atravessando frames

A partir da v0.4.4, as telas de autenticacao do Power BI/Microsoft sao tratadas pelo DOM achatado do Chrome:

`DOM.getFlattenedDocument(depth=-1, pierce=true)`

Com isso o robo localiza controles mesmo quando estao dentro de frames internos.

Fluxo executado:

1. localizar INPUT de e-mail, focar, digitar a credencial fornecida e enviar;
2. localizar INPUT de senha, focar, digitar a senha fornecida e entrar;
3. localizar o botao Sim da tela Continuar conectado e confirmar;
4. somente depois abrir o link do relatorio.

Cada campo aguarda ate 20 segundos para ficar disponivel ao Chrome DevTools antes de considerar falha.


## Login direto sem foco

Na v0.4.5 o login do Power BI nao usa mais `DOM.focus`, teclado ou ENTER.

O robo:

1. atravessa os frames com `DOM.getFlattenedDocument(pierce=true)`;
2. resolve cada INPUT/BUTTON real com `DOM.resolveNode`;
3. verifica visibilidade por `getBoundingClientRect()` e `getComputedStyle()`;
4. escolhe somente o controle visivel;
5. preenche e-mail/senha pelo setter nativo do input + eventos `input/change`;
6. chama `click()` diretamente no botao visivel;
7. confirma o valor efetivamente aplicado antes de avancar.

Isso elimina o erro `Element is not focusable`.


## Reconciliacao com o Total do BI

No mes corrente, o RoboPrecos nao considera mais suficiente apenas contar lojas.

Ele soma todos os registros capturados de PRECO ERRADO e compara com o rodape Total do proprio visual do Power BI:

- Quantidade Cupons;
- Desconto.

A coleta corrente so e considerada integra quando:

`soma das lojas = Total do visual`

Se qualquer valor divergir, o modulo falha e nao deve seguir para gravacao.

O modo 3 agora imprime todas as lojas coletadas, sem corte nas 15 primeiras, e mostra:

- total de lojas;
- soma dos cupons;
- soma dos descontos;
- Total do visual;
- resultado da reconciliacao;
- IDs de empresa ausentes dentro do intervalo encontrado.


## Redirect pos-login tolerante

Na v0.4.7 o login permanece congelado no fluxo ja validado.

A unica alteracao esta na transicao entre autenticacao concluida e abertura do relatorio:

- aguarda alguns segundos para o redirect automatico do Power BI terminar;
- se o proprio Power BI ja chegar ao relatorio, nao envia nova navegacao;
- se `Page.navigate` retornar `net::ERR_ABORTED` durante redirect SSO, o erro e tratado como transitorio;
- o robo confirma a URL real do navegador antes de considerar falha;
- existe uma unica tentativa final por `window.location.replace` se o redirect anterior cancelar a navegacao.

Nenhuma credencial, etapa de login ou regra de coleta foi alterada.


## Espera de materializacao do visual

Na v0.4.8, depois que o texto do relatorio aparece, o RoboPrecos nao assume que a tabela ja esta pronta.

A leitura do visual tenta por ate 45 segundos:

- grid/tabela via DOM;
- linhas via arvore de acessibilidade do Chrome.

A primeira tentativa vazia nao e mais tratada como falha.

Login, credenciais e navegacao permanecem congelados no fluxo ja validado.
