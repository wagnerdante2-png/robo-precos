# RoboPrecos

RPA local para coleta dos totalizadores de Auditoria de Precos no sistema PDA.

## Objetivo atual

O prototipo v0.1 consulta uma unica loja e le diretamente da tela do PDA:

- OK
- Divergente
- Sem etiqueta
- Total auditado

Nesta etapa ainda nao altera a planilha de controle. Primeiro validamos a coleta diretamente no PDA.

## Premissas

- Windows 10/11
- Google Chrome
- PowerShell nativo do Windows
- sem Python
- sem Selenium
- sem instalacao de bibliotecas
- sem GitHub Actions
- execucao local

## Como o prototipo funciona

1. Execute `RoboPrecos.cmd`.
2. Informe loja, data inicial e data final.
3. Na primeira execucao, informe usuario e senha do PDA.
4. A senha e protegida localmente pelo Windows via DPAPI e nao e enviada ao GitHub.
5. O robo abre uma instancia controlada do Chrome.
6. Acessa diretamente a tela de Auditoria de Preco.
7. Se a sessao estiver expirada, refaz o login.
8. Seleciona loja e periodo.
9. Aciona Pesquisar.
10. Le os quatro totalizadores diretamente do DOM da pagina.
11. Valida que OK + Divergente + Sem etiqueta = Total.
12. Salva o resultado de teste em `output/`.

## Arquivos novos

- `RoboPrecos.cmd`
- `RoboPrecos.ps1`
- `config.precos.example.json`
- `src/cdp.ps1`
- `src/pda.ps1`

A base herdada do `robo-horas` foi preservada para reaproveitamento, mas o novo fluxo usa somente os modulos necessarios.

## Seguranca local

Os arquivos abaixo ficam fora do Git:

- `config.precos.json`
- `data/pda_credential.json`
- `output/`

A credencial PDA nao deve ser adicionada manualmente ao repositorio.

## Proximas etapas apos validar o prototipo

1. percorrer todas as lojas;
2. checkpoint por loja;
3. retomada automatica apos timeout ou falha;
4. consolidacao da coleta;
5. preenchimento das abas ETIQUETAS, DIVERGENCIAS e SEM PRECO na planilha existente;
6. integrar posteriormente os dados de descontos vindos do BI.
