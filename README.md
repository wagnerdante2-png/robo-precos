# Robo de Disparos BI -> WhatsApp

MVP local para Windows, sem instalacao de Python, Selenium ou qualquer dependencia adicional.

## Como funciona

1. De duplo clique em RoboHoras.exe.
2. O robo abre o BI no Google Chrome.
3. Voce exporta o relatorio para Excel.
4. O robo detecta o novo arquivo na pasta Downloads.
5. O Excel e separado por loja.
6. Cada loja e cruzada com o telefone de data/lojas.csv.
7. A mensagem e gerada.
8. Em modo real, o WhatsApp Web e aberto e a mensagem e enviada.
9. Logs, previas e arquivos separados ficam em output/.

## Importante

- Nao existe instalacao.
- Nao e necessario Python.
- Nao e necessario Selenium.
- Nao existe GitHub Actions neste projeto.
- O robo usa PowerShell/.NET ja presentes no Windows e o Excel desktop instalado na maquina.
- O Google Chrome deve estar instalado.
- O envio pelo WhatsApp Web e temporario. Quando a Zenvia for confirmada, somente a camada final de envio sera substituida.

## Primeira execucao

Na primeira execucao, o robo cria automaticamente config.json e data/lojas.csv e abre ambos para edicao.

Em config.json, troque a URL do Power BI pela URL real do relatorio e confira o nome da coluna que identifica a loja.

Formato de data/lojas.csv:

loja,telefone,ativo
ML01,5511999999999,sim
ML02,5511988888888,sim

Use DDI + DDD + numero, somente digitos.

## Seguranca de disparo

O padrao e dryRun = true. Nesse modo o robo processa tudo, mas nao envia mensagens.
Depois da validacao, altere para false.

## Saidas

output/
  lojas/
  previews/
  logs/

## Requisitos ja existentes no PC

- Windows 10/11
- Google Chrome
- Microsoft Excel desktop

Nenhum pacote e instalado pelo robo.

## Estado atual do MVP

A exportacao do BI esta em modo assistido: o robo abre o relatorio e aguarda o usuario clicar em Exportar para Excel. Depois disso o processamento e automatico.

A automacao dos cliques especificos do relatorio sera adicionada depois de identificarmos a tela real do BI.

## Arquitetura futura

Hoje:
BI -> Excel -> separacao por loja -> mensagem -> WhatsApp Web

Depois da confirmacao da Zenvia:
BI -> Excel -> separacao por loja -> mensagem -> Zenvia API
