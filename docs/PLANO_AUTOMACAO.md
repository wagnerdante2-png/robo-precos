# Plano de automacao RoboHoras

## Fase atual - fechar o envio WhatsApp Web

- [x] Executavel portatil sem instalacao
- [x] Leitura e normalizacao do Excel
- [x] Identificacao por loja
- [x] Geracao da mensagem
- [x] Abertura da conversa correta no WhatsApp Web
- [ ] Acionamento automatico validado do botao Enviar
- [ ] Confirmacao visual/log de envio real

## Fase seguinte - multiplos destinatarios

- cadastrar telefone por loja;
- enviar sequencialmente;
- intervalo configuravel entre envios;
- continuar apos falha de uma loja;
- retry isolado das lojas com falha.

## Coleta 100% automatica no Power BI

Nao sera utilizada API do Power BI.

Fluxo alvo:

1. abrir o Power BI no navegador;
2. abrir o relatorio configurado;
3. aplicar filtros necessarios;
4. localizar o visual correto;
5. exportar dados pela interface;
6. detectar o arquivo baixado;
7. validar conteudo antes dos disparos.

## Agendamento futuro

Configuracao preparada em `config.json`:

```json
"schedule": {
  "enabled": false,
  "time": "07:00",
  "days": ["MON", "TUE", "WED", "THU", "FRI", "SAT"],
  "runMissedAsSoonAsPossible": true,
  "preventDuplicateSameDay": true,
  "retryFailedStores": true,
  "requireNetwork": true,
  "lateRunPolicy": "same_day_any_time"
}
```

Comportamento alvo:

- se o PC estiver ligado as 07:00, executar normalmente;
- se o PC estiver desligado, executar assim que o Windows voltar a disponibilizar a tarefa;
- impedir uma segunda execucao completa no mesmo dia;
- em caso de falhas parciais, retentar somente destinatarios pendentes;
- manter estados PENDENTE, EM_EXECUCAO, CONCLUIDO e CONCLUIDO_COM_FALHAS.

O agendamento permanecera desativado ate o fluxo BI -> dados -> WhatsApp estar validado ponta a ponta.
