# GetReportFT Script

Este é um script Python projetado para processar contratos e relatórios fornecidos como argumentos da linha de comando. Ele suporta a especificação de contratos, relatórios e uma data de referência para o processamento.

## Uso

```bash
python GetReportFT.py [OPÇÕES]
```

## Opções

- `-c CONTRATO, --contrato=CONTRATO, --contratos=CONTRATO`: Especifica o contrato ou uma lista de contratos separados por vírgula a serem processados.
- `-r RELATÓRIO, --relatorio=RELATÓRIO, --relatorios=RELATÓRIO`: Especifica o relatório ou uma lista de relatórios separados por vírgula a serem processados.
- `-d DATA, --datareferencia=DATA, --datareferencias=DATA`: Especifica a data de referência para processamento.

## Exemplos

- Para processar os contratos `dnitms` e `msvia`, os relatórios `rel_tst` e `rel_inf`, e a data de referência `202404`:
  ```bash
  python GetReportFT.py -c dnitms,msvia -r rel_tst,rel_inf -d 202404
  ```

- Para processar apenas o contrato `dnitms`:
  ```bash
  python GetReportFT.py --contrato=dnitms
  ```

- Para processar o contrato `msvia`, os relatórios `rel_inf` e `rel_flx`, e a data de referência `202405`:
  ```bash
  python GetReportFT.py -c msvia --relatorio=rel_inf,rel_flx -d 202405
  ```

- Para obter ajuda sobre como usar o script:
  ```bash
  python GetReportFT.py --help
  ```

## Observações

- Este script depende de argumentos de linha de comando para especificar os contratos, relatórios e data de referência.
- As configurações de exibição de mensagens durante a execução do script podem ser ajustadas definindo a variável `SHOW_RESPONSE` no código-fonte.