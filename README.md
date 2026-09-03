# Controle CS — Operação de Trocas

Dashboard local para acompanhar a caixa `trocas@disktrans.com.br` sem publicar o conteúdo dos e-mails.

## Abrir o painel

1. Execute a coleta em `executar_analise.bat` (ou use os relatórios já existentes em `saida`).
2. Dê dois cliques em `abrir_dashboard.bat`.
3. O navegador abrirá `http://localhost:8765`.

O próprio `executar_analise.bat` também atualiza o relatório preliminar e os dados do painel ao final da coleta.

O painel lê `saida/dashboard_data.json`, que é gerado por `gerar_dashboard_data.py`. Os dados reais e o banco SQLite estão ignorados pelo Git por conterem informações de clientes. Para atualizar o painel após uma nova coleta, execute:

```powershell
py -3.12 gerar_dashboard_data.py
```

## Atualização automática

Depois da primeira coleta histórica, execute `instalar_atualizacao_automatica.bat` uma única vez. O Windows criará uma tarefa que roda a cada 30 minutos e executa `atualizar_automatico.bat`.

Essa rotina:

- lê apenas mensagens desde a última coleta, com margem de dois dias para sincronizações tardias;
- não apaga nem altera e-mails no Outlook;
- atualiza o SQLite, os CSVs, o relatório preliminar e o JSON local;
- grava o histórico em `saida\atualizacao.log`;
- ignora uma nova rodada se a anterior ainda estiver em andamento.

O Outlook clássico precisa estar aberto e sincronizado no computador. O arquivo `parar_atualizacao_automatica.bat` remove a tarefa quando necessário. A revisão semântica dos lotes para Claude continua sendo uma etapa manual; a atualização automática recalcula os indicadores e as classificações heurísticas locais.

## Conteúdo

- Visão geral: casos, recebidos, taxa de resposta, backlog, SLA e volume.
- Motivos & gargalos: classificação heurística, concentração por cliente/unidade e sinais de risco.
- Atendimentos: tabela pesquisável por assunto, cliente, unidade e status.

Motivos, gargalos, reclamações e qualidade são triagens automáticas. Para o diagnóstico executivo, use os lotes em `saida/lotes_claude` e o `GUIA_CLAUDE.md`.

## Publicação

O frontend pode ser versionado no GitHub, mas não publique a pasta `saida` nem `casos_para_analise.jsonl` sem remover dados pessoais e obter autorização da empresa. A configuração atual é deliberadamente local.
