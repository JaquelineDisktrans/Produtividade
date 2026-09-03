# Analisador de atendimento do Outlook

Este programa transforma os e-mails do Outlook instalado no computador em indicadores de atendimento. Ele **somente le** Caixa de Entrada e Itens Enviados: nao apaga, move nem envia mensagens.

## Como usar

1. Abra o Outlook e espere ele terminar de sincronizar.
2. Feche o arquivo Excel, se ele estiver usando relatórios antigos.
3. Dê dois cliques em `executar_analise.bat`.
4. Espere a mensagem `Pronto`. Para 60 mil mensagens, a primeira execução pode levar bastante tempo. Não feche o Outlook ou a janela preta durante a leitura.

O arquivo já está configurado para a caixa `trocas@disktrans.com.br` e para mensagens de **1º de julho de 2026 em diante**. Ele falha com uma mensagem clara em vez de analisar a conta padrão caso essa caixa não esteja configurada no Outlook.

Os arquivos estarão em `saida`:

- `resumo.csv`: totais e tempo médio de primeira resposta;
- `conversas.csv`: uma linha por conversa, com assunto, status e primeira resposta;
- `assuntos_consolidados.csv`: agrupamento dos assuntos normalizados;
- `casos_para_analise.jsonl`: uma linha por conversa, incluindo o texto das mensagens para classificar motivo, gargalo, reclamação e qualidade;
- `outlook_atendimento.sqlite`: base técnica local. Não é necessário abri-la.

## Atualização automática

Depois que a primeira leitura histórica terminar, execute `instalar_atualizacao_automatica.bat` uma vez. O Windows passará a executar a atualização a cada 30 minutos, enquanto o usuário estiver conectado e o Outlook clássico estiver aberto e sincronizado. A rotina incremental lê uma janela de dois dias, reaproveita o SQLite e evita reler todo o histórico. O arquivo `saida\atualizacao.log` registra cada execução. Para desativar, execute `parar_atualizacao_automatica.bat`.

Arquivos `.csv` abrem no Excel. O separador usado é `;` e a codificação é compatível com acentos no Excel.

O arquivo JSONL pode ficar grande porque contém o corpo das mensagens. Ele permanece no computador e serve para a etapa de análise qualitativa. Não o envie para serviços externos sem aprovação da empresa; os e-mails podem conter dados pessoais e informações de clientes.

## Como o status é calculado

Uma conversa é considerada **Respondida** quando existe uma mensagem em Itens Enviados na mesma conversa, enviada depois do primeiro e-mail recebido. O intervalo entre as duas datas é o **tempo de primeira resposta**. Se não existir mensagem enviada depois, ela aparece como **Sem resposta**.

O Outlook agrupa mensagens pelo identificador interno de conversa. Em mensagens sem esse identificador, o programa usa o assunto sem prefixos como `RE:`, `RES:`, `FW:` e `ENC:`. Por isso, e-mails independentes que têm exatamente o mesmo assunto podem eventualmente ficar no mesmo grupo; os detalhes podem ser conferidos em `conversas.csv`.

## Observações importantes

- A caixa `trocas@disktrans.com.br` precisa estar configurada no Outlook. Ela pode ser compartilhada; não precisa ser a conta padrão.
- Respostas automáticas enviadas pelo Outlook contam como resposta. Se a caixa usa muitas delas, me peça para ajustar a regra e excluí-las.
- A data selecionada é aplicada tanto aos recebidos quanto aos enviados. Para calcular respostas a mensagens logo antes da data inicial, rode um período com alguma margem anterior.
