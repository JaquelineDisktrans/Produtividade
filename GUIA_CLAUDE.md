# Guia para fazer a análise qualitativa no Claude

O coletor calcula os indicadores objetivos localmente e cria `saida/casos_para_analise.jsonl`. Esse arquivo tem uma linha por conversa, com o texto das mensagens. O script `dividir_para_claude.py` o separa em lotes para não tentar enviar 60 mil e-mails de uma vez.

## Preparar os lotes

Depois que `executar_analise.bat` terminar, abra PowerShell na pasta `analise_outlook` e execute:

```powershell
py -3.12 dividir_para_claude.py
```

Serão criados arquivos em `saida\lotes_claude`. O padrão é 100 casos por lote; se os e-mails forem longos, use 30 ou 50:

```powershell
py -3.12 dividir_para_claude.py --casos-por-arquivo 50
```

## Etapa 1 — classificação de cada lote

Envie um lote por vez ao Claude com esta instrução, anexando o arquivo JSONL:

> Você é analista sênior de operações de atendimento. Analise cada objeto `case_id` deste lote. Não invente dados: use `null` quando algo não estiver identificável. Diferencie `observado` de `inferido` e inclua `confianca` (`alta`, `media` ou `baixa`). Responda exclusivamente com um JSONL válido, uma linha por caso, preservando o `case_id`.
>
> Para cada caso, produza: `case_id`, `id_caso_identificado`, `cliente`, `unidade`, `solicitante`, `assunto`, `motivo`, `submotivo`, `equipamento_modelo`, `patrimonio_identificacao`, `cidade_regiao`, `responsavel_disktrans`, `status_aparente`, `data_conclusao`, `principal_gargalo`, `houve_reclamacao`, `problema_portal_cadastro`, `contato_evitalvel`, `contato_necessario_ou_evitavel`, `cobrancas_cliente`, `numero_interacoes`, `qualidade` (notas de 1 a 5 para clareza, cordialidade, objetividade, urgencia, ownership, proatividade, resolucao, transparencia, previsao, profissionalismo e fechamento), `proximo_passo_claro`, `responsavel_claro`, `prazo_claro`, `evidencias` (indices das mensagens e pequenos trechos), `observado`, `inferido`, `confianca`.
>
> Motivos devem usar esta taxonomia quando aplicável: nova solicitação de troca; cobrança de troca; equipamento quebrado/parado; equipamento inadequado; problema recorrente; pedido de status; pedido de prazo; alteração/cancelamento; problema cadastral; equipamento não localizado na base; unidade não cadastrada; problema no Portal; dificuldade de acesso ao Portal; reclamação de atraso; reclamação de qualidade; solicitação administrativa; outros. Marque como potencialmente evitável cobrança por falta de retorno, pedido de status/previsão, repetição de informação, cadastro/Portal e ausência de comunicação proativa.

Salve cada resposta como `analise_lote_0001.jsonl`, `analise_lote_0002.jsonl` etc. Não peça ao Claude um relatório geral a cada lote; primeiro preserve a classificação estruturada.

## Etapa 2 — relatório executivo consolidado

Depois de classificar todos os lotes, forneça ao Claude:

- `resumo.csv`;
- `conversas.csv`;
- `assuntos_consolidados.csv`;
- todos os arquivos `analise_lote_*.jsonl`.

Use esta instrução final:

> Consolide os arquivos fornecidos em uma análise completa da operação de atendimento da Trocas Disktrans. Use os CSVs para métricas numéricas e os JSONL classificados para motivos, gargalos, qualidade e inferências. Não some respostas da mesma conversa como novas solicitações. Não invente: sinalize `não identificável`, `estimado` ou `baixa confiabilidade` quando necessário. Diferencie sempre dado observado de inferência e explique a metodologia e limitações da caixa de e-mails.
>
> Entregue: (A) base estruturada dos atendimentos; (B) tabela de indicadores; (C) gráficos/tendências ou tabelas equivalentes; (D) diagnóstico executivo; (E) problemas encontrados; (F) plano de ação priorizado em quick wins (0–30 dias), melhorias estruturais (30–90 dias) e mudanças sistêmicas (90+ dias). Inclua volume por dia/semana/mês, dia da semana e horário, média/mediana/P75/P90/P95 de primeira resposta, faixas de SLA, conclusão, interações, cobranças, sem resposta, backlog/aging, resolução, reincidência, taxonomia de motivos, contatos necessários versus evitáveis, gargalos, jornada, qualidade, produtividade com limitações, Top 10 motivos/retrabalho/reclamações/atrasos, clientes/unidades críticas, problemas sistêmicos/processuais/de comunicação, automação, autosserviço, Portal, atividades elimináveis e comunicação proativa.
>
> Termine obrigatoriamente com a seção **O que os dados estão tentando nos dizer**, contendo 5 a 10 conclusões executivas. Para cada recomendação informe impacto no cliente, redução de volume, redução de tempo, redução de retrabalho e facilidade de implantação.

## Privacidade

`casos_para_analise.jsonl` contém conteúdo de e-mails e dados pessoais. Mantenha-o local ou use somente um ambiente Claude autorizado pela empresa. Não envie o banco SQLite inteiro; envie os arquivos necessários e, se a política exigir, anonimize nomes, e-mails, telefones e patrimônios antes do upload.
