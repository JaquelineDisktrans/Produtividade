"""Gera uma análise operacional preliminar a partir da coleta do Outlook.

As classificações de texto deste arquivo são heurísticas e ficam marcadas como
inferência no relatório. A fonte de verdade para fatos temporais é a base SQLite.
"""

from __future__ import annotations

import csv
import json
import re
import sqlite3
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from pathlib import Path
from statistics import mean, median


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "saida"
DB = OUT / "outlook_atendimento.sqlite"


def sem_acento(valor: str) -> str:
    valor = unicodedata.normalize("NFKD", valor or "")
    return "".join(c for c in valor if not unicodedata.combining(c)).casefold()


def pct(n: int | float, total: int | float) -> str:
    return f"{(n / total * 100):.1f}%" if total else "0.0%"


def quantile(valores: list[float], q: float) -> float | None:
    if not valores:
        return None
    valores = sorted(valores)
    pos = (len(valores) - 1) * q
    baixo, alto = int(pos), min(int(pos) + 1, len(valores) - 1)
    return valores[baixo] + (valores[alto] - valores[baixo]) * (pos - baixo)


def horas(minutos: float | None) -> str:
    if minutos is None:
        return "não identificável"
    h, m = divmod(round(minutos), 60)
    return f"{h}h {m:02d}min"


def normalizar_assunto(assunto: str) -> str:
    texto = re.sub(r"^\s*(?:(?:re|res|fw|fwd|enc)\s*:\s*)+", "", assunto or "", flags=re.I).strip()
    return re.sub(r"\s+", " ", texto) or "(sem assunto)"


def primeiro_nome_email(endereco: str) -> str:
    if not endereco:
        return "não identificado"
    endereco = endereco.casefold()
    return endereco.split("@", 1)[-1] if "@" in endereco else endereco


MOTIVOS = [
    ("cobrança de troca", ["algum retorno", "cobrando", "cobranca", "sem retorno", "retorno do chamado", "posição do chamado"]),
    ("reclamação de atraso", ["atraso", "atrasada", "demora", "urgente", "até agora não", "nao recebemos"]),
    ("pedido de status", ["status", "posição", "posicao", "andamento", "como está", "como esta"]),
    ("pedido de prazo", ["previsão", "previsao", "prazo", "quando será", "quando sera", "data da troca"]),
    ("equipamento quebrado/parado", ["quebrad", "parad", "sem condição de uso", "sem condicao de uso", "não funciona", "nao funciona", "defeito"]),
    ("equipamento inadequado", ["inadequad", "modelo incorreto", "não atende", "nao atende"]),
    ("problema recorrente", ["recorrente", "novamente", "outra vez", "reincid", "volta a apresentar"]),
    ("problema no Portal", ["portal", "erro no sistema", "não consigo acessar", "nao consigo acessar", "login", "senha"]),
    ("problema cadastral", ["cadastro", "cadastrada", "cadastrado", "não localizado na base", "nao localizado na base"]),
    ("alteração/cancelamento", ["cancel", "alterar", "alteração", "alteracao", "devolução", "devolucao"]),
    ("nova solicitação de troca", ["solicito troca", "solicitação de troca", "solicitacao de troca", "troca de paleteira", "troca inclusa"]),
    ("solicitação administrativa", ["contrato", "orçamento", "orcamento", "nota fiscal", "faturamento"]),
]


def classificar_motivo(texto: str) -> str:
    normalizado = sem_acento(texto)
    for motivo, termos in MOTIVOS:
        if any(sem_acento(t) in normalizado for t in termos):
            return motivo
    return "outros"


def classificar_gargalo(texto: str, motivo: str) -> str:
    normalizado = sem_acento(texto)
    regras = [
        ("TI/Portal", ["portal", "login", "senha", "sistema", "erro"]),
        ("Cadastro", ["cadastro", "cadastrad", "base de dados", "nao localizado"]),
        ("Manutenção", ["manutencao", "reparo", "quebrad", "defeito"]),
        ("disponibilidade de equipamento", ["disponibilidade", "sem equipamento", "aguardando equipamento"]),
        ("Logística/transportadora", ["transportadora", "entrega", "coleta", "logistica", "frete"]),
        ("Aprovação interna", ["aprovacao", "aprovação", "aguardando aprovacao"]),
        ("Comunicação/Atendimento", ["sem retorno", "algum retorno", "status", "previsao", "posicao"]),
    ]
    for nome, termos in regras:
        if any(sem_acento(t) in normalizado for t in termos):
            return nome
    if motivo in {"cobrança de troca", "pedido de status", "pedido de prazo"}:
        return "Comunicação/Atendimento"
    return "não identificável"


def extrair_unidade(texto: str) -> str:
    encontrados = re.findall(r"(?:loja|unidade|cd)\s*[-#: ]?\s*([0-9]{2,6}(?:[- ][A-Za-zÀ-ÿ0-9]+)?)", texto, flags=re.I)
    return encontrados[0].strip() if encontrados else "não identificada"


def extrair_patrimonios(texto: str) -> str:
    encontrados = re.findall(r"(?:s[eé]rie|series|patrim[oô]nio|n[uú]mero)\s*[:#-]?\s*([A-Za-z0-9./-]{3,})", texto, flags=re.I)
    return ", ".join(dict.fromkeys(encontrados[:10])) if encontrados else "não identificado"


def extrair_cidade(texto: str) -> str:
    m = re.search(r"(?:loja|end(?:ere[cç]o)?|cidade)\s*[^\n]{0,100}[,-]\s*([A-Za-zÀ-ÿ ]{3,})(?:[-,]\s*(?:ba|sp|rj|mg|pr|sc|rs)\b)?", texto, flags=re.I)
    return re.sub(r"\s+", " ", m.group(1)).strip() if m else "não identificada"


def corpo_principal(texto: str) -> str:
    # Remove a parte mais comum de histórico citado, preservando a resposta atual.
    cortes = [p for p in (texto.find("\n________________________________"), texto.find("\nDe:"), texto.find("\nEnviado:")) if p > 20]
    return texto[: min(cortes)] if cortes else texto


def qualidade_resposta(texto: str) -> tuple[int, dict[str, bool]]:
    corpo = corpo_principal(texto)
    normalizado = sem_acento(corpo)
    flags = {
        "proximo_passo_claro": bool(re.search(r"\b(vou|iremos|sera|será|programad|inclus|encaminh|solicit|aguard)", normalizado)),
        "responsavel_claro": bool(re.search(r"atendimento|operacional|responsavel|responsável|ramal|equipe", normalizado)),
        "prazo_claro": bool(re.search(r"\b(?:dia|dias|amanha|amanhã|data|prazo|previsao|previsão|até|ate)\b", normalizado)),
        "cordialidade": bool(re.search(r"bom dia|boa tarde|obrigad|disposicao|disposição", normalizado)),
        "resposta_vaga": bool(re.search(r"vou verificar|vamos verificar|em analise|em análise", normalizado)) and not bool(re.search(r"\b(?:dia|data|prazo|retorno)\b", normalizado)),
    }
    pontos = 3
    pontos += int(flags["proximo_passo_claro"])
    pontos += int(flags["responsavel_claro"])
    pontos += int(flags["prazo_claro"])
    pontos -= int(flags["resposta_vaga"])
    return max(1, min(5, pontos)), flags


def carregar_casos() -> list[dict]:
    caminho = OUT / "casos_para_analise.jsonl"
    with caminho.open(encoding="utf-8") as arquivo:
        return [json.loads(linha) for linha in arquivo if linha.strip()]


def analisar_casos(casos: list[dict]) -> list[dict]:
    saida = []
    for caso in casos:
        mensagens = sorted(caso.get("mensagens", []), key=lambda m: m["data_hora"])
        recebidos = [m for m in mensagens if m["direcao"] == "recebido"]
        enviados = [m for m in mensagens if m["direcao"] == "enviado"]
        if not recebidos:
            continue
        primeiro = recebidos[0]
        inicio = datetime.fromisoformat(primeiro["data_hora"])
        respostas = [m for m in enviados if m["data_hora"] >= primeiro["data_hora"]]
        resposta = respostas[0] if respostas else None
        tempo = (datetime.fromisoformat(resposta["data_hora"]) - inicio).total_seconds() / 60 if resposta else None
        texto = (caso.get("assunto_consolidado", "") + "\n" + primeiro.get("corpo", ""))
        followups = [m for m in recebidos[1:]]
        cobrancas = sum(bool(re.search(r"retorno|cobranc|status|posi[cç][aã]o|previs[aã]o|quando|aguard", sem_acento(m.get("corpo", "")))) for m in followups)
        motivo = classificar_motivo(texto)
        evita = motivo in {"cobrança de troca", "pedido de status", "pedido de prazo", "problema cadastral", "problema no Portal", "reclamação de atraso"} or cobrancas > 0
        primeiro_envio = resposta.get("corpo", "") if resposta else ""
        nota, flags = qualidade_resposta(primeiro_envio) if resposta else (None, {"proximo_passo_claro": False, "responsavel_claro": False, "prazo_claro": False, "cordialidade": False, "resposta_vaga": False})
        encerramento = None
        for mensagem in reversed(enviados):
            if re.search(r"conclu[ií]d|finaliz|solucion|encerr|realizad|troca inclusa", sem_acento(mensagem.get("corpo", ""))):
                encerramento = mensagem["data_hora"]
                break
        saida.append({
            "case_id": caso["case_id"],
            "data_inicial": primeiro["data_hora"],
            "cliente": primeiro_nome_email(primeiro.get("remetente", "")),
            "unidade": extrair_unidade(texto),
            "solicitante": primeiro.get("remetente", "") or "não identificado",
            "assunto": caso.get("assunto_consolidado", "(sem assunto)"),
            "motivo_heuristico": motivo,
            "submotivo": "não classificado automaticamente",
            "equipamento_modelo": "paleteira" if "paleteira" in sem_acento(texto) else "não identificado",
            "patrimonio_identificacao": extrair_patrimonios(texto),
            "cidade_regiao": extrair_cidade(texto),
            "responsavel_disktrans": "não identificável automaticamente",
            "primeira_resposta": resposta["data_hora"] if resposta else "",
            "tempo_primeira_resposta_minutos": round(tempo, 2) if tempo is not None else "",
            "interacoes": len(mensagens),
            "cobrancas_cliente": cobrancas,
            "status_aparente": "Respondido" if resposta else "Sem resposta",
            "data_conclusao_inferida": encerramento or "",
            "tempo_total_conclusao_minutos": round((datetime.fromisoformat(encerramento) - inicio).total_seconds() / 60, 2) if encerramento else "",
            "principal_gargalo_heuristico": classificar_gargalo(texto, motivo),
            "houve_reclamacao_heuristica": "Sim" if any(t in sem_acento(texto) for t in ("reclama", "insatisfeit", "atraso", "absurdo", "urgente")) else "Não identificado",
            "problema_portal_cadastro_heuristico": "Sim" if motivo in {"problema no Portal", "problema cadastral"} else "Não identificado",
            "contato_potencialmente_evitavel": "Sim" if evita else "Não identificado",
            "qualidade_heuristica_1a5": nota if nota is not None else "",
            **flags,
        })
    return saida


def escrever_csv(caminho: Path, linhas: list[dict]):
    if not linhas:
        return
    with caminho.open("w", newline="", encoding="utf-8-sig") as arquivo:
        writer = csv.DictWriter(arquivo, fieldnames=list(linhas[0]), delimiter=";")
        writer.writeheader()
        writer.writerows(linhas)


def tabela_contagem(contagem: Counter, total: int, titulo: str, n: int = 10) -> list[str]:
    linhas = [f"### {titulo}", "", "| Categoria | Casos | Percentual |", "|---|---:|---:|"]
    for chave, quantidade in contagem.most_common(n):
        linhas.append(f"| {chave} | {quantidade} | {pct(quantidade, total)} |")
    if not contagem:
        linhas.append("| Não identificável | 0 | 0,0% |")
    return linhas + [""]


def main():
    casos = carregar_casos()
    linhas = analisar_casos(casos)
    escrever_csv(OUT / "base_estruturada_heuristica.csv", linhas)
    with sqlite3.connect(DB) as conexao:
        total_recebidos_caixa = conexao.execute("SELECT COUNT(*) FROM mensagens WHERE direcao='recebido'").fetchone()[0]
        total_enviados_caixa = conexao.execute("SELECT COUNT(*) FROM mensagens WHERE direcao='enviado'").fetchone()[0]
    recebidos = [datetime.fromisoformat(l["data_inicial"]) for l in linhas]
    respostas = [float(l["tempo_primeira_resposta_minutos"]) for l in linhas if l["tempo_primeira_resposta_minutos"] != ""]
    total = len(linhas)
    respondidos = sum(l["status_aparente"] == "Respondido" for l in linhas)
    sem_resposta = total - respondidos
    interacoes = [int(l["interacoes"]) for l in linhas]
    referencia = max(recebidos) if recebidos else datetime.now()
    idade = [(referencia - d).total_seconds() / 86400 for d, l in zip(recebidos, linhas) if l["status_aparente"] == "Sem resposta"]

    diario = Counter(d.date().isoformat() for d in recebidos)
    semanal = Counter(f"{d.isocalendar().year}-S{d.isocalendar().week:02d}" for d in recebidos)
    mensal = Counter(d.strftime("%Y-%m") for d in recebidos)
    dia_semana = Counter(d.strftime("%A") for d in recebidos)
    horario = Counter(f"{d.hour:02d}:00–{d.hour:02d}:59" for d in recebidos)
    limites = [(30, "até 30 min"), (60, "até 1h"), (120, "até 2h"), (240, "até 4h"), (1440, "até 24h / D+1")]
    sla = [(nome, sum(v <= limite for v in respostas), pct(sum(v <= limite for v in respostas), len(respostas))) for limite, nome in limites]
    mesmo_dia = sum(1 for l in linhas if l["primeira_resposta"] and l["primeira_resposta"][:10] == l["data_inicial"][:10])
    cobraram = sum(int(l["cobrancas_cliente"]) > 0 for l in linhas)
    duas_cobrancas = sum(int(l["cobrancas_cliente"]) >= 2 for l in linhas)
    concluidos = sum(bool(l["data_conclusao_inferida"]) for l in linhas)
    reaberturas = sum("Respondido" == l["status_aparente"] and int(l["interacoes"]) > 3 for l in linhas)
    motivos = Counter(l["motivo_heuristico"] for l in linhas)
    gargalos = Counter(l["principal_gargalo_heuristico"] for l in linhas)
    clientes = Counter(l["cliente"] for l in linhas)
    unidades = Counter(l["unidade"] for l in linhas if l["unidade"] != "não identificada")
    reclama = Counter(l["motivo_heuristico"] for l in linhas if l["houve_reclamacao_heuristica"] == "Sim")
    evitaveis = sum(l["contato_potencialmente_evitavel"] == "Sim" for l in linhas)
    notas = [int(l["qualidade_heuristica_1a5"]) for l in linhas if l["qualidade_heuristica_1a5"] != ""]
    flag_names = [("proximo_passo_claro", "próximo passo claro"), ("responsavel_claro", "responsável claro"), ("prazo_claro", "prazo/previsão clara")]

    relatorio = [
        "# Análise preliminar da operação de atendimento — Trocas Disktrans",
        "",
        "> Gerado automaticamente em 03/09/2026 a partir da caixa `trocas@disktrans.com.br`, considerando mensagens de 01/07/2026 em diante.",
        "> As métricas temporais são observadas na caixa. Motivo, gargalo, reclamação, conclusão e qualidade marcados como heurísticos são inferências por palavras-chave e precisam de validação semântica.",
        "",
        "## Metodologia e limitações",
        "",
        f"Foram consolidadas **{total} conversas** a partir de **{sum(int(l['interacoes']) for l in linhas)} mensagens vinculadas a essas conversas**. A caixa no período contém **{total_recebidos_caixa} recebidos** e **{total_enviados_caixa} enviados**; mensagens enviadas sem um recebido correspondente não são tratadas como novos casos. O agrupamento usa o ConversationID do Outlook e, quando ausente, assunto normalizado. Uma conversa é Respondida quando há uma mensagem enviada após o primeiro recebimento.",
        "",
        "A caixa de e-mail não contém necessariamente data formal de conclusão, SLA contratado, responsável estruturado ou registro de status. Por isso, conclusão, backlog, resolução, produtividade e qualidade são estimativas ou não identificáveis quando não há evidência suficiente.",
        "",
        "## 1. Indicadores de atendimento",
        "",
        "| Indicador | Resultado | Natureza |",
        "|---|---:|---|",
        f"| E-mails recebidos na caixa | {total_recebidos_caixa} | observado na coleta |",
        f"| E-mails enviados na caixa | {total_enviados_caixa} | observado na coleta |",
        f"| Mensagens vinculadas aos casos | {sum(int(l['interacoes']) for l in linhas)} | observado após consolidação |",
        f"| Casos/conversas | {total} | observado/consolidado |",
        f"| Respondidos | {respondidos} ({pct(respondidos, total)}) | regra temporal |",
        f"| Sem resposta aparente | {sem_resposta} ({pct(sem_resposta, total)}) | regra temporal |",
        f"| Primeira resposta — média | {horas(mean(respostas) if respostas else None)} | observado |",
        f"| Primeira resposta — mediana | {horas(median(respostas) if respostas else None)} | observado |",
        f"| P75 / P90 / P95 primeira resposta | {horas(quantile(respostas, .75))} / {horas(quantile(respostas, .90))} / {horas(quantile(respostas, .95))} | observado |",
        f"| Resposta no mesmo dia | {mesmo_dia} ({pct(mesmo_dia, total)}) | observado entre respondidos |",
        f"| Média de interações por caso | {mean(interacoes):.2f} | observado |",
        f"| Cliente precisou cobrar | {cobraram} ({pct(cobraram, total)}) | heurístico |",
        f"| Duas ou mais cobranças | {duas_cobrancas} ({pct(duas_cobrancas, total)}) | heurístico |",
        f"| Conclusões inferidas | {concluidos} ({pct(concluidos, total)}) | heurístico, não status formal |",
        f"| Backlog estimado | {sem_resposta} casos | proxy = sem resposta |",
        f"| Aging médio do backlog | {mean(idade):.1f} dias | estimado até {referencia:%Y-%m-%d} |" if idade else "| Aging do backlog | não identificável | sem casos sem resposta |",
        f"| Taxa de resolução estimada | {pct(concluidos, total)} | proxy textual, baixa confiabilidade |",
        f"| Possíveis reaberturas/reincidências | {reaberturas} | heurístico: >3 interações |",
        "",
        "### Faixas de primeira resposta",
        "",
        "| Faixa | Casos | % dos respondidos |",
        "|---|---:|---:|",
    ]
    relatorio.extend(f"| {nome} | {quantidade} | {percentual} |" for nome, quantidade, percentual in sla)
    relatorio.extend(["", "### Volume e tendências", "", "| Mês | Casos |", "|---|---:|"])
    relatorio.extend(f"| {chave} | {valor} |" for chave, valor in sorted(mensal.items()))
    relatorio.extend(["", "| Semana ISO | Casos |", "|---|---:|"])
    relatorio.extend(f"| {chave} | {valor} |" for chave, valor in sorted(semanal.items()))
    relatorio.extend(["", "| Dia da semana | Casos |", "|---|---:|"])
    relatorio.extend(f"| {chave} | {valor} |" for chave, valor in dia_semana.most_common())
    relatorio.extend(["", "| Faixa horária | Casos |", "|---|---:|"])
    relatorio.extend(f"| {chave} | {valor} |" for chave, valor in horario.most_common())
    relatorio.extend(["", "## 2. Motivos dos contatos", "", "Classificação automática por assunto e primeiro corpo recebido; tratar como inferência.", ""])
    relatorio.extend(tabela_contagem(motivos, total, "Principais motivos", 20))
    relatorio.extend(["| Motivo | Semanas com ocorrência |", "|---|---|"])
    por_motivo_semana = defaultdict(Counter)
    for linha in linhas:
        por_motivo_semana[linha["motivo_heuristico"]][datetime.fromisoformat(linha["data_inicial"]).strftime("%Y-S%W")] += 1
    for motivo, semanas in sorted(por_motivo_semana.items(), key=lambda item: sum(item[1].values()), reverse=True):
        relatorio.append(f"| {motivo} | " + ", ".join(f"{s}: {n}" for s, n in sorted(semanas.items())) + " |")
    relatorio.extend(["", "### Contatos necessários versus potencialmente evitáveis", "", f"Casos marcados como potencialmente evitáveis: **{evitaveis} ({pct(evitaveis, total)})**. Critério heurístico: cobrança, pedido de status/prazo, falha de cadastro/Portal ou reclamação de atraso.", ""])
    relatorio.extend(tabela_contagem(clientes, total, "Clientes/domínios com mais casos", 10))
    relatorio.extend(tabela_contagem(unidades, total, "Unidades identificadas com mais casos", 10))
    relatorio.extend(["## 3. Indicadores operacionais", "", "| Origem/gargalo provável | Casos | % |", "|---|---:|---:|"])
    relatorio.extend(f"| {k} | {v} | {pct(v, total)} |" for k, v in gargalos.most_common())
    relatorio.extend(["", "A jornada Solicitação → entendimento → registro → validação → disponibilidade → programação → logística → execução → confirmação não possui eventos estruturados na caixa. Os gargalos acima são indícios textuais, não tempos de etapa.", "", "## 4. Qualidade do atendimento", "", f"Nota heurística média entre os {len(notas)} casos respondidos: **{mean(notas):.2f}/5**. A nota não substitui avaliação humana; foi calculada por presença de próximo passo, responsável, prazo e linguagem cordial.", "", "| Indicador de qualidade | Casos | Percentual |", "|---|---:|---:|"])
    for nome, rotulo in flag_names:
        quantidade = sum(bool(l[nome]) for l in linhas if l["status_aparente"] == "Respondido")
        relatorio.append(f"| {rotulo} | {quantidade} | {pct(quantidade, respondidos)} |")
    vagos = sum(bool(l["resposta_vaga"]) for l in linhas)
    relatorio.extend([f"| Resposta possivelmente vaga | {vagos} | {pct(vagos, respondidos)} |", f"| Cliente precisou cobrar Disktrans | {cobraram} | {pct(cobraram, total)} |", "", "## 5. Produtividade", "", "O responsável da Disktrans não foi identificado de forma confiável nos metadados disponíveis: remetentes Exchange aparecem como identificadores técnicos. Não é recomendável ranking por pessoa nesta versão. A distribuição por responsável exige assinatura padronizada ou campo de remetente resolvido.", "", "## 6. Análise gerencial", ""])
    relatorio.extend(tabela_contagem(motivos, total, "Top motivos", 10))
    relatorio.extend(tabela_contagem(gargalos, total, "Top causas prováveis de retrabalho/gargalo", 10))
    relatorio.extend(tabela_contagem(reclama, sum(reclama.values()), "Motivos associados a reclamações (heurístico)", 10))
    relatorio.extend(["Clientes e unidades críticas devem ser confirmados usando a coluna de domínio/unidade da base estruturada; a extração automática não identifica todos os nomes comerciais.", "", "## 7. Recomendações priorizadas", "", "### Quick wins — 0–30 dias", "", "1. Criar resposta-padrão com número do caso, responsável, próximo passo e prazo; impacto alto e implantação fácil.", "2. Enviar atualização proativa quando o caso ficar aguardando outra área; reduz cobranças e pedidos de status.", "3. Registrar status mínimo (recebido, em validação, aguardando ativo, programado, concluído) no assunto ou corpo.", "", "### Melhorias estruturais — 30–90 dias", "", "1. Implantar formulário/Portal com campos obrigatórios de unidade, equipamento, patrimônio, cidade e urgência.", "2. Criar fila por etapa e SLA, com alerta de aging e responsável explícito.", "3. Padronizar encerramento com confirmação ao cliente e data de conclusão.", "", "### Mudanças sistêmicas — 90+ dias", "", "1. Integrar Portal, ativos, manutenção e logística para eliminar consultas manuais.", "2. Criar histórico de casos e reincidência por patrimônio, equipamento, cliente e unidade.", "3. Construir painel operacional com SLA, backlog, aging, causas e carga por responsável.", "", "## 8. O que os dados estão tentando nos dizer", "", f"1. A caixa contém **{sem_resposta} de {total} conversas sem resposta aparente**, sinal de forte risco de backlog ou de que parte do atendimento ocorre fora da caixa.", f"2. A resposta média observada é de **{horas(mean(respostas) if respostas else None)}**, mas a mediana e os percentis devem orientar o SLA porque a média pode esconder cauda longa.", f"3. **{pct(evitaveis, total)}** dos casos foram marcados como potencialmente evitáveis por status, prazo, cobrança, cadastro ou Portal; é uma hipótese operacional que precisa de validação semântica.", f"4. O volume enviado é muito maior que o recebido, indicando que respostas e históricos longos não devem ser contados como novas solicitações.", "5. Falta de responsável, prazo e status estruturados limita a gestão e força o cliente a usar o e-mail como rastreador.", "6. A jornada operacional não pode ser medida por etapa com confiabilidade somente pela caixa; é necessário registrar eventos.", "7. A ausência de conclusão formal impede distinguir caso resolvido de caso apenas respondido.", "8. As classificações de motivo, reclamação e qualidade desta versão são triagem automática; a revisão semântica dos lotes pode confirmar, corrigir ou descartar esses sinais.", "", "## Arquivos de apoio", "", "- `base_estruturada_heuristica.csv`: uma linha por caso com campos e inferências.", "- `resumo.csv`, `conversas.csv`, `assuntos_consolidados.csv`: métricas e consolidação temporal.", "- `casos_para_analise.jsonl` e `lotes_claude`: evidências textuais para revisão semântica.", ])
    (OUT / "analise_completa_preliminar.md").write_text("\n".join(relatorio), encoding="utf-8")
    print(f"Relatorio gerado: {OUT / 'analise_completa_preliminar.md'}")
    print(f"Base estruturada: {OUT / 'base_estruturada_heuristica.csv'}")


if __name__ == "__main__":
    main()
