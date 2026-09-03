"""Prepara dados agregados para o dashboard local do Controle CS."""

from __future__ import annotations

import json
import sqlite3
from collections import Counter
from datetime import datetime
from pathlib import Path

from gerar_relatorio import analisar_casos, carregar_casos, quantile


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "saida"


def tempo_legivel(minutos):
    if minutos is None:
        return "Não identificável"
    horas, mins = divmod(round(minutos), 60)
    return f"{horas}h {mins:02d}min"


def counter_rows(counter, total=None, label="label"):
    return [
        {label: chave, "value": valor, "percent": round((valor / total * 100), 1) if total else 0}
        for chave, valor in counter.most_common()
    ]


def main():
    casos = carregar_casos()
    linhas = analisar_casos(casos)
    with sqlite3.connect(OUT / "outlook_atendimento.sqlite") as conn:
        recebidos_caixa = conn.execute("SELECT COUNT(*) FROM mensagens WHERE direcao='recebido'").fetchone()[0]
        enviados_caixa = conn.execute("SELECT COUNT(*) FROM mensagens WHERE direcao='enviado'").fetchone()[0]
    respostas = [float(l["tempo_primeira_resposta_minutos"]) for l in linhas if l["tempo_primeira_resposta_minutos"] != ""]
    recebidos = [datetime.fromisoformat(l["data_inicial"]) for l in linhas]
    respondidos = sum(l["status_aparente"] == "Respondido" for l in linhas)
    sem_resposta = len(linhas) - respondidos
    referencia = max(recebidos).isoformat() if recebidos else None
    dados = {
        "meta": {
            "titulo": "Controle CS",
            "caixa": "trocas@disktrans.com.br",
            "periodo": "01/07/2026 em diante",
            "gerado_em": datetime.now().isoformat(timespec="seconds"),
            "referencia_backlog": referencia,
            "observacao": "Motivos, gargalos e qualidade são heurísticas e devem ser validados semanticamente.",
        },
        "kpis": {
            "casos": len(linhas),
            "recebidos": recebidos_caixa,
            "enviados": enviados_caixa,
            "respondidos": respondidos,
            "sem_resposta": sem_resposta,
            "taxa_resposta": round(respondidos / len(linhas) * 100, 1) if linhas else 0,
            "tempo_medio": round(sum(respostas) / len(respostas), 2) if respostas else None,
            "tempo_medio_legivel": tempo_legivel(sum(respostas) / len(respostas)) if respostas else "Não identificável",
            "tempo_mediano": round(sorted(respostas)[len(respostas) // 2], 2) if respostas else None,
            "tempo_mediano_legivel": tempo_legivel(sorted(respostas)[len(respostas) // 2]) if respostas else "Não identificável",
            "p75": round(quantile(respostas, .75), 2) if respostas else None,
            "p90": round(quantile(respostas, .90), 2) if respostas else None,
            "p95": round(quantile(respostas, .95), 2) if respostas else None,
            "cobranca": sum(int(l["cobrancas_cliente"]) > 0 for l in linhas),
            "evitaveis": sum(l["contato_potencialmente_evitavel"] == "Sim" for l in linhas),
        },
        "volume": {
            "diario": counter_rows(Counter(l["data_inicial"][:10] for l in linhas), label="date"),
            "semanal": counter_rows(Counter(datetime.fromisoformat(l["data_inicial"]).strftime("%Y-S%W") for l in linhas), label="week"),
            "mensal": counter_rows(Counter(l["data_inicial"][:7] for l in linhas), label="month"),
            "dia_semana": counter_rows(Counter(datetime.fromisoformat(l["data_inicial"]).strftime("%A") for l in linhas), label="day"),
            "hora": counter_rows(Counter(f"{datetime.fromisoformat(l['data_inicial']).hour:02d}:00" for l in linhas), label="hour"),
        },
        "motivos": counter_rows(Counter(l["motivo_heuristico"] for l in linhas), len(linhas), "label"),
        "gargalos": counter_rows(Counter(l["principal_gargalo_heuristico"] for l in linhas), len(linhas), "label"),
        "clientes": counter_rows(Counter(l["cliente"] for l in linhas if l["cliente"] != "não identificado"), len(linhas), "label")[:10],
        "unidades": counter_rows(Counter(l["unidade"] for l in linhas if l["unidade"] != "não identificada"), len(linhas), "label")[:10],
        "casos": [
            {
                "id": l["case_id"],
                "data": l["data_inicial"],
                "cliente": l["cliente"],
                "unidade": l["unidade"],
                "assunto": l["assunto"],
                "motivo": l["motivo_heuristico"],
                "status": l["status_aparente"],
                "primeira_resposta": l["primeira_resposta"],
                "tempo_resposta": l["tempo_primeira_resposta_minutos"],
                "interacoes": l["interacoes"],
                "cobrancas": l["cobrancas_cliente"],
                "gargalo": l["principal_gargalo_heuristico"],
                "reclamacao": l["houve_reclamacao_heuristica"],
                "evitavel": l["contato_potencialmente_evitavel"],
                "qualidade": l["qualidade_heuristica_1a5"],
            }
            for l in linhas
        ],
    }
    destino = OUT / "dashboard_data.json"
    destino.write_text(json.dumps(dados, ensure_ascii=False), encoding="utf-8")
    print(f"Dados do dashboard gerados em {destino}")


if __name__ == "__main__":
    main()
