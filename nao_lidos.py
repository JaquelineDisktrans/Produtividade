"""Registra o volume de nao lidos da Caixa de Entrada de trocas@disktrans.com.br.

Le o numero de nao lidos direto do Outlook classico (mesmo numero em negrito que
o Outlook mostra ao lado da pasta) e grava um historico em saida/nao_lidos.json:

- "atual": ultima leitura (valor + horario) -> alimenta o "agora" do painel.
- "dias": por dia, a foto das 18h e das 23:59 -> quanto sobrou no fim do expediente
  e no fim do dia.
- "amostras": ultimas leituras do "agora" (para uma linha de tendencia leve).

Uso:
    py -3.12 nao_lidos.py                 # marco "agora" (usado na rodada de 30min)
    py -3.12 nao_lidos.py --marco 18h     # foto das 18h (tarefa agendada)
    py -3.12 nao_lidos.py --marco 2359    # foto das 23:59 (tarefa agendada)

Nunca levanta excecao para o chamador: em erro, registra no proprio arquivo de log
e sai com codigo 0, para nao derrubar a atualizacao automatica.
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "saida"
ARQ = OUT / "nao_lidos.json"
LOG = OUT / "nao_lidos.log"

CAIXA_PADRAO = "trocas@disktrans.com.br"
OL_FOLDER_INBOX = 6
MAX_AMOSTRAS = 800


def log(msg: str) -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    with LOG.open("a", encoding="utf-8") as f:
        f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n")


def ler_nao_lidos() -> tuple[int, int]:
    """Retorna (nao_lidos, total) da Caixa de Entrada. So funciona no Windows/Outlook."""
    # Reaproveita a conexao e a busca de caixa ja usadas pelo coletor.
    from outlook_atendimento import abrir_outlook, localizar_caixa

    namespace = abrir_outlook()
    caixa = localizar_caixa(namespace, CAIXA_PADRAO)
    inbox = caixa.GetDefaultFolder(OL_FOLDER_INBOX)
    nao_lidos = int(inbox.UnReadItemCount)
    try:
        total = int(inbox.Items.Count)
    except Exception:
        total = None
    return nao_lidos, total


def carregar() -> dict:
    if ARQ.exists():
        try:
            return json.loads(ARQ.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"atual": {}, "dias": {}, "amostras": []}


def salvar(dados: dict) -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    ARQ.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")


def registrar(marco: str) -> int:
    agora = datetime.now()
    try:
        nao_lidos, total = ler_nao_lidos()
    except Exception as erro:
        log(f"ERRO ao ler nao lidos (marco={marco}): {erro}")
        return 0

    dados = carregar()
    dados.setdefault("dias", {})
    dados.setdefault("amostras", [])

    registro = {"valor": nao_lidos, "total": total, "em": agora.isoformat(timespec="seconds")}
    dados["atual"] = registro

    dados["amostras"].append({"ts": registro["em"], "valor": nao_lidos})
    dados["amostras"] = dados["amostras"][-MAX_AMOSTRAS:]

    if marco in ("18h", "2359"):
        dia = agora.strftime("%Y-%m-%d")
        chave = "as18" if marco == "18h" else "as2359"
        dia_reg = dados["dias"].setdefault(dia, {})
        dia_reg[chave] = nao_lidos
        dia_reg[f"{chave}_em"] = registro["em"]

    salvar(dados)
    log(f"OK marco={marco} nao_lidos={nao_lidos} total={total}")
    print(f"Nao lidos ({marco}): {nao_lidos}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Registra nao lidos da Caixa de Entrada.")
    parser.add_argument(
        "--marco",
        choices=["agora", "18h", "2359"],
        default="agora",
        help="agora = leitura corrente; 18h/2359 = foto do dia nesses horarios.",
    )
    args = parser.parse_args()
    return registrar(args.marco)


if __name__ == "__main__":
    raise SystemExit(main())
