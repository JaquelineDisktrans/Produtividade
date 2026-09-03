"""Divide casos_para_analise.jsonl em lotes de tamanho seguro para revisão."""

from __future__ import annotations

import argparse
from pathlib import Path


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--entrada", default="saida/casos_para_analise.jsonl")
    parser.add_argument("--saida", default="saida/lotes_claude")
    parser.add_argument("--casos-por-arquivo", type=int, default=100)
    args = parser.parse_args()
    entrada = Path(args.entrada)
    saida = Path(args.saida)
    if not entrada.exists():
        raise SystemExit(f"Arquivo nao encontrado: {entrada}")
    if args.casos_por_arquivo < 1:
        raise SystemExit("--casos-por-arquivo deve ser maior que zero")
    saida.mkdir(parents=True, exist_ok=True)
    numero = 0
    quantidade = 0
    arquivo = None
    try:
        for linha in entrada.open(encoding="utf-8"):
            if not quantidade:
                numero += 1
                arquivo = (saida / f"lote_{numero:04d}.jsonl").open("w", encoding="utf-8")
            arquivo.write(linha)
            quantidade += 1
            if quantidade >= args.casos_por_arquivo:
                arquivo.close()
                arquivo = None
                quantidade = 0
    finally:
        if arquivo is not None:
            arquivo.close()
    print(f"Criados {numero} lote(s) em {saida.resolve()}")


if __name__ == "__main__":
    main()
