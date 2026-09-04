"""Atualiza a base local do Outlook sem reler todo o histórico a cada execução."""

from __future__ import annotations

import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "saida"
LOCK = OUT / ".atualizacao_em_andamento"
LOG = OUT / "atualizacao.log"


def escrever_log(arquivo, mensagem: str):
    arquivo.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {mensagem}\n")
    arquivo.flush()


def executar(arquivo_log, script: str, *argumentos: str):
    comando = [sys.executable, str(ROOT / script), *argumentos]
    escrever_log(arquivo_log, f"Iniciando: {' '.join(comando)}")
    resultado = subprocess.run(
        comando,
        cwd=ROOT,
        stdout=arquivo_log,
        stderr=subprocess.STDOUT,
        text=True,
    )
    if resultado.returncode != 0:
        raise RuntimeError(f"{script} terminou com código {resultado.returncode}.")


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    try:
        descritor = os.open(str(LOCK), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        # Uma execução anterior ainda está trabalhando; a próxima rodada do
        # Agendador do Windows fará nova tentativa.
        return 0

    try:
        with os.fdopen(descritor, "w", encoding="utf-8") as marcador, LOG.open(
            "a", encoding="utf-8"
        ) as arquivo_log:
            marcador.write(str(os.getpid()))
            escrever_log(arquivo_log, "--- Atualização automática iniciada ---")
            executar(
                arquivo_log,
                "outlook_atendimento.py",
                "--caixa",
                "trocas@disktrans.com.br",
                "--incremental",
            )
            executar(arquivo_log, "gerar_relatorio.py")
            executar(arquivo_log, "gerar_dashboard_data.py")
            escrever_log(arquivo_log, "Atualização concluída com sucesso.")
        return 0
    except Exception as erro:
        with LOG.open("a", encoding="utf-8") as arquivo_log:
            escrever_log(arquivo_log, f"ERRO: {erro}")
        return 1
    finally:
        try:
            LOCK.unlink()
        except FileNotFoundError:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
