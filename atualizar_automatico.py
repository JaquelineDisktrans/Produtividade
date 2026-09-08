"""Atualiza a base local do Outlook sem reler todo o histórico a cada execução."""

from __future__ import annotations

import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "saida"
LOCK = OUT / ".atualizacao_em_andamento"
LOG = OUT / "atualizacao.log"

# Se uma rodada for interrompida, a trava fica para tras e bloquearia todas as
# proximas. Passado esse tempo, consideramos a trava obsoleta e seguimos.
TRAVA_OBSOLETA_SEGUNDOS = 2 * 60 * 60  # 2 horas


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


def capturar_nao_lidos(arquivo_log, marco: str):
    """Registra o volume de nao lidos. Nunca interrompe a atualização."""
    try:
        resultado = subprocess.run(
            [sys.executable, str(ROOT / "nao_lidos.py"), "--marco", marco],
            cwd=ROOT,
            stdout=arquivo_log,
            stderr=subprocess.STDOUT,
            text=True,
        )
        if resultado.returncode != 0:
            escrever_log(arquivo_log, f"Captura de nao lidos retornou {resultado.returncode} (ignorado).")
    except Exception as erro:
        escrever_log(arquivo_log, f"Captura de nao lidos ignorada: {erro}")


def publicar_online(arquivo_log):
    """Envia o saida/dashboard_data.json ao GitHub. Nunca interrompe a atualização."""
    script = ROOT / "publicar_dados.bat"
    if not script.exists():
        escrever_log(arquivo_log, "publicar_dados.bat ausente; publicação online ignorada.")
        return
    escrever_log(arquivo_log, "Publicando dados no GitHub...")
    try:
        resultado = subprocess.run(
            [str(script)],
            cwd=ROOT,
            stdout=arquivo_log,
            stderr=subprocess.STDOUT,
            text=True,
            shell=True,
        )
        if resultado.returncode != 0:
            escrever_log(
                arquivo_log,
                f"Publicação retornou código {resultado.returncode} (ignorado).",
            )
    except Exception as erro:
        escrever_log(arquivo_log, f"Publicação online ignorada: {erro}")


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    # Remove trava obsoleta de uma rodada anterior que foi interrompida, para
    # não bloquear as próximas execuções indefinidamente.
    if LOCK.exists():
        try:
            if time.time() - LOCK.stat().st_mtime > TRAVA_OBSOLETA_SEGUNDOS:
                LOCK.unlink()
        except OSError:
            pass
    try:
        descritor = os.open(str(LOCK), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        # Uma execução anterior ainda está trabalhando (trava recente); a próxima
        # rodada do Agendador do Windows fará nova tentativa.
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
            capturar_nao_lidos(arquivo_log, "agora")
            executar(arquivo_log, "gerar_dashboard_data.py")
            publicar_online(arquivo_log)
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
