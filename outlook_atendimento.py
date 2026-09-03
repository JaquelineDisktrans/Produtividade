"""Gera indicadores de atendimento a partir do Outlook desktop.

O programa somente le mensagens do Outlook. Nenhuma mensagem e enviada, movida
ou alterada. Os dados intermediarios ficam em SQLite e os relatorios em CSV.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sqlite3
import sys
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path
from statistics import mean

try:
    import pythoncom
    import win32com.client
except ImportError:
    print("Dependencia ausente: pywin32. Execute o arquivo executar_analise.bat.")
    raise SystemExit(1)


OL_FOLDER_INBOX = 6
OL_FOLDER_SENT_MAIL = 5
OL_MAIL = 43
CORPO_MAX_CHARS = 60000
SCHEMA = """
CREATE TABLE IF NOT EXISTS mensagens (
    entry_id TEXT PRIMARY KEY,
    direcao TEXT NOT NULL,
    conversa_id TEXT,
    assunto_original TEXT,
    assunto_normalizado TEXT,
    data_hora TEXT NOT NULL,
    remetente TEXT,
    destinatarios TEXT,
    cc TEXT,
    anexos INTEGER DEFAULT 0,
    corpo_texto TEXT
);
CREATE INDEX IF NOT EXISTS ix_mensagens_conversa ON mensagens(conversa_id, data_hora);
CREATE INDEX IF NOT EXISTS ix_mensagens_assunto ON mensagens(assunto_normalizado, data_hora);
"""


def garantir_schema(conn: sqlite3.Connection):
    conn.executescript(SCHEMA)
    colunas = {linha[1] for linha in conn.execute("PRAGMA table_info(mensagens)")}
    extras = {
        "destinatarios": "TEXT",
        "cc": "TEXT",
        "anexos": "INTEGER DEFAULT 0",
        "corpo_texto": "TEXT",
    }
    for nome, tipo in extras.items():
        if nome not in colunas:
            conn.execute(f"ALTER TABLE mensagens ADD COLUMN {nome} {tipo}")
    conn.commit()


def normalizar_assunto(assunto: str | None) -> str:
    """Remove prefixos comuns para consolidar RE:, RES:, FW: e ENC:."""
    texto = (assunto or "(sem assunto)").strip()
    padrao = re.compile(r"^\s*(?:(?:re|res|fw|fwd|enc)\s*:\s*)+", re.IGNORECASE)
    while True:
        novo = padrao.sub("", texto).strip()
        if novo == texto:
            break
        texto = novo
    return re.sub(r"\s+", " ", texto) or "(sem assunto)"


def iso_data(valor) -> str:
    """Converte a data COM do Outlook em texto ISO, sem depender do idioma local."""
    try:
        return valor.strftime("%Y-%m-%dT%H:%M:%S")
    except (AttributeError, ValueError):
        return str(valor)


def smtp_do_remetente(item) -> str:
    """Tenta obter SMTP mesmo para remetentes internos do Exchange."""
    try:
        endereco = item.SenderEmailAddress or ""
        if item.SenderEmailType != "EX":
            return endereco
        return item.PropertyAccessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x5D01001E"
        ) or endereco
    except Exception:
        return ""


def texto_do_item(item) -> str:
    try:
        corpo = (item.Body or "").replace("\x00", "").replace("\r\n", "\n").strip()
        if len(corpo) > CORPO_MAX_CHARS:
            corpo = corpo[:CORPO_MAX_CHARS] + "\n[corpo truncado pelo coletor]"
        return corpo
    except Exception:
        return ""


def abrir_outlook():
    pythoncom.CoInitialize()
    try:
        # Em execucao pelo duplo clique, o Outlook e um processo interativo
        # acessivel. Tentamos anexar primeiro; caso ele nao esteja registrado
        # no ROT, o Dispatch usa a abertura normal do Outlook classico.
        try:
            app = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            app = win32com.client.Dispatch("Outlook.Application")
        return app.GetNamespace("MAPI")
    except Exception as erro:
        raise RuntimeError(
            "Nao foi possivel conectar ao Outlook classico. Abra-o pelo menu Iniciar "
            "(nao o 'Novo Outlook'), confirme que a caixa Trocas esta sincronizada, "
            "deixe a janela aberta e execute novamente."
        ) from erro


def listar_caixas(namespace) -> list[str]:
    return [namespace.Stores.Item(i).DisplayName for i in range(1, namespace.Stores.Count + 1)]


def localizar_caixa(namespace, termo: str):
    """Localiza uma conta/caixa configurada, sem cair silenciosamente na conta padrao."""
    procurado = termo.strip().casefold()
    encontrados = []
    for indice in range(1, namespace.Stores.Count + 1):
        caixa = namespace.Stores.Item(indice)
        if procurado in (caixa.DisplayName or "").casefold():
            encontrados.append(caixa)
    if len(encontrados) == 1:
        return encontrados[0]
    disponiveis = "\n- ".join(listar_caixas(namespace)) or "(nenhuma caixa encontrada)"
    if not encontrados:
        raise RuntimeError(
            f"Nao encontrei a caixa '{termo}' no Outlook. Caixas disponiveis:\n- {disponiveis}"
        )
    raise RuntimeError(
        f"Encontrei mais de uma caixa para '{termo}'. Use um nome mais especifico.\n- {disponiveis}"
    )


def restringir_itens_por_data(itens, campo_data: str, data_inicio: datetime | None):
    """Tenta filtrar no próprio Outlook para que as atualizações sejam incrementais.

    O formato de data aceito pelo Outlook varia conforme o idioma da instalação.
    A validação do primeiro item evita perder mensagens caso a instalação rejeite
    o formato brasileiro; nesse cenário o coletor volta para a coleção original.
    """
    if not data_inicio:
        return itens
    try:
        filtro = data_inicio.strftime("%d/%m/%Y %H:%M")
        filtrados = itens.Restrict(f"[{campo_data}] >= '{filtro}'")
        if filtrados.Count == 0:
            # Se há item recente na coleção original, o filtro provavelmente não
            # foi interpretado. Repetir a leitura completa é mais seguro que
            # omitir mensagens novas.
            try:
                ultimo = itens.Item(1)
                data_ultimo = getattr(ultimo, campo_data).replace(tzinfo=None)
                if data_ultimo >= data_inicio:
                    return itens
            except Exception:
                pass
        else:
            try:
                primeiro = filtrados.Item(1)
                data_primeiro = getattr(primeiro, campo_data).replace(tzinfo=None)
                if data_primeiro < data_inicio - timedelta(minutes=1):
                    return itens
            except Exception:
                pass
        return filtrados
    except Exception:
        return itens


def coletar_pasta(conn: sqlite3.Connection, pasta, direcao: str, data_inicio: datetime | None):
    """Le uma pasta, em streaming, e grava somente metadados essenciais."""
    try:
        itens = pasta.Items
        campo_data = "[ReceivedTime]" if direcao == "recebido" else "[SentOn]"
        nome_campo_data = campo_data.strip("[]")
        itens.Sort(campo_data, True)
        total_geral = itens.Count
        itens = restringir_itens_por_data(itens, nome_campo_data, data_inicio)
        total = itens.Count
    except Exception as erro:
        raise RuntimeError(f"Nao foi possivel ler a pasta {pasta.Name}.") from erro

    if total != total_geral:
        print(f"Lendo {pasta.Name}: {total:,} itens na janela incremental (de {total_geral:,}).")
    else:
        print(f"Lendo {pasta.Name}: {total:,} itens encontrados.")
    inseridos = 0
    ignorados = 0
    lote = []
    existentes = {linha[0] for linha in conn.execute("SELECT entry_id FROM mensagens")}
    # A colecao COM e indexada a partir de 1.
    for indice in range(1, total + 1):
        try:
            item = itens.Item(indice)
            if item.Class != OL_MAIL:
                ignorados += 1
                continue
            entry_id = item.EntryID
            if entry_id in existentes:
                continue
            data = item.ReceivedTime if direcao == "recebido" else item.SentOn
            if data_inicio and data.replace(tzinfo=None) < data_inicio:
                continue
            assunto = item.Subject or "(sem assunto)"
            conversa_id = (item.ConversationID or "").strip()
            # Alguns itens nao possuem ConversationID. O assunto vira a chave de reserva.
            chave = conversa_id or "assunto:" + normalizar_assunto(assunto).casefold()
            try:
                anexos = int(item.Attachments.Count)
            except Exception:
                anexos = 0
            lote.append(
                (
                    entry_id,
                    direcao,
                    chave,
                    assunto,
                    normalizar_assunto(assunto),
                    iso_data(data),
                    smtp_do_remetente(item) if direcao == "recebido" else (item.SenderEmailAddress or ""),
                    item.To or "",
                    item.CC or "",
                    anexos,
                    texto_do_item(item),
                )
            )
            if len(lote) >= 250:
                conn.executemany(
                    "INSERT OR IGNORE INTO mensagens "
                    "(entry_id, direcao, conversa_id, assunto_original, assunto_normalizado, "
                    "data_hora, remetente, destinatarios, cc, anexos, corpo_texto) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", lote
                )
                conn.commit()
                inseridos += len(lote)
                existentes.update(linha[0] for linha in lote)
                lote.clear()
            if indice % 1000 == 0:
                print(f"  {indice:,}/{total:,} lidos")
        except Exception:
            # Itens corrompidos ou em sincronizacao nao impedem o restante da analise.
            ignorados += 1
    if lote:
        conn.executemany(
            "INSERT OR IGNORE INTO mensagens "
            "(entry_id, direcao, conversa_id, assunto_original, assunto_normalizado, "
            "data_hora, remetente, destinatarios, cc, anexos, corpo_texto) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", lote
        )
        conn.commit()
        inseridos += len(lote)
        existentes.update(linha[0] for linha in lote)
    print(f"  Concluido: {inseridos:,} processados; {ignorados:,} ignorados.")


def minutos_entre(inicio: str, fim: str) -> float:
    return (datetime.fromisoformat(fim) - datetime.fromisoformat(inicio)).total_seconds() / 60


def formatar_minutos(valor: float | None) -> str:
    if valor is None:
        return ""
    horas, minutos = divmod(round(valor), 60)
    return f"{horas}h {minutos:02d}min"


def exportar_casos_para_analise(conn: sqlite3.Connection, pasta_saida: Path):
    """Exporta uma linha JSON por conversa para classificacao semantica em lotes."""
    consulta = (
        "SELECT conversa_id, direcao, assunto_original, assunto_normalizado, data_hora, "
        "remetente, destinatarios, cc, anexos, corpo_texto FROM mensagens "
        "WHERE conversa_id IN (SELECT conversa_id FROM mensagens WHERE direcao='recebido') "
        "ORDER BY conversa_id, data_hora"
    )
    destino = pasta_saida / "casos_para_analise.jsonl"
    with destino.open("w", encoding="utf-8") as arquivo:
        chave_atual = None
        caso = None
        for row in conn.execute(consulta):
            chave, direcao, assunto, assunto_norm, data, remetente, para, cc, anexos, corpo = row
            if chave != chave_atual:
                if caso is not None:
                    arquivo.write(json.dumps(caso, ensure_ascii=False) + "\n")
                chave_atual = chave
                caso = {
                    "case_id": chave,
                    "assunto_consolidado": assunto_norm,
                    "mensagens": [],
                }
            caso["mensagens"].append({
                "direcao": direcao,
                "data_hora": data,
                "assunto": assunto,
                "remetente": remetente or "",
                "destinatarios": para or "",
                "cc": cc or "",
                "anexos": anexos or 0,
                "corpo": corpo or "",
            })
        if caso is not None:
            arquivo.write(json.dumps(caso, ensure_ascii=False) + "\n")
    return destino


def exportar(conn: sqlite3.Connection, pasta_saida: Path):
    registros = conn.execute(
        "SELECT conversa_id, direcao, assunto_original, assunto_normalizado, data_hora, remetente "
        "FROM mensagens ORDER BY conversa_id, data_hora"
    ).fetchall()
    conversas = defaultdict(list)
    for registro in registros:
        conversas[registro[0]].append(registro)

    linhas_conversas = []
    for chave, mensagens in conversas.items():
        recebidos = [m for m in mensagens if m[1] == "recebido"]
        enviados = [m for m in mensagens if m[1] == "enviado"]
        if not recebidos:
            continue
        primeira = recebidos[0]
        respostas = [m for m in enviados if m[4] >= primeira[4]]
        resposta = respostas[0] if respostas else None
        minutos = minutos_entre(primeira[4], resposta[4]) if resposta else None
        linhas_conversas.append({
            "conversa_id": chave,
            "assunto": primeira[3],
            "primeiro_email_recebido": primeira[4],
            "primeira_resposta_enviada": resposta[4] if resposta else "",
            "status": "Respondido" if resposta else "Sem resposta",
            "tempo_primeira_resposta_minutos": round(minutos, 2) if minutos is not None else "",
            "tempo_primeira_resposta_legivel": formatar_minutos(minutos),
            "emails_recebidos": len(recebidos),
            "emails_enviados": len(enviados),
            "remetente_primeiro_email": primeira[5],
        })

    linhas_conversas.sort(key=lambda linha: linha["primeiro_email_recebido"], reverse=True)
    with (pasta_saida / "conversas.csv").open("w", newline="", encoding="utf-8-sig") as arquivo:
        campos = list(linhas_conversas[0]) if linhas_conversas else ["conversa_id", "assunto", "status"]
        writer = csv.DictWriter(arquivo, fieldnames=campos, delimiter=";")
        writer.writeheader()
        writer.writerows(linhas_conversas)

    assuntos = defaultdict(list)
    for linha in linhas_conversas:
        assuntos[linha["assunto"]].append(linha)
    linhas_assuntos = []
    for assunto, linhas in assuntos.items():
        tempos = [float(l["tempo_primeira_resposta_minutos"]) for l in linhas if l["tempo_primeira_resposta_minutos"] != ""]
        linhas_assuntos.append({
            "assunto_consolidado": assunto,
            "conversas": len(linhas),
            "emails_recebidos": sum(l["emails_recebidos"] for l in linhas),
            "respondidas": sum(l["status"] == "Respondido" for l in linhas),
            "sem_resposta": sum(l["status"] == "Sem resposta" for l in linhas),
            "tempo_medio_primeira_resposta_minutos": round(mean(tempos), 2) if tempos else "",
            "tempo_medio_primeira_resposta_legivel": formatar_minutos(mean(tempos)) if tempos else "",
        })
    linhas_assuntos.sort(key=lambda linha: linha["conversas"], reverse=True)
    with (pasta_saida / "assuntos_consolidados.csv").open("w", newline="", encoding="utf-8-sig") as arquivo:
        campos = list(linhas_assuntos[0]) if linhas_assuntos else ["assunto_consolidado", "conversas"]
        writer = csv.DictWriter(arquivo, fieldnames=campos, delimiter=";")
        writer.writeheader()
        writer.writerows(linhas_assuntos)

    total_recebidos = sum(1 for r in registros if r[1] == "recebido")
    total_enviados = sum(1 for r in registros if r[1] == "enviado")
    respondidas = sum(l["status"] == "Respondido" for l in linhas_conversas)
    tempos = [float(l["tempo_primeira_resposta_minutos"]) for l in linhas_conversas if l["tempo_primeira_resposta_minutos"] != ""]
    resumo = [
        ("E-mails recebidos", total_recebidos),
        ("E-mails enviados", total_enviados),
        ("Conversas recebidas", len(linhas_conversas)),
        ("Conversas respondidas", respondidas),
        ("Conversas sem resposta", len(linhas_conversas) - respondidas),
        ("Percentual respondido", f"{(respondidas / len(linhas_conversas) * 100):.2f}%" if linhas_conversas else "0%"),
        ("Tempo medio primeira resposta", formatar_minutos(mean(tempos)) if tempos else "Sem respostas"),
    ]
    with (pasta_saida / "resumo.csv").open("w", newline="", encoding="utf-8-sig") as arquivo:
        writer = csv.writer(arquivo, delimiter=";")
        writer.writerow(["indicador", "valor"])
        writer.writerows(resumo)
    exportar_casos_para_analise(conn, pasta_saida)
    return total_recebidos, len(linhas_conversas)


def main():
    parser = argparse.ArgumentParser(description="Analisa atendimento no Outlook desktop.")
    parser.add_argument(
        "--desde", default="2026-07-01",
        help="Data inicial no formato AAAA-MM-DD. Padrao: 2026-07-01.",
    )
    parser.add_argument(
        "--caixa", default="trocas@disktrans.com.br",
        help="Nome ou e-mail da caixa configurada no Outlook. Padrao: trocas@disktrans.com.br.",
    )
    parser.add_argument("--saida", default="saida", help="Pasta onde os relatorios serao gravados.")
    parser.add_argument("--limpar-base", action="store_true", help="Apaga somente o banco local antes de iniciar.")
    parser.add_argument(
        "--incremental",
        action="store_true",
        help="Lê apenas mensagens desde a última coleta, com margem de dois dias para sincronizações tardias.",
    )
    args = parser.parse_args()
    try:
        data_inicio = datetime.strptime(args.desde, "%Y-%m-%d") if args.desde else None
    except ValueError:
        parser.error("--desde deve usar o formato AAAA-MM-DD, por exemplo 2026-01-01.")

    pasta_programa = Path(__file__).resolve().parent
    pasta_saida = (pasta_programa / args.saida).resolve()
    pasta_saida.mkdir(parents=True, exist_ok=True)
    banco = pasta_saida / "outlook_atendimento.sqlite"
    if args.limpar_base and banco.exists():
        banco.unlink()
    conn = sqlite3.connect(banco)
    garantir_schema(conn)
    try:
        namespace = abrir_outlook()
        caixa = localizar_caixa(namespace, args.caixa)
        print(f"Caixa selecionada: {caixa.DisplayName}; periodo: {args.desde} em diante.")
        limites = {"recebido": data_inicio, "enviado": data_inicio}
        if args.incremental:
            for direcao in limites:
                ultima = conn.execute(
                    "SELECT MAX(data_hora) FROM mensagens WHERE direcao = ?", (direcao,)
                ).fetchone()[0]
                if ultima:
                    try:
                        limites[direcao] = max(
                            data_inicio,
                            datetime.fromisoformat(ultima) - timedelta(days=2),
                        )
                    except ValueError:
                        pass
            print(
                "Modo incremental: recebidos desde "
                f"{limites['recebido']:%Y-%m-%d %H:%M} e enviados desde "
                f"{limites['enviado']:%Y-%m-%d %H:%M}."
            )
        coletar_pasta(conn, caixa.GetDefaultFolder(OL_FOLDER_INBOX), "recebido", limites["recebido"])
        coletar_pasta(conn, caixa.GetDefaultFolder(OL_FOLDER_SENT_MAIL), "enviado", limites["enviado"])
        total, conversas = exportar(conn, pasta_saida)
        print(f"\nPronto. {total:,} e-mails recebidos e {conversas:,} conversas analisadas.")
        print(f"Abra a pasta de resultados: {pasta_saida}")
    except RuntimeError as erro:
        print(f"\nERRO: {erro}", file=sys.stderr)
        return 2
    finally:
        conn.close()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main() or 0)
