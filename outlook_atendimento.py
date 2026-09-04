"""Gera indicadores de atendimento a partir do Outlook desktop.

O programa somente le mensagens do Outlook. Nenhuma mensagem e enviada, movida,
marcada ou alterada. Os dados intermediarios ficam em SQLite e os relatorios em CSV.

Correcoes importantes desta versao:
- Sem piso de periodo artificial: por padrao coleta TODO o historico acessivel.
  Use --desde/--ate apenas se quiser recortar.
- Varre TODAS as pastas da caixa (subpastas de recebidos + Itens Enviados), e nao
  somente a raiz da Caixa de Entrada. Assim e-mails ja arquivados em subpastas
  voltam a ser contabilizados. Itens Excluidos, Lixo Eletronico, Rascunhos e Caixa
  de Saida sao ignorados por padrao.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sqlite3
import sys
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path
from statistics import mean

# O win32com so existe no Windows com Outlook. A etapa de AGREGACAO/EXPORT nao
# depende dele, entao a ausencia nao pode derrubar o modulo: apenas marcamos a
# indisponibilidade e falhamos com mensagem clara quando a COLETA for chamada.
try:
    import pythoncom
    import win32com.client
    OUTLOOK_DISPONIVEL = True
except ImportError:
    pythoncom = None
    win32com = None
    OUTLOOK_DISPONIVEL = False


OL_FOLDER_INBOX = 6
OL_FOLDER_SENT_MAIL = 5
OL_FOLDER_DELETED = 3
OL_FOLDER_OUTBOX = 4
OL_FOLDER_DRAFTS = 16
OL_FOLDER_JUNK = 23
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
    if not OUTLOOK_DISPONIVEL:
        raise RuntimeError(
            "Dependencia ausente: pywin32/Outlook nao disponiveis neste ambiente. "
            "A coleta so roda no Windows com o Outlook classico. Use executar_analise.bat."
        )
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


def id_pasta_padrao(caixa, tipo: int) -> str | None:
    """EntryID de uma pasta padrao da caixa (Inbox, Sent, Deleted...), se existir."""
    try:
        return caixa.GetDefaultFolder(tipo).EntryID
    except Exception:
        return None


def restringir_itens_por_data(itens, campo_data: str, data_inicio: datetime | None):
    """Tenta filtrar no proprio Outlook para acelerar coletas incrementais.

    O formato de data aceito pelo Outlook varia conforme o idioma da instalacao.
    A validacao do primeiro item evita perder mensagens caso a instalacao rejeite
    o formato brasileiro; nesse cenario o coletor volta para a colecao original.
    """
    if not data_inicio:
        return itens
    try:
        filtro = data_inicio.strftime("%d/%m/%Y %H:%M")
        filtrados = itens.Restrict(f"[{campo_data}] >= '{filtro}'")
        if filtrados.Count == 0:
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


def coletar_pasta(
    conn: sqlite3.Connection,
    pasta,
    direcao: str,
    data_inicio: datetime | None,
    data_fim: datetime | None,
    existentes: set,
):
    """Le uma pasta, em streaming, e grava somente metadados essenciais.

    Percorre em ordem de data. Itens invalidos, sem assunto ou corrompidos nao
    interrompem a execucao: sao contados como ignorados e o coletor segue.
    """
    try:
        itens = pasta.Items
        campo_data = "[ReceivedTime]" if direcao == "recebido" else "[SentOn]"
        nome_campo_data = campo_data.strip("[]")
        # Ordem crescente de data (ponto 7: percorrer em ordem de data).
        itens.Sort(campo_data, False)
        total_geral = itens.Count
        itens = restringir_itens_por_data(itens, nome_campo_data, data_inicio)
        total = itens.Count
    except Exception:
        # Pasta sem itens de e-mail (ex.: calendario/contatos) ou inacessivel.
        return 0, 0

    if total == 0:
        return 0, 0

    # Feedback para pastas grandes (ex.: Itens Enviados), evitando que a leitura
    # pareca travada durante os minutos em que percorre milhares de itens.
    if total >= 200:
        print(f"  Lendo {pasta.Name} [{direcao}]: {total:,} itens...", flush=True)

    inseridos = 0
    ignorados = 0
    lote = []
    for indice in range(1, total + 1):
        try:
            if indice % 1000 == 0:
                print(f"    {pasta.Name}: {indice:,}/{total:,}...", flush=True)
            item = itens.Item(indice)
            if getattr(item, "Class", None) != OL_MAIL:
                ignorados += 1
                continue
            entry_id = item.EntryID
            if entry_id in existentes:
                continue
            data = item.ReceivedTime if direcao == "recebido" else item.SentOn
            data_naive = data.replace(tzinfo=None)
            if data_inicio and data_naive < data_inicio:
                continue
            if data_fim and data_naive > data_fim:
                continue
            assunto = item.Subject or "(sem assunto)"
            conversa_id = (item.ConversationID or "").strip()
            # Sem ConversationID, o assunto normalizado vira chave de reserva.
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
            existentes.add(entry_id)
            if len(lote) >= 250:
                _gravar_lote(conn, lote)
                inseridos += len(lote)
                lote.clear()
        except Exception:
            # Itens corrompidos ou em sincronizacao nao impedem o restante.
            ignorados += 1
    if lote:
        _gravar_lote(conn, lote)
        inseridos += len(lote)
    return inseridos, ignorados


def _gravar_lote(conn: sqlite3.Connection, lote: list):
    conn.executemany(
        "INSERT OR IGNORE INTO mensagens "
        "(entry_id, direcao, conversa_id, assunto_original, assunto_normalizado, "
        "data_hora, remetente, destinatarios, cc, anexos, corpo_texto) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", lote
    )
    conn.commit()


def coletar_caixa(
    conn: sqlite3.Connection,
    caixa,
    data_inicio: datetime | None,
    data_fim: datetime | None,
):
    """Varre TODAS as pastas relevantes da caixa (recursivo).

    Classificacao:
    - subtree de Itens Enviados -> 'enviado'
    - demais pastas -> 'recebido'
    Excluidas: Itens Excluidos, Lixo Eletronico, Rascunhos e Caixa de Saida.
    """
    id_enviados = id_pasta_padrao(caixa, OL_FOLDER_SENT_MAIL)
    ids_excluir = {
        id_pasta_padrao(caixa, OL_FOLDER_DELETED),
        id_pasta_padrao(caixa, OL_FOLDER_OUTBOX),
        id_pasta_padrao(caixa, OL_FOLDER_DRAFTS),
        id_pasta_padrao(caixa, OL_FOLDER_JUNK),
    }
    ids_excluir.discard(None)

    existentes = {linha[0] for linha in conn.execute("SELECT entry_id FROM mensagens")}
    totais = {"recebido": 0, "enviado": 0}
    ignorados_total = 0

    try:
        raiz = caixa.GetRootFolder()
    except Exception as erro:
        raise RuntimeError(f"Nao foi possivel abrir a caixa {caixa.DisplayName}.") from erro

    pilha = [(raiz, None)]  # (pasta, direcao_forcada)
    while pilha:
        pasta, direcao_forcada = pilha.pop()
        try:
            eid = pasta.EntryID
            nome = pasta.Name
        except Exception:
            continue
        if eid in ids_excluir:
            continue
        direcao = direcao_forcada
        if direcao is None:
            direcao = "enviado" if eid == id_enviados else "recebido"

        inseridos, ignorados = coletar_pasta(
            conn, pasta, direcao, data_inicio, data_fim, existentes
        )
        if inseridos or ignorados:
            print(f"  {nome} [{direcao}]: {inseridos:,} novos; {ignorados:,} ignorados.", flush=True)
        totais[direcao] += inseridos
        ignorados_total += ignorados

        # Empilha subpastas; sob Enviados herda 'enviado'.
        try:
            subpastas = pasta.Folders
            for i in range(1, subpastas.Count + 1):
                try:
                    sub = subpastas.Item(i)
                except Exception:
                    continue
                heranca = "enviado" if direcao == "enviado" else None
                pilha.append((sub, heranca))
        except Exception:
            pass

    print(
        f"Coleta concluida: {totais['recebido']:,} recebidos novos, "
        f"{totais['enviado']:,} enviados novos, {ignorados_total:,} ignorados."
    )
    return totais


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


def _escrever_csv(caminho: Path, campos: list[str], linhas: list[dict]):
    with caminho.open("w", newline="", encoding="utf-8-sig") as arquivo:
        writer = csv.DictWriter(arquivo, fieldnames=campos, delimiter=";")
        writer.writeheader()
        writer.writerows(linhas)


def exportar(conn: sqlite3.Connection, pasta_saida: Path):
    """Constroi conversas, assuntos consolidados e resumo mensal.

    Agrupamento por ConversationID (ponto 8). Todas as saidas de nivel de conversa
    trazem colunas de periodo (mes_ano, ano, mes) para o front-end filtrar meses.
    """
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
        data_receb = primeira[4]
        mes_ano = data_receb[:7]
        linhas_conversas.append({
            "conversa_id": chave,
            "assunto_original": primeira[2],
            "assunto_consolidado": primeira[3],
            "data_recebimento": data_receb,
            "mes_ano": mes_ano,
            "ano": data_receb[:4],
            "mes": data_receb[5:7],
            "primeiro_email_recebido": data_receb,
            "primeira_resposta_enviada": resposta[4] if resposta else "",
            "status": "Respondido" if resposta else "Sem resposta",
            "tempo_primeira_resposta_minutos": round(minutos, 2) if minutos is not None else "",
            "tempo_primeira_resposta_legivel": formatar_minutos(minutos),
            "emails_recebidos": len(recebidos),
            "emails_enviados": len(enviados),
            "remetente_primeiro_email": primeira[5],
        })

    linhas_conversas.sort(key=lambda linha: linha["primeiro_email_recebido"], reverse=True)

    # conversas_detalhadas.csv (requerido) + conversas.csv (compat do pipeline atual)
    campos_conv = list(linhas_conversas[0]) if linhas_conversas else [
        "conversa_id", "assunto_consolidado", "data_recebimento", "mes_ano", "status"
    ]
    _escrever_csv(pasta_saida / "conversas_detalhadas.csv", campos_conv, linhas_conversas)
    _escrever_csv(pasta_saida / "conversas.csv", campos_conv, linhas_conversas)

    # assuntos_consolidados.csv (agregado por assunto)
    assuntos = defaultdict(list)
    for linha in linhas_conversas:
        assuntos[linha["assunto_consolidado"]].append(linha)
    linhas_assuntos = []
    for assunto, linhas in assuntos.items():
        tempos = [float(l["tempo_primeira_resposta_minutos"]) for l in linhas if l["tempo_primeira_resposta_minutos"] != ""]
        datas = sorted(l["data_recebimento"] for l in linhas)
        linhas_assuntos.append({
            "assunto_consolidado": assunto,
            "conversas": len(linhas),
            "emails_recebidos": sum(l["emails_recebidos"] for l in linhas),
            "respondidas": sum(l["status"] == "Respondido" for l in linhas),
            "sem_resposta": sum(l["status"] == "Sem resposta" for l in linhas),
            "primeiro_mes": datas[0][:7],
            "ultimo_mes": datas[-1][:7],
            "tempo_medio_primeira_resposta_minutos": round(mean(tempos), 2) if tempos else "",
            "tempo_medio_primeira_resposta_legivel": formatar_minutos(mean(tempos)) if tempos else "",
        })
    linhas_assuntos.sort(key=lambda linha: linha["conversas"], reverse=True)
    campos_ass = list(linhas_assuntos[0]) if linhas_assuntos else ["assunto_consolidado", "conversas"]
    _escrever_csv(pasta_saida / "assuntos_consolidados.csv", campos_ass, linhas_assuntos)

    # resumo_mensal.csv (uma linha por mes)
    por_mes = defaultdict(list)
    for linha in linhas_conversas:
        por_mes[linha["mes_ano"]].append(linha)
    # e-mails recebidos/enviados por mes direto da base (nao so de conversas com recebido)
    recebidos_mes = dict(conn.execute(
        "SELECT substr(data_hora,1,7), COUNT(*) FROM mensagens WHERE direcao='recebido' GROUP BY 1"
    ).fetchall())
    enviados_mes = dict(conn.execute(
        "SELECT substr(data_hora,1,7), COUNT(*) FROM mensagens WHERE direcao='enviado' GROUP BY 1"
    ).fetchall())
    todos_meses = sorted(set(por_mes) | set(recebidos_mes) | set(enviados_mes))
    linhas_mensal = []
    for mes in todos_meses:
        convs = por_mes.get(mes, [])
        tempos = [float(l["tempo_primeira_resposta_minutos"]) for l in convs if l["tempo_primeira_resposta_minutos"] != ""]
        respondidas = sum(l["status"] == "Respondido" for l in convs)
        linhas_mensal.append({
            "mes_ano": mes,
            "ano": mes[:4],
            "mes": mes[5:7],
            "conversas": len(convs),
            "emails_recebidos": recebidos_mes.get(mes, 0),
            "emails_enviados": enviados_mes.get(mes, 0),
            "respondidas": respondidas,
            "sem_resposta": len(convs) - respondidas,
            "tempo_medio_primeira_resposta_minutos": round(mean(tempos), 2) if tempos else "",
            "tempo_medio_primeira_resposta_legivel": formatar_minutos(mean(tempos)) if tempos else "",
        })
    campos_mensal = [
        "mes_ano", "ano", "mes", "conversas", "emails_recebidos", "emails_enviados",
        "respondidas", "sem_resposta", "tempo_medio_primeira_resposta_minutos",
        "tempo_medio_primeira_resposta_legivel",
    ]
    _escrever_csv(pasta_saida / "resumo_mensal.csv", campos_mensal, linhas_mensal)

    # resumo.csv (compat: indicadores globais)
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
    return total_recebidos, len(linhas_conversas), linhas_mensal


def mes_seguinte(mes: str) -> str:
    ano, m = int(mes[:4]), int(mes[5:7])
    if m == 12:
        return f"{ano + 1}-01"
    return f"{ano}-{m + 1:02d}"


def relatorio_validacao(linhas_mensal: list[dict]):
    """Relatorio de validacao por mes (ponto 12). Destaca meses sem registros."""
    print("\n" + "=" * 60)
    print("RELATORIO DE VALIDACAO - mensagens e conversas por mes")
    print("=" * 60)
    if not linhas_mensal:
        print("Nenhum registro encontrado na base.")
        return
    print(f"{'Mes':<9}{'Conversas':>10}{'Recebidos':>11}{'Enviados':>10}")
    print("-" * 60)
    meses_presentes = {l["mes_ano"] for l in linhas_mensal}
    indexado = {l["mes_ano"]: l for l in linhas_mensal}
    primeiro = min(meses_presentes)
    ultimo = max(meses_presentes)
    # Percorre o intervalo continuo para revelar buracos (meses sem registros).
    cursor = primeiro
    meses_vazios = []
    while cursor <= ultimo:
        l = indexado.get(cursor)
        if l is None:
            print(f"{cursor:<9}{0:>10}{0:>11}{0:>10}   <== SEM REGISTROS")
            meses_vazios.append(cursor)
        else:
            print(f"{l['mes_ano']:<9}{l['conversas']:>10}{l['emails_recebidos']:>11}{l['emails_enviados']:>10}")
            if l["emails_recebidos"] == 0:
                print(f"{'':<9}{'':>10}{'':>11}{'':>10}   <== SEM E-MAILS RECEBIDOS")
        cursor = mes_seguinte(cursor)
    print("-" * 60)
    tot_conv = sum(l["conversas"] for l in linhas_mensal)
    tot_rec = sum(l["emails_recebidos"] for l in linhas_mensal)
    tot_env = sum(l["emails_enviados"] for l in linhas_mensal)
    print(f"{'TOTAL':<9}{tot_conv:>10}{tot_rec:>11}{tot_env:>10}")
    print(f"Periodo coberto: {primeiro} a {ultimo}.")
    if meses_vazios:
        print(f"ATENCAO: meses sem NENHUM registro: {', '.join(meses_vazios)}")
    else:
        print("Todos os meses do intervalo possuem registros.")
    print("=" * 60)


def _parse_data(valor: str | None, fim_do_dia: bool, parser) -> datetime | None:
    if not valor:
        return None
    try:
        d = datetime.strptime(valor, "%Y-%m-%d")
        return d.replace(hour=23, minute=59, second=59) if fim_do_dia else d
    except ValueError:
        parser.error("As datas devem usar o formato AAAA-MM-DD, por exemplo 2026-01-01.")


def main():
    parser = argparse.ArgumentParser(description="Analisa atendimento no Outlook desktop.")
    parser.add_argument(
        "--desde", default=None,
        help="Data inicial AAAA-MM-DD (opcional). Sem valor: coleta todo o historico acessivel.",
    )
    parser.add_argument(
        "--ate", default=None,
        help="Data final AAAA-MM-DD (opcional). Sem valor: ate a mensagem mais recente.",
    )
    parser.add_argument(
        "--caixa", default="trocas@disktrans.com.br",
        help="Nome ou e-mail da caixa configurada no Outlook. Padrao: trocas@disktrans.com.br.",
    )
    parser.add_argument("--saida", default="saida", help="Pasta onde os relatorios serao gravados.")
    parser.add_argument("--limpar-base", action="store_true", help="Apaga somente o banco local antes de iniciar.")
    parser.add_argument(
        "--reconstruir",
        action="store_true",
        help="Recoleta TODO o historico numa base temporaria e so substitui a base boa "
             "quando a coleta termina inteira (troca atomica). Fechar a janela no meio nao "
             "corrompe a base nem o painel anteriores.",
    )
    parser.add_argument(
        "--incremental",
        action="store_true",
        help="Le apenas mensagens desde a ultima coleta, com margem de dois dias para sincronizacoes tardias.",
    )
    parser.add_argument(
        "--somente-exportar",
        action="store_true",
        help="Nao acessa o Outlook: apenas reexporta CSVs e o relatorio a partir do banco local.",
    )
    args = parser.parse_args()
    data_inicio = _parse_data(args.desde, fim_do_dia=False, parser=parser)
    data_fim = _parse_data(args.ate, fim_do_dia=True, parser=parser)

    pasta_programa = Path(__file__).resolve().parent
    pasta_saida = (pasta_programa / args.saida).resolve()
    pasta_saida.mkdir(parents=True, exist_ok=True)
    banco_final = pasta_saida / "outlook_atendimento.sqlite"

    # Em --reconstruir, a coleta vai para uma base temporaria e so substitui a
    # base boa apos terminar inteira. Assim, interromper a coleta (fechar a
    # janela) nunca deixa a base/painel pela metade.
    usar_staging = args.reconstruir and not args.somente_exportar
    banco_coleta = (pasta_saida / "outlook_atendimento.rebuild.sqlite") if usar_staging else banco_final

    # Remove sobras de uma reconstrucao anterior interrompida.
    resto_staging = pasta_saida / "outlook_atendimento.rebuild.sqlite"
    if usar_staging and resto_staging.exists():
        resto_staging.unlink()
    if args.limpar_base and not usar_staging and banco_final.exists():
        banco_final.unlink()

    conn = sqlite3.connect(banco_coleta)
    garantir_schema(conn)
    swap_ok = False
    try:
        if not args.somente_exportar:
            namespace = abrir_outlook()
            caixa = localizar_caixa(namespace, args.caixa)
            periodo = "todo o historico" if not data_inicio else f"{args.desde} em diante"
            if data_fim:
                periodo += f" ate {args.ate}"
            print(f"Caixa selecionada: {caixa.DisplayName}; periodo: {periodo}.")
            limite_inicio = data_inicio
            # Incremental so faz sentido sobre a base existente, nunca sobre staging.
            if args.incremental and not usar_staging:
                ultima = conn.execute("SELECT MAX(data_hora) FROM mensagens").fetchone()[0]
                if ultima:
                    try:
                        base = datetime.fromisoformat(ultima) - timedelta(days=2)
                        limite_inicio = max(base, data_inicio) if data_inicio else base
                    except ValueError:
                        pass
                if limite_inicio:
                    print(f"Modo incremental: coletando desde {limite_inicio:%Y-%m-%d %H:%M}.")
            print("Varrendo todas as pastas da caixa (recebidos + enviados)...")
            coletar_caixa(conn, caixa, limite_inicio, data_fim)

            if usar_staging:
                # Troca atomica: so chega aqui se a coleta terminou por completo.
                conn.commit()
                conn.close()
                os.replace(banco_coleta, banco_final)
                swap_ok = True
                print("Base reconstruida por completo (troca atomica concluida).")
                conn = sqlite3.connect(banco_final)

        total, conversas, linhas_mensal = exportar(conn, pasta_saida)
        relatorio_validacao(linhas_mensal)
        print(f"\nPronto. {total:,} e-mails recebidos e {conversas:,} conversas analisadas.")
        print(f"Abra a pasta de resultados: {pasta_saida}")
    except RuntimeError as erro:
        print(f"\nERRO: {erro}", file=sys.stderr)
        return 2
    finally:
        try:
            conn.close()
        except Exception:
            pass
        # Coleta interrompida ou com falha antes da troca: descarta a base
        # temporaria e preserva a base boa e o painel anteriores.
        if usar_staging and not swap_ok and resto_staging.exists():
            try:
                resto_staging.unlink()
            except OSError:
                pass
        if OUTLOOK_DISPONIVEL and not args.somente_exportar:
            pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main() or 0)
