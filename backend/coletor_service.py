from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
import re
import shutil
import subprocess
import time

from decimal import Decimal, InvalidOperation

from sqlalchemy import func, or_, select
from sqlalchemy.orm import Session

from backend.config import get_settings
from backend.models import Caixa, Coletor, EstoqueAgregado, LeituraColetor, SessaoConferencia


WEDGE_RE = re.compile(r'DCS-Wedge.*?\+\+wedge\("(?P<codigo>[^"]+)"')
LOGCAT_TIME_RE = re.compile(r"(?P<data>\d{2}-\d{2})\s+(?P<hora>\d{2}:\d{2}:\d{2}\.\d{3})")


@dataclass
class AdbResult:
    ok: bool
    comando: list[str]
    stdout: str
    stderr: str


def normalizar_codigo(codigo: str) -> str:
    codigo = (codigo or "").strip()
    codigo = re.sub(r"^LA,", "", codigo, flags=re.IGNORECASE)
    return re.sub(r"[\s\x1d\x1e\x1f]+", "", codigo)


def _decimal_kg(value: str | None) -> Decimal | None:
    if not value:
        return None
    try:
        return Decimal(value.replace(",", ".")).quantize(Decimal("0.001"))
    except (InvalidOperation, ValueError):
        return None


def _formatar_data_gs1(value: str | None) -> str | None:
    if not value or len(value) != 6:
        return None
    try:
        data = datetime.strptime(value, "%y%m%d")
    except ValueError:
        return None
    return data.strftime("%d/%m/%Y")


def _codigos_produto_conhecidos(db: Session) -> set[str]:
    agregados = db.scalars(select(EstoqueAgregado.produto_codigo).where(EstoqueAgregado.status == "ativo")).all()
    caixas = db.scalars(select(Caixa.descricao).limit(5000)).all()
    codigos = {str(codigo).lstrip("0") for codigo in agregados if codigo}
    for descricao in caixas:
        match = re.match(r"^\s*(\d+)\s*-", descricao or "")
        if match:
            codigos.add(match.group(1).lstrip("0"))
    return {codigo for codigo in codigos if codigo}


def interpretar_etiqueta(codigo_normalizado: str, codigos_conhecidos: set[str]) -> dict:
    codigo = normalizar_codigo(codigo_normalizado)
    validade_match = list(re.finditer(r"15(?P<data>\d{6})", codigo))
    validade_pos = validade_match[-1].start() if validade_match else -1
    validade = _formatar_data_gs1(validade_match[-1].group("data")) if validade_match else None

    posicoes_37 = [match.start() for match in re.finditer("37", codigo)]
    produto_codigo = None
    peso_lido = None

    if posicoes_37:
        prefixo_produto = re.sub(r"\D", "", codigo[2:posicoes_37[0]])
        candidatos = []
        for tamanho in range(8, 2, -1):
            if len(prefixo_produto) >= tamanho:
                candidatos.append(prefixo_produto[-tamanho:].lstrip("0"))
        produto_codigo = next((c for c in candidatos if c in codigos_conhecidos), None)
        if not produto_codigo and candidatos:
            produto_codigo = candidatos[-1]

        pos_peso = max((pos for pos in posicoes_37 if validade_pos == -1 or pos < validade_pos), default=posicoes_37[-1])
        fim_peso = validade_pos if validade_pos > pos_peso else len(codigo)
        peso_texto_match = re.match(r"(?P<peso>\d+(?:\.\d+)?)", codigo[pos_peso + 2 : fim_peso])
        if peso_texto_match:
            peso_lido = _decimal_kg(peso_texto_match.group("peso"))

    return {
        "produto_codigo": produto_codigo,
        "peso_lido_kg": peso_lido,
        "validade_lida": validade,
    }


def _parse_logcat_datetime(linha: str) -> datetime:
    match = LOGCAT_TIME_RE.search(linha)
    if not match:
        return datetime.now()

    agora = datetime.now()
    texto = f"{agora.year}-{match.group('data')} {match.group('hora')}"
    try:
        data_hora = datetime.strptime(texto, "%Y-%m-%d %H:%M:%S.%f")
    except ValueError:
        return agora

    if data_hora > agora + timedelta(days=1):
        data_hora = data_hora.replace(year=agora.year - 1)
    return data_hora


def extrair_leituras_logcat(conteudo: str) -> list[dict]:
    leituras = []
    for match in WEDGE_RE.finditer(conteudo):
        inicio_linha = conteudo.rfind("\n", 0, match.start()) + 1
        fim_linha = conteudo.find("\n", match.end())
        if fim_linha == -1:
            fim_linha = len(conteudo)
        linha = conteudo[inicio_linha:fim_linha]
        codigo = match.group("codigo")
        leituras.append(
            {
                "data_hora": _parse_logcat_datetime(linha),
                "codigo_lido": codigo,
                "codigo_normalizado": normalizar_codigo(codigo),
                "raw_line": linha.strip(),
            }
        )
    return leituras


def _recalcular_resumo_sessao(db: Session, sessao: SessaoConferencia) -> None:
    rows = db.execute(
        select(LeituraColetor.status, func.count(LeituraColetor.id))
        .where(LeituraColetor.sessao_id == sessao.id)
        .group_by(LeituraColetor.status)
    ).all()
    totais = {status: int(total or 0) for status, total in rows}
    sessao.total_leituras = sum(totais.values())
    sessao.baixadas = sum(total for status, total in totais.items() if status.startswith("baixada_"))
    sessao.duplicadas = totais.get("duplicada", 0) + totais.get("ja_baixada", 0)
    sessao.nao_encontradas = totais.get("nao_encontrada", 0) + totais.get("divergencia_saldo_insuficiente", 0)
    sessao.ignoradas = totais.get("ignorada_sessao_inativa", 0)


def _encontrar_caixa(db: Session, codigo_lido: str, codigo_normalizado: str) -> Caixa | None:
    variantes = {codigo_lido, codigo_normalizado, f"LA,{codigo_normalizado}"}
    variantes = {valor for valor in variantes if valor}
    if not variantes:
        return None

    return db.scalar(
        select(Caixa)
        .where(
            or_(
                Caixa.codigo_barras_original.in_(variantes),
                Caixa.codigo_caixa.in_(variantes),
            )
        )
        .order_by(Caixa.id.asc())
    )


def _produto_descricao(db: Session, produto_codigo: str | None) -> str | None:
    if not produto_codigo:
        return None
    agregado = db.scalar(
        select(EstoqueAgregado.produto_descricao)
        .where(EstoqueAgregado.produto_codigo == produto_codigo, EstoqueAgregado.status == "ativo")
        .order_by(EstoqueAgregado.importado_em.desc())
        .limit(1)
    )
    if agregado:
        return agregado
    caixa = db.scalar(select(Caixa.descricao).where(Caixa.descricao.ilike(f"{produto_codigo} -%")).limit(1))
    return caixa


def _encontrar_caixa_por_dados(
    db: Session,
    produto_codigo: str | None,
    produto_descricao: str | None,
    peso: Decimal | None,
    validade: str | None,
) -> Caixa | None:
    if not produto_codigo or peso is None or not validade:
        return None
    condicoes_produto = [Caixa.descricao.ilike(f"{produto_codigo} -%")]
    if produto_descricao:
        condicoes_produto.append(Caixa.descricao.ilike(produto_descricao))
    query = (
        select(Caixa)
        .where(Caixa.status == "em_estoque", Caixa.peso_kg == peso, Caixa.validade == validade, or_(*condicoes_produto))
        .order_by(Caixa.id.asc())
        .limit(1)
    )
    return db.scalar(query)


def _maior_saldo_agregado(
    db: Session,
    produto_codigo: str | None,
    validade: str | None = None,
) -> EstoqueAgregado | None:
    if not produto_codigo:
        return None
    query = select(EstoqueAgregado).where(
        EstoqueAgregado.status == "ativo",
        EstoqueAgregado.produto_codigo == produto_codigo,
        EstoqueAgregado.saldo_disponivel_atual_kg > 0,
    )
    if validade:
        query = query.where(EstoqueAgregado.validade == validade)
    return db.scalar(query.order_by(EstoqueAgregado.saldo_disponivel_atual_kg.desc(), EstoqueAgregado.id.asc()).limit(1))


def _baixar_agregado(agregado: EstoqueAgregado, peso: Decimal | None) -> bool:
    if peso is None or peso <= 0:
        return False
    if agregado.saldo_disponivel_atual_kg < peso:
        return False
    agregado.saldo_disponivel_atual_kg = (agregado.saldo_disponivel_atual_kg - peso).quantize(Decimal("0.001"))
    return True


def _baixar_por_regra(db: Session, codigo_lido: str, codigo_normalizado: str) -> tuple[Caixa | None, EstoqueAgregado | None, str, str | None, dict]:
    codigos_conhecidos = _codigos_produto_conhecidos(db)
    dados = interpretar_etiqueta(codigo_normalizado, codigos_conhecidos)
    dados["produto_descricao"] = _produto_descricao(db, dados["produto_codigo"])

    caixa = _encontrar_caixa(db, codigo_lido, codigo_normalizado)
    if caixa is None:
        caixa = _encontrar_caixa_por_dados(
            db,
            dados["produto_codigo"],
            dados["produto_descricao"],
            dados["peso_lido_kg"],
            dados["validade_lida"],
        )

    if caixa is not None:
        if caixa.status == "em_estoque":
            caixa.status = "baixada_wms"
            return caixa, None, "baixada_caixa_exata", "caixa_exata", dados
        return caixa, None, "ja_baixada", "caixa_exata", dados

    agregado = _maior_saldo_agregado(db, dados["produto_codigo"], dados["validade_lida"])
    if agregado is not None:
        if _baixar_agregado(agregado, dados["peso_lido_kg"]):
            dados["produto_descricao"] = agregado.produto_descricao
            return None, agregado, "baixada_agregado_produto_validade", "produto_validade_agregado", dados
        return None, agregado, "divergencia_saldo_insuficiente", "produto_validade_agregado", dados

    agregado = _maior_saldo_agregado(db, dados["produto_codigo"])
    if agregado is not None:
        if _baixar_agregado(agregado, dados["peso_lido_kg"]):
            dados["produto_descricao"] = agregado.produto_descricao
            return None, agregado, "baixada_agregado_produto", "produto_agregado", dados
        return None, agregado, "divergencia_saldo_insuficiente", "produto_agregado", dados

    return None, None, "nao_encontrada", None, dados


def registrar_leituras(
    db: Session,
    coletor: Coletor,
    leituras: list[dict],
    sessao: SessaoConferencia | None,
) -> dict:
    resultado = {
        "total": 0,
        "baixadas": 0,
        "duplicadas": 0,
        "nao_encontradas": 0,
        "ignoradas": 0,
    }

    for dados in leituras:
        ja_importada = db.scalar(
            select(func.count())
            .select_from(LeituraColetor)
            .where(
                LeituraColetor.coletor_id == coletor.id,
                LeituraColetor.sessao_id == (sessao.id if sessao else None),
                LeituraColetor.data_hora == dados["data_hora"],
                LeituraColetor.codigo_normalizado == dados["codigo_normalizado"],
            )
        )
        if ja_importada:
            continue

        resultado["total"] += 1
        caixa = None
        agregado = None
        regra_baixa = None
        dados_etiqueta = {}
        status = "ignorada_sessao_inativa"

        if sessao and sessao.status == "ativa":
            caixa, agregado, status, regra_baixa, dados_etiqueta = _baixar_por_regra(
                db,
                dados["codigo_lido"],
                dados["codigo_normalizado"],
            )
            if status == "nao_encontrada":
                resultado["nao_encontradas"] += 1
            elif status == "divergencia_saldo_insuficiente":
                resultado["nao_encontradas"] += 1
            elif status.startswith("baixada_"):
                resultado["baixadas"] += 1
            else:
                resultado["duplicadas"] += 1
        else:
            resultado["ignoradas"] += 1

        leitura = LeituraColetor(
            coletor_id=coletor.id,
            sessao_id=sessao.id if sessao else None,
            caixa_id=caixa.id if caixa else None,
            estoque_agregado_id=agregado.id if agregado else None,
            data_hora=dados["data_hora"],
            codigo_lido=dados["codigo_lido"],
            codigo_normalizado=dados["codigo_normalizado"],
            status=status,
            regra_baixa=regra_baixa,
            produto_codigo=dados_etiqueta.get("produto_codigo"),
            produto_descricao=dados_etiqueta.get("produto_descricao"),
            peso_lido_kg=dados_etiqueta.get("peso_lido_kg"),
            validade_lida=dados_etiqueta.get("validade_lida"),
            peso_baixado_kg=dados_etiqueta.get("peso_lido_kg") if status.startswith("baixada_agregado") else None,
            raw_line=dados.get("raw_line"),
        )
        db.add(leitura)
        coletor.ultimo_bipe = dados["data_hora"]
        coletor.ultimo_codigo_lido = dados["codigo_lido"]

    coletor.ultimo_contato = datetime.now()
    coletor.status = "online"
    if sessao:
        db.flush()
        _recalcular_resumo_sessao(db, sessao)
    return resultado


def importar_logcat(db: Session, coletor: Coletor, conteudo: str, sessao: SessaoConferencia | None) -> dict:
    leituras = extrair_leituras_logcat(conteudo)
    resultado = registrar_leituras(db, coletor, leituras, sessao)
    resultado["linhas_extraidas"] = len(leituras)
    return resultado


def reprocessar_leituras_sessao(db: Session, sessao: SessaoConferencia) -> dict:
    resultado = {"reprocessadas": 0, "baixadas": 0, "continuam_nao_encontradas": 0, "divergencias": 0}
    rows = db.scalars(
        select(LeituraColetor).where(
            LeituraColetor.sessao_id == sessao.id,
            LeituraColetor.status.in_(["nao_encontrada", "divergencia_saldo_insuficiente"]),
        )
    ).all()

    for leitura in rows:
        resultado["reprocessadas"] += 1
        caixa, agregado, status, regra_baixa, dados_etiqueta = _baixar_por_regra(
            db,
            leitura.codigo_lido,
            leitura.codigo_normalizado,
        )
        if status == "nao_encontrada":
            resultado["continuam_nao_encontradas"] += 1
            continue
        if status == "divergencia_saldo_insuficiente":
            resultado["divergencias"] += 1
        if status.startswith("baixada_"):
            resultado["baixadas"] += 1
        leitura.caixa_id = caixa.id if caixa else None
        leitura.estoque_agregado_id = agregado.id if agregado else None
        leitura.status = status
        leitura.regra_baixa = regra_baixa
        leitura.produto_codigo = dados_etiqueta.get("produto_codigo")
        leitura.produto_descricao = dados_etiqueta.get("produto_descricao")
        leitura.peso_lido_kg = dados_etiqueta.get("peso_lido_kg")
        leitura.validade_lida = dados_etiqueta.get("validade_lida")
        leitura.peso_baixado_kg = dados_etiqueta.get("peso_lido_kg") if status.startswith("baixada_agregado") else None

    db.flush()
    _recalcular_resumo_sessao(db, sessao)
    return resultado


def alvo_adb(coletor: Coletor) -> str:
    if coletor.ip_atual:
        return f"{coletor.ip_atual}:5555"
    return coletor.serial


def resolver_adb() -> str:
    settings = get_settings()
    candidatos = [
        settings.adb_path,
        shutil.which("adb"),
        r"C:\platform-tools\adb.exe",
        str(Path.home() / "AppData" / "Local" / "Android" / "Sdk" / "platform-tools" / "adb.exe"),
        str(Path.home() / "Downloads" / "platform-tools" / "adb.exe"),
    ]
    for candidato in candidatos:
        if candidato and Path(candidato).exists():
            return candidato
    return "adb"


def executar_adb(args: list[str], timeout: int = 20) -> AdbResult:
    comando = [resolver_adb(), *args]
    try:
        proc = subprocess.run(comando, capture_output=True, text=True, timeout=timeout, check=False)
    except FileNotFoundError:
        return AdbResult(False, comando, "", "adb.exe nao encontrado. Configure ADB_PATH no .env ou instale em C:\\platform-tools\\adb.exe")
    except subprocess.TimeoutExpired as exc:
        return AdbResult(False, comando, exc.stdout or "", exc.stderr or "timeout executando adb")
    return AdbResult(proc.returncode == 0, comando, proc.stdout, proc.stderr)


def listar_status_adb() -> dict[str, str]:
    resultado = executar_adb(["devices"], timeout=8)
    if not resultado.ok:
        return {}

    dispositivos: dict[str, str] = {}
    for linha in resultado.stdout.splitlines():
        partes = linha.split()
        if len(partes) < 2 or partes[0].lower() == "list":
            continue
        dispositivos[partes[0]] = partes[1].lower()
    return dispositivos


def limpar_logcat_coletor(coletor: Coletor) -> AdbResult:
    alvo = alvo_adb(coletor)
    if coletor.ip_atual:
        executar_adb(["connect", alvo], timeout=12)
    return executar_adb(["-s", alvo, "logcat", "-c"], timeout=20)


def reconectar_coletor_ip(coletor: Coletor) -> AdbResult:
    if not coletor.ip_atual:
        return AdbResult(False, ["adb", "connect"], "", "coletor sem IP cadastrado")
    resultado = conectar_validar_rede(coletor.ip_atual, coletor.serial)
    return AdbResult(
        resultado["ok"],
        [resolver_adb(), "connect", f"{coletor.ip_atual}:5555"],
        resultado.get("stdout", ""),
        resultado.get("stderr", "") or resultado.get("erro", ""),
    )


def conectar_validar_rede(ip_atual: str, serial_esperado: str, tentativas: int = 5) -> dict:
    alvo_rede = f"{ip_atual}:5555"
    saidas = []
    erros = []

    for tentativa in range(1, tentativas + 1):
        connect = executar_adb(["connect", alvo_rede], timeout=15)
        saidas.append(connect.stdout)
        erros.append(connect.stderr)

        if "10013" in (connect.stderr or "") or "10013" in (connect.stdout or ""):
            executar_adb(["kill-server"], timeout=10)
            executar_adb(["start-server"], timeout=10)
            time.sleep(1)
            connect = executar_adb(["connect", alvo_rede], timeout=15)
            saidas.append(connect.stdout)
            erros.append(connect.stderr)

        serial_cmd = executar_adb(["-s", alvo_rede, "shell", "getprop", "ro.serialno"], timeout=12)
        serial_rede = (serial_cmd.stdout or "").strip()
        saidas.append(serial_cmd.stdout)
        erros.append(serial_cmd.stderr)

        if serial_cmd.ok and serial_rede == serial_esperado:
            return {
                "ok": True,
                "alvo_rede": alvo_rede,
                "serial_rede": serial_rede,
                "stdout": "".join(saidas),
                "stderr": "".join(erros),
            }

        if tentativa < tentativas:
            time.sleep(2)

    return {
        "ok": False,
        "alvo_rede": alvo_rede,
        "erro": f"nao foi possivel validar o coletor {serial_esperado} em {alvo_rede}",
        "stdout": "".join(saidas),
        "stderr": "".join(erros),
    }


def _extrair_ip_dispositivo(alvo: str) -> str | None:
    for args in (
        ["-s", alvo, "shell", "ip", "route", "get", "8.8.8.8"],
        ["-s", alvo, "shell", "ip", "-f", "inet", "addr", "show", "wlan0"],
        ["-s", alvo, "shell", "ip", "-f", "inet", "addr"],
    ):
        ip_cmd = executar_adb(args, timeout=20)
        ip_match = re.search(r"\bsrc\s+(\d+\.\d+\.\d+\.\d+)\b", ip_cmd.stdout)
        if not ip_match:
            ip_match = re.search(r"inet\s+(\d+\.\d+\.\d+\.\d+)/", ip_cmd.stdout)
        if ip_match and not ip_match.group(1).startswith("127."):
            return ip_match.group(1)
    return None


def restaurar_adb_usb(serial: str | None = None, validar_rede: bool = True) -> dict:
    devices = executar_adb(["devices", "-l"], timeout=12)
    if not devices.ok:
        return {"ok": False, "erro": devices.stderr or devices.stdout, "devices": devices.stdout}

    dispositivos_usb = []
    for linha in devices.stdout.splitlines():
        partes = linha.split()
        if len(partes) < 2 or partes[1] != "device":
            continue
        alvo = partes[0]
        if ":" in alvo:
            continue
        dispositivos_usb.append(alvo)

    if not dispositivos_usb:
        return {"ok": False, "erro": "nenhum coletor USB encontrado", "devices": devices.stdout}
    if len(dispositivos_usb) > 1 and not serial:
        return {"ok": False, "erro": "mais de um coletor USB conectado; informe o serial", "devices": devices.stdout}

    alvo = serial if serial and serial in dispositivos_usb else dispositivos_usb[0]
    serial_cmd = executar_adb(["-s", alvo, "shell", "getprop", "ro.serialno"], timeout=12)
    serial_descoberto = (serial_cmd.stdout or "").strip() or alvo
    if serial and serial_descoberto != serial:
        return {
            "ok": False,
            "erro": f"serial USB encontrado ({serial_descoberto}) diferente do informado ({serial})",
            "serial": serial_descoberto,
            "devices": devices.stdout,
        }

    ip_atual = _extrair_ip_dispositivo(alvo)
    if not ip_atual:
        return {
            "ok": False,
            "erro": "nao foi possivel descobrir o IP do coletor USB",
            "serial": serial_descoberto,
            "devices": devices.stdout,
        }

    tcpip = executar_adb(["-s", alvo, "tcpip", "5555"], timeout=8)
    if not validar_rede:
        return {
            "ok": tcpip.ok,
            "serial": serial_descoberto,
            "alvo_usb": alvo,
            "alvo_rede": f"{ip_atual}:5555",
            "ip_atual": ip_atual,
            "stdout": tcpip.stdout,
            "stderr": tcpip.stderr,
            "erro": None if tcpip.ok else (tcpip.stderr or tcpip.stdout or "falha ao ativar adb tcpip"),
        }

    time.sleep(2)
    rede = conectar_validar_rede(ip_atual, serial_descoberto)
    return {
        "ok": tcpip.ok and rede["ok"],
        "serial": serial_descoberto,
        "alvo_usb": alvo,
        "alvo_rede": rede.get("alvo_rede"),
        "ip_atual": ip_atual,
        "stdout": tcpip.stdout + (rede.get("stdout") or ""),
        "stderr": tcpip.stderr + (rede.get("stderr") or ""),
        "erro": None if rede["ok"] else rede.get("erro"),
    }
