from datetime import datetime
import re

from fastapi import Depends, FastAPI, Query, Request
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from sqlalchemy import func, or_, select
from sqlalchemy.orm import Session

from backend.coletor_service import importar_logcat, limpar_logcat_coletor, reprocessar_leituras_sessao, restaurar_adb_usb
from backend.coletor_service import executar_adb, alvo_adb, listar_status_adb
from backend.coletor_service import reconectar_coletor_ip
from backend.config import get_settings
from backend.database import get_db, init_db
from backend.estoque_agregado_service import importar_estoque_xlsx
from backend.models import Caixa, Coletor, EstoqueAgregado, LeituraColetor, Pallet, Pesagem, Recebimento, SessaoConferencia
from backend.repository import backfill_caixas_pesagens_existentes
from monitor import carregar_recebimento, encerrar_recebimento, iniciar_recebimento

settings = get_settings()
app = FastAPI(title="PesoCtrl API", version="0.1.0")
from backend.inventario import router as inventario_router

app.include_router(inventario_router)


class ColetorPayload(BaseModel):
    nome: str
    serial: str
    ip_atual: str | None = None


class SessaoPayload(BaseModel):
    operador: str | None = None
    observacao: str | None = None


class LogcatPayload(BaseModel):
    conteudo: str


class RestaurarUsbPayload(BaseModel):
    serial: str | None = None


class CodigoBarrasPayload(BaseModel):
    codigo_barras_original: str


class ImportarEstoquePayload(BaseModel):
    nome_arquivo: str | None = None
    arquivo_base64: str

app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origin_list,
    allow_origin_regex=r"https?://(localhost|127\.0\.0\.1|192\.168\.\d+\.\d+|10\.\d+\.\d+\.\d+|172\.(1[6-9]|2\d|3[0-1])\.\d+\.\d+)(:\d+)?",
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.on_event("startup")
def startup() -> None:
    init_db()
    backfill_caixas_pesagens_existentes()


@app.get("/api/health")
def health(db: Session = Depends(get_db)) -> dict:
    db.execute(select(1))
    return {"status": "ok", "time": datetime.now().isoformat(timespec="seconds")}


@app.get("/api/recebimento-global/status")
def status_recebimento_global() -> dict:
    return carregar_recebimento()


@app.post("/api/recebimento-global/iniciar")
def iniciar_recebimento_global_api() -> dict:
    ok, msg = iniciar_recebimento()
    return {"ok": ok, "mensagem": msg, "status": carregar_recebimento()}


@app.post("/api/recebimento-global/finalizar")
def finalizar_recebimento_global_api() -> dict:
    ok, msg = encerrar_recebimento()
    return {"ok": ok, "mensagem": msg, "status": carregar_recebimento()}


def _serializar_coletor(row: Coletor) -> dict:
    ativa = next((sessao for sessao in row.sessoes if sessao.status == "ativa"), None)
    ultima = max(row.sessoes, key=lambda sessao: sessao.inicio, default=None)
    return {
        "id": row.id,
        "nome": row.nome,
        "serial": row.serial,
        "ip_atual": row.ip_atual,
        "status": row.status,
        "ultimo_contato": row.ultimo_contato,
        "ultimo_bipe": row.ultimo_bipe,
        "ultimo_codigo_lido": row.ultimo_codigo_lido,
        "sessao_ativa_id": ativa.id if ativa else None,
        "ultima_sessao_id": ultima.id if ultima else None,
    }


def _serializar_sessao(row: SessaoConferencia) -> dict:
    return {
        "id": row.id,
        "coletor_id": row.coletor_id,
        "inicio": row.inicio,
        "fim": row.fim,
        "status": row.status,
        "operador": row.operador,
        "observacao": row.observacao,
        "log_limpo": row.log_limpo,
        "total_leituras": row.total_leituras,
        "baixadas": row.baixadas,
        "duplicadas": row.duplicadas,
        "nao_encontradas": row.nao_encontradas,
        "ignoradas": row.ignoradas,
    }


def _coletor_por_ip_requisicao(request: Request, db: Session) -> Coletor | None:
    ip_cliente = request.client.host if request.client else None
    if not ip_cliente:
        return None
    return db.scalar(select(Coletor).where(Coletor.ip_atual == ip_cliente))


def _sessao_ativa_coletor(db: Session, coletor_id: int) -> SessaoConferencia | None:
    return db.scalar(
        select(SessaoConferencia).where(
            SessaoConferencia.coletor_id == coletor_id,
            SessaoConferencia.status == "ativa",
        )
    )


def _proximo_nome_coletor(db: Session) -> str:
    nomes = db.scalars(select(Coletor.nome)).all()
    maior_numero = 0
    for nome in nomes:
        match = re.search(r"coletor\D*(\d+)\s*$", nome or "", re.IGNORECASE)
        if match:
            maior_numero = max(maior_numero, int(match.group(1)))
    if maior_numero == 0:
        maior_numero = len(nomes)
    return f"Coletor {maior_numero + 1:02d}"


def _status_coletor_por_adb(coletor: Coletor, dispositivos: dict[str, str]) -> str:
    alvos = []
    if coletor.ip_atual:
        alvos.append(f"{coletor.ip_atual}:5555")
    if coletor.serial:
        alvos.append(coletor.serial)

    for alvo in alvos:
        if dispositivos.get(alvo) == "device":
            return "online"
    for alvo in alvos:
        if alvo in dispositivos:
            return "offline"
    return "offline"


def _sincronizar_status_adb(db: Session, rows: list[Coletor]) -> None:
    dispositivos = listar_status_adb()
    mudou = False
    agora = datetime.now()
    for row in rows:
        status = _status_coletor_por_adb(row, dispositivos)
        if row.status != status:
            row.status = status
            row.ultimo_contato = agora
            mudou = True
    if mudou:
        db.commit()


def _capturar_logcat_para_sessao(db: Session, sessao: SessaoConferencia) -> dict:
    coletor = db.get(Coletor, sessao.coletor_id)
    if coletor is None:
        return {"ok": False, "erro": "coletor nao encontrado"}

    alvo = alvo_adb(coletor)
    resultado_adb = executar_adb(["-s", alvo, "logcat", "-d", "-v", "time", "-s", "DCS-Wedge"], timeout=20)
    if not resultado_adb.ok:
        return {"ok": False, "erro": resultado_adb.stderr or resultado_adb.stdout or "falha ao capturar logcat"}

    resultado = importar_logcat(db, coletor, resultado_adb.stdout, sessao)
    return {"ok": True, **resultado, "sessao": _serializar_sessao(sessao)}


@app.get("/api/coletor-operacional/status")
def status_coletor_operacional(request: Request, db: Session = Depends(get_db)) -> dict:
    coletor = _coletor_por_ip_requisicao(request, db)
    if coletor is None:
        ip_cliente = request.client.host if request.client else "-"
        return {"ok": False, "erro": f"nenhum coletor cadastrado com o IP {ip_cliente}"}
    sessao = _sessao_ativa_coletor(db, coletor.id)
    return {
        "ok": True,
        "coletor": _serializar_coletor(coletor),
        "sessao": _serializar_sessao(sessao) if sessao else None,
    }


@app.post("/api/coletor-operacional/iniciar")
def iniciar_coletor_operacional(request: Request, db: Session = Depends(get_db)) -> dict:
    coletor = _coletor_por_ip_requisicao(request, db)
    if coletor is None:
        ip_cliente = request.client.host if request.client else "-"
        return {"ok": False, "erro": f"nenhum coletor cadastrado com o IP {ip_cliente}"}

    ativa = _sessao_ativa_coletor(db, coletor.id)
    if ativa:
        return {"ok": False, "erro": "ja existe conferencia ativa para este coletor", "sessao": _serializar_sessao(ativa)}

    limpeza = limpar_logcat_coletor(coletor)
    aviso = None
    if not limpeza.ok:
        aviso = limpeza.stderr or limpeza.stdout or "falha ao limpar logcat"

    sessao = SessaoConferencia(
        coletor_id=coletor.id,
        operador=coletor.nome,
        observacao="Tela operacional do coletor",
        log_limpo=limpeza.ok,
    )
    coletor.status = "online"
    coletor.ultimo_contato = datetime.now()
    db.add(sessao)
    db.commit()
    db.refresh(sessao)
    resposta = {"ok": True, "coletor": _serializar_coletor(coletor), "sessao": _serializar_sessao(sessao)}
    if aviso:
        resposta["aviso"] = f"Conferencia iniciada, mas nao foi possivel limpar o logcat via ADB: {aviso}"
    return resposta


@app.post("/api/coletor-operacional/finalizar")
def finalizar_coletor_operacional(request: Request, db: Session = Depends(get_db)) -> dict:
    coletor = _coletor_por_ip_requisicao(request, db)
    if coletor is None:
        ip_cliente = request.client.host if request.client else "-"
        return {"ok": False, "erro": f"nenhum coletor cadastrado com o IP {ip_cliente}"}

    sessao = _sessao_ativa_coletor(db, coletor.id)
    if sessao is None:
        return {"ok": False, "erro": "nao existe conferencia ativa para este coletor"}

    captura = _capturar_logcat_para_sessao(db, sessao)
    if not captura.get("ok"):
        db.commit()
        return captura

    sessao.status = "encerrada"
    sessao.fim = datetime.now()
    db.commit()
    return {
        "ok": True,
        "coletor": _serializar_coletor(coletor),
        "sessao": _serializar_sessao(sessao),
        "captura": {key: value for key, value in captura.items() if key not in {"ok", "sessao"}},
    }


@app.get("/api/coletores")
def listar_coletores(db: Session = Depends(get_db)) -> dict:
    rows = db.scalars(select(Coletor).order_by(Coletor.nome.asc())).all()
    _sincronizar_status_adb(db, rows)
    return {"total": len(rows), "items": [_serializar_coletor(row) for row in rows]}


@app.post("/api/coletores")
def salvar_coletor(payload: ColetorPayload, db: Session = Depends(get_db)) -> dict:
    row = db.scalar(select(Coletor).where(Coletor.serial == payload.serial))
    if row is None:
        row = Coletor(nome=payload.nome, serial=payload.serial)
        db.add(row)
    row.nome = payload.nome
    row.ip_atual = payload.ip_atual
    db.commit()
    db.refresh(row)
    return _serializar_coletor(row)


@app.delete("/api/coletores/{coletor_id}")
def excluir_coletor(coletor_id: int, db: Session = Depends(get_db)) -> dict:
    row = db.get(Coletor, coletor_id)
    if row is None:
        return {"ok": False, "erro": "coletor nao encontrado"}
    ativa = db.scalar(
        select(SessaoConferencia).where(
            SessaoConferencia.coletor_id == coletor_id,
            SessaoConferencia.status == "ativa",
        )
    )
    if ativa:
        return {"ok": False, "erro": "finalize a conferencia ativa antes de excluir"}

    sessoes = db.scalar(
        select(func.count())
        .select_from(SessaoConferencia)
        .where(SessaoConferencia.coletor_id == coletor_id)
    ) or 0
    leituras = db.scalar(
        select(func.count())
        .select_from(LeituraColetor)
        .where(LeituraColetor.coletor_id == coletor_id)
    ) or 0
    if sessoes or leituras:
        return {
            "ok": False,
            "erro": "nao e possivel excluir coletor com historico de conferencias ou leituras; use Criar auto para atualizar serial/IP",
        }

    db.delete(row)
    db.commit()
    return {"ok": True}


@app.post("/api/coletores/{coletor_id}/limpar-log")
def limpar_log_coletor(coletor_id: int, db: Session = Depends(get_db)) -> dict:
    coletor = db.get(Coletor, coletor_id)
    if coletor is None:
        return {"ok": False, "erro": "coletor nao encontrado"}
    resultado = limpar_logcat_coletor(coletor)
    coletor.status = "online" if resultado.ok else "offline"
    coletor.ultimo_contato = datetime.now()
    db.commit()
    return {
        "ok": resultado.ok,
        "comando": " ".join(resultado.comando),
        "stdout": resultado.stdout,
        "stderr": resultado.stderr,
    }


@app.post("/api/coletores/{coletor_id}/reconectar-ip")
def reconectar_ip_coletor(coletor_id: int, db: Session = Depends(get_db)) -> dict:
    coletor = db.get(Coletor, coletor_id)
    if coletor is None:
        return {"ok": False, "erro": "coletor nao encontrado"}
    resultado = reconectar_coletor_ip(coletor)
    coletor.status = "online" if resultado.ok else "offline"
    coletor.ultimo_contato = datetime.now()
    db.commit()
    return {
        "ok": resultado.ok,
        "comando": " ".join(resultado.comando),
        "stdout": resultado.stdout,
        "stderr": resultado.stderr,
        "erro": None if resultado.ok else (resultado.stderr or resultado.stdout),
    }


@app.post("/api/coletores/{coletor_id}/iniciar-conferencia")
def iniciar_conferencia(coletor_id: int, payload: SessaoPayload, db: Session = Depends(get_db)) -> dict:
    coletor = db.get(Coletor, coletor_id)
    if coletor is None:
        return {"ok": False, "erro": "coletor nao encontrado"}

    ativa = db.scalar(
        select(SessaoConferencia).where(
            SessaoConferencia.coletor_id == coletor_id,
            SessaoConferencia.status == "ativa",
        )
    )
    if ativa:
        return {"ok": False, "erro": "ja existe conferencia ativa para este coletor", "sessao": _serializar_sessao(ativa)}

    limpeza = limpar_logcat_coletor(coletor)
    aviso = None
    if not limpeza.ok:
        aviso = limpeza.stderr or limpeza.stdout or "falha ao limpar logcat"

    sessao = SessaoConferencia(
        coletor_id=coletor.id,
        operador=payload.operador,
        observacao=payload.observacao,
        log_limpo=limpeza.ok,
    )
    coletor.status = "online"
    coletor.ultimo_contato = datetime.now()
    db.add(sessao)
    db.commit()
    db.refresh(sessao)
    resposta = {"ok": True, "sessao": _serializar_sessao(sessao)}
    if aviso:
        resposta["aviso"] = f"Conferencia iniciada, mas nao foi possivel limpar o logcat via ADB: {aviso}"
    return resposta


@app.post("/api/sessoes-conferencia/{sessao_id}/importar-logcat")
def importar_logcat_sessao(sessao_id: int, payload: LogcatPayload, db: Session = Depends(get_db)) -> dict:
    sessao = db.get(SessaoConferencia, sessao_id)
    if sessao is None:
        return {"ok": False, "erro": "sessao nao encontrada"}
    coletor = db.get(Coletor, sessao.coletor_id)
    if coletor is None:
        return {"ok": False, "erro": "coletor nao encontrado"}
    resultado = importar_logcat(db, coletor, payload.conteudo, sessao)
    db.commit()
    return {"ok": True, **resultado, "sessao": _serializar_sessao(sessao)}


@app.post("/api/sessoes-conferencia/{sessao_id}/capturar-logcat")
def capturar_logcat_sessao(sessao_id: int, db: Session = Depends(get_db)) -> dict:
    sessao = db.get(SessaoConferencia, sessao_id)
    if sessao is None:
        return {"ok": False, "erro": "sessao nao encontrada"}
    coletor = db.get(Coletor, sessao.coletor_id)
    if coletor is None:
        return {"ok": False, "erro": "coletor nao encontrado"}

    alvo = alvo_adb(coletor)
    resultado_adb = executar_adb(["-s", alvo, "logcat", "-d", "-v", "time", "-s", "DCS-Wedge"], timeout=20)
    if not resultado_adb.ok:
        return {"ok": False, "erro": resultado_adb.stderr or resultado_adb.stdout or "falha ao capturar logcat"}

    resultado = importar_logcat(db, coletor, resultado_adb.stdout, sessao)
    db.commit()
    return {"ok": True, **resultado, "sessao": _serializar_sessao(sessao)}


@app.post("/api/sessoes-conferencia/{sessao_id}/finalizar")
def finalizar_conferencia(sessao_id: int, db: Session = Depends(get_db)) -> dict:
    sessao = db.get(SessaoConferencia, sessao_id)
    if sessao is None:
        return {"ok": False, "erro": "sessao nao encontrada"}
    sessao.status = "encerrada"
    sessao.fim = datetime.now()
    db.commit()
    return {"ok": True, "sessao": _serializar_sessao(sessao)}


@app.post("/api/sessoes-conferencia/{sessao_id}/reprocessar")
def reprocessar_conferencia(sessao_id: int, db: Session = Depends(get_db)) -> dict:
    sessao = db.get(SessaoConferencia, sessao_id)
    if sessao is None:
        return {"ok": False, "erro": "sessao nao encontrada"}
    resultado = reprocessar_leituras_sessao(db, sessao)
    db.commit()
    return {"ok": True, **resultado, "sessao": _serializar_sessao(sessao)}


@app.get("/api/sessoes-conferencia/{sessao_id}/leituras")
def listar_leituras_sessao(
    sessao_id: int,
    limit: int = Query(100, ge=1, le=500),
    db: Session = Depends(get_db),
) -> dict:
    rows = db.scalars(
        select(LeituraColetor)
        .where(LeituraColetor.sessao_id == sessao_id)
        .order_by(LeituraColetor.data_hora.desc(), LeituraColetor.id.desc())
        .limit(limit)
    ).all()
    total = db.scalar(select(func.count()).select_from(LeituraColetor).where(LeituraColetor.sessao_id == sessao_id)) or 0
    return {
        "total": total,
        "items": [
            {
                "id": row.id,
                "data_hora": row.data_hora,
                "codigo_lido": row.codigo_lido,
                "status": row.status,
                "caixa_id": row.caixa_id,
                "estoque_agregado_id": row.estoque_agregado_id,
                "regra_baixa": row.regra_baixa,
                "produto_codigo": row.produto_codigo,
                "produto_descricao": row.produto_descricao,
                "peso_lido_kg": float(row.peso_lido_kg) if row.peso_lido_kg is not None else None,
                "validade_lida": row.validade_lida,
                "peso_baixado_kg": float(row.peso_baixado_kg) if row.peso_baixado_kg is not None else None,
            }
            for row in rows
        ],
    }


@app.get("/api/saidas")
def listar_saidas(
    busca: str | None = None,
    nome: str | None = None,
    validade: str | None = None,
    peso: float | None = None,
    limit: int = Query(100, ge=1, le=500),
    offset: int = Query(0, ge=0),
    db: Session = Depends(get_db),
) -> dict:
    query = select(LeituraColetor).where(LeituraColetor.status != "ignorada_sessao_inativa")
    count_query = select(func.count()).select_from(LeituraColetor).where(LeituraColetor.status != "ignorada_sessao_inativa")

    filters = []
    if busca:
        pattern = f"%{busca}%"
        filters.append(
            or_(
                LeituraColetor.codigo_lido.ilike(pattern),
                LeituraColetor.produto_codigo.ilike(pattern),
                LeituraColetor.produto_descricao.ilike(pattern),
                LeituraColetor.status.ilike(pattern),
                LeituraColetor.regra_baixa.ilike(pattern),
            )
        )
    if nome:
        filters.append(LeituraColetor.produto_descricao.ilike(f"%{nome}%"))
    if validade:
        filters.append(LeituraColetor.validade_lida.ilike(f"%{validade}%"))
    if peso is not None:
        filters.append(LeituraColetor.peso_lido_kg == peso)

    for condition in filters:
        query = query.where(condition)
        count_query = count_query.where(condition)

    total = db.scalar(count_query) or 0
    rows = db.scalars(query.order_by(LeituraColetor.data_hora.desc(), LeituraColetor.id.desc()).limit(limit).offset(offset)).all()
    return {
        "total": total,
        "items": [
            {
                "id": row.id,
                "data_hora": row.data_hora,
                "codigo_lido": row.codigo_lido,
                "status": row.status,
                "caixa_id": row.caixa_id,
                "estoque_agregado_id": row.estoque_agregado_id,
                "regra_baixa": row.regra_baixa,
                "produto_codigo": row.produto_codigo,
                "produto_descricao": row.produto_descricao,
                "peso_lido_kg": float(row.peso_lido_kg) if row.peso_lido_kg is not None else None,
                "validade_lida": row.validade_lida,
                "peso_baixado_kg": float(row.peso_baixado_kg) if row.peso_baixado_kg is not None else None,
                "sessao_id": row.sessao_id,
                "coletor": row.coletor.nome if row.coletor else None,
            }
            for row in rows
        ],
    }


@app.post("/api/caixas/{caixa_id}/codigo-barras")
def atualizar_codigo_barras_caixa(caixa_id: int, payload: CodigoBarrasPayload, db: Session = Depends(get_db)) -> dict:
    caixa = db.get(Caixa, caixa_id)
    if caixa is None:
        return {"ok": False, "erro": "caixa nao encontrada"}
    codigo = "".join((payload.codigo_barras_original or "").split())
    codigo = codigo.replace("\x1d", "")
    caixa.codigo_barras_original = codigo
    db.commit()
    return {"ok": True, "id": caixa.id, "codigo_barras_original": caixa.codigo_barras_original}


@app.post("/api/coletores/restaurar-usb")
def restaurar_usb(payload: RestaurarUsbPayload, db: Session = Depends(get_db)) -> dict:
    resultado = restaurar_adb_usb(payload.serial)
    if resultado.get("ok") and resultado.get("serial"):
        coletor = db.scalar(select(Coletor).where(Coletor.serial == resultado["serial"]))
        if coletor is None:
            resultado["ok"] = False
            resultado["erro"] = f"serial {resultado['serial']} nao esta cadastrado em coletores"
            return resultado
        coletor.ip_atual = resultado.get("ip_atual") or coletor.ip_atual
        coletor.status = "online"
        coletor.ultimo_contato = datetime.now()
        db.commit()
        resultado["coletor_id"] = coletor.id
        resultado["nome"] = coletor.nome
    return resultado


@app.post("/api/coletores/criar-automatico-usb")
def criar_coletor_automatico_usb(payload: RestaurarUsbPayload, db: Session = Depends(get_db)) -> dict:
    resultado = restaurar_adb_usb(payload.serial, validar_rede=False)
    if not resultado.get("ok") or not resultado.get("serial"):
        return resultado

    coletor = db.scalar(select(Coletor).where(Coletor.serial == resultado["serial"]))
    criado = False
    if coletor is None:
        coletor = Coletor(nome=_proximo_nome_coletor(db), serial=resultado["serial"])
        db.add(coletor)
        criado = True

    coletor.ip_atual = resultado.get("ip_atual") or coletor.ip_atual
    coletor.status = "online"
    coletor.ultimo_contato = datetime.now()
    db.commit()
    db.refresh(coletor)

    return {
        **resultado,
        "ok": True,
        "criado": criado,
        "coletor": _serializar_coletor(coletor),
        "coletor_id": coletor.id,
        "nome": coletor.nome,
    }


@app.get("/api/recebimentos")
def listar_recebimentos(
    limit: int = Query(50, ge=1, le=500),
    offset: int = Query(0, ge=0),
    db: Session = Depends(get_db),
) -> dict:
    total = db.scalar(select(func.count()).select_from(Recebimento)) or 0
    rows = db.scalars(
        select(Recebimento).order_by(Recebimento.inicio.desc()).limit(limit).offset(offset)
    ).all()
    return {
        "total": total,
        "items": [
            {
                "id": row.id,
                "recebimento_id": row.recebimento_id,
                "inicio": row.inicio,
                "fim": row.fim,
                "status": row.status,
                "total_pallets": row.total_pallets,
            }
            for row in rows
        ],
    }


@app.get("/api/pallets")
def listar_pallets(
    recebimento_id: str | None = None,
    status: str | None = None,
    limit: int = Query(100, ge=1, le=500),
    offset: int = Query(0, ge=0),
    db: Session = Depends(get_db),
) -> dict:
    query = select(Pallet)
    count_query = select(func.count()).select_from(Pallet)
    if recebimento_id:
        query = query.where(Pallet.recebimento_id_texto == recebimento_id)
        count_query = count_query.where(Pallet.recebimento_id_texto == recebimento_id)
    if status:
        query = query.where(Pallet.status == status)
        count_query = count_query.where(Pallet.status == status)

    total = db.scalar(count_query) or 0
    rows = db.scalars(query.order_by(Pallet.inicio.desc()).limit(limit).offset(offset)).all()
    return {
        "total": total,
        "items": [
            {
                "id": row.id,
                "recebimento_id": row.recebimento_id_texto,
                "sessao_id": row.sessao_id,
                "numero_pallet": row.numero_pallet,
                "impressora": row.impressora_nome,
                "descricao_produto": row.descricao_produto,
                "inicio": row.inicio,
                "fim": row.fim,
                "status": row.status,
                "arquivo_excel": row.arquivo_excel,
            }
            for row in rows
        ],
    }


@app.get("/api/pesagens")
def listar_pesagens(
    busca: str | None = None,
    nome: str | None = None,
    recebimento_id: str | None = None,
    numero_pallet: int | None = None,
    validade: str | None = None,
    peso: float | None = None,
    limit: int = Query(100, ge=1, le=500),
    offset: int = Query(0, ge=0),
    db: Session = Depends(get_db),
) -> dict:
    query = select(Pesagem)
    count_query = select(func.count()).select_from(Pesagem)

    filters = []
    if busca:
        pattern = f"%{busca}%"
        filters.append(or_(Pesagem.descricao.ilike(pattern), Pesagem.impressora_nome.ilike(pattern)))
    if nome:
        filters.append(Pesagem.descricao.ilike(f"%{nome}%"))
    if recebimento_id:
        filters.append(Pesagem.recebimento_id == recebimento_id)
    if numero_pallet:
        filters.append(Pesagem.numero_pallet == numero_pallet)
    if validade:
        filters.append(Pesagem.validade.ilike(f"%{validade}%"))
    if peso is not None:
        filters.append(Pesagem.peso_kg == peso)

    for condition in filters:
        query = query.where(condition)
        count_query = count_query.where(condition)

    total = db.scalar(count_query) or 0
    rows = db.scalars(query.order_by(Pesagem.data_hora.desc()).limit(limit).offset(offset)).all()
    return {
        "total": total,
        "items": [
            {
                "id": row.id,
                "data_hora": row.data_hora,
                "recebimento_id": row.recebimento_id,
                "numero_pallet": row.numero_pallet,
                "impressora": row.impressora_nome,
                "descricao": row.descricao,
                "peso_kg": float(row.peso_kg) if row.peso_kg is not None else None,
                "qtd_etiquetas": row.qtd_etiquetas,
                "peso_total_kg": float(row.peso_total_kg) if row.peso_total_kg is not None else None,
                "validade": row.validade,
            }
            for row in rows
        ],
    }


@app.get("/api/estoque/resumo")
def resumo_estoque(
    busca: str | None = None,
    nome: str | None = None,
    validade: str | None = None,
    peso: float | None = None,
    limit: int = Query(100, ge=1, le=500),
    offset: int = Query(0, ge=0),
    db: Session = Depends(get_db),
) -> dict:
    acumulado: dict[tuple[str, str | None], dict] = {}

    caixas_rows = db.execute(
        select(
            Caixa.descricao,
            Caixa.validade,
            func.count(Caixa.id).label("qtd_etiquetas"),
            func.sum(Caixa.peso_kg).label("peso_total_kg"),
            func.count(func.distinct(Caixa.pesagem_id)).label("leituras"),
        )
        .where(Caixa.status == "em_estoque")
        .group_by(Caixa.descricao, Caixa.validade)
    ).all()
    for row in caixas_rows:
        key = (row.descricao or "", row.validade)
        acumulado[key] = {
            "descricao": row.descricao,
            "validade": row.validade,
            "qtd_etiquetas": int(row.qtd_etiquetas or 0),
            "peso_total_kg": float(row.peso_total_kg or 0),
            "leituras": int(row.leituras or 0),
        }

    agregados_rows = db.execute(
        select(
            EstoqueAgregado.produto_codigo,
            EstoqueAgregado.produto_descricao,
            EstoqueAgregado.validade,
            func.sum(EstoqueAgregado.saldo_disponivel_atual_kg).label("peso_total_kg"),
        )
        .where(EstoqueAgregado.status == "ativo", EstoqueAgregado.saldo_disponivel_atual_kg > 0)
        .group_by(EstoqueAgregado.produto_codigo, EstoqueAgregado.produto_descricao, EstoqueAgregado.validade)
    ).all()
    for row in agregados_rows:
        descricao = f"{row.produto_codigo} - {row.produto_descricao}" if row.produto_codigo else row.produto_descricao
        key = (descricao or "", row.validade)
        item = acumulado.setdefault(
            key,
            {"descricao": descricao, "validade": row.validade, "qtd_etiquetas": 0, "peso_total_kg": 0.0, "leituras": 0},
        )
        item["peso_total_kg"] += float(row.peso_total_kg or 0)

    itens = list(acumulado.values())
    if busca:
        termo = busca.lower()
        itens = [item for item in itens if termo in (item["descricao"] or "").lower()]
    if nome:
        termo = nome.lower()
        itens = [item for item in itens if termo in (item["descricao"] or "").lower()]
    if validade:
        termo = validade.lower()
        itens = [item for item in itens if termo in (item["validade"] or "").lower()]
    if peso is not None:
        itens = [item for item in itens if round(item["peso_total_kg"], 3) == round(float(peso), 3)]

    itens.sort(key=lambda item: ((item["descricao"] or ""), item["validade"] or ""))
    total = len(itens)
    rows = itens[offset : offset + limit]
    return {
        "total": total,
        "items": rows,
    }


@app.post("/api/estoque-agregado/importar")
def importar_estoque_agregado(payload: ImportarEstoquePayload, db: Session = Depends(get_db)) -> dict:
    return importar_estoque_xlsx(db, payload.arquivo_base64, payload.nome_arquivo)


@app.get("/api/caixas")
def listar_caixas(
    busca: str | None = None,
    nome: str | None = None,
    validade: str | None = None,
    peso: float | None = None,
    limit: int = Query(100, ge=1, le=500),
    offset: int = Query(0, ge=0),
    db: Session = Depends(get_db),
) -> dict:
    query = select(Caixa)
    count_query = select(func.count()).select_from(Caixa)

    filters = []
    if busca:
        pattern = f"%{busca}%"
        filters.append(or_(Caixa.descricao.ilike(pattern), Caixa.codigo_caixa.ilike(pattern), Caixa.codigo_barras_original.ilike(pattern), Caixa.impressora_nome.ilike(pattern)))
    if nome:
        filters.append(Caixa.descricao.ilike(f"%{nome}%"))
    if validade:
        filters.append(Caixa.validade.ilike(f"%{validade}%"))
    if peso is not None:
        filters.append(Caixa.peso_kg == peso)

    for condition in filters:
        query = query.where(condition)
        count_query = count_query.where(condition)

    total = db.scalar(count_query) or 0
    rows = db.scalars(query.order_by(Caixa.data_hora.desc(), Caixa.id.desc()).limit(limit).offset(offset)).all()
    return {
        "total": total,
        "items": [
            {
                "id": row.id,
                "codigo_caixa": row.codigo_caixa,
                "codigo_barras_original": row.codigo_barras_original,
                "data_hora": row.data_hora,
                "recebimento_id": row.recebimento_id,
                "numero_pallet": row.numero_pallet,
                "impressora": row.impressora_nome,
                "descricao": row.descricao,
                "peso_kg": float(row.peso_kg) if row.peso_kg is not None else None,
                "validade": row.validade,
                "status": row.status,
            }
            for row in rows
        ],
    }
