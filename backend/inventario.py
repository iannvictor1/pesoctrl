from datetime import datetime
from decimal import Decimal
import json
from pathlib import Path

from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import Response
from pydantic import BaseModel, Field
from sqlalchemy import select, func
from sqlalchemy.orm import Session

from backend.database import get_db
from backend.models import Inventario, InventarioPallet, InventarioCaixa
from backend.coletor_service import (
    normalizar_codigo, interpretar_etiqueta, _codigos_produto_conhecidos,
    _encontrar_caixa, _produto_descricao,
)

router = APIRouter(prefix="/api/inventarios", tags=["Inventário"])
ENDERECOS = json.loads((Path(__file__).parent / "data" / "enderecos_inventario.json").read_text(encoding="utf-8"))


class InventarioPayload(BaseModel):
    operador_nome: str = Field(min_length=1, max_length=120)


class PalletPayload(BaseModel):
    unitizador: str = Field(min_length=1, max_length=120)


class EnderecoPayload(BaseModel):
    codigo: str = Field(min_length=1, max_length=120)


class LeituraPayload(BaseModel):
    coletor_nome: str | None = Field(default=None, max_length=120)
    leitura_id: str | None = Field(default=None, min_length=1, max_length=80)
    codigo: str = Field(min_length=1, max_length=512)
    descricao: str | None = Field(default=None, max_length=500)
    validade: str | None = None
    peso_kg: Decimal | None = Field(default=None, gt=0, max_digits=12, decimal_places=3)


def obter_inventario(db, inventario_id, aberto=False):
    row = db.scalar(select(Inventario).where(Inventario.id == inventario_id).with_for_update())
    if row is None:
        raise HTTPException(404, "Inventário não encontrado.")
    if aberto and row.status != "aberto":
        raise HTTPException(409, "Este inventário está encerrado.")
    return row


def obter_pallet(db, inventario_id, pallet_id):
    row = db.get(InventarioPallet, pallet_id)
    if row is None or row.inventario_id != inventario_id:
        raise HTTPException(404, "Pallet não pertence a este inventário.")
    if row.status != "aberto":
        raise HTTPException(409, "Este pallet está encerrado.")
    return row


def campos(row):
    return {column.name: getattr(row, column.name) for column in row.__table__.columns}


@router.get("")
def listar(db: Session = Depends(get_db)):
    return [campos(row) for row in db.scalars(select(Inventario).order_by(Inventario.id.desc()))]


@router.post("")
def criar(payload: InventarioPayload, db: Session = Depends(get_db)):
    operador_nome = payload.operador_nome.strip()
    if not operador_nome:
        raise HTTPException(422, "Informe o nome do operador do inventário.")
    row = Inventario(operador_nome=operador_nome)
    db.add(row)
    db.commit()
    return campos(row)


@router.get("/resumos")
def resumos(db: Session = Depends(get_db)):
    pallets = select(
        InventarioPallet.inventario_id.label("inventario_id"),
        func.count(InventarioPallet.id).label("total_pallets"),
    ).group_by(InventarioPallet.inventario_id).subquery()
    caixas = select(
        InventarioCaixa.inventario_id.label("inventario_id"),
        func.count(InventarioCaixa.id).label("total_caixas"),
        func.sum(InventarioCaixa.peso_kg).label("peso_total_kg"),
    ).group_by(InventarioCaixa.inventario_id).subquery()
    rows = db.execute(select(
        Inventario,
        func.coalesce(pallets.c.total_pallets, 0),
        func.coalesce(caixas.c.total_caixas, 0),
        func.coalesce(caixas.c.peso_total_kg, 0),
    ).outerjoin(pallets, pallets.c.inventario_id == Inventario.id)
      .outerjoin(caixas, caixas.c.inventario_id == Inventario.id)
      .order_by(Inventario.id.desc())).all()
    return [{**campos(row), "total_pallets": np, "total_caixas": nc, "peso_total_kg": peso}
            for row, np, nc, peso in rows]


@router.get("/{inventario_id}")
def detalhe(inventario_id: int, db: Session = Depends(get_db)):
    row = db.get(Inventario, inventario_id)
    if row is None:
        raise HTTPException(404, "Inventário não encontrado.")
    pallets = db.scalars(select(InventarioPallet).where(InventarioPallet.inventario_id == inventario_id).order_by(InventarioPallet.id)).all()
    caixas = db.scalars(select(InventarioCaixa).where(InventarioCaixa.inventario_id == inventario_id).order_by(InventarioCaixa.id.desc())).all()
    return {**campos(row), "pallets": [{**campos(p), "caixas": [campos(c) for c in caixas if c.pallet_id == p.id]} for p in pallets]}


@router.get("/{inventario_id}/exportar")
def exportar(inventario_id: int, db: Session = Depends(get_db)):
    from backend.inventario_excel import gerar_planilha

    # Mesmo bloqueio das gravações para capturar um retrato consistente.
    obter_inventario(db, inventario_id)
    dados = detalhe(inventario_id, db)
    db.commit()
    return Response(
        content=gerar_planilha(dados),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="inventario_{inventario_id}.xlsx"', "Cache-Control": "no-store"},
    )


@router.post("/{inventario_id}/encerrar")
def encerrar(inventario_id: int, db: Session = Depends(get_db)):
    row = obter_inventario(db, inventario_id, True)
    if db.scalar(select(InventarioPallet.id).where(InventarioPallet.inventario_id == inventario_id, InventarioPallet.status == "aberto")):
        raise HTTPException(409, "Encerre os pallets abertos antes de encerrar o inventário.")
    row.status, row.fim = "encerrado", datetime.now()
    db.commit()
    return campos(row)


@router.post("/{inventario_id}/pallets")
def criar_pallet(inventario_id: int, payload: PalletPayload, db: Session = Depends(get_db)):
    obter_inventario(db, inventario_id, True)
    unitizador = payload.unitizador.strip()
    if not unitizador:
        raise HTTPException(422, "Bipe o unitizador do pallet.")
    row = InventarioPallet(inventario_id=inventario_id, unitizador=unitizador)
    db.add(row)
    db.commit()
    return campos(row)


@router.post("/{inventario_id}/pallets/{pallet_id}/endereco")
def identificar_endereco(inventario_id: int, pallet_id: int, payload: EnderecoPayload, db: Session = Depends(get_db)):
    obter_inventario(db, inventario_id, True)
    row = obter_pallet(db, inventario_id, pallet_id)
    codigo = payload.codigo.strip()
    endereco = ENDERECOS.get(codigo)
    if endereco is None:
        raise HTTPException(422, "Código de endereço não encontrado na planilha cadastrada. Confira a etiqueta e bipe novamente.")
    if row.endereco_codigo and row.endereco_codigo != codigo:
        raise HTTPException(409, "Este pallet já possui endereço. Bipe um unitizador para criar outro pallet.")
    row.endereco_codigo, row.endereco = codigo, endereco
    db.commit()
    return campos(row)


@router.post("/{inventario_id}/pallets/{pallet_id}/encerrar")
def encerrar_pallet(inventario_id: int, pallet_id: int, db: Session = Depends(get_db)):
    obter_inventario(db, inventario_id, True)
    row = obter_pallet(db, inventario_id, pallet_id)
    row.status, row.fim = "encerrado", datetime.now()
    db.commit()
    return campos(row)


@router.post("/{inventario_id}/pallets/{pallet_id}/caixas")
def adicionar(inventario_id: int, pallet_id: int, payload: LeituraPayload, db: Session = Depends(get_db)):
    obter_inventario(db, inventario_id, True)
    pallet = obter_pallet(db, inventario_id, pallet_id)
    if not pallet.unitizador or not pallet.endereco_codigo:
        raise HTTPException(409, "Bipe o unitizador e o endereço antes de registrar caixas.")
    codigo = normalizar_codigo(payload.codigo)
    if not codigo:
        raise HTTPException(422, "Leia o código da caixa.")
    if payload.leitura_id:
        existente = db.scalar(select(InventarioCaixa).where(InventarioCaixa.inventario_id == inventario_id, InventarioCaixa.leitura_id == payload.leitura_id))
        if existente:
            if existente.codigo != codigo or existente.pallet_id != pallet_id:
                raise HTTPException(409, "Identificador de leitura já utilizado para outra caixa.")
            return campos(existente)
    caixa = _encontrar_caixa(db, payload.codigo, codigo)
    dados = interpretar_etiqueta(codigo, _codigos_produto_conhecidos(db))
    descricao = (payload.descricao or (caixa.descricao if caixa else None) or _produto_descricao(db, dados['produto_codigo']) or '').strip()
    peso = payload.peso_kg or (caixa.peso_kg if caixa else None) or dados['peso_lido_kg']
    validade = payload.validade or (caixa.validade if caixa else None) or dados['validade_lida']
    try:
        data = datetime.strptime(validade or '', '%Y-%m-%d' if '-' in (validade or '') else '%d/%m/%Y')
        validade = data.strftime('%d/%m/%Y')
    except ValueError:
        validade = None
    if not descricao or not validade or peso is None or not peso.is_finite() or peso <= 0 or peso >= Decimal('1000000000'):
        raise HTTPException(422, "Não foi possível identificar nome, validade e peso. Complete os dados da caixa abaixo e tente novamente.")
    row = InventarioCaixa(inventario_id=inventario_id, pallet_id=pallet_id, codigo=codigo, leitura_id=payload.leitura_id, coletor_nome=payload.coletor_nome, descricao=descricao, validade=validade, peso_kg=peso)
    db.add(row)
    db.commit()
    return campos(row)
