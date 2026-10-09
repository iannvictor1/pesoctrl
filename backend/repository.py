from datetime import datetime
from decimal import Decimal
import uuid

from sqlalchemy import func, select
from sqlalchemy.exc import SQLAlchemyError

from backend.database import SessionLocal, init_db
from backend.models import Caixa, Pallet, Pesagem, Recebimento


def _parse_datetime(value: str | None) -> datetime:
    if not value:
        return datetime.now()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y%m%d_%H%M%S"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass
    return datetime.now()


def _decimal_or_none(value) -> Decimal | None:
    if value is None or value == "":
        return None
    return Decimal(str(value)).quantize(Decimal("0.001"))


def _gerar_codigo_caixa(data_hora: datetime, numero_pallet: int | None, sequencia: int) -> str:
    data = data_hora.strftime("%Y%m%d%H%M%S")
    pallet = str(numero_pallet or 0).zfill(4)
    sufixo = uuid.uuid4().hex[:6].upper()
    return f"CX-{data}-P{pallet}-{str(sequencia).zfill(3)}-{sufixo}"


def _criar_caixas_para_pesagem(db, pesagem: Pesagem, codigo_barras_original: str | None = None) -> None:
    existentes = db.scalar(select(func.count()).select_from(Caixa).where(Caixa.pesagem_id == pesagem.id)) or 0
    if existentes:
        return

    qtd = max(int(pesagem.qtd_etiquetas or 1), 1)
    for sequencia in range(1, qtd + 1):
        db.add(
            Caixa(
                codigo_caixa=_gerar_codigo_caixa(pesagem.data_hora, pesagem.numero_pallet, sequencia),
                pesagem_id=pesagem.id,
                sequencia=sequencia,
                data_hora=pesagem.data_hora,
                recebimento_id=pesagem.recebimento_id,
                sessao_id=pesagem.sessao_id,
                numero_pallet=pesagem.numero_pallet,
                impressora_id=pesagem.impressora_id,
                impressora_nome=pesagem.impressora_nome,
                descricao=pesagem.descricao,
                codigo_barras_original=codigo_barras_original,
                peso_kg=pesagem.peso_kg,
                validade=pesagem.validade,
            )
        )


def backfill_caixas_pesagens_existentes() -> None:
    init_db()
    with SessionLocal() as db:
        pesagens = db.scalars(
            select(Pesagem)
            .outerjoin(Caixa, Caixa.pesagem_id == Pesagem.id)
            .group_by(Pesagem.id)
            .having(func.count(Caixa.id) == 0)
        ).all()
        for pesagem in pesagens:
            _criar_caixas_para_pesagem(db, pesagem)
        db.commit()


def salvar_recebimento_iniciado(recebimento_id: str, inicio: str | None = None) -> None:
    init_db()
    with SessionLocal() as db:
        rec = db.scalar(select(Recebimento).where(Recebimento.recebimento_id == recebimento_id))
        if rec is None:
            db.add(Recebimento(recebimento_id=recebimento_id, inicio=_parse_datetime(inicio), status="ativo"))
        else:
            rec.status = "ativo"
            rec.fim = None
        db.commit()


def salvar_recebimento_encerrado(recebimento_id: str, total_pallets: int) -> None:
    init_db()
    with SessionLocal() as db:
        rec = db.scalar(select(Recebimento).where(Recebimento.recebimento_id == recebimento_id))
        if rec is None:
            rec = Recebimento(recebimento_id=recebimento_id, inicio=datetime.now())
            db.add(rec)
        rec.status = "encerrado"
        rec.fim = datetime.now()
        rec.total_pallets = int(total_pallets or 0)
        db.commit()


def salvar_pallet_iniciado(cfg: dict, status: dict) -> None:
    init_db()
    recebimento_id = status.get("recebimento_id")
    if not recebimento_id:
        return

    with SessionLocal() as db:
        rec = db.scalar(select(Recebimento).where(Recebimento.recebimento_id == recebimento_id))
        pallet = db.scalar(
            select(Pallet).where(
                Pallet.recebimento_id_texto == recebimento_id,
                Pallet.numero_pallet == int(status.get("numero_pallet") or 0),
                Pallet.impressora_id == cfg["id"],
            )
        )
        if pallet is None:
            pallet = Pallet(
                recebimento_db_id=rec.id if rec else None,
                recebimento_id_texto=recebimento_id,
                sessao_id=status.get("sessao_id") or "",
                numero_pallet=int(status.get("numero_pallet") or 0),
                impressora_id=cfg["id"],
                impressora_nome=cfg["nome_amigavel"],
                descricao_produto=status.get("descricao_produto"),
                inicio=_parse_datetime(status.get("inicio_sessao")),
                status="ativo",
            )
            db.add(pallet)
        else:
            pallet.status = "ativo"
            pallet.fim = None
            pallet.sessao_id = status.get("sessao_id") or pallet.sessao_id
        db.commit()


def salvar_pallet_encerrado(cfg: dict, status: dict, arquivo_excel: str | None) -> None:
    init_db()
    with SessionLocal() as db:
        pallet = db.scalar(select(Pallet).where(Pallet.sessao_id == (status.get("sessao_id") or "")))
        if pallet:
            pallet.status = "encerrado"
            pallet.fim = datetime.now()
            pallet.descricao_produto = status.get("descricao_produto") or pallet.descricao_produto
            pallet.arquivo_excel = arquivo_excel
            db.commit()


def salvar_pesagem(cfg: dict, status: dict, registro: dict, arquivo_fila: str, arquivo_historico: str) -> None:
    init_db()
    with SessionLocal() as db:
        pallet = None
        if status.get("sessao_id"):
            pallet = db.scalar(select(Pallet).where(Pallet.sessao_id == status.get("sessao_id")))

        pesagem = Pesagem(
            data_hora=_parse_datetime(registro.get("DataHora")),
            recebimento_id=status.get("recebimento_id"),
            pallet_db_id=pallet.id if pallet else None,
            sessao_id=status.get("sessao_id"),
            numero_pallet=status.get("numero_pallet"),
            impressora_id=cfg["id"],
            impressora_nome=cfg["nome_amigavel"],
            descricao=registro.get("Descricao") or "",
            peso_kg=_decimal_or_none(registro.get("PesoKG")),
            qtd_etiquetas=int(registro.get("QtdEtiq") or 1),
            peso_total_kg=_decimal_or_none(registro.get("PesoTotalKG")),
            validade=registro.get("Validade") or None,
            arquivo_fila=arquivo_fila,
            arquivo_historico=arquivo_historico,
        )
        db.add(pesagem)
        db.flush()
        _criar_caixas_para_pesagem(db, pesagem, registro.get("CodigoBarrasOriginal") or None)
        if pallet and not pallet.descricao_produto and pesagem.descricao:
            pallet.descricao_produto = pesagem.descricao
        db.commit()


def tentar_salvar(acao, *args, **kwargs) -> tuple[bool, str | None]:
    try:
        acao(*args, **kwargs)
        return True, None
    except SQLAlchemyError as exc:
        return False, str(exc)
    except Exception as exc:
        return False, str(exc)
