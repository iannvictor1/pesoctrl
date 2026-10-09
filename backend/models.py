from datetime import datetime
from decimal import Decimal

from sqlalchemy import Boolean, DateTime, ForeignKey, Integer, Numeric, String, Text, UniqueConstraint
from sqlalchemy.orm import Mapped, mapped_column, relationship

from backend.database import Base


class Inventario(Base):
    __tablename__ = "inventarios"
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    operador_nome: Mapped[str | None] = mapped_column(String(120), nullable=True)
    inicio: Mapped[datetime] = mapped_column(DateTime, default=datetime.now)
    fim: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    status: Mapped[str] = mapped_column(String(20), default="aberto")


class InventarioPallet(Base):
    __tablename__ = "inventario_pallets"
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    inventario_id: Mapped[int] = mapped_column(ForeignKey("inventarios.id"), index=True)
    unitizador: Mapped[str | None] = mapped_column(String(120), nullable=True)
    endereco_codigo: Mapped[str | None] = mapped_column(String(120), nullable=True)
    endereco: Mapped[str | None] = mapped_column(String(500), nullable=True)
    inicio: Mapped[datetime] = mapped_column(DateTime, default=datetime.now)
    fim: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    status: Mapped[str] = mapped_column(String(20), default="aberto")


class InventarioCaixa(Base):
    __tablename__ = "inventario_caixas"
    __table_args__ = (UniqueConstraint("inventario_id", "leitura_id", name="uq_inventario_leitura"),)
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    inventario_id: Mapped[int] = mapped_column(ForeignKey("inventarios.id"), index=True)
    pallet_id: Mapped[int] = mapped_column(ForeignKey("inventario_pallets.id"), index=True)
    codigo: Mapped[str] = mapped_column(String(512))
    leitura_id: Mapped[str | None] = mapped_column(String(80), nullable=True)
    coletor_nome: Mapped[str | None] = mapped_column(String(120), nullable=True)
    descricao: Mapped[str] = mapped_column(Text)
    validade: Mapped[str] = mapped_column(String(20))
    peso_kg: Mapped[Decimal] = mapped_column(Numeric(12, 3))
    data_hora: Mapped[datetime] = mapped_column(DateTime, default=datetime.now)


class Recebimento(Base):
    __tablename__ = "recebimentos"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, index=True)
    recebimento_id: Mapped[str] = mapped_column(String(40), unique=True, index=True)
    inicio: Mapped[datetime] = mapped_column(DateTime, default=datetime.now, index=True)
    fim: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    status: Mapped[str] = mapped_column(String(20), default="ativo", index=True)
    total_pallets: Mapped[int] = mapped_column(Integer, default=0)

    pallets: Mapped[list["Pallet"]] = relationship(back_populates="recebimento")


class Pallet(Base):
    __tablename__ = "pallets"
    __table_args__ = (
        UniqueConstraint("recebimento_id_texto", "numero_pallet", "impressora_id", name="uq_pallet_operacao"),
    )

    id: Mapped[int] = mapped_column(Integer, primary_key=True, index=True)
    recebimento_db_id: Mapped[int | None] = mapped_column(ForeignKey("recebimentos.id"), nullable=True)
    recebimento_id_texto: Mapped[str] = mapped_column(String(40), index=True)
    sessao_id: Mapped[str] = mapped_column(String(40), index=True)
    numero_pallet: Mapped[int] = mapped_column(Integer, index=True)
    impressora_id: Mapped[str] = mapped_column(String(50), index=True)
    impressora_nome: Mapped[str] = mapped_column(String(120))
    descricao_produto: Mapped[str | None] = mapped_column(Text, nullable=True)
    inicio: Mapped[datetime] = mapped_column(DateTime, default=datetime.now, index=True)
    fim: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    status: Mapped[str] = mapped_column(String(20), default="ativo", index=True)
    arquivo_excel: Mapped[str | None] = mapped_column(Text, nullable=True)

    recebimento: Mapped[Recebimento | None] = relationship(back_populates="pallets")
    pesagens: Mapped[list["Pesagem"]] = relationship(back_populates="pallet")


class Pesagem(Base):
    __tablename__ = "pesagens"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, index=True)
    data_hora: Mapped[datetime] = mapped_column(DateTime, default=datetime.now, index=True)
    recebimento_id: Mapped[str | None] = mapped_column(String(40), index=True)
    pallet_db_id: Mapped[int | None] = mapped_column(ForeignKey("pallets.id"), nullable=True)
    sessao_id: Mapped[str | None] = mapped_column(String(40), index=True)
    numero_pallet: Mapped[int | None] = mapped_column(Integer, index=True)
    impressora_id: Mapped[str] = mapped_column(String(50), index=True)
    impressora_nome: Mapped[str] = mapped_column(String(120))
    descricao: Mapped[str] = mapped_column(Text, index=True)
    peso_kg: Mapped[Decimal | None] = mapped_column(Numeric(12, 3), nullable=True)
    qtd_etiquetas: Mapped[int] = mapped_column(Integer, default=1)
    peso_total_kg: Mapped[Decimal | None] = mapped_column(Numeric(12, 3), nullable=True)
    validade: Mapped[str | None] = mapped_column(String(20), nullable=True, index=True)
    arquivo_fila: Mapped[str | None] = mapped_column(Text, nullable=True)
    arquivo_historico: Mapped[str | None] = mapped_column(Text, nullable=True)

    pallet: Mapped[Pallet | None] = relationship(back_populates="pesagens")
    caixas: Mapped[list["Caixa"]] = relationship(back_populates="pesagem")


class Caixa(Base):
    __tablename__ = "caixas"
    __table_args__ = (
        UniqueConstraint("pesagem_id", "sequencia", name="uq_caixa_pesagem_sequencia"),
    )

    id: Mapped[int] = mapped_column(Integer, primary_key=True, index=True)
    codigo_caixa: Mapped[str] = mapped_column(String(40), unique=True, index=True)
    pesagem_id: Mapped[int] = mapped_column(ForeignKey("pesagens.id"), index=True)
    sequencia: Mapped[int] = mapped_column(Integer)
    data_hora: Mapped[datetime] = mapped_column(DateTime, default=datetime.now, index=True)
    recebimento_id: Mapped[str | None] = mapped_column(String(40), index=True)
    sessao_id: Mapped[str | None] = mapped_column(String(40), index=True)
    numero_pallet: Mapped[int | None] = mapped_column(Integer, index=True)
    impressora_id: Mapped[str] = mapped_column(String(50), index=True)
    impressora_nome: Mapped[str] = mapped_column(String(120))
    descricao: Mapped[str] = mapped_column(Text, index=True)
    codigo_barras_original: Mapped[str | None] = mapped_column(Text, nullable=True, index=True)
    peso_kg: Mapped[Decimal | None] = mapped_column(Numeric(12, 3), nullable=True, index=True)
    validade: Mapped[str | None] = mapped_column(String(20), nullable=True, index=True)
    status: Mapped[str] = mapped_column(String(20), default="em_estoque", index=True)

    pesagem: Mapped[Pesagem] = relationship(back_populates="caixas")


class EstoqueAgregado(Base):
    __tablename__ = "estoque_agregado"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, index=True)
    importacao_id: Mapped[str] = mapped_column(String(40), index=True)
    arquivo_origem: Mapped[str | None] = mapped_column(Text, nullable=True)
    importado_em: Mapped[datetime] = mapped_column(DateTime, default=datetime.now, index=True)
    deposito: Mapped[str | None] = mapped_column(String(120), nullable=True, index=True)
    endereco: Mapped[str | None] = mapped_column(Text, nullable=True)
    produto_codigo: Mapped[str] = mapped_column(String(40), index=True)
    produto_descricao: Mapped[str] = mapped_column(Text, index=True)
    unitizador: Mapped[str | None] = mapped_column(String(120), nullable=True, index=True)
    condicao: Mapped[str | None] = mapped_column(String(80), nullable=True)
    situacao: Mapped[str | None] = mapped_column(String(80), nullable=True)
    saldo_fisico_kg: Mapped[Decimal] = mapped_column(Numeric(12, 3), default=0)
    saldo_reservado_kg: Mapped[Decimal] = mapped_column(Numeric(12, 3), default=0)
    saldo_disponivel_inicial_kg: Mapped[Decimal] = mapped_column(Numeric(12, 3), default=0)
    saldo_disponivel_atual_kg: Mapped[Decimal] = mapped_column(Numeric(12, 3), default=0, index=True)
    bloqueado_movimentacao_kg: Mapped[Decimal] = mapped_column(Numeric(12, 3), default=0)
    validade: Mapped[str | None] = mapped_column(String(20), nullable=True, index=True)
    lote: Mapped[str | None] = mapped_column(String(120), nullable=True, index=True)
    status: Mapped[str] = mapped_column(String(20), default="ativo", index=True)


class Coletor(Base):
    __tablename__ = "coletores"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, index=True)
    nome: Mapped[str] = mapped_column(String(80), unique=True, index=True)
    serial: Mapped[str] = mapped_column(String(120), unique=True, index=True)
    ip_atual: Mapped[str | None] = mapped_column(String(60), nullable=True)
    status: Mapped[str] = mapped_column(String(20), default="offline", index=True)
    ultimo_contato: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    ultimo_bipe: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    ultimo_codigo_lido: Mapped[str | None] = mapped_column(Text, nullable=True)
    criado_em: Mapped[datetime] = mapped_column(DateTime, default=datetime.now)

    sessoes: Mapped[list["SessaoConferencia"]] = relationship(back_populates="coletor")
    leituras: Mapped[list["LeituraColetor"]] = relationship(back_populates="coletor")


class SessaoConferencia(Base):
    __tablename__ = "sessoes_conferencia"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, index=True)
    coletor_id: Mapped[int] = mapped_column(ForeignKey("coletores.id"), index=True)
    inicio: Mapped[datetime] = mapped_column(DateTime, default=datetime.now, index=True)
    fim: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    status: Mapped[str] = mapped_column(String(20), default="ativa", index=True)
    operador: Mapped[str | None] = mapped_column(String(120), nullable=True)
    observacao: Mapped[str | None] = mapped_column(Text, nullable=True)
    log_limpo: Mapped[bool] = mapped_column(Boolean, default=False)
    total_leituras: Mapped[int] = mapped_column(Integer, default=0)
    baixadas: Mapped[int] = mapped_column(Integer, default=0)
    duplicadas: Mapped[int] = mapped_column(Integer, default=0)
    nao_encontradas: Mapped[int] = mapped_column(Integer, default=0)
    ignoradas: Mapped[int] = mapped_column(Integer, default=0)

    coletor: Mapped[Coletor] = relationship(back_populates="sessoes")
    leituras: Mapped[list["LeituraColetor"]] = relationship(back_populates="sessao")


class LeituraColetor(Base):
    __tablename__ = "leituras_coletor"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, index=True)
    coletor_id: Mapped[int] = mapped_column(ForeignKey("coletores.id"), index=True)
    sessao_id: Mapped[int | None] = mapped_column(ForeignKey("sessoes_conferencia.id"), nullable=True, index=True)
    caixa_id: Mapped[int | None] = mapped_column(ForeignKey("caixas.id"), nullable=True, index=True)
    estoque_agregado_id: Mapped[int | None] = mapped_column(ForeignKey("estoque_agregado.id"), nullable=True, index=True)
    data_hora: Mapped[datetime] = mapped_column(DateTime, default=datetime.now, index=True)
    codigo_lido: Mapped[str] = mapped_column(Text, index=True)
    codigo_normalizado: Mapped[str] = mapped_column(Text, index=True)
    status: Mapped[str] = mapped_column(String(40), default="recebida", index=True)
    regra_baixa: Mapped[str | None] = mapped_column(String(60), nullable=True, index=True)
    produto_codigo: Mapped[str | None] = mapped_column(String(40), nullable=True, index=True)
    produto_descricao: Mapped[str | None] = mapped_column(Text, nullable=True)
    peso_lido_kg: Mapped[Decimal | None] = mapped_column(Numeric(12, 3), nullable=True)
    validade_lida: Mapped[str | None] = mapped_column(String(20), nullable=True, index=True)
    peso_baixado_kg: Mapped[Decimal | None] = mapped_column(Numeric(12, 3), nullable=True)
    origem: Mapped[str] = mapped_column(String(40), default="logcat")
    raw_line: Mapped[str | None] = mapped_column(Text, nullable=True)
    criado_em: Mapped[datetime] = mapped_column(DateTime, default=datetime.now)

    coletor: Mapped[Coletor] = relationship(back_populates="leituras")
    sessao: Mapped[SessaoConferencia | None] = relationship(back_populates="leituras")
