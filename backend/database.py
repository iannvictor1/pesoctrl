from collections.abc import Generator

from sqlalchemy import create_engine, inspect, text
from sqlalchemy.orm import DeclarativeBase, Session, sessionmaker

from backend.config import get_settings


class Base(DeclarativeBase):
    pass


settings = get_settings()
engine = create_engine(settings.database_url, pool_pre_ping=True)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False, expire_on_commit=False)


def get_db() -> Generator[Session, None, None]:
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def init_db() -> None:
    from backend import models  # noqa: F401

    Base.metadata.create_all(bind=engine)
    ensure_runtime_schema()


def ensure_runtime_schema() -> None:
    inspector = inspect(engine)
    tables = inspector.get_table_names()
    if "caixas" not in tables:
        return

    dialect = engine.dialect.name
    text_type = "TEXT" if dialect != "mysql" else "LONGTEXT"
    numeric_type = "NUMERIC(12, 3)"
    int_type = "INTEGER"
    with engine.begin() as conn:
        if "inventarios" in tables:
            inventario_columns = {c["name"] for c in inspector.get_columns("inventarios")}
            if "operador_nome" not in inventario_columns:
                conn.execute(text("ALTER TABLE inventarios ADD COLUMN operador_nome VARCHAR(120)"))
        if "inventario_pallets" in tables:
            pallet_columns = {c["name"] for c in inspector.get_columns("inventario_pallets")}
            for name, size in [("unitizador", 120), ("endereco_codigo", 120), ("endereco", 500)]:
                if name not in pallet_columns:
                    conn.execute(text(f"ALTER TABLE inventario_pallets ADD COLUMN {name} VARCHAR({size})"))
        if "inventario_caixas" in tables:
            inv_columns = {c["name"] for c in inspector.get_columns("inventario_caixas")}
            if "leitura_id" not in inv_columns:
                conn.execute(text("ALTER TABLE inventario_caixas ADD COLUMN leitura_id VARCHAR(80)"))
            if "coletor_nome" not in inv_columns:
                conn.execute(text("ALTER TABLE inventario_caixas ADD COLUMN coletor_nome VARCHAR(120)"))
            if dialect == "postgresql":
                conn.execute(text("ALTER TABLE inventario_caixas DROP CONSTRAINT IF EXISTS uq_inventario_codigo"))
                conn.execute(text("CREATE UNIQUE INDEX IF NOT EXISTS uq_inventario_leitura ON inventario_caixas (inventario_id, leitura_id)"))
        columns = {column["name"] for column in inspector.get_columns("caixas")}
        if "codigo_barras_original" not in columns:
            conn.execute(text(f"ALTER TABLE caixas ADD COLUMN codigo_barras_original {text_type}"))
            if dialect == "postgresql":
                conn.execute(text("CREATE INDEX IF NOT EXISTS ix_caixas_codigo_barras_original ON caixas (codigo_barras_original)"))

        if "leituras_coletor" in tables:
            leitura_columns = {column["name"] for column in inspector.get_columns("leituras_coletor")}
            additions = {
                "estoque_agregado_id": int_type,
                "regra_baixa": "VARCHAR(60)",
                "produto_codigo": "VARCHAR(40)",
                "produto_descricao": text_type,
                "peso_lido_kg": numeric_type,
                "validade_lida": "VARCHAR(20)",
                "peso_baixado_kg": numeric_type,
            }
            for name, column_type in additions.items():
                if name not in leitura_columns:
                    conn.execute(text(f"ALTER TABLE leituras_coletor ADD COLUMN {name} {column_type}"))
