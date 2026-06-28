import os

from sqlalchemy import create_engine, inspect, text
from sqlalchemy.orm import sessionmaker, DeclarativeBase

DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./inventory.db")

engine = create_engine(
    DATABASE_URL,
    connect_args={"check_same_thread": False},
)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)


class Base(DeclarativeBase):
    pass


def ensure_schema(engine_to_fix=engine):
    inspector = inspect(engine_to_fix)
    if "products" in inspector.get_table_names():
        product_columns = {column["name"] for column in inspector.get_columns("products")}
        if "unit_price" not in product_columns:
            with engine_to_fix.begin() as conn:
                conn.execute(text("ALTER TABLE products ADD COLUMN unit_price NUMERIC(12, 2) NOT NULL DEFAULT 0"))
        if "substrate_category" not in product_columns:
            with engine_to_fix.begin() as conn:
                conn.execute(text("ALTER TABLE products ADD COLUMN substrate_category TEXT NOT NULL DEFAULT ''"))
        if "surface_type" not in product_columns:
            with engine_to_fix.begin() as conn:
                conn.execute(text("ALTER TABLE products ADD COLUMN surface_type TEXT NOT NULL DEFAULT ''"))

    inspector = inspect(engine_to_fix)
    if "order_lines" in inspector.get_table_names():
        order_line_columns = {column["name"] for column in inspector.get_columns("order_lines")}
        if "inbound_qty" not in order_line_columns:
            with engine_to_fix.begin() as conn:
                conn.execute(text("ALTER TABLE order_lines ADD COLUMN inbound_qty INTEGER NOT NULL DEFAULT 0"))
        if "unit_price" not in order_line_columns:
            with engine_to_fix.begin() as conn:
                conn.execute(text("ALTER TABLE order_lines ADD COLUMN unit_price NUMERIC(12, 2) NOT NULL DEFAULT 0"))
        if "total_price" not in order_line_columns:
            with engine_to_fix.begin() as conn:
                conn.execute(text("ALTER TABLE order_lines ADD COLUMN total_price NUMERIC(12, 2) NOT NULL DEFAULT 0"))
        if "substrate_category" not in order_line_columns:
            with engine_to_fix.begin() as conn:
                conn.execute(text("ALTER TABLE order_lines ADD COLUMN substrate_category TEXT NOT NULL DEFAULT ''"))
        if "surface_type" not in order_line_columns:
            with engine_to_fix.begin() as conn:
                conn.execute(text("ALTER TABLE order_lines ADD COLUMN surface_type TEXT NOT NULL DEFAULT ''"))

    if "inbound_orders" in inspector.get_table_names():
        inbound_columns = {column["name"] for column in inspector.get_columns("inbound_orders")}
        if "related_order_id" not in inbound_columns:
            with engine_to_fix.begin() as conn:
                conn.execute(text("ALTER TABLE inbound_orders ADD COLUMN related_order_id INTEGER"))

    inspector = inspect(engine_to_fix)
    if "outbound_orders" in inspector.get_table_names():
        outbound_columns = {column["name"] for column in inspector.get_columns("outbound_orders")}
        if "related_inbound_id" not in outbound_columns:
            with engine_to_fix.begin() as conn:
                conn.execute(text("ALTER TABLE outbound_orders ADD COLUMN related_inbound_id INTEGER"))

    inspector = inspect(engine_to_fix)
    if "outbound_order_lines" in inspector.get_table_names():
        outbound_line_columns = {column["name"] for column in inspector.get_columns("outbound_order_lines")}
        if "unit_price" not in outbound_line_columns:
            with engine_to_fix.begin() as conn:
                conn.execute(text("ALTER TABLE outbound_order_lines ADD COLUMN unit_price NUMERIC(12, 2) NOT NULL DEFAULT 0"))
        if "total_price" not in outbound_line_columns:
            with engine_to_fix.begin() as conn:
                conn.execute(text("ALTER TABLE outbound_order_lines ADD COLUMN total_price NUMERIC(12, 2) NOT NULL DEFAULT 0"))


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
