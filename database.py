import os
from collections.abc import Generator

from sqlalchemy import create_engine
from sqlalchemy.orm import DeclarativeBase, Session, sessionmaker


def database_url() -> str:
    value = os.getenv("DATABASE_URL", "sqlite:///./dues_desk.db")
    if value.startswith("postgres://"):
        return "postgresql+psycopg://" + value.removeprefix("postgres://")
    if value.startswith("postgresql://"):
        return "postgresql+psycopg://" + value.removeprefix("postgresql://")
    return value


DATABASE_URL = database_url()
IS_SQLITE = DATABASE_URL.startswith("sqlite:")
if os.getenv("APP_ENV") == "production" and not DATABASE_URL.startswith("postgresql+psycopg://"):
    raise RuntimeError("Hosted production requires DATABASE_URL to use PostgreSQL.")

engine = create_engine(
    DATABASE_URL,
    connect_args={"check_same_thread": False} if IS_SQLITE else {},
    pool_pre_ping=not IS_SQLITE,
)
SessionLocal = sessionmaker(bind=engine, autoflush=False, expire_on_commit=False)


class Base(DeclarativeBase):
    pass


def get_db() -> Generator[Session, None, None]:
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
