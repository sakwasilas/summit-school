from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, scoped_session, declarative_base

DATABASE_URL = "postgresql+psycopg2://summit_249y_user:Pmiu7B4lZaFaLpcpHPynChPClsgWpcEm@dpg-d76j771aae7s73c6dspg-a.oregon-postgres.render.com:5432/summit_249y"

engine = create_engine(
    DATABASE_URL,
    pool_pre_ping=True,
    echo=True
)

SessionLocal = scoped_session(sessionmaker(bind=engine))

Base = declarative_base()