from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, scoped_session, declarative_base

DATABASE_URL = "postgresql+psycopg2://cpajuly_user:pr0MTct7AOSmN1b8Wk4YBB9f2lYgaJyN@dpg-d9b2gv3tqb8s73a5sjeg-a.oregon-postgres.render.com:5432/cpajuly"

engine = create_engine(
    DATABASE_URL,
    pool_pre_ping=True,
    echo=True
)

SessionLocal = scoped_session(sessionmaker(bind=engine))

Base = declarative_base()