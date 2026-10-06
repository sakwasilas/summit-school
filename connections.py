from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, scoped_session, declarative_base

DATABASE_URL = "mysql+pymysql://root:2480@localhost:3306/exams_db"

engine = create_engine(
    DATABASE_URL,
    pool_pre_ping=True,
    echo=True
)

SessionLocal = scoped_session(sessionmaker(bind=engine))

Base = declarative_base()
