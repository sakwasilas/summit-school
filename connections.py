from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, scoped_session, declarative_base

# ✅ MySQL database URL (using PyMySQL)
DATABASE_URL = "mysql+pymysql://root:2480@localhost/exams_db"

engine = create_engine(
    DATABASE_URL,
    pool_pre_ping=True,
    echo=True
)

# Scoped session
SessionLocal = scoped_session(sessionmaker(bind=engine))

# Base class for ORM models
Base = declarative_base()