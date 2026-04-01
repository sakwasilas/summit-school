# from sqlalchemy import create_engine
# from sqlalchemy.orm import sessionmaker, scoped_session, declarative_base

# # ✅ MySQL database URL (using PyMySQL)
# DATABASE_URL = "mysql+pymysql://root:2480@localhost/exams_db"

# engine = create_engine(
#     DATABASE_URL,
#     pool_pre_ping=True,
#     echo=True
# )

# # Scoped session
# SessionLocal = scoped_session(sessionmaker(bind=engine))

# # Base class for ORM models
# Base = declarative_base()

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, scoped_session, declarative_base

# ✅ PostgreSQL database URL (Render)
DATABASE_URL = "postgresql://summit_249y_user:Pmiu7B4lZaFaLpcpHPynChPClsgWpcEm@dpg-d76j771aae7s73c6dspg-a.oregon-postgres.render.com:5432/summit_249y"

# Create the SQLAlchemy engine
engine = create_engine(
    DATABASE_URL,
    pool_pre_ping=True,
    echo=True  # Set to False in production
)

# Scoped session
SessionLocal = scoped_session(sessionmaker(bind=engine))

# Base class for ORM models
Base = declarative_base()