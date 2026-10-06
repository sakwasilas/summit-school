# from sqlalchemy import create_engine
# from sqlalchemy.orm import sessionmaker, scoped_session, declarative_base

# DATABASE_URL = "mysql+pymysql://root:2480@localhost:3306/exams_db"

# engine = create_engine(
#     DATABASE_URL,
#     pool_pre_ping=True,
#     echo=True
# )

# SessionLocal = scoped_session(sessionmaker(bind=engine))

# Base = declarative_base()

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, scoped_session, declarative_base

DATABASE_URL = "postgresql://summit_school_user:s1pSq54ogYX5M0JLH1N2gdTc7rwrqXeC@dpg-db29otvlot8c73edlps0-a.oregon-postgres.render.com/summit_school"

engine = create_engine(
    DATABASE_URL,
    pool_pre_ping=True,
    echo=True
)

SessionLocal = scoped_session(sessionmaker(bind=engine))

Base = declarative_base()
