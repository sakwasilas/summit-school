from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, scoped_session, declarative_base

# ✅ Render PostgreSQL database URL
DATABASE_URL = "postgresql+psycopg2://kasneb_exams_user:6MRZVYQH4IlueDPteIajtSNfRbBdCUl7@dpg-d6tv3lvdiees73d74k3g-a.oregon-postgres.render.com:5432/kasneb_exams"

engine = create_engine(DATABASE_URL)

# Create a scoped session
Session = scoped_session(sessionmaker(bind=engine))
SessionLocal = Session

# Base class for ORM models
Base = declarative_base()