from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, scoped_session, declarative_base

# ✅ Updated database URL for `summit_2_594c`
DATABASE_URL = "postgresql://summit_2_594c_user:nONVlxoLsg6mhZgHsQfZveamzOuxNcop@dpg-d4arcia4d50c73crc21g-a.oregon-postgres.render.com:5432/summit_2_594c"

# Create SQLAlchemy engine
engine = create_engine(DATABASE_URL)

# Create a scoped session
Session = scoped_session(sessionmaker(bind=engine))
SessionLocal = Session

# Base class for ORM models
Base = declarative_base()
