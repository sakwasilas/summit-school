# import os
# from dotenv import load_dotenv
# from sqlalchemy import create_engine
# from sqlalchemy.orm import sessionmaker, scoped_session, declarative_base

# load_dotenv()

# DATABASE_URL = os.getenv("DATABASE_URL")

# engine = create_engine(
#     DATABASE_URL,
#     pool_pre_ping=True,
#     echo=True
# )

# SessionLocal = scoped_session(sessionmaker(bind=engine))

# Base = declarative_base()

import os
from dotenv import load_dotenv
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, scoped_session, declarative_base

load_dotenv()

DATABASE_URL = os.getenv("DATABASE_URL")

if not DATABASE_URL:
    raise RuntimeError(
        "DATABASE_URL is not set. "
        "Locally: add it to your .env file. "
        "On Render: add it under Environment."
    )

# MySQL on Aiven requires SSL. The `ssl: {}` connect_arg tells PyMySQL to
# negotiate TLS with the server's default CA bundle.
connect_args = {}
if DATABASE_URL.startswith("mysql"):
    connect_args["ssl"] = {}

engine = create_engine(
    DATABASE_URL,
    pool_pre_ping=True,
    pool_recycle=280,
    echo=False,                # set to True temporarily if you need to see SQL
    connect_args=connect_args,
)

SessionLocal = scoped_session(
    sessionmaker(
        bind=engine,
        autoflush=False,
        autocommit=False,
    )
)

Base = declarative_base()