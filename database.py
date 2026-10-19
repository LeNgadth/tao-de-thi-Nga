
import os
from sqlalchemy import create_engine, Column, Integer, String, Text, inspect, text
from sqlalchemy.orm import declarative_base, sessionmaker

DATABASE_URL = os.getenv(
    "DATABASE_URL",
    "sqlite:///ngan_hang_cau_hoi.db"
)

connect_args = (
    {"check_same_thread": False}
    if DATABASE_URL.startswith("sqlite")
    else {}
)

engine = create_engine(
    DATABASE_URL,
    connect_args=connect_args,
    pool_pre_ping=True
)

SessionLocal = sessionmaker(
    autocommit=False,
    autoflush=False,
    bind=engine
)

Base = declarative_base()


class CauHoi(Base):
    __tablename__ = "cau_hoi"

    id = Column(Integer, primary_key=True, index=True)
    mon_hoc = Column(String(100), nullable=False, index=True)
    khoi_lop = Column(String(10), nullable=False, index=True)
    dang_cau_hoi = Column(String(20), nullable=False, index=True)
    muc_do = Column(String(20), nullable=False, index=True)
    noi_dung = Column(Text, nullable=False)
    dap_an_a = Column(Text, nullable=True)
    dap_an_b = Column(Text, nullable=True)
    dap_an_c = Column(Text, nullable=True)
    dap_an_d = Column(Text, nullable=True)
    dap_an_dung = Column(Text, nullable=True)
    giai_thich = Column(Text, nullable=True)


Base.metadata.create_all(bind=engine)

# Bổ sung cột giải thích nếu bảng cũ chưa có
if "cau_hoi" in inspect(engine).get_table_names():
    columns = {
        col["name"]
        for col in inspect(engine).get_columns("cau_hoi")
    }

    if "giai_thich" not in columns:
        with engine.begin() as conn:
            conn.execute(
                text("ALTER TABLE cau_hoi ADD COLUMN giai_thich TEXT")
            )
