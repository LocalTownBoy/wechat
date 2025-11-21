import os
from pathlib import Path
from typing import List
from urllib.parse import quote_plus

from fastapi import Depends, FastAPI, Form, Request
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy import Column, Integer, String, create_engine, text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import Session, sessionmaker

BASE_DIR = Path(__file__).resolve().parent

DB_HOST = os.getenv("DB_HOST", "10.11.108.216")
DB_PORT = os.getenv("DB_PORT", "3306")
DB_USER = os.getenv("DB_USER", "root")
DB_PASSWORD = quote_plus(os.getenv("DB_PASSWORD", "123123Wwb"))
DB_NAME = os.getenv("DB_NAME", "papers")

DATABASE_URL = (
    f"mysql+pymysql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}?charset=utf8mb4"
)

# 确保数据库存在
_bootstrap_engine = create_engine(
    f"mysql+pymysql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/",
    pool_pre_ping=True,
)
with _bootstrap_engine.connect() as conn:
    conn.execute(text(f"CREATE DATABASE IF NOT EXISTS `{DB_NAME}` DEFAULT CHARACTER SET utf8mb4"))

# 连接具体数据库
engine = create_engine(DATABASE_URL, pool_pre_ping=True)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()


class Paper(Base):
    __tablename__ = "papers"

    id = Column(Integer, primary_key=True, index=True)
    title = Column(String(255), nullable=False)
    author = Column(String(255), nullable=False)
    section = Column(String(255), nullable=False)


Base.metadata.create_all(bind=engine)

app = FastAPI(title="Paper Collector")
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


@app.get("/papers", response_class=HTMLResponse)
def show_form(request: Request, db: Session = Depends(get_db)):
    papers: List[Paper] = db.query(Paper).order_by(Paper.id.desc()).all()
    return templates.TemplateResponse(
        "paper_form.html",
        {
            "request": request,
            "papers": papers,
            "message": None,
        },
    )


@app.post("/papers", response_class=HTMLResponse)
def submit_paper(
    request: Request,
    title: str = Form(...),
    author: str = Form(...),
    section: str = Form(...),
    db: Session = Depends(get_db),
):
    paper = Paper(title=title.strip(), author=author.strip(), section=section.strip())
    db.add(paper)
    db.commit()
    db.refresh(paper)
    papers: List[Paper] = db.query(Paper).order_by(Paper.id.desc()).all()
    return templates.TemplateResponse(
        "paper_form.html",
        {
            "request": request,
            "papers": papers,
            "message": "已保存！",
        },
    )
