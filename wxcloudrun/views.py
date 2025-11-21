import datetime
import json
import logging
import os
from urllib.parse import quote_plus

from django.conf import settings
from django.http import JsonResponse
from django.shortcuts import render
from openpyxl import Workbook, load_workbook
from sqlalchemy import Column, Integer, String, create_engine, text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import Session, sessionmaker
from wxcloudrun.models import Counters


logger = logging.getLogger('log')
EXCEL_FILE_NAME = 'push_messages.xlsx'

# SQLAlchemy 配置，用于论文收集
DB_HOST = os.getenv("DB_HOST", "10.11.108.216")
DB_PORT = os.getenv("DB_PORT", "3306")
DB_USER = os.getenv("DB_USER", "root")
DB_PASSWORD = quote_plus(os.getenv("DB_PASSWORD", "123123Wwb.,"))
DB_NAME = os.getenv("DB_NAME", "papers")

DATABASE_URL = (
    f"mysql+pymysql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}?charset=utf8mb4"
)

_bootstrap_engine = create_engine(
    f"mysql+pymysql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/",
    pool_pre_ping=True,
)
with _bootstrap_engine.connect() as conn:
    conn.execute(text(f"CREATE DATABASE IF NOT EXISTS `{DB_NAME}` DEFAULT CHARACTER SET utf8mb4"))

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


def index(request, _):
    """
    获取主页

     `` request `` 请求对象
    """

    return render(request, 'index.html')


def counter(request, _):
    """
    获取当前计数

     `` request `` 请求对象
    """

    rsp = JsonResponse({'code': 0, 'errorMsg': ''}, json_dumps_params={'ensure_ascii': False})
    if request.method == 'GET' or request.method == 'get':
        rsp = get_count()
    elif request.method == 'POST' or request.method == 'post':
        rsp = update_count(request)
    else:
        rsp = JsonResponse({'code': -1, 'errorMsg': '请求方式错误'},
                            json_dumps_params={'ensure_ascii': False})
    logger.info('response result: {}'.format(rsp.content.decode('utf-8')))
    return rsp


def push_msg(request, _):
    """
    接收微信公众号推送的JSON消息并写入Excel
    """

    if request.method not in ['POST', 'post']:
        return JsonResponse({'code': -1, 'errorMsg': '请求方式错误'},
                            json_dumps_params={'ensure_ascii': False}, status=405)

    if not request.body:
        return JsonResponse({'code': -1, 'errorMsg': '请求体为空'},
                            json_dumps_params={'ensure_ascii': False})

    try:
        body_unicode = request.body.decode('utf-8')
        payload = json.loads(body_unicode)
    except json.JSONDecodeError:
        return JsonResponse({'code': -1, 'errorMsg': '请求体不是合法JSON'},
                            json_dumps_params={'ensure_ascii': False})

    try:
        file_path = _append_to_excel(payload)
    except Exception as exc:  # pylint: disable=broad-except
        logger.exception('failed to write push message to excel: %s', exc)
        return JsonResponse({'code': -1, 'errorMsg': '写入Excel失败'},
                            json_dumps_params={'ensure_ascii': False})

    return JsonResponse({'code': 0, 'data': {'file': file_path}},
                        json_dumps_params={'ensure_ascii': False})


def push_msg_list(request, _):
    """
    展示Excel中已存储的推送消息
    """
    excel_path = getattr(settings, 'PUSH_MSG_EXCEL_PATH',
                         os.path.join(settings.BASE_DIR, EXCEL_FILE_NAME))
    messages = []
    error = None

    if not os.path.exists(excel_path):
        error = 'Excel文件不存在，请先调用 /push/msg 写入数据'
    else:
        try:
            workbook = load_workbook(excel_path)
            sheet = workbook.active
            for received_at, payload in sheet.iter_rows(min_row=2, max_col=2, values_only=True):
                if received_at is None and payload is None:
                    continue
                payload_text = payload
                if isinstance(payload, str):
                    try:
                        payload_text = json.dumps(json.loads(payload), ensure_ascii=False, indent=2)
                    except json.JSONDecodeError:
                        payload_text = payload
                messages.append({'received_at': received_at, 'payload': payload_text})
        except Exception as exc:  # pylint: disable=broad-except
            logger.exception('failed to load excel: %s', exc)
            error = '读取Excel失败，请检查日志'

    return render(request, 'push_messages.html', {
        'messages': messages,
        'excel_path': excel_path,
        'error': error,
    })


def papers(request, _):
    """
    收集论文信息（标题、作者、章节）并存入 MySQL，返回列表页面
    """
    message = None
    error = None
    session: Session = SessionLocal()

    if request.method in ['POST', 'post']:
        title = request.POST.get('title', '').strip()
        author = request.POST.get('author', '').strip()
        section = request.POST.get('section', '').strip()

        if not title or not author or not section:
            error = '请完整填写标题、作者和章节'
        else:
            try:
                paper = Paper(title=title, author=author, section=section)
                session.add(paper)
                session.commit()
                message = '已保存！'
            except Exception as exc:  # pylint: disable=broad-except
                session.rollback()
                logger.exception('failed to save paper: %s', exc)
                error = '保存失败，请检查日志'

    try:
        papers = session.query(Paper).order_by(Paper.id.desc()).all()
    except Exception as exc:  # pylint: disable=broad-except
        logger.exception('failed to load papers: %s', exc)
        papers = []
        if not error:
            error = '读取数据失败，请检查日志'
    finally:
        session.close()

    return render(request, 'paper_form.html', {
        'papers': papers,
        'message': message,
        'error': error,
    })


def get_count():
    """
    获取当前计数
    """

    try:
        data = Counters.objects.get(id=1)
    except Counters.DoesNotExist:
        return JsonResponse({'code': 0, 'data': 0},
                    json_dumps_params={'ensure_ascii': False})
    return JsonResponse({'code': 0, 'data': data.count},
                        json_dumps_params={'ensure_ascii': False})


def update_count(request):
    """
    更新计数，自增或者清零

    `` request `` 请求对象
    """

    logger.info('update_count req: {}'.format(request.body))

    body_unicode = request.body.decode('utf-8')
    body = json.loads(body_unicode)

    if 'action' not in body:
        return JsonResponse({'code': -1, 'errorMsg': '缺少action参数'},
                            json_dumps_params={'ensure_ascii': False})

    if body['action'] == 'inc':
        try:
            data = Counters.objects.get(id=1)
        except Counters.DoesNotExist:
            data = Counters()
        data.id = 1
        data.count += 1
        data.save()
        return JsonResponse({'code': 0, "data": data.count},
                    json_dumps_params={'ensure_ascii': False})
    elif body['action'] == 'clear':
        try:
            data = Counters.objects.get(id=1)
            data.delete()
        except Counters.DoesNotExist:
            logger.info('record not exist')
        return JsonResponse({'code': 0, 'data': 0},
                    json_dumps_params={'ensure_ascii': False})
    else:
        return JsonResponse({'code': -1, 'errorMsg': 'action参数错误'},
                    json_dumps_params={'ensure_ascii': False})


def _append_to_excel(payload):
    """
    将推送消息写入Excel，按时间顺序追加
    """
    excel_path = getattr(settings, 'PUSH_MSG_EXCEL_PATH',
                         os.path.join(settings.BASE_DIR, EXCEL_FILE_NAME))
    directory = os.path.dirname(excel_path)

    if directory and not os.path.exists(directory):
        os.makedirs(directory, exist_ok=True)

    workbook = _load_or_create_workbook(excel_path)
    sheet = workbook.active
    sheet.append([
        datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        json.dumps(payload, ensure_ascii=False)
    ])
    workbook.save(excel_path)
    return excel_path


def _load_or_create_workbook(excel_path):
    """
    获取已存在的工作簿，或创建新的并初始化表头
    """
    if os.path.exists(excel_path):
        return load_workbook(excel_path)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = 'messages'
    sheet.append(['received_at', 'payload'])
    return workbook
