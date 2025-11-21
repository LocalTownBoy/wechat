import datetime
import json
import logging
import os
import re

from django.conf import settings
from django.http import JsonResponse
from django.shortcuts import render
from openpyxl import Workbook, load_workbook
import pdfplumber
from wxcloudrun.models import Counters, Paper


logger = logging.getLogger('log')
EXCEL_FILE_NAME = 'push_messages.xlsx'


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
    收集论文信息（标题、作者、章节）并存入数据库，返回列表页面
    """
    message = None
    error = None
    parsed = {}

    if request.method in ['POST', 'post']:
        title = request.POST.get('title', '').strip()
        author = request.POST.get('author', '').strip()
        section = request.POST.get('section', '').strip()
        pdf_file = request.FILES.get('pdf')

        if pdf_file:
            try:
                parsed = _parse_pdf(pdf_file)
                title = title or parsed.get('title', '')
                author = author or parsed.get('author', '')
                section = section or parsed.get('sections_text', '')
            except Exception as exc:  # pylint: disable=broad-except
                logger.exception('failed to parse pdf: %s', exc)
                error = 'PDF 解析失败，请检查文件格式'

        if not title or not author or not section:
            error = '请完整填写标题、作者和章节'
        else:
            try:
                Paper.objects.create(title=title, author=author, section=section)
                message = '已保存！'
            except Exception as exc:  # pylint: disable=broad-except
                logger.exception('failed to save paper: %s', exc)
                error = '保存失败，请检查日志'

    try:
        papers = Paper.objects.order_by('-id').all()
    except Exception as exc:  # pylint: disable=broad-except
        logger.exception('failed to load papers: %s', exc)
        papers = []
        if not error:
            error = '读取数据失败，请检查日志'

    return render(request, 'paper_form.html', {
        'papers': papers,
        'message': message,
        'error': error,
        'parsed': parsed,
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


def _parse_pdf(pdf_file):
    """
    解析上传的 PDF，返回标题、作者、章节列表的最佳猜测
    """
    try:
        pdf_file.seek(0)
    except Exception:  # pylint: disable=broad-except
        pass

    filename = getattr(pdf_file, 'name', '') or ''
    base_title = os.path.splitext(os.path.basename(filename))[0] if filename else ''

    with pdfplumber.open(pdf_file) as pdf:
        # 尝试读取元数据
        info = pdf.metadata or {}
        title = ''
        author = ''

        raw_title = info.get('Title')
        raw_author = info.get('Author')
        if raw_title:
            title = str(raw_title).strip()
        if raw_author:
            author = str(raw_author).strip()

        # 抽取前几页文本用于解析章节/备用标题
        texts = []
        for page in pdf.pages[:5]:
            page_text = page.extract_text() or ''
            texts.append(page_text)
        full_text = '\n'.join(texts)

    lines = [ln.strip() for ln in full_text.splitlines() if ln.strip()]

    if not title:
        if lines:
            title = lines[0][:255]
        elif base_title:
            title = base_title[:255]

    if not author:
        for ln in lines[:5]:
            m = re.search(r'(作者|Author)[:：]\s*(.+)', ln)
            if m:
                author = m.group(2).strip()[:255]
                break

    sections = []
    section_pattern = re.compile(r'^(\d+(\.\d+)*)\s+(.+)|^(摘要|Abstract|引言|绪论|结论|参考文献)', re.IGNORECASE)
    for ln in lines:
        if len(sections) >= 15:
            break
        if section_pattern.match(ln):
            sections.append(ln[:255])

    dedup_sections = []
    for sec in sections:
        if sec not in dedup_sections:
            dedup_sections.append(sec)

    sections_text = '; '.join(dedup_sections[:10])

    return {
        'title': title,
        'author': author,
        'sections': dedup_sections,
        'sections_text': sections_text,
    }
