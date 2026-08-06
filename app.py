import os
import shutil
import time
import re
import logging
import mimetypes

from contextlib import asynccontextmanager
from fastapi import FastAPI, UploadFile, File, Form, \
    HTTPException, status, Request, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from openpyxl import load_workbook
from zipfile import ZipFile

mimetypes.add_type('text/css', '.css')
mimetypes.add_type('application/javascript', '.js')

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("server.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


def cyrillic_to_latin(text: str) -> str:
    text = text.strip().lower()
    # Словарь соответствия русских и английских букв (ГОСТ-ориентированный)
    rules = {
        'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'yo',
        'ж': 'zh', 'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm',
        'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u',
        'ф': 'f', 'х': 'kh', 'ц': 'ts', 'ч': 'ch', 'ш': 'sh', 'щ': 'shch',
        'ы': 'y', 'э': 'e', 'ю': 'yu', 'я': 'ya', 'ъ': '', 'ь': ''
    }
    output = []
    for char in text:
        if char in rules:
            output.append(rules[char])
        else:
            output.append(char)
    translated_text = ''.join(output)
    translated_text = translated_text.replace(' ', '_').replace('-', '_')
    translated_text = re.sub(r'[^a-z0-9_]', '', translated_text)
    return translated_text if translated_text else 'campaign'


@asynccontextmanager
async def lifespan(app: FastAPI):
    # старт сервера
    logger.info("Запуск сервера, проверка временных директорий...")

    if os.path.exists(TEMP_DIR):
        try:
            shutil.rmtree(TEMP_DIR)
            os.makedirs(TEMP_DIR, exist_ok=True)
            logger.info(
                "Временная папка web_tmp успешно очищена при старте сервера.")
        except Exception as e:
            logger.error(f"Не удалось очистить web_tmp при старте: {e}")
    else:
        os.makedirs(TEMP_DIR, exist_ok=True)
    yield

    # остановка сервера
    logger.info("Сервер останавливается. Подчищаем остаточные файлы...")
    if os.path.exists(TEMP_DIR):
        try:
            shutil.rmtree(TEMP_DIR)
            logger.info("Папка web_tmp успешно удалена перед выключением.")
        except Exception as e:
            logger.error(f"Ошибка очистки при выключении: {e}")

app = FastAPI(title="UTM Tag Generator for DOCX", lifespan=lifespan)
app.mount("/static", StaticFiles(directory="static"), name="static")

templates = Jinja2Templates(directory="templates")
TEMP_DIR = "web_tmp"
MAX_FILE_SIZE = 10 * 1024 * 1024
VALID_MIME_TYPE = "application/vnd.openxmlformats-officedocument.\
    wordprocessingml.document"


@app.get("/", response_class=HTMLResponse)
async def read_index(request: Request):
    return templates.TemplateResponse(request=request, name="index.html",
                                      context={"request": request})


@app.get('/favicon.ico', include_in_schema=False)
async def favicon():
    return FileResponse(
        os.path.join('static', 'favicon.svg'),
        media_type='image/svg+xml'
    )


def modify_url(source_url: str, campaign: str, source: str, medium: str, content: str, term: str) -> str:
    """Блок раздельной сборки: принимает оригинальный URL и безопасно приклеивает 
    к нему только заполненные UTM-параметры без дублирования знаков '?'.
    """
    source_url = source_url.strip()
    if not source_url:
        return source_url

    params = [
        f'utm_source={source}',
        f'utm_medium={medium}',
        f'utm_campaign={campaign}'
    ]

    if content and content.strip():
        params.append(f'utm_content={content.strip()}')
    if term and term.strip():
        params.append(f'utm_term={term.strip()}')

    utm_string = '&'.join(params)
    if '?' in source_url:
        if source_url.endswith('?') or source_url.endswith('&'):
            return f'{source_url}{utm_string}'
        else:
            return f'{source_url}&{utm_string}'
    else:
        return f'{source_url}?{utm_string}'


def add_utm_to_docx(input_path: str, campaign: str, source: str, medium: str, content: str, term: str, output_path: str):
    """Открывает файл .docx, проходит по всем параграфам, таблицам и связям,
    находит гиперссылки и обновляет их через функцию modify_url.
    """
    doc = Document(input_path)

    # Разметка скрытых связей документа (основной массив гиперссылок в Word)
    rels = doc.part.rels
    for rel_id in rels:
        rel = rels[rel_id]
        if rel.reltype == RT.HYPERLINK:
            # Вызываем нашу новую функцию раздельной сборки
            new_url = modify_url(
                source_url=rel.target_ref,
                campaign=campaign,
                source=source,
                medium=medium,
                content=content,
                term=term
            )
            rel._target = new_url

    # Разметка текстовых полей параграфов и таблиц (для защиты структуры документа)
    for p in doc.paragraphs:
        if 'hyperlink' in p._element.xml:
            for node in p._element.xpath('.//w:hyperlink'):
                r_id = node.get('{http://openxmlformats.org}id')
                if r_id and r_id in rels:
                    node_rel = rels[r_id]
                    new_url = modify_url(
                        source_url=node_rel.target_ref,
                        campaign=campaign,
                        source=source,
                        medium=medium,
                        content=content,
                        term=term
                    )
                    node_rel._target = new_url

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if 'hyperlink' in p._element.xml:
                        for node in p._element.xpath('.//w:hyperlink'):
                            r_id = node.get('{http://openxmlformats.org}id')
                            if r_id and r_id in rels:
                                node_rel = rels[r_id]
                                new_url = modify_url(
                                    source_url=node_rel.target_ref,
                                    campaign=campaign,
                                    source=source,
                                    medium=medium,
                                    content=content,
                                    term=term
                                )
                                node_rel._target = new_url
    doc.save(output_path)


def remove_file(path: str):
    # удаление файлов после окончания работы
    try:
        if os.path.exists(path):
            os.remove(path)
            logger.info(
                f"Успешная генерация архива для кампании")
    except Exception as e:
        logger.error(f"Ошибка обработки файла: {str(e)}")


@app.post('/generate-utm/')
async def generate_utm_api(
    background_tasks: BackgroundTasks,
    campaign_name: str = Form(...),
    platforms: list[str] = Form(None),
    custom_platform: str = Form(None),
    file: UploadFile = File(...)
):
    clean_campaign = cyrillic_to_latin(campaign_name)

    # Валидация площадок
    final_platforms = []
    if platforms:
        final_platforms.extend(platforms)
    if custom_platform and custom_platform.strip():
        if not re.match(r'^[a-zа-я0-9_,\- ]+$', custom_platform.strip().lower()):
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail='Названия площадок содержат недопустимые спецсимволы! Разрешены только буквы, цифры и запятые.'
            )
        custom_list = [p.strip()
                       for p in custom_platform.lower().split(',') if p.strip()]

        for platform_item in custom_list:
            clean_custom = cyrillic_to_latin(platform_item)
            # Добавляем в список, если такой площадки там еще нет (защита от дубликатов)
            if clean_custom not in final_platforms:
                final_platforms.append(clean_custom)

    if not final_platforms:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST,
                            detail='Выберите или введите хотя бы одну площадку.')

    # Проверки безопасности
    if file.size > MAX_FILE_SIZE:
        raise HTTPException(
            status_code=status.HTTP_413_CONTENT_TOO_LARGE, detail='Файл слишком большой. Не более 10 МБ.')
    if file.content_type != VALID_MIME_TYPE:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST, detail='Некорректный тип файла.')

    session_id = f'session_{int(time.time())}'
    session_dir = os.path.join(TEMP_DIR, session_id)
    os.makedirs(session_dir, exist_ok=True)

    uploaded_file_path = os.path.join(session_dir, 'source_article.docx')
    with open(uploaded_file_path, 'wb') as buffer:
        shutil.copyfileobj(file.file, buffer)

    created_files = []
    raw_filename = os.path.splitext(file.filename)[0]
    safe_base_name = ''.join(
        [c for c in raw_filename if c.isalnum() or c in (' ', '_', '-')]).strip()
    if not safe_base_name:
        safe_base_name = 'article'

    try:
        for platform in final_platforms:
            out_filename = f'{safe_base_name}_{platform}.docx'
            out_file_path = os.path.join(session_dir, out_filename)
            add_utm_to_docx(uploaded_file_path, clean_campaign,
                            platform, out_file_path)
            created_files.append(out_file_path)

        zip_path = os.path.join(TEMP_DIR, f'{session_id}_packed.zip')
        with ZipFile(zip_path, 'w') as zipf:
            for f_path in created_files:
                zipf.write(f_path, os.path.basename(f_path))

        shutil.rmtree(session_dir)
        background_tasks.add_task(remove_file, zip_path)

        return FileResponse(
            path=zip_path,
            filename=f'{safe_base_name}_utm_packed.zip',
            media_type='application/zip'
        )
    except Exception as e:
        logger.error(f'Ошибка обработки: {e}')
        if os.path.exists(session_dir):
            shutil.rmtree(session_dir)
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_CONTENT, detail='Ошибка обработки документа.')


# массовая обработка
@app.post('/generate-utm-batch/')
async def generate_utm_batch_api(
    background_tasks: BackgroundTasks,
    excel_file: UploadFile = File(..., description="Файл манифеста .xlsx"),
    zip_file: UploadFile = File(..., description="Архив со статьями .zip")
):
    # Проверка форматов на бэкенде
    if not excel_file.filename.endswith('.xlsx'):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST, detail='Манифест должен быть в формате .xlsx')
    if not zip_file.filename.endswith('.zip'):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST, detail='Архив файлов должен быть в формате .zip')

    session_id = f'batch_{int(time.time())}'
    session_dir = os.path.join(TEMP_DIR, session_id)

    extracted_dir = os.path.join(session_dir, 'extracted')
    output_dir = os.path.join(session_dir, 'output')
    os.makedirs(extracted_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    excel_path = os.path.join(session_dir, 'manifest.xlsx')
    with open(excel_path, 'wb') as buffer:
        shutil.copyfileobj(excel_file.file, buffer)

    zip_path = os.path.join(session_dir, 'input.zip')
    with open(zip_path, 'wb') as buffer:
        shutil.copyfileobj(zip_file.file, buffer)

    with ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extracted_dir)

    try:
        wb = load_workbook(excel_path, read_only=True)
        sheet = wb.active
        created_files = []

        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue

            filename = str(row[0]).strip()
            campaign_raw = str(row[1]).strip()
            sources_raw = str(row[2]).strip()

            clean_campaign = cyrillic_to_latin(campaign_raw)
            platforms = [cyrillic_to_latin(
                p) for p in sources_raw.split(',') if p.strip()]

            input_docx_path = os.path.join(extracted_dir, filename)

            if not os.path.exists(input_docx_path):
                logger.warning(
                    f'Файл {filename} указан в Excel, но отсутствует в ZIP-архиве.')
                continue

            base_name = os.path.splitext(filename)[0]

            for platform in platforms:
                out_filename = f'{base_name}_{platform}.docx'
                out_file_path = os.path.join(output_dir, out_filename)
                add_utm_to_docx(input_docx_path, clean_campaign,
                                platform, out_file_path)
                created_files.append(out_file_path)

        wb.close()

        if not created_files:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST, detail='Не удалось обработать ни один файл.')

        final_zip_path = os.path.join(TEMP_DIR, f'{session_id}_final.zip')
        with ZipFile(final_zip_path, 'w') as zipf:
            for f_path in created_files:
                zipf.write(f_path, os.path.basename(f_path))

        shutil.rmtree(session_dir)
        background_tasks.add_task(remove_file, final_zip_path)

        return FileResponse(
            path=final_zip_path,
            filename='mass_utm_packed.zip',
            media_type='application/zip'
        )

    except Exception as e:
        logger.error(f'Ошибка пакетной обработки: {e}')
        if os.path.exists(session_dir):
            shutil.rmtree(session_dir)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail='Ошибка сервера при пакетной обработке.')
