# Code Wiki проекта (Flask) — «Цифровой Цех / ТП‑Конструктор»

## 1. Назначение и общий обзор

Проект представляет собой небольшое веб‑приложение на Flask, которое:

- Отдаёт несколько HTML‑страниц (Jinja2-шаблоны, фактически статические).
- Предоставляет API для распознавания спецификации/BOM из DOCX (`python-docx`).
- Позволяет загружать DOCX для простого просмотра в браузере.
- Позволяет сохранять HTML‑снимок (snapshot) сформированного ТП на сервер и получать ссылку на него.

Основной фокус кода — серверная функция распознавания таблицы спецификации из DOCX и фронтенд‑логика «ТП‑конструктора», которая использует этот API.

## 2. Структура репозитория

```
/workspace
  app/
    app.py                 # Flask backend (роуты + логика распознавания DOCX)
    templates/
      index.html           # главная страница «ЕПК / цифровой цех»
      tp_constructor.html  # «ТП‑конструктор» (основная логика на JS)
      recognizer.html      # тестовая страница распознавания DOCX
  requirements.txt         # зависимости Python
  Procfile                 # запуск через gunicorn (типично для Heroku)
```

Код: [app.py](file:///workspace/app/app.py), шаблоны: [templates](file:///workspace/app/templates).

## 3. Архитектура выполнения (runtime architecture)

### 3.1 Компоненты

- **Browser UI**
  - [index.html](file:///workspace/app/templates/index.html) — «витрина» процесса/навигация.
  - [tp_constructor.html](file:///workspace/app/templates/tp_constructor.html) — UI «конструктора ТП», включает:
    - импорт/обновление спецификации из DOCX через `POST /api/parse_bom_docx`;
    - ведение изделий/операций/переходов на клиенте;
    - экспорт HTML и сохранение на сервер через `POST /api/save_tp_html`.
  - [recognizer.html](file:///workspace/app/templates/recognizer.html) — минимальный UI для отладки распознавания.

- **Flask backend**
  - Единый модуль: [app.py](file:///workspace/app/app.py).
  - Роуты HTML‑страниц + API + «песочница» для загрузки/просмотра DOCX.

- **Файловое хранилище (локально на сервере)**
  - `app/static/shared_docs/` — загруженные DOCX (для `/share`).
  - `app/static/shared_docs/snapshots/` — сохранённые HTML‑снимки ТП (для `/api/save_tp_html`).

### 3.2 Потоки данных

#### A) Распознавание спецификации (DOCX → BOM JSON)

1. Пользователь выбирает `.docx` в UI (страница [recognizer.html](file:///workspace/app/templates/recognizer.html) или [tp_constructor.html](file:///workspace/app/templates/tp_constructor.html)).
2. Браузер отправляет `multipart/form-data` на `POST /api/parse_bom_docx`.
3. Backend читает DOCX (`python-docx`), ищет «лучшую» таблицу, определяет колонки по заголовкам и возвращает JSON:
   - `rows`: строки спецификации
   - `columns`: карта найденных индексов колонок
   - `product_name` / `product_designation`: метаданные изделия

#### B) Публикация/просмотр DOCX (загрузка → ссылка просмотра)

1. Пользователь загружает `.docx` через `POST /share`.
2. Сервер сохраняет файл в `app/static/shared_docs/` и делает редирект на `/share/<doc_id>`.
3. Страница `/share/<doc_id>` делает упрощённый HTML‑рендер документа (абзацы + таблицы).

#### C) Сохранение HTML‑снимка ТП (UI → файл на сервере → ссылка)

1. В UI генерируется HTML snapshot (на клиенте) и отправляется на `POST /api/save_tp_html`.
2. Сервер сохраняет файл в `app/static/shared_docs/snapshots/`.
3. Возвращает JSON со ссылкой `url`, которую можно открыть/скачать.

## 4. Основные модули и их ответственность

### 4.1 Backend: [app.py](file:///workspace/app/app.py)

Файл объединяет:

- Инициализацию Flask‑приложения.
- Роутинг HTML‑страниц.
- API для распознавания DOCX и сохранения HTML.
- Мини‑страницы для «общего доступа» и просмотра загруженных DOCX.

### 4.2 Frontend: [templates](file:///workspace/app/templates)

- [index.html](file:///workspace/app/templates/index.html) — UI‑страница «ЕПК / цифровой цех». Логика в основном UI/визуализация и навигация.
- [tp_constructor.html](file:///workspace/app/templates/tp_constructor.html) — главный функциональный экран. Практически вся доменная логика (изделия, BOM‑дерево, операции, экспорт) реализована на клиенте.
- [recognizer.html](file:///workspace/app/templates/recognizer.html) — отладочный экран распознавания, полезен для теста API.

## 5. Роуты и API (контракт)

### 5.1 HTML‑страницы

- `GET /` и `GET /index.html` → [index](file:///workspace/app/app.py#L15-L22)
- `GET /tp_constructor` и `GET /tp_constructor.html` → [tp_constructor](file:///workspace/app/app.py#L23-L27)
- `GET /recognizer` и `GET /recognizer.html` → [recognizer](file:///workspace/app/app.py#L28-L31)

### 5.2 «Общий доступ» к DOCX

- `GET /share` — список загруженных DOCX + форма загрузки → [share_doc](file:///workspace/app/app.py#L81-L155)
- `POST /share` — загрузка `.docx`, сохранение в `app/static/shared_docs/`, редирект на просмотр → [share_doc](file:///workspace/app/app.py#L81-L155)
- `GET /share/<doc_id>` — просмотр загруженного файла в виде простого HTML → [view_shared_doc](file:///workspace/app/app.py#L157-L204)

### 5.3 API

- `POST /api/save_tp_html` — сохранить HTML snapshot на сервер
  - Вход: `multipart/form-data` с полем `file` (HTML).
  - Выход (успех): `{"ok": true, "url": "...", "id": "...", "name": "..."}`.
  - Реализация: [save_tp_html](file:///workspace/app/app.py#L206-L233)

- `POST /api/parse_bom_docx` — распознать спецификацию/BOM из DOCX
  - Вход: `multipart/form-data` с полем `file` (DOCX).
  - Выход (успех): `{"ok": true, "rows": [...], "columns": {...}, "product_name": "...", "product_designation": "..."}`.
  - Реализация: [parse_bom_docx](file:///workspace/app/app.py#L234-L560)

- `GET /api/status` — простой health/status endpoint
  - Выход: `{"status":"online","message":"...","process":"..."}`.
  - Реализация: [get_status](file:///workspace/app/app.py#L562-L569)

## 6. Ключевые функции (backend)

### 6.1 CORS для всех ответов

- [add_cors_headers](file:///workspace/app/app.py#L8-L13)
  - Добавляет `Access-Control-Allow-Origin: *`, а также методы/заголовки.
  - Важно: в текущем виде открывает доступ всем источникам (подходит для демо/прототипа, но требует ужесточения для продакшена).

### 6.2 Директории хранения

- [UPLOAD_DIR / SNAPSHOT_DIR](file:///workspace/app/app.py#L33-L35)
- [ensure_upload_dir](file:///workspace/app/app.py#L36-L41)
- [ensure_snapshot_dir](file:///workspace/app/app.py#L42-L47)

Поведение: директории создаются «по месту» при обращении к функционалу.

### 6.3 Превью DOCX в HTML

- [docx_to_simple_html](file:///workspace/app/app.py#L48-L79)
  - Парсит DOCX и конвертирует:
    - непустые абзацы → `<p>…</p>`
    - таблицы → `<table>…</table>`
  - Это не полноценный рендер Word‑документа, а «быстрый просмотр».

### 6.4 Загрузка и просмотр DOCX по ссылке

- [share_doc](file:///workspace/app/app.py#L81-L155)
  - `POST`: валидирует расширение `.docx`, генерирует `doc_id`, сохраняет файл, делает редирект на просмотр.
  - `GET`: строит простую HTML‑страницу со списком последних загрузок.

- [view_shared_doc](file:///workspace/app/app.py#L157-L204)
  - Ищет файл по префиксу `<doc_id>__`, вызывает `docx_to_simple_html`, отдаёт страницу просмотра.

### 6.5 Сохранение HTML‑снимка ТП

- [save_tp_html](file:///workspace/app/app.py#L206-L233)
  - Принимает файл HTML, сохраняет в `SNAPSHOT_DIR`, возвращает URL в `static/…`.

### 6.6 Распознавание спецификации/BOM из DOCX

Основной endpoint: [parse_bom_docx](file:///workspace/app/app.py#L234-L560)

Ключевые идеи алгоритма:

1. **Нормализация заголовков** (внутренняя `norm`): приводит строку к нижнему регистру, убирает лишние пробелы/неразрывные пробелы.
2. **Поиск колонок по синонимам** (`synonyms`, `find_col_index`):
   - `code`: код/обозначение/позиция/артикул…
   - `name`: наименование/номенклатура…
   - `qty`: количество…
   - `unit`: единица измерения…
   - `level`: уровень (если присутствует)
3. **Выбор “лучшей” таблицы**:
   - перебираются таблицы в документе;
   - пробуются первые 1–3 строки как “строка заголовков”;
   - строится score по наличию ожидаемых колонок и числу распарсенных строк.
4. **Парсинг строк**:
   - `qty` пытается приводиться к `float` (с поддержкой `,` как разделителя).
   - `level` берётся из колонки, только если там целое число.
5. **Восстановление уровней иерархии, если level отсутствует**:
   - сначала из отступов абзацев внутри ячеек (`cell_indent_pt`, `infer_levels_from_indents`);
   - иначе — из структуры кода (количество `.` в обозначении) (`infer_levels_from_designation`).
6. **Извлечение метаданных изделия**:
   - `product_name` — пытается найти строку со словом “спецификац…”, иначе берёт первую значимую строку (`extract_product_name`);
   - `product_designation` — ищет строку “обознач…/шифр…”, иначе по regex‑паттерну кода (`extract_product_designation`);
   - если не найдено — использует первую строку BOM как fallback.

## 7. Ключевая логика (frontend, tp_constructor.html)

Файл [tp_constructor.html](file:///workspace/app/templates/tp_constructor.html) — большой монолитный HTML+CSS+JS. Ниже — ориентиры по крупным подсистемам, чтобы быстрее разбираться в коде.

### 7.1 Настройка “сервера распознавания”

Функции (локальное хранилище base URL и проверка `/api/status`):

- [getRecognizerBaseUrl / setRecognizerBaseUrl / buildRecognizerUrl](file:///workspace/app/templates/tp_constructor.html#L1541-L1565)
- [testRecognizerServer](file:///workspace/app/templates/tp_constructor.html#L1574-L1589)

Идея: по умолчанию используется текущий origin (`/api/...`), но можно задать внешний base URL (например, если распознавание вынесено на отдельный сервер).

### 7.2 Импорт/обновление спецификации из DOCX

Опорные функции:

- [parseDocxToBomTree](file:///workspace/app/templates/tp_constructor.html#L1956-L1997) — отправляет DOCX на `POST /api/parse_bom_docx`, затем преобразует строки в BOM‑дерево.
- [handleCreateTpDocx](file:///workspace/app/templates/tp_constructor.html#L1896-L1955) — создание ТП по загруженному DOCX.
- [handleUpdateSpecDocx](file:///workspace/app/templates/tp_constructor.html#L1838-L1895) — обновление спецификации для выбранного изделия.

### 7.3 Построение и работа с BOM

Парсинг табличного текста и построение дерева:

- [detectBomDelimiter](file:///workspace/app/templates/tp_constructor.html#L2431-L2440)
- [buildBomHeaderMap / parseBomRow](file:///workspace/app/templates/tp_constructor.html#L2441-L2525)
- [buildBomTree / buildBomFromRows](file:///workspace/app/templates/tp_constructor.html#L2526-L2599)
- [renderBomPreview](file:///workspace/app/templates/tp_constructor.html#L2600-L2612)

Сопоставление “старой” и “новой” BOM при обновлении (сохранение метаданных):

- [flattenBomForMatch / buildBomMatchIndex](file:///workspace/app/templates/tp_constructor.html#L1628-L1709)
- [scoreBomMatch / matchAndMergeBomMeta](file:///workspace/app/templates/tp_constructor.html#L1712-L1837)

### 7.4 Управление изделиями и метаданными

- [selectProduct / deleteProduct](file:///workspace/app/templates/tp_constructor.html#L1998-L2027)
- [updateProductMeta / ensureProductMeta](file:///workspace/app/templates/tp_constructor.html#L2191-L2238)
- [ensureBomVersions / getActiveBomVersion / setActiveBomVersion](file:///workspace/app/templates/tp_constructor.html#L2316-L2403)

### 7.5 Экспорт HTML и сохранение snapshot на сервер

- [showTpHtmlExportOverlay / closeTpHtmlExportOverlay](file:///workspace/app/templates/tp_constructor.html#L2841-L2864)
- [buildTpConstructorSnapshotHtml](file:///workspace/app/templates/tp_constructor.html#L2965-L3138)
- [saveTpHtmlSnapshotToServer](file:///workspace/app/templates/tp_constructor.html#L2926-L2963) — вызывает backend `POST /api/save_tp_html`
- [exportCurrentTpHtml](file:///workspace/app/templates/tp_constructor.html#L3143-L3211)

## 8. Зависимости и связи (dependency relationships)

### 8.1 Python зависимости

См. [requirements.txt](file:///workspace/requirements.txt):

- `flask` — веб‑фреймворк и шаблонизатор (Jinja2).
- `python-docx` — чтение DOCX (используется в `docx_to_simple_html` и `parse_bom_docx`).
- `gunicorn` — WSGI‑сервер для запуска в продакшене/на платформе.

### 8.2 Frontend зависимости

- Font Awesome подключается с CDN: `https://cdnjs.cloudflare.com/.../font-awesome/...` (см. [index.html](file:///workspace/app/templates/index.html), [tp_constructor.html](file:///workspace/app/templates/tp_constructor.html)).
- Остальная логика — нативный JS (без сборщиков/пакетных менеджеров).

### 8.3 Взаимозависимости компонентов

- `tp_constructor.html` ↔ `POST /api/parse_bom_docx` и `POST /api/save_tp_html`
- `recognizer.html` ↔ `POST /api/parse_bom_docx`
- `/share` и `/share/<doc_id>` ↔ локальная директория `app/static/shared_docs/`

## 9. Инструкции по запуску

### 9.1 Запуск локально (development)

1. Установить Python 3.10+ (рекомендуется).
2. Создать виртуальное окружение и установить зависимости:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

3. Запустить приложение:

```bash
python app/app.py
```

4. Открыть в браузере:

- `http://localhost:5000/` — главная
- `http://localhost:5000/tp_constructor` — ТП‑конструктор
- `http://localhost:5000/recognizer` — тест распознавания
- `http://localhost:5000/share` — загрузка/просмотр DOCX

Порт можно задать переменной окружения `PORT` (см. [app.py](file:///workspace/app/app.py#L571-L586)).

### 9.2 Запуск через gunicorn (production / Procfile)

См. [Procfile](file:///workspace/Procfile):

```bash
gunicorn --chdir app app:app --bind 0.0.0.0:${PORT:-5000}
```

## 10. Практические заметки (для сопровождения)

- В проекте нет аутентификации/авторизации; загрузка файлов и CORS открыты для всех.
- Загруженные DOCX и HTML‑snapshots пишутся на диск; при деплое стоит продумать:
  - очистку старых файлов,
  - лимиты размера,
  - изоляцию/права доступа,
  - перенос в объектное хранилище (S3‑совместимое) при необходимости.

