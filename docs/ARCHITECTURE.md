# Архитектура Avtomat (ТП‑Конструктор)

## TL;DR
- Приложение состоит из одного Flask‑сервера и HTML/JS UI.
- Доменная логика конструктора в основном на клиенте (внутри `tp_constructor.html`).
- Сервер предоставляет API для распознавания спецификаций (BOM) и для загрузки/выгрузки файлов.
- У приложения нет БД; состояние пользователя хранится в `localStorage`, серверные файлы пишутся в `app/static/...` (эпhemeral на Railway).

## Компоненты

### Backend (Flask)
- Единственный модуль: [app/app.py](../app/app.py)
- Ответственность:
  - отдача страниц (`/tp_constructor`, `/recognizer`, `/share`)
  - API распознавания BOM из DOC/DOCX и Excel
  - API выгрузки BOM в Excel/DOC
  - API загрузки вложений/снапшотов в файловую систему контейнера

Запуск:
- dev: `python app/app.py` (учитывает `PORT`)
- prod: `gunicorn --chdir app app:app --bind 0.0.0.0:$PORT` ([Procfile](../Procfile))

### Frontend (HTML/JS)
- Основной UI: [app/templates/tp_constructor.html](../app/templates/tp_constructor.html)
  - монолитный файл: layout + CSS + большой inline‑JS
  - хранит состояние “в браузере” (localStorage) и работает без серверной БД
- Отладка распознавания: [app/templates/recognizer.html](../app/templates/recognizer.html)
- Лэндинг/вход: [app/templates/index.html](../app/templates/index.html)
- Общие стили/тема: [app/static/ui.css](../app/static/ui.css), [app/static/ui.js](../app/static/ui.js)

## Хранение данных (as-is)

### Клиент: localStorage
Ключевые значения:
- `tp_constructor_autosave_v1` — автосейв состояния конструктора (включая изделия, BOM‑версии, операции и справочники).
- `recognizerBaseUrl` — базовый URL сервера распознавания (можно вынести распознавание на отдельный сервис).
- `theme` — тема UI (light/dark).

Формат автосейва (верхний уровень) формируется в `__getAutosavePayload()`:
- [tp_constructor.html:L1880-L1927](../app/templates/tp_constructor.html#L1880-L1927)

Важно:
- localStorage — это “пер‑браузер/пер‑устройство” хранилище. Нет общей командной консистентности, нет версионирования, возможны “призраки старой версии” из-за старых ключей.

### Сервер: файловая система контейнера
Сервер сохраняет runtime‑файлы прямо в `app/static/`:
- shared DOCX: `app/static/shared_docs/` (эндпоинты `/share`, `/share/<id>`)
- HTML‑снапшоты: `app/static/shared_docs/snapshots/` (`POST /api/save_tp_html`)
- вложения: `app/static/attachments/<YYYYMMDD>/...` (`POST /api/upload_attachment`)

См. константы путей в [app.py:L155-L178](../app/app.py#L155-L178).

Ограничение для Railway:
- файловая система обычно ephemeral → при рестарте/редеплое файлы могут исчезнуть
- при горизонтальном масштабировании файлы будут “локальными” для инстанса

## Безопасность (as-is)
- CORS включён для всех ответов: `Access-Control-Allow-Origin: *` ([app.py:L125-L130](../app/app.py#L125-L130)).
- Аутентификации/авторизации нет.
- Загрузка файлов доступна без контроля доступа.
- Есть фильтрация опасных расширений в `/api/upload_attachment` ([app.py:L409-L416](../app/app.py#L409-L416)), но этого недостаточно для корпоративной среды.

## Зависимости
- Python: `flask`, `python-docx`, `gunicorn` ([requirements.txt](../requirements.txt))
- Frontend: Font Awesome подключается по CDN (в шаблонах)

## Точки интеграции для SELMA HUB (кратко)
- Выделить устойчивое серверное хранилище: Postgres/ORM + объектное хранилище (S3‑совместимое) вместо `app/static/...`.
- Ввести authN/authZ (SSO, роли) и ограничить CORS.
- Разделить API и UI: либо оставить UI как “встроенный модуль”, либо вынести в отдельный фронтенд SELMA и использовать только API.

