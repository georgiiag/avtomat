# Avtomat — ТП‑Конструктор (Flask)

Avtomat — небольшое веб‑приложение для подготовки технологического процесса (ТП) по спецификации изделия (BOM) с акцентом на быстрый “вайб‑прототипинг”: UI живёт в браузере, сервер даёт API для распознавания спецификаций и загрузки файлов.

## Назначение (производство)
- Загрузить спецификацию изделия (DOCX/DOC или XLSX/XLSM) → получить BOM‑дерево.
- На основе BOM собрать/отредактировать маршрут/ТП как набор операций, переходов и прикреплений.
- Экспортировать/сохранить результат (HTML‑снапшот, выгрузка BOM в Excel/DOC).

## Технологии
- Backend: Python + Flask (один модуль [app/app.py](app/app.py))
- Frontend: HTML + встроенный JavaScript (монолит [app/templates/tp_constructor.html](app/templates/tp_constructor.html))
- WSGI: gunicorn ([Procfile](Procfile))
- Парсинг DOCX: python-docx
- Хранение состояния: localStorage (клиент) + файловая система контейнера (загрузки в `app/static/...`)

## Быстрый старт (локально)
```bash
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe app\app.py
```

Открыть:
- http://127.0.0.1:5000/ (редирект на конструктор)
- http://127.0.0.1:5000/tp_constructor
- http://127.0.0.1:5000/recognizer
- http://127.0.0.1:5000/share

## Документация для интеграции в SELMA HUB
Эти документы рассчитаны на разработчика SELMA HUB и ИИ‑агента, который будет планировать интеграцию.
- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) — архитектура, компоненты, хранение данных и ограничения.
- [docs/API.md](docs/API.md) — полный контракт эндпоинтов + примеры запросов.
- [docs/DOMAIN_MODEL.md](docs/DOMAIN_MODEL.md) — сущности и то, как они представлены в текущем коде (особенно в localStorage).
- [docs/INTEGRATION_SELMA_HUB.md](docs/INTEGRATION_SELMA_HUB.md) — план интеграции и маппинг на корпоративный стек (БД, auth, storage, сервисы).
- [docs/OPERATIONS.md](docs/OPERATIONS.md) — запуск, деплой, переменные окружения, нюансы Railway.

