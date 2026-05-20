# Ops / Runbook

## Запуск локально

### Требования
- Python 3.12+ (желательно; фактически будет работать на 3.10+)

### Команды (Windows)
```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe app\app.py
```

По умолчанию:
- слушает `0.0.0.0`
- порт `PORT` (если задан), иначе `5000`

## Запуск в прод (gunicorn)
Команда соответствует [Procfile](../Procfile):
```bash
gunicorn --chdir app app:app --bind 0.0.0.0:$PORT
```

## Переменные окружения
- `PORT` — порт веб‑сервера (Railway задаёт автоматически).
- `HOST` — хост для dev‑режима (по умолчанию `0.0.0.0`).

## Railway
- Источник деплоя: ветка `main` (автодеплой после пуша).
- Важно: файловая система контейнера ephemeral.

### Что ломается на ephemeral FS
Функции, которые пишут данные на диск:
- `/api/upload_attachment` → `app/static/attachments/<YYYYMMDD>/...`
- `/api/save_tp_html` → `app/static/shared_docs/snapshots/...`
- `/share` → `app/static/shared_docs/...`

На Railway это означает:
- файлы могут пропадать после рестарта/редеплоя
- при нескольких инстансах файлы будут разъезжаться по инстансам

Рекомендация для корпоративной интеграции:
- заменить хранение на object storage (S3/MinIO) и хранить метаданные в БД.

## Диагностика

### Быстрая проверка
- `GET /api/status` должен вернуть `200` и JSON `{"status":"online", ...}`
- `GET /tp_constructor` должен отдавать HTML

### Типовые проблемы
- `ModuleNotFoundError: flask` → зависимости не установлены в активном интерпретаторе; ставить через `python -m pip install -r requirements.txt` ровно тем Python, которым запускаете.
- “Кажется старая версия UI” → очистить localStorage/кеш браузера (в UI основное состояние в localStorage).

