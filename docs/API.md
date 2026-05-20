# API Avtomat (Flask)

База: `<host>` (например, `http://127.0.0.1:5000` или Railway‑домен).

Формат ошибок (общее правило):  
`{"ok": false, "error": "<message>"}` с кодом `400/404/500` (если не указано иначе).

## Health / Info

### GET /api/status
Возвращает “онлайн” статус (заглушка).

Ответ 200:
```json
{"status":"online","message":"Система готова к работе","process":"Ожидание команды"}
```

## UI (HTML)

### GET /
Редирект на `/tp_constructor`.

### GET /tp_constructor
Главный экран конструктора ТП (HTML).

### GET /recognizer
Экран отладки распознавания спецификации (HTML).

## BOM recognition (parse)

### POST /api/parse_bom_docx
Распознаёт спецификацию из `.docx` или `.doc`.

Вход:
- `multipart/form-data`
- поле `file`: `.docx` или `.doc`

Пример:
```bash
curl -sS -X POST "<host>/api/parse_bom_docx" \
  -F "file=@spec.docx" | jq .
```

Успех 200:
```json
{
  "ok": true,
  "rows": [
    {"level":0,"code":"...","name":"...","qty":1,"unit":"шт"},
    {"level":1,"code":"...","name":"...","qty":2,"unit":"шт"}
  ],
  "columns": {"level":0,"code":2,"name":3,"qty":4,"unit":5},
  "product_name": "…",
  "product_designation": "…",
  "tables_used": 1
}
```

Примечания:
- Для `.doc` выполняется попытка конвертации в `.docx` (LibreOffice/soffice, либо MS Word, если доступно) — см. [app.py](../app/app.py).
- Если пользователь загрузит временный файл Word `~$...`, вернётся ошибка 400.

### POST /api/parse_bom_excel
Распознаёт спецификацию из `.xlsx` или `.xlsm` (через `openpyxl`, если библиотека установлена в окружении).

Вход:
- `multipart/form-data`
- поле `file`: `.xlsx` или `.xlsm`

Пример:
```bash
curl -sS -X POST "<host>/api/parse_bom_excel" \
  -F "file=@spec.xlsx" | jq .
```

Успех 200 (пример структуры):
```json
{
  "ok": true,
  "rows": [{"level":0,"code":"...","name":"...","qty":1,"unit":"шт"}],
  "columns": {"level":0,"code":1,"name":2,"qty":3,"unit":4},
  "product_name": "…",
  "product_designation": "…",
  "sheet": "Спецификация",
  "sheets": ["Спецификация"]
}
```

## BOM export

### POST /api/export_bom_excel
Экспортирует BOM в `.xlsx` и отдаёт файлом.

Вход: `application/json`
```json
{
  "columns": [{"key":"code","label":"Обозначение"},{"key":"name","label":"Наименование"}],
  "rows": [["ABC-1","Деталь",1,"шт"]],
  "filename": "spec.xlsx"
}
```

Выход: файл (`Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`).

### POST /api/export_bom_doc
Экспортирует BOM в `.doc` (фактически RTF‑контент) и отдаёт файлом.

Вход: `application/json`
```json
{
  "columns": [{"key":"code","label":"Обозначение"},{"key":"name","label":"Наименование"}],
  "rows": [["ABC-1","Деталь",1,"шт"]],
  "filename": "spec.doc"
}
```

Выход: файл (`Content-Type: application/msword`).

## File storage / attachments

### POST /api/upload_attachment
Загрузка вложений (картинки/PDF/Word и т.п.) в `app/static/attachments/<YYYYMMDD>/...`.

Вход:
- `multipart/form-data`
- поле `file`: бинарный файл

Пример:
```bash
curl -sS -X POST "<host>/api/upload_attachment" \
  -F "file=@photo.png" | jq .
```

Успех 200:
```json
{"ok": true, "url": "/static/attachments/20260520/att-...__photo.png", "name": "photo.png", "ext": "png", "size": 12345}
```

### POST /api/save_tp_html
Сохраняет HTML‑снапшот ТП в `app/static/shared_docs/snapshots/`.

Вход:
- `multipart/form-data`
- поле `file`: HTML

Успех 200:
```json
{"ok": true, "url": "/static/shared_docs/snapshots/tp-...__report.html", "id": "tp-...", "name": "report.html"}
```

## Public share (DOCX viewer)

### GET|POST /share
Мини‑страница “общего доступа”: загрузка `.docx` и список последних файлов.

POST:
- `multipart/form-data`, поле `file` (`.docx`)
- успешный ответ: `302 Redirect` на `/share/<doc_id>`

### GET /share/<doc_id>
Отображает загруженный DOCX как упрощённый HTML (абзацы + таблицы) и даёт ссылку на скачивание из `/static/shared_docs/...`.

