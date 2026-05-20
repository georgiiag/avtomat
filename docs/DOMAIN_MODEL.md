# Доменная модель (как она представлена сейчас)

Этот документ описывает сущности “как есть” в Avtomat. Это не нормализованная БД‑модель, а фактическая модель данных, которая живёт в браузере и частично сериализуется в localStorage.

## Ключевая идея
- “Источник правды” для ТП‑конструктора — **состояние в браузере**.
- Сервер хранит только файлы (вложения/снапшоты) и выполняет распознавание спецификаций.

## Сущности (термины)

### Изделие (Product)
“Изделие” — контейнер для BOM и данных ТП (операций). В UI это элемент списка слева (“изделия/папки”).

Хранение:
- переменная `products` (объект `{ [productId]: product }`) в JS
- сериализуется в `localStorage.tp_constructor_autosave_v1`

Типовые поля (по использованию в UI):
- `id` (ключ в `products`)
- `folderId`, `folderName` (группировка в UI)
- `isTpProduct` (маркер типа в UI)
- BOM:
  - `bom`: дерево узлов
  - `bomVersions`: список версий BOM (для сценария “новая версия спецификации”)
  - активная версия определяется функциями `getActiveBomVersion/setActiveBomVersion`
- ТП:
  - `opData`: операции (map по `opId`)
  - `nextOpCode`: генерация кода операций (шагом)

Ссылки на код:
- сбор автосейва: [tp_constructor.html:L1880-L1927](../app/templates/tp_constructor.html#L1880-L1927)

### Узел спецификации (BOM Node)
Нода BOM — элемент дерева спецификации. Узлы строятся из распознанных строк (`/api/parse_bom_docx` или `/api/parse_bom_excel`), либо из ручной вставки табличного текста.

Содержимое узла (по факту парсинга и UI):
- `level` (уровень иерархии)
- `code` (обозначение/код)
- `name` (наименование)
- `qty` (количество, может быть number или string)
- `unit` (ед.изм)
- плюс UI‑метаданные/вложения (могут добавляться пользователем в карточке узла)

### Операция (Operation)
Операция — “кубик” на блок‑схеме ТП.

Хранение:
- `product.opData[opId]`

Типовые поля (по UI‑форме редактирования):
- `code` (например, `010`, `015`, …)
- `name`
- `batch`, `timeOp`, `workers`, `specialty`
- `centerId` (привязка к рабочему центру)
- `opFiles` (прикрепления к операции; сохраняются как `{name,type,url}`)
- `transitions` (переходы; список структур с текстом/ресурсами/файлами)

### Переход (Transition)
Переход — строка внутри операции (шаг/инструкция), может иметь ссылки на ресурсы и вложения.

Типовые поля:
- `text`
- `toolingId`, `instrumentId`, `materialId`
- `files`: `[{name,type,url}, ...]`

### Справочники (Registries)
В UI есть справочники:
- Working centers (рабочие центры)
- Tooling (оснастка)
- Instruments (инструмент)
- Materials (материалы)
- Manufacturers (производители)

Хранение:
- в JS массивах (например, `workingCentersRegistry`, `toolingRegistry`, …)
- сериализуются в автосейв вместе с `products`

Seed‑данные:
- backend пытается подгрузить “сид” оснастки из Excel‑файла, если в окружении есть `openpyxl` и файл присутствует рядом с проектом: [app.py:L10-L123](../app/app.py#L10-L123).

## Автосейв (localStorage schema)
Ключ: `tp_constructor_autosave_v1`

Верхнеуровневая структура:
- `v`: версия схемы (сейчас `1`)
- `ts`: timestamp
- `currentProductId`
- `currentUserId`
- `users`: список пользователей (UI‑модель)
- `productOrder`: порядок отображения
- `productFolderCollapsedById`: состояние “свернутости” папок
- `products`: объект изделий (с BOM и opData)
- `bomExpandedByKey`: набор раскрытых узлов BOM (per‑product/per‑version)
- `bomRootExpandedByKey`: раскрыт ли root узла

Формирование: [tp_constructor.html:L1880-L1927](../app/templates/tp_constructor.html#L1880-L1927)

## Что это значит для интеграции в SELMA HUB
- В SELMA HUB нужно нормализовать эти сущности в БД (как минимум: Product, BomVersion, BomItem, Operation, Transition, Attachment, RegistryEntry).
- Нужен механизм версионирования и совместной работы вместо localStorage.
- Нужна стратегия миграции существующих данных из localStorage (если важно сохранять результаты пользователей).

