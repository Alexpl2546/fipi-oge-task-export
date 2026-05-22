# FIPI OGE Task Exporter

Локальный выгрузчик заданий из открытого банка тестовых заданий ФИПИ ОГЭ для личного образовательного использования.

Проект не обходит CAPTCHA, блокировки, авторизацию или ограничения доступа. Если сайт возвращает страницу защиты, HTTP 403/429/5xx или другие явные признаки ограничения доступа, скрипт останавливается и пишет понятную ошибку.

## Установка

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Быстрый запуск

Перед полной выгрузкой обязательно сделайте малую проверку на 10-20 заданий:

```powershell
python -m src.fipi_export --url "https://oge.fipi.ru/bank/index.php?proj=DE0E276E497AB3784C3FC4CC20248DC0" --out data_test --delay 2 --limit 20
python -m src.validate_export --data data_test/tasks.jsonl
```

Полная выгрузка после проверки:

```powershell
python -m src.fipi_export --url "https://oge.fipi.ru/bank/index.php?proj=DE0E276E497AB3784C3FC4CC20248DC0" --out data --delay 2 --resume
```

## CLI

- `--url` — ссылка на банк с параметром `proj`.
- `--out` — директория выгрузки, по умолчанию `data`.
- `--delay` — задержка между HTTP-запросами, по умолчанию 2 секунды.
- `--limit 10` — выгрузить только первые 10 новых заданий.
- `--resume` — не дублировать задания, уже записанные в `tasks.jsonl`.
- `--format jsonl,csv,sqlite` — итоговые форматы. JSONL всегда пишется для resume.
- `--headless true/false` — совместимый параметр; Playwright сейчас не используется.
- `--debug` — сохранить диагностические HTML-страницы в `data/debug`.
- `--diagnose-tls` — проверить TLS-соединение в режимах default, certifi и insecure.
- `--tls-mode certifi|system|insecure` — выбрать режим проверки TLS. По умолчанию используется `certifi`.
- `--insecure` — устаревший алиас для `--tls-mode insecure`. Не рекомендуется для обычной работы.

## TLS

Все рабочие HTTPS-запросы идут через единый HTTP-клиент.

- `--tls-mode certifi` — режим по умолчанию. Если установлен `requests`, используется `requests.Session()` и для каждого запроса явно передаётся `verify=certifi.where()`. Если `requests` недоступен, используется `urllib` с `ssl.create_default_context(cafile=certifi.where())`.
- `--tls-mode system` — использует системное хранилище сертификатов Windows через пакет `truststore`. Скрипт вызывает `truststore.inject_into_ssl()` и выполняет запросы через `urllib` с системным SSL context.
- `--tls-mode insecure` — отключает TLS-проверку только явно: для `requests` передаётся `verify=False`, для `urllib` используется unverified SSL context.

Диагностика TLS:

```powershell
python -m src.fipi_export --diagnose-tls --url "https://oge.fipi.ru/bank/index.php?proj=DE0E276E497AB3784C3FC4CC20248DC0"
```

Диагностика проверяет `default`, `certifi`, `system truststore` и `insecure`. Если `certifi` падает, а `system` проходит, запускайте выгрузку с `--tls-mode system`:

```powershell
python -m src.fipi_export --url "https://oge.fipi.ru/bank/index.php?proj=DE0E276E497AB3784C3FC4CC20248DC0" --out data --delay 2 --resume --tls-mode system
```

Если TLS-проверка падает, проверьте доверенные сертификаты ОС/Python, установку `certifi` и `truststore`, переменные `SSL_CERT_FILE`, `REQUESTS_CA_BUNDLE`, `CURL_CA_BUNDLE` и наличие корпоративного TLS inspection сертификата. В текущем окружении `--tls-mode insecure` может потребоваться только как временный диагностический обход, чтобы подтвердить, что проблема локальная и связана именно с цепочкой сертификатов; для обычной выгрузки он не рекомендуется.

## Результаты

- `data/tasks.jsonl` — основной формат, одно задание на строку.
- `data/tasks.csv` — табличная версия в `UTF-8-SIG`, чтобы Excel корректнее открывал кириллицу.
- `data/tasks.sqlite` — SQLite-база с таблицей `tasks`.
- `data/images/` — локальные копии изображений и вложений, найденных в заданиях.
- `data/markdown/` — отдельный Markdown-файл для каждого задания.

`data/` исключена из Git, кроме служебного `.gitkeep`; пример малой выгрузки находится в `examples/sample_tasks.jsonl`.

## Validation

Проверка выгрузки:

```powershell
python -m src.validate_export --data data/tasks.jsonl
```

Валидатор проверяет:

- уникальность `qid`;
- уникальность URL;
- непустой текст задания;
- корректность URL ФИПИ;
- существование локальных изображений из `local_images`;
- совпадение количества записей в JSONL, CSV и SQLite.

Если последняя строка JSONL повреждена после сбоя, валидатор покажет номер строки. Сам выгрузчик при `--resume` игнорирует повреждённые JSONL-строки, логирует предупреждение и продолжает без дублирования корректно прочитанных `qid` и URL.

## Excel Export

Основной источник для Excel-экспорта — `data_full/tasks.jsonl`. Права на исходные задания, изображения и сопутствующие материалы сохраняются за ФИПИ/правообладателями; не публикуйте выгруженные данные в Git.

Preview на 20 заданий со встроенными изображениями:

```powershell
py -m src.export_excel --data data_full/tasks.jsonl --out data_full/fipi_oge_math_preview.xlsx --mode embedded --limit 20
```

Полная версия со встроенными изображениями:

```powershell
py -m src.export_excel --data data_full/tasks.jsonl --out data_full/fipi_oge_math_full_embedded.xlsx --mode embedded
```

Лёгкая версия со ссылками на локальные изображения:

```powershell
py -m src.export_excel --data data_full/tasks.jsonl --out data_full/fipi_oge_math_index.xlsx --mode links
```

Embedded-файл может быть тяжёлым и открываться медленнее, потому что изображения физически встраиваются в workbook. Links-версия компактнее: она сохраняет карточки заданий и кликабельные локальные ссылки на изображения.

В Excel создаются листы `Общая информация`, `Перечень заданий`, `Задания` и, только при ошибках обработки, `Ошибки`. На листе `Перечень заданий` есть содержательные колонки `Тип задания` и `Тема`: они берутся из исходных данных, если такие поля заполнены, иначе определяются воспроизводимой локальной эвристикой по тексту задания. Лист `Задания` очищен от технических полей и предназначен для чтения заданий как учебного сборника: заголовок карточки ведёт на источник ФИПИ, ниже идут текст, варианты ответа и связанные изображения или ссылки на них.

У листа `Задания` есть два режима отображения:

- `--task-view structured` — редактируемый текст задания, варианты ответа и изображения отдельными блоками.
- `--task-view rendered` — HTML задания рендерится в PNG через Playwright и визуально ближе к сайту ФИПИ: inline-формулы, таблицы и картинки остаются на своих местах. В режиме `--mode embedded` PNG-карточки вставляются в Excel; rendered-файл может быть тяжёлым. В режиме `--mode links` создаются ссылки на локальные rendered PNG.

Для rendered-режима после установки зависимостей может потребоваться установить Chromium для Playwright:

```powershell
py -m playwright install chromium
```

Preview rendered:

```powershell
py -m src.export_excel --data data_full/tasks.jsonl --out data_full/fipi_oge_math_preview_rendered.xlsx --mode embedded --task-view rendered --limit 20
```

Полная rendered-версия:

```powershell
py -m src.export_excel --data data_full/tasks.jsonl --out data_full/fipi_oge_math_full_rendered.xlsx --mode embedded --task-view rendered
```

## Site Analysis

Стартовая страница проекта:

```text
https://oge.fipi.ru/bank/index.php?proj=DE0E276E497AB3784C3FC4CC20248DC0
```

Она возвращает HTML в кодировке `windows-1251` и содержит iframe/контейнер для загрузки списка заданий:

```text
questions.php?proj=DE0E276E497AB3784C3FC4CC20248DC0&init_filter_themes=1
questions.php?proj=DE0E276E497AB3784C3FC4CC20248DC0&page=N
```

Задания уже присутствуют в HTML ответа `questions.php`; JavaScript в основном управляет размером iframe, статусами решения, избранным и вставкой картинок через функции вида `ShowPictureQ(...)`. Поэтому выбран обычный HTTP и `lxml`; Playwright не нужен для базовой выгрузки.

Список заданий формируется постраничным чтением `questions.php?proj=...&page=N`. Важная деталь: пагинация начинается с `page=0`, а не с `page=1`. На одной странице обычно 10 блоков `div.qblock`; общий счётчик и номер страницы возвращаются внизу HTML вызовом `setQCount(3884, currentPage, 10)`. Для проекта `DE0E276E497AB3784C3FC4CC20248DC0` при проверке сайт сообщил 3884 задания, значит корректный обход идёт от `page=0` до `page=388` включительно.

Публичный идентификатор берётся из `id="q45B944"` и дублируется как `Номер: 45B944`. Внутренний GUID хранится в скрытом поле `input name="guid"`. Метаданные находятся в соседнем блоке `div` с id вида `i45B944`: там извлекаются `КЭС` и `Тип ответа`.

Текст задания берётся из `qblock`, при этом `script`/`style` удаляются только из чистого текста. Оригинальный HTML-фрагмент сохраняется в поле `html`, чтобы не терять MathML, таблицы и изображения. Правильные ответы и пояснения на просмотренных страницах не выдаются как открытый статический HTML, поэтому поля `answer` и `explanation` сохраняются как `null`.

Изображения встречаются как обычные `<img src="...">` и как строковые аргументы JavaScript-функций `ShowPicture...`. Скрипт извлекает оба варианта, скачивает файлы в один поток и не перекачивает уже существующие файлы.

Обнаруженные ограничения: сайт может иметь проблемы с цепочкой TLS-сертификатов в некоторых локальных окружениях; при 403, повторном 429 и 5xx выгрузчик останавливается. Агрессивный параллельный сбор, обход CAPTCHA, прокси-ротация и имитация обхода блокировок не реализуются.

## Legal And Technical Limits

Материалы берутся из открытого банка ФИПИ и предназначены для локального личного образовательного использования. Права на задания, изображения и сопутствующие материалы сохраняются за правообладателями. Не публикуйте полную выгруженную базу как отдельный открытый датасет и сохраняйте указание на источник ФИПИ.
