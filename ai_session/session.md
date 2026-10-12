# Запись сессии с AI-агентом (Claude Code)

Автоматический экспорт сессии агента по заданию: запросы, ответы агента, вызовы инструментов и их результаты (длинные результаты сокращены, полная версия — в session.jsonl).

## 👤 Запрос (19:28:13)

Привет. Тестовое задание: автоматизировать чистку файла заявок. Работаем в папке /Users/[скрыто]/zayavki-test (используй абсолютные пути, все файлы клади только туда). Там лежат zayavki.xlsx (входные данные) и ZADANIE.txt (текст задания) — прочитай задание целиком.

Коротко, что нужно: CLI-скрипт на Python. На вход — xlsx с колонками Имя, Телефон, Дата заявки, Источник. На выход — чистый файл и отдельный список проблемных строк.
1. Телефон → +7XXXXXXXXXX (10 цифр после +7). Не телефон → «нет телефона».
2. Дата → ГГГГ-ММ-ДД. Нераспознанную дату пометить как проблемную, строку не выбрасывать.
3. Имя → с заглавной буквы, без лишних пробелов.
4. Дубли по телефону убрать, оставить одну запись (продумай и опиши, какую именно оставляем и почему).
5. Проблемные строки (нет телефона, битая дата, нет имени) — в отдельный список с причиной.
Скрипт будут запускать на ДРУГОМ файле той же структуры, поэтому никакого хардкода под конкретные строки — правила должны быть общими.

Как работаем, по шагам:
1. Сначала посмотри на данные: выведи все 48 строк и составь список всех встречающихся форматов телефонов, дат и имён, найди пропуски и дубли. Покажи мне эту сводку перед тем, как писать код.
2. Предложи правила нормализации для каждого поля (включая спорные случаи: 8 vs +7 vs 7, номера без кода, лишние/недостающие цифры, добавочные, даты с двузначным годом, ДД.ММ vs ММ/ДД, даты текстом, Excel-даты числом, будущие даты и т.п.). Спорные решения зафиксируй явно.
3. Напиши скрипт clean_zayavki.py: аргументы — входной файл и папка/файлы вывода (argparse, --help). Выход: clean.xlsx (и clean.csv) + problems.xlsx (и problems.csv) с колонкой «Причина». Внимание: системный python3 из Homebrew (3.14) может быть сломан (pyexpat) — если упрёшься, сделай venv на другой версии (например /opt/homebrew/bin/python3.12 -m venv .venv) и requirements.txt.
4. Проверка: напиши тесты (pytest) на функции нормализации, отдельно сгенерируй ДРУГОЙ тестовый файл той же структуры со своими «грязными» случаями и прогони скрипт на нём. Сверь вручную итог на исходном файле: сколько строк было, сколько дублей убрано, сколько проблемных — числа должны сходиться. Если находишь ошибку в своём коде — исправляй и прямо отмечай, что было не так.
5. Напиши README.md: как установить и запустить (2–3 команды), какие правила применяются, как разрешаются спорные случаи.
6. В конце дай отчёт: что сделано, итоговые числа (вход/чистые/дубли/проблемные), какие ошибки нашёл и исправил по ходу, какие случаи остались спорными.

Не трогай ничего за пределами /Users/[скрыто]/zayavki-test. Никаких сетевых сервисов кроме pip install нужных пакетов.

## 👤 Запрос (19:28:13)

<system-reminder>
Your final report is delivered through SubagentHandback: when your work is complete, call SubagentHandback({message: <your full report>}). The call ends your run, so make it your last step. Only a SubagentHandback call reaches your caller as your result; plain text you write at the end is not delivered.
</system-reminder>

**🔧 Bash** List task folder and read assignment text

```
ls -la /Users/[скрыто]/zayavki-test; cat /Users/[скрыто]/zayavki-test/ZADANIE.txt
```

<details><summary>результат</summary>

```
total 24
drwxr-xr-x@  4 [скрыто]  staff   128 Oct 11 22:27 .
drwxr-x---+ 89 [скрыто]  staff  2848 Oct 11 22:26 ..
-rw-r--r--@  1 [скрыто]  staff  2487 Oct 11 22:27 ZADANIE.txt
-rw-r--r--@  1 [скрыто]  staff  7181 Oct 11 22:26 zayavki.xlsx
Тестовое задание UNICA — AI-автоматизация (текст из zadanie.pdf)

Файл zayavki.xlsx — заявки за сентябрь 2026 года, 48 строк, колонки: Имя, Телефон, Дата заявки, Источник.

Написать скрипт (запуск из командной строки): на вход — файл с заявками, на выход — чистый файл и список проблемных строк.
1. Телефон → единый формат +7XXXXXXXXXX (10 цифр после +7). Если это не телефон — считать «нет телефона».
2. Дата заявки → единый формат ГГГГ-ММ-ДД. Нераспознаваемую дату пометить как проблемную, строку не выбрасывать.
3. Имя → с заглавной буквы, без лишних пробелов.
4. Дубли — убрать повторы по телефону (одинаковый номер = один человек), оставить одну запись.
5. Проблемные строки (нет телефона, битая дата, нет имени) — вынести в отдельный список, не терять.

Скрипт запустят на другом файле такой же структуры — он должен работать на любых похожих данных.
Не веб-форма и не бот. Чистить руками в Excel не считается.

Прислать одним архивом:
1. Скрипт и короткую инструкцию запуска.
2. Результат: чистый файл и список проблемных строк.
3. Запись работы с AI — все сессии по заданию, как есть (можно заменить на «[скрыто]» пароли, ключи, почту, имя пользователя в путях).
4. 5–7 строк своими словами: как разбили задачу на шаги; что поручали агенту; где агент ошибся или что пришлось поправить; как проверили, что результат верный.

Срок — до 12 октября включительно. Ссылку на папку (например, Яндекс Диск) прислать в чат HH.
```
</details>

**🔧 Bash** Check available Python versions and packages

```
ls /opt/homebrew/bin/python3* 2>/dev/null; which python3; python3 --version; python3 -c "import openpyxl, pandas; print('ok')" 2>&1 | tail -2; /opt/homebrew/bin/python3.12 --version 2>&1
```

<details><summary>результат</summary>

```
/opt/homebrew/bin/python3
/opt/homebrew/bin/python3-config
/opt/homebrew/bin/python3.12
/opt/homebrew/bin/python3.12-config
/opt/homebrew/bin/python3.13
/opt/homebrew/bin/python3.13-config
/opt/homebrew/bin/python3.14
/opt/homebrew/bin/python3.14-config
/opt/homebrew/bin/python3.9
/opt/homebrew/bin/python3.9-config
/opt/homebrew/bin/python3
Python 3.14.5
ok
Python 3.12.11
```
</details>

**🔧 Bash** Create Python 3.12 venv and install openpyxl and pytest

```
cd /Users/[скрыто]/zayavki-test && /opt/homebrew/bin/python3.12 -m venv .venv && .venv/bin/pip install -q openpyxl pytest 2>&1 | tail -3; .venv/bin/python -c "import openpyxl, pytest, xml.parsers.expat; print(openpyxl.__version__, pytest.__version__)"
```

<details><summary>результат</summary>

```
[notice] A new release of pip is available: 25.1.1 -> 26.2.1
[notice] To update, run: /Users/[скрыто]/zayavki-test/.venv/bin/python3.12 -m pip install --upgrade pip
3.1.5 9.1.1
```
</details>

### 🤖 Агент (19:28:42)

Now let me dump all rows with raw values and cell types.

**🔧 Bash** Dump all rows of the input workbook with cell types

```
cd /Users/[скрыто]/zayavki-test && .venv/bin/python -c "
import openpyxl
wb = openpyxl.load_workbook('zayavki.xlsx')
print(wb.sheetnames)
for ws in wb:
    print(ws.title, ws.max_row, ws.max_column, ws.dimensions)
    for i,row in enumerate(ws.iter_rows(), start=1):
        print(i, [ (repr(c.value), type(c.value).__name__, c.number_format) if c.data_type!='s' else repr(c.value) for c in row])
"
```

<details><summary>результат</summary>

```
['Заявки']
Заявки 49 4 A1:D49
1 ["'Имя'", "'Телефон'", "'Дата заявки'", "'Источник'"]
2 ["'сергей волков'", "'8 (916) 245-17-83'", "'09/03/2026'", "'сайт'"]
3 ["'Волков Сергей'", "'+7 916 245 17 83'", "'03.09.2026'", "'email'"]
4 ["'ПАВЛОВА ирина'", ('89032218840', 'int', 'General'), "'05.09.2026'", "'звонок'"]
5 ["'Ирина Павлова'", "'+7 (903) 221-88-40'", "'09/05/2026'", "'сайт'"]
6 ["'Кузнецов  Андрей'", "'7 925 600 31 12'", "'6 сентября 2026'", "'звонок'"]
7 ["'Андрей Кузнецов'", ('79256003112', 'int', 'General'), ('datetime.datetime(2026, 9, 6, 0, 0)', 'datetime', 'DD.MM.YYYY'), "'email'"]
8 ["'Морозова Елена'", "'+7 985 112-90-47'", "'07/09/2026'", "'telegram'"]
9 ["'Белов Никита'", "'+7 985 112-90-47'", "'2026-09-08'", "'email'"]
10 ["'Соколов Дмитрий'", ('None', 'NoneType', '@'), "'08.09.2026'", "'звонок'"]
11 ["'влад'", "'не знаю'", "'09/09/2026'", "'telegram'"]
12 ["'Новикова Татьяна'", "'тел.: 8 917 400 52 19 (после 18:00)'", "'10.09.2026'", "'звонок'"]
13 ["'Татьяна Новикова'", "'+79174005219'", "'09/10/2026'", "'сайт'"]
14 ["'Егорова Анна'", "'8 (926) 377-41-05'", "'хх.09.2026'", "'email'"]
15 ["'Фёдоров Илья'", "'+7 909 515 66 30'", ('None', 'NoneType', '@'), "'telegram'"]
16 ["'Фёдоров Илья'", "'89095156630'", "'12 сентября'", "'звонок'"]
17 ["'Захарова Ольга'", "'+7 977 241 08 66'", "'13.09.26'", "'email'"]
18 [('None', 'NoneType', '@'), "'+7 999 820-14-73'", "'14.09.2026'", "'звонок'"]
19 ["'Лебедев Максим'", "'8 (800) 707-46-19'", "'15.09.2026'", "'звонок'"]
20 ["'Орлова Светлана'", "'88007074619'", "'15.09.2026'", "'звонок'"]
21 ["'Тихонов Артём'", "'8-800-707-46-19'", "'16.09.2026'", "'звонок'"]
22 ["'Гусева Марина'", "'8 800 707 46 19'", "'17.09.2026'", "'звонок'"]
23 ["'Медведев Роман'", "'+7 800 707-46-19'", "'18.09.2026'", "'звонок'"]
24 ["'Петров Алексей'", "'+7 915 640 22 81'", "'16.09.2026'", "'email'"]
25 ["'Пeтров Алексей'", "'8 915 640-22-81'", "'09/16/2026'", "'сайт'"]
26 ["'Королёва Дарья'", "'+7 926 118 75 02'", "'17/09/2026'", "'telegram'"]
27 ["'Королева Дарья'", ('89261187502', 'int', 'General'), "'17.09.2026'", "'звонок'"]
28 ["'Смирнов\\xa0Олег'", "'+7 903 774 19 26'", "'18.09.2026'", "'email'"]
29 ["'Смирнов Олег\\u200b'", "'8 903 774 19 26'", "'09/18/2026'", "'сайт'"]
30 ["'Карпова Юлия'", "'+7 977 503 82 14'", "'03.09.2026'", "'email'"]
31 ["'Карпова Юлия'", "'+7 977 503-82-14'", "'21.09.2026'", "'звонок'"]
32 ["'Николаев Павел'", "'+7 916 832 05 57'", "'19.09.2062'", "'email'"]
33 ["'Степанова Вера'", "'8 925 019 64
… [сокращено, всего 3765 символов]
```
</details>

**🔧 Bash** Scan cells for hidden characters, mixed alphabets, spaces

```
cd /Users/[скрыто]/zayavki-test && .venv/bin/python -c "
import openpyxl, unicodedata, re
wb = openpyxl.load_workbook('zayavki.xlsx')
ws = wb.active
for i,row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    for col, v in zip('ИТДС', row):
        if isinstance(v, str):
            odd = [(ch, hex(ord(ch)), unicodedata.name(ch,'?')) for ch in v if not (ch.isascii() and ch.isprintable()) and not ('Ѐ'<=ch<='ӿ')]
            latin_in_cyr = col=='И' and re.search('[А-Яа-яЁё]', v) and re.search('[A-Za-z]', v)
            dbl = '  ' in v or v!=v.strip()
            if odd or latin_in_cyr or dbl:
                print(i, col, repr(v), odd, 'LATIN-MIX' if latin_in_cyr else '', 'SPACES' if dbl else '')
print('merged:', ws.merged_cells.ranges, 'hidden rows:', [r for r,d in ws.row_dimensions.items() if d.hidden])
print('sources:', sorted({r[3] for r in ws.iter_rows(min_row=2, values_only=True)}, key=str))
"
```

<details><summary>результат</summary>

```
6 И 'Кузнецов  Андрей' []  SPACES
25 И 'Пeтров Алексей' [] LATIN-MIX 
28 И 'Смирнов\xa0Олег' [('\xa0', '0xa0', 'NO-BREAK SPACE')]  
29 И 'Смирнов Олег\u200b' [('\u200b', '0x200b', 'ZERO WIDTH SPACE')]  
merged: set() hidden rows: []
sources: ['email', 'telegram', 'звонок', 'сайт']
```
</details>

### 🤖 Агент (19:34:53)

Data dump done. Before writing code, let me check a couple of facts I'll rely on (Excel serial epoch, `\s` matching NBSP) quickly.

**🔧 Bash** Verify regex whitespace and Excel serial date behavior

```
cd /Users/[скрыто]/zayavki-test && .venv/bin/python -c "
import re, datetime
print(bool(re.match(r'\s', '\xa0')), bool(re.match(r'\s', '​')))
from openpyxl.utils.datetime import from_excel
print(from_excel(46268), datetime.date(1899,12,30)+datetime.timedelta(46268))
print('17.09'.split('.'))
"
```

<details><summary>результат</summary>

```
True False
2026-09-03 00:00:00 2026-09-03
['17', '09']
```
</details>

### 🤖 Агент (19:34:57)

Now the main script.

**🔧 Write** /Users/[скрыто]/zayavki-test/clean_zayavki.py

```
#!/usr/bin/env python3
"""Чистка файла заявок (Имя, Телефон, Дата заявки, Источник).

Вход:  xlsx с колонками Имя / Телефон / Дата заявки / Источник
       (колонки ищутся по заголовку, порядок не важен).
Выход: clean.xlsx + clean.csv            — чистые уникальные записи;
       problems.xlsx + problems.csv      — проблемные строки с колонкой «Причина»;
       duplicates.xlsx + duplicates.csv  — удалённые дубли (что и с чем склеено).

Каждая входная строка попадает ровно в один из трёх файлов:
    входных строк = чистых + проблемных + удалённых дублей.

Правила нормализации и спорные решения описаны в README.md.
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from statistics import median_low

try:
    import openpyxl
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    sys.exit("Не установлен openpyxl: pip install -r requirements.txt")

NO_PHONE = "нет телефона"

# ---------------------------------------------------------------------------
# Общие утилиты
# ---------------------------------------------------------------------------

# Невидимые символы, которые \s не ловит: zero-width space/joiners, BOM, мягкий перенос.
_INVISIBLE = dict.fromkeys(map(ord, "​‌‍⁠﻿­"), None)


def clean_text(value) -> str:
    """Строка без невидимых символов, с одиночными обычными пробелами."""
    if value is None:
        return ""
    s = unicodedata.normalize("NFC", str(value)).translate(_INVISIBLE)
    return re.sub(r"\s+", " ", s).strip()  # \s в Python ловит и NBSP, и табы


def show(value) -> str:
    """Исходное значение ячейки в читаемом виде (для отчётов)."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M") if (value.hour or value.minute) else value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


# ---------------------------------------------------------------------------
# Имя
# ---------------------------------------------------------------------------

# Латинские буквы, неотличимые на глаз от кириллических («Пeтров» с латинской e).
_HOMOGLYPHS = str.maketrans("aeopcxykAEOPCXYKBMHT", "аеорсхукАЕОРСХУКВМНТ")
_CYR = re.compile(r"[а-яё]", re.I)
_LAT = re.compile(r"[a-z]", re.I)
_LAT_HOMOGLYPH_ONLY = re.compile(r"^[^a-z]*(?:[aeopcxykbmht][^a-z]*)+$", re.I)

NAME_PLACEHOLDERS = {
    "нет", "не указано", "не указан", "неизвестно", "не знаю", "без имени", "аноним",
    "имя", "клиент", "n/a", "na", "none", "null", "unknown", "noname",
}


def _fix_homoglyphs(word: str) -> str:
    # Меняем латиницу на кириллицу только в «смешанных» словах: есть кириллица,
    # а латинские буквы — только двойники кириллических. «John» не трогаем.
    if _CYR.search(word) and _LAT.search(word) and _LAT_HOMOGLYPH_ONLY.match(word):
        return word.translate(_HOMOGLYPHS)
    return word


def _cap(part: str) -> str:
    return part[:1].upper() + part[1:].lower()


def clean_name(raw) -> str | None:
    """Имя с заглавной буквы (каждое слово и часть через дефис), без лишних пробелов.
    None — если имени нет: пусто, нет ни одной буквы или заглушка вроде «нет»."""
    s = clean_text(raw).strip(" ,;:_-")
    if not s or not re.search(r"[^\W\d_]", s):
        return None
    if s.lower().strip(".") in NAME_PLACEHOLDERS:
        return None
    words = [_fix_homoglyphs(w) for w in s.split(" ")]
    return " ".join("-".join(_cap(p) for p in w.split("-")) for w in words)


def _name_tokens(name: str | None) -> list[str]:
    if not name:
        return []
    return [t for t in re.split(r"[\s.\-]+", name.lower().replace("ё", "е")) if t]


def names_
… [сокращено, всего 27520 символов]
```

<details><summary>результат</summary>

```
File created successfully at: /Users/[скрыто]/zayavki-test/clean_zayavki.py (file state is current in your context — no need to Read it back)
```
</details>

**🔧 Bash** Apply fixes, show help, run on the input file

```
cd /Users/[скрыто]/zayavki-test && .venv/bin/python - <<'EOF'
import re
p='clean_zayavki.py'
s=open(p,encoding='utf-8').read()
reps=[
('_RE_YEAR_WORD = re.compile(r"\\s*(?:года|год|г\\.?)$")',
 '_RE_YEAR_WORD = re.compile(r"(?<=\\d)\\s*(?:года|год|г)\\.?$")  # «2026 г.», но не «авг»'),
('_LAT_HOMOGLYPH_ONLY = re.compile(r"^[^a-z]*(?:[aeopcxykbmht][^a-z]*)+$", re.I)',
 '_LAT_HOMOGLYPH_ONLY = re.compile(r"^[^a-zA-Z]*(?:[aeopcxykAEOPCXYKBMHT][^a-zA-Z]*)+$")'),
('''    if had_plus and not digits.startswith("7"):
        return None, f"иностранный номер (+{digits[:3]}…)"
''','''    if had_plus and not digits.startswith("7"):
        return None, f"иностранный номер (+{digits[:3]}…)"
    if had_plus and n != 11:
        return None, f"не хватает цифр ({n})" if n < 11 else f"лишние цифры ({n})"
'''),
]
for a,b in reps:
    assert s.count(a)==1, a
    s=s.replace(a,b)
open(p,'w',encoding='utf-8').write(s)
EOF
.venv/bin/python clean_zayavki.py --help && .venv/bin/python clean_zayavki.py zayavki.xlsx -o output --today 2026-10-11
```

<details><summary>результат</summary>

```
usage: clean_zayavki.py [-h] [-o OUT_DIR] [--clean CLEAN]
                        [--problems PROBLEMS] [--duplicates DUPLICATES]
                        [--sheet SHEET] [--today TODAY] [--csv-sep CSV_SEP]
                        input

Чистка файла заявок: телефоны -> +7XXXXXXXXXX, даты -> ГГГГ-ММ-ДД, имена с
заглавной, дубли по телефону убираются, проблемные строки — в отдельный файл с
причиной.

positional arguments:
  input                 входной .xlsx (колонки Имя, Телефон, Дата заявки,
                        Источник)

options:
  -h, --help            show this help message and exit
  -o OUT_DIR, --out-dir OUT_DIR
                        папка для результатов (по умолчанию: output рядом со
                        входным файлом)
  --clean CLEAN         путь к чистому файлу (по умолчанию <out-
                        dir>/clean.xlsx)
  --problems PROBLEMS   путь к файлу проблемных (по умолчанию <out-
                        dir>/problems.xlsx)
  --duplicates DUPLICATES
                        путь к отчёту об удалённых дублях (по умолчанию <out-
                        dir>/duplicates.xlsx)
  --sheet SHEET         имя листа (по умолчанию первый)
  --today TODAY         «сегодня» для проверки дат в будущем, ГГГГ-ММ-ДД (по
                        умолчанию текущая дата)
  --csv-sep CSV_SEP     разделитель CSV (по умолчанию «;» — для русского
                        Excel)

Пример: python clean_zayavki.py zayavki.xlsx -o output
Вход: zayavki.xlsx (лист «Заявки»), сегодня = 2026-10-11
Строк с данными на входе: 48
  чистых уникальных записей: 11
  удалено дублей по телефону: 14
  проблемных строк:           23 — битая дата: 21, нет телефона: 8, нет имени: 1
Сверка: 11 + 14 + 23 = 48 == вход, OK
Файлы:
  output/clean.xlsx  (+ clean.csv)
  output/problems.xlsx  (+ problems.csv)
  output/duplicates.xlsx  (+ duplicates.csv)
```
</details>

### 🤖 Агент (19:37:01)

Mismatch: I expected 22 / 14 / 12 from the hand count, but got 11 / 14 / 23 with 21 "битая дата". Let me look at the reasons.

**🔧 Bash** Show problem rows with their reasons

```
cd /Users/[скрыто]/zayavki-test && cut -d';' -f1,4,6,9 output/problems.csv | head -30
```

<details><summary>результат</summary>

```
﻿Строка;Дата заявки;Причина;Дата (исходная)
10;2026-09-08;нет телефона (пусто);08.09.2026
11;2026-09-09;нет телефона (не номер: «не знаю»);09/09/2026
14;;битая дата (не распознана: «хх.09.2026»);хх.09.2026
15;;битая дата (пусто);
17;;битая дата (несуществующая дата: «13.09.26»);13.09.26
18;;"битая дата (несуществующая дата: «14.09.2026»);+7 999 820-14-73
19;;"нет телефона (номер 8-80X — сервисная линия организации, не телефон клиента: «8 (800) 707-46-19»);8 (800) 707-46-19
20;;"нет телефона (номер 8-80X — сервисная линия организации, не телефон клиента: «88007074619»);88007074619
21;;"нет телефона (номер 8-80X — сервисная линия организации, не телефон клиента: «8-800-707-46-19»);8-800-707-46-19
22;;"нет телефона (номер 8-80X — сервисная линия организации, не телефон клиента: «8 800 707 46 19»);8 800 707 46 19
23;;"нет телефона (номер 8-80X — сервисная линия организации, не телефон клиента: «+7 800 707-46-19»);+7 800 707-46-19
24;;битая дата (несуществующая дата: «16.09.2026»);16.09.2026
26;;битая дата (несуществующая дата: «17/09/2026»);17/09/2026
28;;битая дата (несуществующая дата: «18.09.2026»);18.09.2026
32;;битая дата (несуществующая дата: «19.09.2062»);19.09.2062
33;;битая дата (несуществующая дата: «20/09/2026»);20/09/2026
34;;битая дата (несуществующая дата: «09/21/2026»);09/21/2026
40;;битая дата (несуществующая дата: «09/24/2026»);09/24/2026
41;;битая дата (несуществующая дата: «24.09.2026»);24.09.2026
43;;"нет телефона (не номер: «abc»);abc
44;;битая дата (несуществующая дата: «09/25/2026»);09/25/2026
48;;битая дата (несуществующая дата: «09/27/2026»);09/27/2026
49;;битая дата (несуществующая дата: «28.09.2026»);28.09.2026
```
</details>

**🔧 Edit** /Users/[скрыто]/zayavki-test/clean_zayavki.py

```
            # _make(год, МЕСЯЦ, ДЕНЬ): для ДД/ММ месяц = b, для ММ/ДД месяц = a
            d = _make(y, a, b) if order == "MDY" else _make(y, b, a)
```

**🔧 Bash** Rerun on input and print all three outputs

```
cd /Users/[скрыто]/zayavki-test && .venv/bin/python clean_zayavki.py zayavki.xlsx -o output --today 2026-10-11 && .venv/bin/python -c "
import csv
for f in ['clean','problems','duplicates']:
    print('=====', f)
    for r in csv.reader(open(f'output/{f}.csv', encoding='utf-8-sig'), delimiter=';'):
        print(' | '.join(r))
"
```

<details><summary>результат</summary>

```
The file /Users/[скрыто]/zayavki-test/clean_zayavki.py has been updated successfully.
```
</details>

<details><summary>результат</summary>

```
Вход: zayavki.xlsx (лист «Заявки»), сегодня = 2026-10-11
Строк с данными на входе: 48
  чистых уникальных записей: 22
  удалено дублей по телефону: 14
  проблемных строк:           12 — нет телефона: 8, битая дата: 3, нет имени: 1
Сверка: 22 + 14 + 12 = 48 == вход, OK
Файлы:
  output/clean.xlsx  (+ clean.csv)
  output/problems.xlsx  (+ problems.csv)
  output/duplicates.xlsx  (+ duplicates.csv)
===== clean
Имя | Телефон | Дата заявки | Источник | Строка | Примечание
Сергей Волков | +79162451783 | 2026-09-03 | сайт | 2 | дубли по телефону: стр. 3; дата «09/03/2026» неоднозначна, принят формат ММ/ДД: все однозначные даты источника «сайт» записаны как ММ/ДД
Павлова Ирина | +79032218840 | 2026-09-05 | звонок | 4 | дубли по телефону: стр. 5; дата «05.09.2026» неоднозначна, принят формат ДД.ММ: все однозначные даты источника «звонок» записаны как ДД.ММ
Кузнецов Андрей | +79256003112 | 2026-09-06 | звонок | 6 | дубли по телефону: стр. 7
Морозова Елена | +79851129047 | 2026-09-07 | telegram | 8 | дубли по телефону: стр. 9 «Белов Никита» — другое имя, проверьте; дата «07/09/2026» неоднозначна, принят формат ДД/ММ: все однозначные даты источника «telegram» записаны как ДД/ММ
Новикова Татьяна | +79174005219 | 2026-09-10 | звонок | 12 | дубли по телефону: стр. 13; номер извлечён из текста «тел.: 8 917 400 52 19 (после 18:00)»; дата «10.09.2026» неоднозначна, принят формат ДД.ММ: все однозначные даты источника «звонок» записаны как ДД.ММ
Захарова Ольга | +79772410866 | 2026-09-13 | email | 17 | 
Петров Алексей | +79156402281 | 2026-09-16 | email | 24 | дубли по телефону: стр. 25
Королёва Дарья | +79261187502 | 2026-09-17 | telegram | 26 | дубли по телефону: стр. 27
Смирнов Олег | +79037741926 | 2026-09-18 | email | 28 | дубли по телефону: стр. 29
Карпова Юлия | +79775038214 | 2026-09-03 | email | 30 | дубли по телефону: стр. 31; дата «03.09.2026» неоднозначна, принят формат ДД.ММ: все однозначные даты источника «email» записаны как ДД.ММ
Степанова Вера | +79250196438 | 2026-09-20 | telegram | 33 | 
Яковлев Георгий | +79054467091 | 2026-09-21 | сайт | 34 | дата «09/21/2026» записана как ММ/ДД
Макарова Нина | +79992135860 | 2026-09-04 | telegram | 35 | дата «04/09/2026» неоднозначна, принят формат ДД/ММ: все однозначные даты источника «telegram» записаны как ДД/ММ
Зайцев Кирилл | +79176823340 | 2026-09-04 | сайт | 36 | дата «09/04/2026» неоднозначна, принят формат ММ/ДД: все однозначные даты источника «сайт» записаны как ММ/ДД
Борисова Алёна | +79857304215 | 2026-09-22 | 
… [сокращено, всего 8182 символов]
```
</details>

**🔧 Read** /Users/[скрыто]/zayavki-test/clean_zayavki.py

```
{"file_path": "/Users/[скрыто]/zayavki-test/clean_zayavki.py", "offset": 340, "limit": 50}
```

<details><summary>результат</summary>

```
340	    problem = check_range(d, today)
341	    return DateResult(None if problem else d, problem)
342	
343	
344	# ---------------------------------------------------------------------------
345	# Записи, разрешение неоднозначных дат, дедупликация
346	# ---------------------------------------------------------------------------
347	
348	@dataclass
349	class Record:
350	    row: int                       # номер строки в исходном Excel (заголовок = 1)
351	    raw_name: object
352	    raw_phone: object
353	    raw_date: object
354	    raw_source: object
355	    name: str | None = None
356	    phone: PhoneResult = None
357	    date_res: DateResult = None
358	    source: str = ""
359	    problems: list[str] = field(default_factory=list)
360	    notes: list[str] = field(default_factory=list)
361	    kept_row: int | None = None    # для удалённого дубля — какая строка оставлена
362	
363	    @property
364	    def date(self) -> date | None:
365	        return self.date_res.value
366	
367	
368	def build_record(row: int, name, phone, dt, source, today: date) -> Record:
369	    rec = Record(row, name, phone, dt, source)
370	    rec.name = clean_name(name)
371	    rec.phone = normalize_phone(phone)
372	    rec.date_res = parse_date(dt, today)
373	    rec.source = clean_text(source).lower()
374	    return rec
375	
376	
377	def resolve_ambiguous_dates(records: list[Record], today: date) -> None:
378	    """«09/03/2026» — 9 марта или 3 сентября? Порядок по шагам:
379	    1) если другой вариант невозможен (будущее/слишком старая) — берём возможный;
380	    2) если у этого источника все однозначные даты с тем же разделителем
381	       записаны одним порядком (сайт: 09/16, 09/24 -> ММ/ДД) — берём его;
382	    3) для «/» без такой подсказки — вариант, ближайший к медиане остальных дат файла;
383	    4) иначе — ДД.ММ (российская норма)."""
384	    evidence: dict[tuple[str, str], Counter] = defaultdict(Counter)
385	    for r in records:
386	        if r.date_res.order_evidence:
387	            evidence[(r.source, r.date_res.sep)][r.date_res.order_evidence] += 1
388	    known = sorted(r.date for r in records if r.date)
389	    ref = median_low(known) if known else None
```
</details>

**🔧 Read** /Users/[скрыто]/zayavki-test/clean_zayavki.py

```
{"file_path": "/Users/[скрыто]/zayavki-test/clean_zayavki.py", "offset": 389, "limit": 35}
```

<details><summary>результат</summary>

```
389	    ref = median_low(known) if known else None
390	
391	    for r in records:
392	        dr = r.date_res
393	        if not dr.ambiguous:
394	            continue
395	        raw = clean_text(r.raw_date)
396	        dmy, mdy = dr.ambiguous
397	        options = {k: v for k, v in (("DMY", dmy), ("MDY", mdy)) if not check_range(v, today)}
398	        votes = evidence[(r.source, dr.sep)]
399	        how = None
400	        if not options:
401	            dr.problem = check_range(dmy, today)
402	            continue
403	        if len(options) == 1:
404	            choice = next(iter(options))
405	            how = "второй вариант невозможен (будущее/слишком старая дата)"
406	        elif votes["DMY"] and not votes["MDY"]:
407	            choice, how = "DMY", f"все однозначные даты источника «{r.source}» записаны как ДД{dr.sep}ММ"
408	        elif votes["MDY"] and not votes["DMY"]:
409	            choice, how = "MDY", f"все однозначные даты источника «{r.source}» записаны как ММ{dr.sep}ДД"
410	        elif dr.sep == "/" and ref is not None:
411	            choice = min(options, key=lambda k: (abs((options[k] - ref).days), k != "DMY"))
412	            how = f"вариант ближе к остальным датам файла (медиана {ref.isoformat()})"
413	        else:
414	            choice = "DMY"
415	        dr.value = options[choice]
416	        if how:
417	            fmt = f"ДД{dr.sep}ММ" if choice == "DMY" else f"ММ{dr.sep}ДД"
418	            dr.note = f"дата «{raw}» неоднозначна, принят формат {fmt}: {how}"
419	
420	
421	def _collect_problems(rec: Record) -> None:
422	    if rec.phone.value is None:
423	        rec.problems.append(f"{NO_PHONE} ({rec.phone.problem})")
```
</details>

**🔧 Bash** Rewrite ambiguous-date resolver and add dup's own problems to notes

```
cd /Users/[скрыто]/zayavki-test && .venv/bin/python - <<'EOF'
p='clean_zayavki.py'
s=open(p,encoding='utf-8').read()
start=s.index('def resolve_ambiguous_dates(')
end=s.index('def _collect_problems(')
new='''def resolve_ambiguous_dates(records: list[Record], today: date) -> None:
    """Даты вида «09/03/2026», где и день, и месяц <= 12: 9 марта или 3 сентября?

    Сигнал «источника»: если у этого источника ВСЕ однозначные даты с тем же
    разделителем записаны одним порядком (сайт: 09/16, 09/24 -> ММ/ДД; telegram:
    17/09, 22/09 -> ДД/ММ), неоднозначные даты этого источника читаем так же.

    «.» и «-» (российская запись): ДД.ММ, если только сигнал источника не говорит
    обратное. Дату не «переворачиваем» из-за того, что она в будущем — это проблема.

    «/» (может быть и американской записью ММ/ДД), по шагам:
      1) сигнал источника;
      2) если один из вариантов невозможен (в будущем / раньше 2000 г.) — другой;
      3) вариант, ближайший к медиане остальных дат файла;
      4) ДД/ММ (российская норма)."""
    evidence: dict[tuple[str, str], Counter] = defaultdict(Counter)
    for r in records:
        if r.date_res.order_evidence:
            evidence[(r.source, r.date_res.sep)][r.date_res.order_evidence] += 1
    known = sorted(r.date for r in records if r.date)
    ref = median_low(known) if known else None

    for r in records:
        dr = r.date_res
        if not dr.ambiguous:
            continue
        raw = clean_text(r.raw_date)
        cand = dict(zip(("DMY", "MDY"), dr.ambiguous))
        votes = evidence[(r.source, dr.sep)]
        by_source = next((k for k in ("DMY", "MDY") if votes[k] and not votes[k == "DMY" and "MDY" or "DMY"]), None)
        how = None
        if by_source:
            choice = by_source
            how = f"все однозначные даты источника «{r.source or 'без источника'}» с «{dr.sep}» записаны так"
            if dr.sep != "/" and choice == "DMY":
                how = None  # обычная российская запись, пояснять нечего
        elif dr.sep != "/":
            choice = "DMY"
        else:
            feasible = [k for k in cand if not check_range(cand[k], today)]
            if len(feasible) == 1:
                choice, how = feasible[0], "второй вариант невозможен (в будущем или слишком старый)"
            elif feasible and ref is not None:
                choice = min(feasible, key=lambda k: (abs((cand[k] - ref).days), k != "DMY"))
                how = f"этот вариант ближе к остальным датам файла (медиана {ref.isoformat()})"
            else:
                choice, how = "DMY", "подсказок нет, по умолчанию российский порядок"
        problem = check_range(cand[choice], today)
        if problem:
            dr.problem = problem
            continue
        dr.value = cand[choice]
        if how:
            fmt = f"ДД{dr.sep}ММ" if choice == "DMY" else f"ММ{dr.sep}ДД"
            dr.note = f"дата «{raw}» неоднозначна, принят формат {fmt}: {how}"


'''
s=s[:start]+new+s[end:]

# duplicates: show the duplicate's own problems in its note
a='''            dup.notes[:0] = [f"дубль по телефону {keeper.phone.value}, оставлена строка {keeper.row}"]'''
b='''            head = [f"дубль по телефону {keeper.phone.value}, оставлена строка {keeper.row}"]
            if dup.problems:
                head.append("у самой строки: " + "; ".join(dup.problems))
            dup.notes[:0] = head'''
assert s.count(a)==1
s=s.replace(a,b)
open(p,'w',encoding='utf-8').write(s)
EOF
grep -n "by_source = next" clean_zayavki.py
```

<details><summary>результат</summary>

```
406:        by_source = next((k for k in ("DMY", "MDY") if votes[k] and not votes[k == "DMY" and "MDY" or "DMY"]), None)
```
</details>

### 🤖 Агент (19:37:55)

That `by_source` one-liner is unreadable; let me replace it with plain code.

**🔧 Edit** /Users/[скрыто]/zayavki-test/clean_zayavki.py

```
        by_source = None
        if votes["DMY"] and not votes["MDY"]:
            by_source = "DMY"
        elif votes["MDY"] and not votes["DMY"]:
            by_source = "MDY"
```

<details><summary>результат</summary>

```
<tool_use_error>File has been modified since read, either by the user or by a linter. Read it again before attempting to write it.</tool_use_error>
```
</details>

**🔧 Bash** Rerun and show row numbers with notes

```
cd /Users/[скрыто]/zayavki-test && .venv/bin/python clean_zayavki.py zayavki.xlsx -o output --today 2026-10-11 && .venv/bin/python -c "
import csv
for f in ['clean','problems','duplicates']:
    print('=====', f)
    for r in csv.reader(open(f'output/{f}.csv', encoding='utf-8-sig'), delimiter=';'):
        print(r[0], '|', r[-1])
"
```

<details><summary>результат</summary>

```
Вход: zayavki.xlsx (лист «Заявки»), сегодня = 2026-10-11
Строк с данными на входе: 48
  чистых уникальных записей: 22
  удалено дублей по телефону: 14
  проблемных строк:           12 — нет телефона: 8, битая дата: 3, нет имени: 1
Сверка: 22 + 14 + 12 = 48 == вход, OK
Файлы:
  output/clean.xlsx  (+ clean.csv)
  output/problems.xlsx  (+ problems.csv)
  output/duplicates.xlsx  (+ duplicates.csv)
===== clean
Имя | Примечание
Сергей Волков | дубли по телефону: стр. 3; дата «09/03/2026» неоднозначна, принят формат ММ/ДД: все однозначные даты источника «сайт» с «/» записаны так
Павлова Ирина | дубли по телефону: стр. 5
Кузнецов Андрей | дубли по телефону: стр. 7
Морозова Елена | дубли по телефону: стр. 9 «Белов Никита» — другое имя, проверьте; дата «07/09/2026» неоднозначна, принят формат ДД/ММ: все однозначные даты источника «telegram» с «/» записаны так
Новикова Татьяна | дубли по телефону: стр. 13; номер извлечён из текста «тел.: 8 917 400 52 19 (после 18:00)»
Захарова Ольга | 
Петров Алексей | дубли по телефону: стр. 25
Королёва Дарья | дубли по телефону: стр. 27
Смирнов Олег | дубли по телефону: стр. 29
Карпова Юлия | дубли по телефону: стр. 31
Степанова Вера | 
Яковлев Георгий | дата «09/21/2026» записана как ММ/ДД
Макарова Нина | дата «04/09/2026» неоднозначна, принят формат ДД/ММ: все однозначные даты источника «telegram» с «/» записаны так
Зайцев Кирилл | дата «09/04/2026» неоднозначна, принят формат ММ/ДД: все однозначные даты источника «сайт» с «/» записаны так
Борисова Алёна | дубли по телефону: стр. 38
Григорьев Денис | 
Кудрявцева Ксения | дата «09/24/2026» записана как ММ/ДД
Романов Виктор | дубли по телефону: стр. 42
Виноградов Лев | дата «09/25/2026» записана как ММ/ДД
Тарасова Мария | дубли по телефону: стр. 46
Белоусов Кирилл | дубли по телефону: стр. 47; дата «09/27/2026» записана как ММ/ДД
Ершова Полина | 
===== problems
Строка | Примечание
10 | 
11 | 
14 | 
15 | дубли по телефону: стр. 16
18 | 
19 | 
20 | 
21 | 
22 | 
23 | 
32 | 
43 | 
===== duplicates
Строка | Примечание
3 | дубль по телефону +79162451783, оставлена строка 2
5 | дубль по телефону +79032218840, оставлена строка 4; дата «09/05/2026» неоднозначна, принят формат ММ/ДД: все однозначные даты источника «сайт» с «/» записаны так
7 | дубль по телефону +79256003112, оставлена строка 6
9 | дубль по телефону +79851129047, оставлена строка 8; имя отличается от оставленной записи («Морозова Елена») — возможно, другой человек
13 | дубль по телефону +79174005219, оставлена строка 12; дата
… [сокращено, всего 3248 символов]
```
</details>

### 🤖 Агент (19:38:01)

The output is right. The Edit was rejected because my Python patch had just changed the file, so I'll re-read it and apply the readability fix again.

**🔧 Read** /Users/[скрыто]/zayavki-test/clean_zayavki.py

```
{"file_path": "/Users/[скрыто]/zayavki-test/clean_zayavki.py", "offset": 400, "limit": 12}
```

<details><summary>результат</summary>

```
400	        dr = r.date_res
401	        if not dr.ambiguous:
402	            continue
403	        raw = clean_text(r.raw_date)
404	        cand = dict(zip(("DMY", "MDY"), dr.ambiguous))
405	        votes = evidence[(r.source, dr.sep)]
406	        by_source = next((k for k in ("DMY", "MDY") if votes[k] and not votes[k == "DMY" and "MDY" or "DMY"]), None)
407	        how = None
408	        if by_source:
409	            choice = by_source
410	            how = f"все однозначные даты источника «{r.source or 'без источника'}» с «{dr.sep}» записаны так"
411	            if dr.sep != "/" and choice == "DMY":
```
</details>

**🔧 Edit** /Users/[скрыто]/zayavki-test/clean_zayavki.py

```
        by_source = None
        if votes["DMY"] and not votes["MDY"]:
            by_source = "DMY"
        elif votes["MDY"] and not votes["DMY"]:
            by_source = "MDY"
```

<details><summary>результат</summary>

```
The file /Users/[скрыто]/zayavki-test/clean_zayavki.py has been updated successfully.
```
</details>

**🔧 Edit** /Users/[скрыто]/zayavki-test/clean_zayavki.py

```
    if n == 10 and not had_plus:
        # Номер без кода страны: 9xx — мобильный, 3xx/4xx/80x–87x — коды регионов и служб.
        # «7…», «88…», «89…» из 10 цифр — это «7»/«8» + номер, в котором не хватает цифры
        # (кодов 88x/89x нет, а 7xx без +7 — это обрезанный номер).
        if digits[0] in "349" or (digits[0] == "8" and digits[1] not in "89"):
            return digits, None
        if digits[0] in "78":
            return None, f"не хватает цифры (10 цифр, начинается с {digits[:2]})"
        return None, f"10 цифр, несуществующий код {digits[:3]}"
```

<details><summary>результат</summary>

```
The file /Users/[скрыто]/zayavki-test/clean_zayavki.py has been updated successfully.
```
</details>

**🔧 Write** /Users/[скрыто]/zayavki-test/tests/make_test_file.py

```
"""Генерирует ДРУГОЙ тестовый файл той же структуры (Имя, Телефон, Дата заявки, Источник)
со своими «грязными» случаями — и ожидаемый результат для каждой строки.

Запуск:  python tests/make_test_file.py [путь.xlsx]
По умолчанию пишет tests/data/zayavki_test2.xlsx. Ожидания используются в тестах
(tests/test_clean_zayavki.py), «сегодня» для них = 2026-10-11.
"""
from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

import openpyxl

TODAY = "2026-10-11"

# (ключ, Имя, Телефон, Дата заявки, Источник, ожидание)
# ожидание:
#   ("clean", имя, телефон, дата, источник)
#   ("problem", подстрока причины)
#   ("dup", ключ оставленной строки)
#   ("empty",)  — пустая строка, пропускается
CASES = [
    ("A1", "  ИВАНОВ   иван ", "+7 (916) 111-22-33 доб. 45", "01.08.2026", "Сайт ",
     ("clean", "Иванов Иван", "+79161112233", "2026-08-01", "сайт")),
    ("A2", "Иван Иванов", 89161112233, datetime(2026, 8, 5, 14, 30), "звонок",
     ("dup", "A1")),  # тот же номер числом, более поздняя заявка
    ("B", "петрова анна-мария", "9162223344", "02.08.2026 14:30", "vk",
     ("clean", "Петрова Анна-Мария", "+79162223344", "2026-08-02", "vk")),
    ("C", "Сидоров Пётр", "8 916 333 44 5", "03.08.2026", "звонок",
     ("problem", "нет телефона (не хватает цифры")),
    ("D", "Kозлов Олег", "+7 916 444 55 66 7", "04.08.2026", "email",  # K — латинская
     ("problem", "нет телефона (лишние цифры (12)")),
    ("E", "John Smith", "+44 20 7946 0958", "Aug 10, 2026", "email",
     ("problem", "нет телефона (иностранный номер")),
    ("F", "Орлов Глеб", "8-916-555-66-77", 46245, "сайт",  # Excel-дата числом
     ("clean", "Орлов Глеб", "+79165556677", "2026-08-11", "сайт")),
    ("G", "Лисицына Ольга", "8 (495) 123-45-67", "5 авг. 2026 г.", "звонок",
     ("clean", "Лисицына Ольга", "+74951234567", "2026-08-05", "звонок")),
    ("H", "", "+7 916 666 77 88", "07.08.2026", "email", ("problem", "нет имени")),
    ("I", "   ", "+7 916 666 77 89", "07.08.2026", "email", ("problem", "нет имени")),
    ("J", "Нет", "+7 916 777 88 99", "08.08.2026", "telegram", ("problem", "нет имени")),
    ("K", "Ковалёв Артур", "8 804 333 22 11", "09.08.2026", "звонок",
     ("problem", "нет телефона (номер 8-80X")),
    ("L", "Попов Ян", "+7 000 000 00 00", "09.08.2026", "звонок",
     ("problem", "нет телефона (несуществующий код")),
    ("M", "Морозов Ефим", "89999999999", "09.08.2026", "звонок",
     ("problem", "нет телефона (номер-заглушка")),
    # «форма» пишет ММ/ДД (08/14 однозначно) -> 08/03 читаем как 3 августа
    ("N1", "Фролов Марк", "+7 917 100 00 01", "08/14/2026", "форма",
     ("clean", "Фролов Марк", "+79171000001", "2026-08-14", "форма")),
    ("N2", "Жукова Лиза", "+7 917 100 00 02", "08/03/2026", "форма",
     ("clean", "Жукова Лиза", "+79171000002", "2026-08-03", "форма")),
    # «vk» пишет ДД/ММ (15/08 однозначно) -> 03/08 читаем как 3 августа
    ("O1", "Комаров Ян", "+7 917 100 00 03", "15/08/2026", "vk",
     ("clean", "Комаров Ян", "+79171000003", "2026-08-15", "vk")),
    ("O2", "Белова Ия", "+7 917 100 00 04", "03/08/2026", "vk",
     ("clean", "Белова Ия", "+79171000004", "2026-08-03", "vk")),
    # у «авито» подсказок нет -> ближе к остальным датам файла (август): 9 августа, не 8 сентября
    ("P", "Носов Глеб", "+7 917 100 00 05", "08/09/2026", "авито",
     ("clean", "Носов Глеб", "+79171000005", "2026-08-09", "авито")),
    ("Q", "Титов Лев", "+7 917 100 00 06", "31.02.2026", "email",
     ("problem", "битая дата (несуществующая дата")),
    ("R", "Уткин Ром", "+7 917 100 00 07", "15.08.2027", "email",
     ("problem", "битая дата (дата в будущем")),
    ("S", "Шилов Ян", "+7 917 100 00 08", "01.01.1990", "email",
     ("problem", "битая дата (неправдоподобно старая")),
    ("T", "Ершов Ян", "+7 917 100 00 09", "10 августа", "звонок",
     ("problem", "битая дата (нет года")),
    ("U", "Дроздова Ада", "моб. 8 916 888 99 00, раб. 8 495 765 43 21", "2026.08.12", "звонок",
     ("clean", "Дрозд
… [сокращено, всего 7153 символов]
```

<details><summary>результат</summary>

```
File created successfully at: /Users/[скрыто]/zayavki-test/tests/make_test_file.py (file state is current in your context — no need to Read it back)
```
</details>

### 🤖 Агент (19:40:05)

Now the pytest suite.

**🔧 Write** /Users/[скрыто]/zayavki-test/tests/test_clean_zayavki.py

```
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pytest

import clean_zayavki as cz
import make_test_file as gen

TODAY = date(2026, 10, 11)
ROOT = Path(__file__).resolve().parent.parent
ORIGINAL = ROOT / "zayavki.xlsx"


# ---------------------------------------------------------------- телефон

@pytest.mark.parametrize("raw, expected", [
    # все форматы из исходного файла
    ("8 (916) 245-17-83", "+79162451783"),
    ("+7 916 245 17 83", "+79162451783"),
    (89032218840, "+79032218840"),                   # число в Excel
    ("+7 (903) 221-88-40", "+79032218840"),
    ("7 925 600 31 12", "+79256003112"),             # 7 без плюса
    (79256003112, "+79256003112"),
    ("+7 985 112-90-47", "+79851129047"),
    ("тел.: 8 917 400 52 19 (после 18:00)", "+79174005219"),  # номер внутри текста
    ("+79174005219", "+79174005219"),
    ("8-916-555-66-77", "+79165556677"),
    # дополнительные случаи
    ("9161234567", "+79161234567"),                  # 10 цифр без кода
    (9161234567, "+79161234567"),
    ("(812) 345-67-89", "+78123456789"),             # СПб без 8
    ("8 (495) 123-45-67", "+74951234567"),
    (89169990011.0, "+79169990011"),                 # float из Excel
    ("+7 916 111-22-33 доб. 45", "+79161112233"),    # добавочный отбрасываем
    ("моб. 8 916 888 99 00, раб. 8 495 765 43 21", "+79168889900"),
    ("8.916.123.45.67", "+79161234567"),
    ("+7 916 123 45 67", "+79161234567"),  # неразрывные пробелы
    ("+7 727 123 45 67", "+77271234567"),            # Казахстан — тоже +7
])
def test_phone_ok(raw, expected):
    res = cz.normalize_phone(raw)
    assert res.value == expected and res.problem is None


@pytest.mark.parametrize("raw, reason", [
    (None, "пусто"),
    ("", "пусто"),
    ("   ", "пусто"),
    ("не знаю", "не номер"),
    ("abc", "не номер"),
    ("8 (800) 707-46-19", "8-80X"),                  # бесплатная линия организации
    ("88007074619", "8-80X"),
    ("+7 800 707-46-19", "8-80X"),
    ("8 804 333 22 11", "8-80X"),
    ("8 916 333 44 5", "не хватает цифры"),          # 10 цифр, начинается с 89
    ("7 925 600 31 1", "не хватает цифры"),
    ("916 123 45", "не хватает цифр (8)"),
    ("+7 916 123 45 6", "не хватает цифр (10)"),
    ("+7 916 444 55 66 7", "лишние цифры (12)"),
    ("8 916 123 45 67 8", "лишние цифры (12)"),
    ("+44 20 7946 0958", "иностранный номер"),
    ("+7 000 000 00 00", "несуществующий код"),
    ("89999999999", "заглушка"),
    ("+7 (812) 3O1-22-33", "не хватает цифр"),        # буква O вместо нуля
    (True, "не номер"),
])
def test_phone_bad(raw, reason):
    res = cz.normalize_phone(raw)
    assert res.value is None
    assert reason in res.problem


def test_phone_notes():
    assert "добавочный" in cz.normalize_phone("+7 916 111-22-33 доб. 45").note
    assert "несколько номеров" in cz.normalize_phone("8 916 888 99 00, 8 495 765 43 21").note
    assert cz.normalize_phone("+7 916 111-22-33").note is None


# ---------------------------------------------------------------- дата

@pytest.mark.parametrize("raw, expected", [
    ("03.09.2026", date(2026, 9, 3)),
    ("13.09.26", date(2026, 9, 13)),                 # двузначный год -> 20YY
    ("2026-09-08", date(2026, 9, 8)),
    ("2026/09/26", date(2026, 9, 26)),
    ("2026.08.12", date(2026, 8, 12)),
    ("17/09/2026", date(2026, 9, 17)),               # однозначно ДД/ММ
    ("09/16/2026", date(2026, 9, 16)),               # однозначно ММ/ДД
    ("09/09/2026", date(2026, 9, 9)),                # день = месяц
    ("6 сентября 2026", date(2026, 9, 6)),
    ("23 сентября 2026", date(2026, 9, 23)),
    ("5 авг. 2026 г.", date(2026, 8, 5)),
    ("1 сент 2026г.", date(2026, 9, 1)),
    ("1 мая 2026 года", date(2026, 5, 1)),
    ("Aug 10, 2026", date(2026, 8, 10)),
    ("10 September 2026", date(2026, 9, 10)),
    ("24.09.2026 14:35", date(2026, 9, 24)),         # время отбрасываем
    ("2026-09-08T14:30:00", date(2026, 9, 8)),
    (datetime(2026, 9, 6), date(2026, 9, 6)),      
… [сокращено, всего 16596 символов]
```

**🔧 Write** /Users/[скрыто]/zayavki-test/pytest.ini

```
[pytest]
testpaths = tests
pythonpath = . tests

```

**🔧 Write** /Users/[скрыто]/zayavki-test/requirements.txt

```
openpyxl>=3.1
pytest>=8

```

<details><summary>результат</summary>

```
File created successfully at: /Users/[скрыто]/zayavki-test/tests/test_clean_zayavki.py (file state is current in your context — no need to Read it back)
```
</details>

<details><summary>результат</summary>

```
File created successfully at: /Users/[скрыто]/zayavki-test/pytest.ini (file state is current in your context — no need to Read it back)
```
</details>

<details><summary>результат</summary>

```
File created successfully at: /Users/[скрыто]/zayavki-test/requirements.txt (file state is current in your context — no need to Read it back)
```
</details>

**🔧 Bash** Run the pytest suite

```
cd /Users/[скрыто]/zayavki-test && .venv/bin/python -m pytest -q 2>&1 | tail -60
```

<details><summary>результат</summary>

```
.........................................F.............................. [ 46%]
........................................................................ [ 92%]
.........F..                                                             [100%]
=================================== FAILURES ===================================
______________________ test_date_ok[03.09.2026-expected0] ______________________

raw = '03.09.2026', expected = datetime.date(2026, 9, 3)

    @pytest.mark.parametrize("raw, expected", [
        ("03.09.2026", date(2026, 9, 3)),
        ("13.09.26", date(2026, 9, 13)),                 # двузначный год -> 20YY
        ("2026-09-08", date(2026, 9, 8)),
        ("2026/09/26", date(2026, 9, 26)),
        ("2026.08.12", date(2026, 8, 12)),
        ("17/09/2026", date(2026, 9, 17)),               # однозначно ДД/ММ
        ("09/16/2026", date(2026, 9, 16)),               # однозначно ММ/ДД
        ("09/09/2026", date(2026, 9, 9)),                # день = месяц
        ("6 сентября 2026", date(2026, 9, 6)),
        ("23 сентября 2026", date(2026, 9, 23)),
        ("5 авг. 2026 г.", date(2026, 8, 5)),
        ("1 сент 2026г.", date(2026, 9, 1)),
        ("1 мая 2026 года", date(2026, 5, 1)),
        ("Aug 10, 2026", date(2026, 8, 10)),
        ("10 September 2026", date(2026, 9, 10)),
        ("24.09.2026 14:35", date(2026, 9, 24)),         # время отбрасываем
        ("2026-09-08T14:30:00", date(2026, 9, 8)),
        (datetime(2026, 9, 6), date(2026, 9, 6)),        # настоящая дата Excel
        (datetime(2026, 8, 20, 15, 45), date(2026, 8, 20)),
        (date(2026, 9, 1), date(2026, 9, 1)),
        (46268, date(2026, 9, 3)),                       # Excel-дата числом
        (46268.75, date(2026, 9, 3)),
        ("46268", date(2026, 9, 3)),
    ])
    def test_date_ok(raw, expected):
        res = cz.parse_date(raw, TODAY)
>       assert res.value == expected and res.problem is None and res.ambiguous is None
E       AssertionError: assert (None == datetime.date(2026, 9, 3))
E        +  where None = DateResult(value=None, problem=None, note=None, ambiguous=(datetime.date(2026, 9, 3), datetime.date(2026, 3, 9)), sep='.', order_evidence=None).value

tests/test_clean_zayavki.py:109: AssertionError
______________ test_other_column_order_title_row_and_extra_column ______________

tmp_path = PosixPath('/private/var/folders/3c/v2_s0c7n0yq9k2w5knm5q14r0000gn/T/pytest-of-[скрыто]/pytest-0/test_other_column_order_title_0')

    def test_other_column_or
… [сокращено, всего 3399 символов]
```
</details>

**🔧 Bash** Fix two test mistakes and rerun pytest

```
cd /Users/[скрыто]/zayavki-test && .venv/bin/python - <<'EOF'
p='tests/test_clean_zayavki.py'
s=open(p,encoding='utf-8').read()
a='    ("03.09.2026", date(2026, 9, 3)),\n    ("13.09.26"'
b='    ("24.09.2026", date(2026, 9, 24)),               # 03.09 неоднозначна, её решает resolve_*\n    ("13.09.26"'
assert s.count(a)==1; s=s.replace(a,b)
a='    ws.save(tmp_path / "in.xlsx")'
assert s.count(a)==1; s=s.replace(a,'    wb.save(tmp_path / "in.xlsx")')
open(p,'w',encoding='utf-8').write(s)
EOF
.venv/bin/python -m pytest -q 2>&1 | tail -5
```

<details><summary>результат</summary>

```
........................................................................ [ 46%]
........................................................................ [ 92%]
............                                                             [100%]
156 passed in 0.62s
```
</details>

**🔧 Bash** Generate test file, run CLI on both files, independent cross-check

```
cd /Users/[скрыто]/zayavki-test && .venv/bin/python tests/make_test_file.py && .venv/bin/python clean_zayavki.py tests/data/zayavki_test2.xlsx -o output_test2 --today 2026-10-11 && echo ---- && .venv/bin/python clean_zayavki.py zayavki.xlsx -o output --today 2026-10-11 && echo ---- && .venv/bin/python -c "
# Независимая сверка: только re.sub по цифрам, без функций скрипта
import openpyxl, re
from collections import Counter
rows = list(openpyxl.load_workbook('zayavki.xlsx').active.iter_rows(min_row=2, values_only=True))
def ph(v):
    if v is None: return None
    s = str(v)
    m = re.search(r'[78]?[\s(]*\d{3}[\s)\-]*\d{3}[\s\-]*\d{2}[\s\-]*\d{2}', s)
    if not m: return None
    d = re.sub(r'\D','',m.group())[-10:]
    return None if d.startswith('80') else d
phones = [ph(r[1]) for r in rows]
valid = [p for p in phones if p]
c = Counter(valid)
print('строк:', len(rows), '| без телефона/8-800:', phones.count(None), '| с телефоном:', len(valid), '| уникальных номеров:', len(c), '| лишних повторов:', len(valid)-len(c))
print('группы >1:', sorted(v for v in c.values() if v>1), 'сумма лишних =', sum(v-1 for v in c.values()))
"
```

<details><summary>результат</summary>

```
/Users/[скрыто]/zayavki-test/tests/data/zayavki_test2.xlsx
Вход: tests/data/zayavki_test2.xlsx (лист «Заявки»), сегодня = 2026-10-11
Строк с данными на входе: 40 (пустых строк пропущено: 1)
  чистых уникальных записей: 18
  удалено дублей по телефону: 4
  проблемных строк:           18 — нет телефона: 8, битая дата: 7, нет имени: 3
Сверка: 18 + 4 + 18 = 40 == вход, OK
Файлы:
  output_test2/clean.xlsx  (+ clean.csv)
  output_test2/problems.xlsx  (+ problems.csv)
  output_test2/duplicates.xlsx  (+ duplicates.csv)
----
Вход: zayavki.xlsx (лист «Заявки»), сегодня = 2026-10-11
Строк с данными на входе: 48
  чистых уникальных записей: 22
  удалено дублей по телефону: 14
  проблемных строк:           12 — нет телефона: 8, битая дата: 3, нет имени: 1
Сверка: 22 + 14 + 12 = 48 == вход, OK
Файлы:
  output/clean.xlsx  (+ clean.csv)
  output/problems.xlsx  (+ problems.csv)
  output/duplicates.xlsx  (+ duplicates.csv)
----
строк: 48 | без телефона/8-800: 8 | с телефоном: 40 | уникальных номеров: 26 | лишних повторов: 14
группы >1: [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2] сумма лишних = 14
```
</details>

**🔧 Bash** Print outputs for the generated test file

```
cd /Users/[скрыто]/zayavki-test && .venv/bin/python -c "
import csv
for f in ['clean','problems','duplicates']:
    print('=====', f)
    for r in csv.reader(open(f'output_test2/{f}.csv', encoding='utf-8-sig'), delimiter=';'):
        print(' | '.join(r))
"
```

<details><summary>результат</summary>

```
===== clean
Имя | Телефон | Дата заявки | Источник | Строка | Примечание
Иванов Иван | +79161112233 | 2026-08-01 | сайт | 2 | дубли по телефону: стр. 3; номер извлечён из текста «+7 (916) 111-22-33 доб. 45»; добавочный номер отброшен
Петрова Анна-Мария | +79162223344 | 2026-08-02 | vk | 4 | 
Орлов Глеб | +79165556677 | 2026-08-11 | сайт | 8 | 
Лисицына Ольга | +74951234567 | 2026-08-05 | звонок | 9 | 
Фролов Марк | +79171000001 | 2026-08-14 | форма | 16 | дата «08/14/2026» записана как ММ/ДД
Жукова Лиза | +79171000002 | 2026-08-03 | форма | 17 | дата «08/03/2026» неоднозначна, принят формат ММ/ДД: все однозначные даты источника «форма» с «/» записаны так
Комаров Ян | +79171000003 | 2026-08-15 | vk | 18 | 
Белова Ия | +79171000004 | 2026-08-03 | vk | 19 | дата «03/08/2026» неоднозначна, принят формат ДД/ММ: все однозначные даты источника «vk» с «/» записаны так
Носов Глеб | +79171000005 | 2026-08-09 | авито | 20 | дата «08/09/2026» неоднозначна, принят формат ММ/ДД: этот вариант ближе к остальным датам файла (медиана 2026-08-12)
Дроздова Ада | +79168889900 | 2026-08-12 | звонок | 25 | номер извлечён из текста «моб. 8 916 888 99 00, раб. 8 495 765 43 21»; в ячейке несколько номеров, взят первый корректный
Зуев Ян | +79169990011 | 2026-08-12 | звонок | 26 | 
Белкина Зоя | +79261002030 | 2026-08-03 | сайт | 29 | дубли по телефону: стр. 28
Гришин Олег | +79262003040 | 2026-08-04 | email | 30 | дубли по телефону: стр. 31 «Гришина Ольга» — другое имя, проверьте
О'нил Шон | +79264005060 | 2026-09-01 | email | 34 | 
Сёмин Ян | +79265006070 | 2026-08-20 | telegram | 35 | 
Аксенов Ян | +79267008090 | 2026-08-16 | email | 37 | 
Сазонов Ян | +79269001020 | 2026-09-11 | почта | 41 | дата «09/11/2026» неоднозначна, принят формат ММ/ДД: второй вариант невозможен (в будущем или слишком старый)
Титова Ада | +79261112200 | 2026-08-13 | email | 42 | 
===== problems
Строка | Имя | Телефон | Дата заявки | Источник | Причина | Имя (исходное) | Телефон (исходный) | Дата (исходная) | Примечание
5 | Сидоров Пётр | нет телефона | 2026-08-03 | звонок | нет телефона (не хватает цифры (10 цифр, начинается с 89): «8 916 333 44 5») | Сидоров Пётр | 8 916 333 44 5 | 03.08.2026 | 
6 | Козлов Олег | нет телефона | 2026-08-04 | email | нет телефона (лишние цифры (12): «+7 916 444 55 66 7») | Kозлов Олег | +7 916 444 55 66 7 | 04.08.2026 | 
7 | John Smith | нет телефона | 2026-08-10 | email | нет телефона (иностранный номер (+442…): «+44 20 7946 0958») | John Smith | +44 20 7946 0958 | Aug 10
… [сокращено, всего 5461 символов]
```
</details>

**🔧 Bash** Tidy reason wording and update tests accordingly

```
cd /Users/[скрыто]/zayavki-test && .venv/bin/python - <<'EOF'
p='clean_zayavki.py'
s=open(p,encoding='utf-8').read()
reps=[
('return None, f"иностранный номер (+{digits[:3]}…)"', 'return None, "иностранный номер (код страны не +7)"'),
('return None, f"не хватает цифр ({n})" if n < 11 else f"лишние цифры ({n})"', 'return None, f"не хватает цифр: {n}" if n < 11 else f"лишние цифры: {n}"'),
('return None, f"не хватает цифры (10 цифр, начинается с {digits[:2]})"', 'return None, f"не хватает цифры: 10 цифр, начинается с {digits[:2]}"'),
('        return None, f"не хватает цифр ({n})"\n    return None, f"лишние цифры ({n})"', '        return None, f"не хватает цифр: {n}"\n    return None, f"лишние цифры: {n}"'),
('return f"дата в будущем ({d.isoformat()})"', 'return f"дата в будущем: {d.isoformat()}"'),
('return f"неправдоподобно старая дата ({d.isoformat()})"', 'return f"неправдоподобно старая дата: {d.isoformat()}"'),
]
for a,b in reps:
    assert s.count(a)==1, a
    s=s.replace(a,b)
open(p,'w',encoding='utf-8').write(s)

p='tests/test_clean_zayavki.py'
s=open(p,encoding='utf-8').read()
for a,b in [('"не хватает цифр (8)"','"не хватает цифр: 8"'),('"не хватает цифр (10)"','"не хватает цифр: 10"'),
            ('"лишние цифры (12)"),\n    ("8 916','"лишние цифры: 12"),\n    ("8 916'),('"лишние цифры (12)"),\n    ("+44','"лишние цифры: 12"),\n    ("+44')]:
    assert s.count(a)==1, a
    s=s.replace(a,b)
open(p,'w',encoding='utf-8').write(s)
p='tests/make_test_file.py'
s=open(p,encoding='utf-8').read()
for a,b in [('"нет телефона (лишние цифры (12)"','"нет телефона (лишние цифры: 12"'),
            ('"битая дата (дата в будущем (2026-11-10)"','"битая дата (дата в будущем: 2026-11-10"')]:
    assert s.count(a)==1, a
    s=s.replace(a,b)
open(p,'w',encoding='utf-8').write(s)
EOF
grep -n "иностранный\|лишние цифры\|не хватает цифр" clean_zayavki.py tests/*.py | head -20; .venv/bin/python -m pytest -q 2>&1 | tail -3
```

<details><summary>результат</summary>

```
tests/make_test_file.py:32:     ("problem", "нет телефона (не хватает цифры")),
tests/make_test_file.py:34:     ("problem", "нет телефона (лишние цифры: 12")),
tests/make_test_file.py:36:     ("problem", "нет телефона (иностранный номер")),
tests/make_test_file.py:96:     ("problem", "нет телефона (не хватает цифр")),
tests/make_test_file.py:98:     ("problem", "нет телефона (не хватает цифры")),
tests/test_clean_zayavki.py:56:    ("8 916 333 44 5", "не хватает цифры"),          # 10 цифр, начинается с 89
tests/test_clean_zayavki.py:57:    ("7 925 600 31 1", "не хватает цифры"),
tests/test_clean_zayavki.py:58:    ("916 123 45", "не хватает цифр: 8"),
tests/test_clean_zayavki.py:59:    ("+7 916 123 45 6", "не хватает цифр: 10"),
tests/test_clean_zayavki.py:60:    ("+7 916 444 55 66 7", "лишние цифры: 12"),
tests/test_clean_zayavki.py:61:    ("8 916 123 45 67 8", "лишние цифры: 12"),
tests/test_clean_zayavki.py:62:    ("+44 20 7946 0958", "иностранный номер"),
tests/test_clean_zayavki.py:65:    ("+7 (812) 3O1-22-33", "не хватает цифр"),        # буква O вместо нуля
clean_zayavki.py:147:        return None, "иностранный номер (код страны не +7)"
clean_zayavki.py:149:        return None, f"не хватает цифр: {n}" if n < 11 else f"лишние цифры: {n}"
clean_zayavki.py:154:        # «7…», «88…», «89…» из 10 цифр — это «7»/«8» + номер, в котором не хватает цифры
clean_zayavki.py:159:            return None, f"не хватает цифры: 10 цифр, начинается с {digits[:2]}"
clean_zayavki.py:162:        return None, f"не хватает цифр: {n}"
clean_zayavki.py:163:    return None, f"лишние цифры: {n}"
........................................................................ [ 92%]
............                                                             [100%]
156 passed in 0.25s
```
</details>

### 🤖 Агент (19:41:58)

156 tests pass. Let me also check whether the system Homebrew python3 (3.14) can run the script, so the README can say so accurately, and inspect the final xlsx cell types.

**🔧 Bash** Check system python, inspect xlsx output, list files

```
cd /Users/[скрыто]/zayavki-test && /opt/homebrew/bin/python3 -c "import pyexpat, openpyxl; print('sys py3.14 ok', openpyxl.__version__)" 2>&1 | tail -2; .venv/bin/python clean_zayavki.py zayavki.xlsx -o output --today 2026-10-11 >/dev/null && .venv/bin/python clean_zayavki.py tests/data/zayavki_test2.xlsx -o output_test2 --today 2026-10-11 > /dev/null && .venv/bin/python -c "
import openpyxl
ws = openpyxl.load_workbook('output/clean.xlsx').active
print(ws.title, ws.dimensions, ws.freeze_panes, ws.auto_filter.ref)
for c in ws[2]: print(repr(c.value), c.number_format, end=' || ')
print()
print({k: round(v.width) for k,v in ws.column_dimensions.items()})
"; ls -la output output_test2 tests tests/data
```

<details><summary>результат</summary>

```
  Referenced from: <1AAB1AFF-0B6A-3B67-86F5-A72CB69F8B1B> /opt/homebrew/Cellar/python@3.14/3.14.5/Frameworks/Python.framework/Versions/3.14/lib/python3.14/lib-dynload/pyexpat.cpython-314-darwin.so
  Expected in:     <850FB22F-BB06-316B-B5F4-FCD1C0385184> /usr/lib/libexpat.1.dylib
Чистые A1:F23 A2 A1:F23
'Сергей Волков' @ || '+79162451783' @ || '2026-09-03' @ || 'сайт' @ || 2 General || 'дубли по телефону: стр. 3; дата «09/03/2026» неоднозначна, принят формат ММ/ДД: все однозначные даты источника «сайт» с «/» записаны так' @ || 
{'A': 19, 'B': 14, 'C': 13, 'D': 10, 'E': 8, 'F': 70}
output:
total 72
drwxr-xr-x@  8 [скрыто]  staff   256 Oct 11 22:36 .
drwxr-xr-x@ 12 [скрыто]  staff   384 Oct 11 22:41 ..
-rw-r--r--@  1 [скрыто]  staff  3286 Oct 11 22:42 clean.csv
-rw-r--r--@  1 [скрыто]  staff  6819 Oct 11 22:42 clean.xlsx
-rw-r--r--@  1 [скрыто]  staff  3795 Oct 11 22:42 duplicates.csv
-rw-r--r--@  1 [скрыто]  staff  6921 Oct 11 22:42 duplicates.xlsx
-rw-r--r--@  1 [скрыто]  staff  2737 Oct 11 22:42 problems.csv
-rw-r--r--@  1 [скрыто]  staff  6482 Oct 11 22:42 problems.xlsx

output_test2:
total 72
drwxr-xr-x@  8 [скрыто]  staff   256 Oct 11 22:41 .
drwxr-xr-x@ 12 [скрыто]  staff   384 Oct 11 22:41 ..
-rw-r--r--@  1 [скрыто]  staff  2500 Oct 11 22:42 clean.csv
-rw-r--r--@  1 [скрыто]  staff  6678 Oct 11 22:42 clean.xlsx
-rw-r--r--@  1 [скрыто]  staff  1194 Oct 11 22:42 duplicates.csv
-rw-r--r--@  1 [скрыто]  staff  5801 Oct 11 22:42 duplicates.xlsx
-rw-r--r--@  1 [скрыто]  staff  3377 Oct 11 22:42 problems.csv
-rw-r--r--@  1 [скрыто]  staff  7062 Oct 11 22:42 problems.xlsx

tests:
total 64
drwxr-xr-x@  5 [скрыто]  staff    160 Oct 11 22:41 .
drwxr-xr-x@ 12 [скрыто]  staff    384 Oct 11 22:41 ..
drwxr-xr-x@  3 [скрыто]  staff     96 Oct 11 22:41 data
-rw-r--r--@  1 [скрыто]  staff   9115 Oct 11 22:41 make_test_file.py
-rw-r--r--@  1 [скрыто]  staff  19032 Oct 11 22:41 test_clean_zayavki.py

tests/data:
total 16
drwxr-xr-x@ 3 [скрыто]  staff    96 Oct 11 22:41 .
drwxr-xr-x@ 5 [скрыто]  staff   160 Oct 11 22:41 ..
-rw-r--r--@ 1 [скрыто]  staff  6685 Oct 11 22:41 zayavki_test2.xlsx
```
</details>

### 🤖 Агент (19:42:09)

The system python3 3.14 is indeed broken (pyexpat can't load libexpat), so the README will point to the 3.12 venv. Now the README.

**🔧 Write** /Users/[скрыто]/zayavki-test/README.md

```
# Чистка файла заявок

`clean_zayavki.py` — CLI-скрипт: берёт xlsx с колонками **Имя, Телефон, Дата заявки, Источник**
и выдаёт чистый файл, список проблемных строк с причиной и отчёт об удалённых дублях.
Правила общие, под конкретные строки ничего не подогнано: скрипт рассчитан на любой файл той же структуры.

## Установка и запуск

```bash
/opt/homebrew/bin/python3.12 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python clean_zayavki.py zayavki.xlsx -o output
.venv/bin/python -m pytest -q          # тесты (156 шт.)
```

Подойдёт любой Python 3.10+ с `openpyxl`. Системный `python3` 3.14 из Homebrew на этой машине
сломан (не грузится `pyexpat`), поэтому нужен venv на 3.12.

Параметры (`--help`):

| параметр | что делает |
|---|---|
| `input` | входной .xlsx |
| `-o / --out-dir` | папка результатов (по умолчанию `output` рядом со входным файлом) |
| `--clean / --problems / --duplicates` | явные пути к выходным .xlsx (рядом пишется .csv с тем же именем) |
| `--sheet` | лист (по умолчанию первый) |
| `--today ГГГГ-ММ-ДД` | «сегодня» для проверки дат в будущем (по умолчанию текущая дата) |
| `--csv-sep` | разделитель CSV (по умолчанию `;` — русский Excel открывает такой CSV двойным кликом; кодировка UTF-8 с BOM) |

## Что получается на выходе

| файл | содержимое |
|---|---|
| `clean.xlsx` / `.csv` | уникальные записи без проблем: Имя, Телефон, Дата заявки, Источник + **Строка** (номер строки в исходнике, заголовок = 1) + **Примечание** (что склеено, как прочитана неоднозначная дата и т.п.) |
| `problems.xlsx` / `.csv` | проблемные строки: нормализованные значения, **Причина**, исходные значения ячеек (чтобы поправить руками) |
| `duplicates.xlsx` / `.csv` | удалённые дубли: какая строка оставлена вместо них и почему |

Каждая входная строка попадает **ровно в один** из трёх файлов. Скрипт печатает сверку:
`входных строк = чистых + дублей + проблемных`. Полностью пустые строки пропускаются, их число тоже печатается.

**На `zayavki.xlsx`: 48 строк → 22 чистых + 14 дублей + 12 проблемных.**
Проблемные: 8 без телефона (из них 5 с номером 8-800), 3 с битой датой, 1 без имени.

## Правила

### Телефон → `+7XXXXXXXXXX`
1. Из ячейки берутся куски, похожие на номер: цифры, разделённые пробелами, дефисами, скобками и точками.
   Текст вокруг отбрасывается («тел.: 8 917 400 52 19 (после 18:00)» → `+79174005219`).
   Короткие числа (< 5 цифр: «18:00», «доб. 45») номером не считаются.
   Если номеров несколько, берётся первый корректный, и об этом пишется примечание.
2. Число в ячейке Excel (`89032218840`, `8.9E10`) обрабатывается так же, как текст.
3. 11 цифр, начинающихся с `8` или `7` (с плюсом или без), → `+7` и последние 10 цифр.
4. 10 цифр без кода страны: `9xx` (мобильный), `3xx`/`4xx`/`80x–87x` (коды регионов и служб) → `+7` + 10 цифр.
   `7…`, `88…`, `89…` из 10 цифр — это `7`/`8` плюс номер с пропущенной цифрой → проблема «не хватает цифры».
5. Всё остальное — **«нет телефона»** с уточнением в «Причине»:
   пусто; не номер («не знаю», «abc»); не хватает или лишние цифры; иностранный код (`+44…`);
   несуществующий код (`+7 0xx`, `1xx`, `2xx`, `5xx`, `6xx`); заглушка из одинаковых цифр (`+7 999 999-99-99`);
   **номер 8-80X** (см. спорные случаи).
6. Добавочный номер («доб. 45», «ext», «#») отбрасывается с пометкой: в формат `+7XXXXXXXXXX` он не помещается.

### Дата → `ГГГГ-ММ-ДД`
- Настоящая дата Excel (ячейка-дата) берётся как есть, время отбрасывается.
- Excel-дата числом (`46268` → 2026-09-03), в том числе сохранённая текстом.
- `ГГГГ-ММ-ДД`, `ГГГГ/ММ/ДД`, `ГГГГ.ММ.ДД`.
- `ДД.ММ.ГГГГ`, `ДД-ММ-ГГГГ`, `ДД/ММ/ГГГГ`, `ММ/ДД/ГГГГ` (порядок см. ниже), год из двух цифр: `26` → `2026`.
- Месяц словом: «6 сентября 2026», «5 авг. 2026 г.», «1 сент 2026г.», «Aug 10, 2026».
- Хвост со временем («24.09.2026 14:35», «…T14:30:00») отбрасывается.
- **Битая дата** (в проблемные, строка не теряется): пусто; не распознана («хх.09.2026»);
  нет года («12 сентября»); несуществующая («31.02.2026»); **в будущем** относительно `-
… [сокращено, всего 9422 символов]
```

<details><summary>результат</summary>

```
File created successfully at: /Users/[скрыто]/zayavki-test/README.md (file state is current in your context — no need to Read it back)
```
</details>

### 🤖 Агент (19:42:55)

Let me re-read the final script once in full to check for stale comments or leftover issues.

**🔧 Read** /Users/[скрыто]/zayavki-test/clean_zayavki.py

```
{"file_path": "/Users/[скрыто]/zayavki-test/clean_zayavki.py", "offset": 120, "limit": 80}
```

<details><summary>результат</summary>

```
120	
121	
122	def name_completeness(name: str | None) -> int:
123	    """Сколько в имени полных слов (не инициалов)."""
124	    return sum(1 for t in (name or "").split() if len(t.strip(".")) > 1 and not t.endswith("."))
125	
126	
127	# ---------------------------------------------------------------------------
128	# Телефон
129	# ---------------------------------------------------------------------------
130	
131	@dataclass
132	class PhoneResult:
133	    value: str | None          # +7XXXXXXXXXX или None
134	    problem: str | None = None  # причина для списка проблемных
135	    note: str | None = None     # пояснение для чистого файла
136	
137	
138	# Кусок текста, похожий на номер: цифры, между ними пробелы/дефисы/скобки/точки.
139	_PHONE_CHUNK = re.compile(r"\+?\d(?:[\s\-.()]*\d)*")
140	_EXT_WORDS = re.compile(r"доб|ext|вн\.|#", re.I)
141	
142	
143	def _national(digits: str, had_plus: bool) -> tuple[str | None, str | None]:
144	    """10 цифр национального номера или (None, причина)."""
145	    n = len(digits)
146	    if had_plus and not digits.startswith("7"):
147	        return None, "иностранный номер (код страны не +7)"
148	    if had_plus and n != 11:
149	        return None, f"не хватает цифр: {n}" if n < 11 else f"лишние цифры: {n}"
150	    if n == 11 and digits[0] in "78":
151	        return digits[1:], None
152	    if n == 10 and not had_plus:
153	        # Номер без кода страны: 9xx — мобильный, 3xx/4xx/80x–87x — коды регионов и служб.
154	        # «7…», «88…», «89…» из 10 цифр — это «7»/«8» + номер, в котором не хватает цифры
155	        # (кодов 88x/89x нет, а 7xx без +7 — это обрезанный номер).
156	        if digits[0] in "349" or (digits[0] == "8" and digits[1] not in "89"):
157	            return digits, None
158	        if digits[0] in "78":
159	            return None, f"не хватает цифры: 10 цифр, начинается с {digits[:2]}"
160	        return None, f"10 цифр, несуществующий код {digits[:3]}"
161	    if n < 10:
162	        return None, f"не хватает цифр: {n}"
163	    return None, f"лишние цифры: {n}"
164	
165	
166	def _check_national(nat: str) -> str | None:
167	    if nat[0] not in "34789":
168	        return f"несуществующий код +7 {nat[:3]}"
169	    if len(set(nat)) == 1:
170	        return "номер-заглушка из одинаковых цифр"
171	    if nat.startswith("80"):
172	        return "номер 8-80X — сервисная линия организации, не телефон клиента"
173	    return None
174	
175	
176	def normalize_phone(raw) -> PhoneResult:
177	    """Приводит 
… [сокращено, всего 3549 символов]
```
</details>

**🔧 Read** /Users/[скрыто]/zayavki-test/clean_zayavki.py

```
{"file_path": "/Users/[скрыто]/zayavki-test/clean_zayavki.py", "offset": 200, "limit": 25}
```

<details><summary>результат</summary>

```
200	        if nat is not None:
201	            problem = _check_national(nat)
202	        if problem is None:
203	            notes = []
204	            if re.search(r"[^\W\d_]", text):
205	                notes.append(f"номер извлечён из текста «{text}»")
206	            if _EXT_WORDS.search(text):
207	                notes.append("добавочный номер отброшен")
208	            if len([c for c in chunks if len(re.sub(r"\D", "", c)) >= 10]) > 1:
209	                notes.append("в ячейке несколько номеров, взят первый корректный")
210	            return PhoneResult("+7" + nat, None, "; ".join(notes) or None)
211	        first_problem = first_problem or f"{problem}: «{text}»"
212	
213	    return PhoneResult(None, first_problem or f"не номер: «{text}»")
214	
215	
216	# ---------------------------------------------------------------------------
217	# Дата
218	# ---------------------------------------------------------------------------
219	
220	MIN_DATE = date(2000, 1, 1)
221	EXCEL_EPOCH = date(1899, 12, 30)  # система дат Excel 1900 (с учётом бага 29.02.1900)
222	
223	_MONTHS = {
224	    "янв": 1, "фев": 2, "мар": 3, "апр": 4, "май": 5, "мая": 5, "июн": 6, "июл": 7,
```
</details>

**🔧 Bash** Final test run and remove generated caches

```
cd /Users/[скрыто]/zayavki-test && .venv/bin/python -m pytest -q 2>&1 | tail -2 && rm -rf .pytest_cache __pycache__ tests/__pycache__ && ls -A
```

<details><summary>результат</summary>

```
............                                                             [100%]
156 passed in 0.26s
.venv
README.md
ZADANIE.txt
clean_zayavki.py
output
output_test2
pytest.ini
requirements.txt
tests
zayavki.xlsx
```
</details>

**🔧 SubagentHandback** 

```
{"message": "## Отчёт: чистка файла заявок (/Users/[скрыто]/zayavki-test)\n\nСкрипт готов. На исходном файле: **48 строк → 22 чистых, 14 дублей убрано, 12 проблемных** (22 + 14 + 12 = 48). Все 156 тестов проходят. Остановиться после шага 1 и дождаться ответа я не мог (работал как сабагент, без возможности переспросить), поэтому сводка по данным идёт первой, а за ней всё остальное.\n\n### 1. Что в данных (все 48 строк)\n- **Телефоны:** встречаются `8 (916) 245-17-83`, `+7 916 245 17 83`, `+7 (903) 221-88-40`, `7 925 600 31 12` (без плюса), `+79174005219`, `8-800-707-46-19`, а также числа в ячейке Excel (`89032218840`, `79256003112`).\n  - Номер внутри текста: «тел.: 8 917 400 52 19 (после 18:00)».\n  - Не телефоны: пусто (стр. 10), «не знаю» (стр. 11), «abc» (стр. 43).\n  - Один номер 8-800 у **5 разных людей** (стр. 19–23).\n  - Один мобильный номер у двух разных людей: Морозова Елена и Белов Никита (стр. 8–9).\n- **Даты:**\n  - через точку `ДД.ММ.ГГГГ`, `13.09.26` (год из двух цифр), ISO `2026-09-08`, `2026/09/26`;\n  - месяц словом: «6 сентября 2026»;\n  - настоящие даты Excel (2 шт.);\n  - через «/» — смесь `ДД/ММ` и `ММ/ДД`. Порядок зависит от источника: «сайт» всегда пишет ММ/ДД (`09/16`, `09/24`), «telegram» — ДД/ММ (`17/09`, `22/09`). Неоднозначных дат 7 (`09/03`, `07/09`, `04/09`, `09/04`…);\n  - битые: пусто (стр. 15), «хх.09.2026», «12 сентября» без года, 19.09.2062 в будущем.\n- **Имена:** разный регистр («сергей волков», «ПАВЛОВА ирина»), двойной пробел, неразрывный пробел, невидимый символ zero-width, латинская «e» в «Пeтров», ё/е («Королёва»/«Королева»), инициал «К. Белоусов», только имя «влад», пустое имя (стр. 18).\n- **Дубли:** 14 пар по телефону, в основном это один человек с переставленными «Имя Фамилия».\n- **Источники:** сайт, email, telegram, звонок — без пропусков.\n\n### 2. Спорные решения\n- **8-800 (и весь диапазон 8-80X)** считаю «нет телефона»: это линия организации, а не клиента. Строки не склеиваются, поэтому 5 разных людей не схлопываются в одного.\n- **Один номер, разные имена** (Морозова/Белов): склеиваются по правилу задания («одинаковый номер = один человек»), но в clean и duplicates стоит пометка «другое имя — проверьте».\n- **Какую запись оставлять из дублей:** сначала запись без проблем, затем самая ранняя заявка (первое обращение — его дата и источник), затем более полное имя, затем строка выше в файле. Поля из разных строк не смешиваются.\n- **Даты через «/»:** порядок берётся по источнику; если подсказки нет — вариант, который не попадает в будущее; затем ближайший к медиане дат файла; иначе ДД/ММ. Даты через точку всегда читаются как ДД.ММ и не «переворачиваются».\n- **Без автоисправления:** дата без года, дата в будущем и номер с лишними или недостающими цифрами идут в проблемные, исходное значение сохраняется рядом.\n\n### 3. Что сделано\n- `clean_zayavki.py`: argparse с `--help`, `-o`, `--clean/--problems/--duplicates`, `--sheet`, `--today`, `--csv-sep`. Колонки ищутся по заголовкам, поэтому порядок колонок не важен.\n- На выходе `clean`, `problems` (с колонкой «Причина») и дополнительно `duplicates` — каждый в `.xlsx` и `.csv`. Каждая входная строка попадает ровно в один файл, скрипт сам печатает сверку.\n- Создан venv на Python 3.12 и `requirements.txt`. Системный python3 3.14 действительно сломан (не грузится pyexpat).\n- Тесты: 156 штук. Кроме проверки функций, есть прогон на исходном файле (каждая строка 2–49 ровно в одном файле, плюс конкретные решения), перепутанный порядок колонок, повторный прогон на собственном clean (ничего не меняется) и коды ошибок.\n- Второй тестовый файл: `tests/make_test_file.py` генерирует 41 строку со своими случаями — добавочный номер, иностранный номер, 8-804, заглушка, Excel-дата числом, 31.02, 1990 год, английский месяц, конфликт имён и другие. Для каждой строки задан ожидаемый результат. Итог: 40 строк с данными → 18 чистых + 4 дубля + 18 проблемных.\n- Ручная сверка исходника. Отдельный скрипт, который просто вытаскивает цифры (без функ
… [сокращено, всего 6516 символов]
```

<details><summary>результат</summary>

```
{"success":true,"message":"Report delivered to your caller."}
```
</details>
