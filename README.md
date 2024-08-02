# Запуск

1. Скачать репозиторий и перейти в корневой католог репо:
   ```bash
   git clone https://github.com/ra-led/docx-html-render.git
   cd docx-html-render
   ```

2. Собрать Docker образы:
   ```bash
   docker-compose build
   ```

3. Запустить полный стэк сервисов:
   ```bash
   docker-compose up -d
   ```
   
Веб приложение будет доступно по адресу `http://localhost:5000`.

# Модуль для парсинга .docx документов doc_parse

Этот модуль предназначен для парсинга документов формата .docx и преобразования их содержимого в HTML и JSON. В данном описании представлен обзор работы модуля, включая основные классы и функции, используемые для обработки документов.

## Структура модуля `doc_parse`

Модуль состоит из следующих файлов:

- `doc_parse/__init__.py`
- `doc_parse/core.py`
- `doc_parse/ml.py`
- `doc_parse/numbering.py`
- `doc_parse/ooxml.py`
- `doc_parse/export_html.py`
- `doc_parse/export_json.py`

## Описание файлов

### `__init__.py`

Этот модуль предоставляет функции для конвертации документов различных форматов. Он включает в себя три основные функции:

- `doc_to_docx`: Преобразует файл формата .doc в формат .docx с использованием библиотеки Aspose.Words. Может быть использован для обновления старых документов .doc до более современного формата .docx.
- `docx_to_html`: Преобразует документ формата .docx в HTML. Полезно для отображения документов в веб-браузерах или интеграции с веб-приложениями.
- `docx_to_json`: Преобразует документ формата .docx в форматированный, согласно спецификации,  JSON. 

### `core.py`

Содержит основные классы для обработки параграфов и таблиц в документе:

- `Node`: Базовый класс для представления узлов в иерархии документа.
- `ParHandler`: Обрабатывает параграфы, извлекая текст, стили и XML-представление.
- `TableHandler`: Обрабатывает таблицы, извлекая их размеры, содержимое и определяя, являются ли они рамками.
- `CellHandler`: Обрабатывает ячейки таблиц, извлекая их содержимое и размеры.
- `TableView`: Представляет таблицу в виде узла иерархии документа.
- `DocRoot`: Представляет корневой узел документа.

### `ml.py`

Содержит класс для классификации текста с использованием модели BERT:

- `BERTTextClassifier`: Классифицирует текст, используя предобученную модель BERT.

### `numbering.py`

Содержит классы и функции для обработки нумерации и стилей в документе:

- `NumberingDB`: Обрабатывает нумерацию и стили параграфов, используя XML-представление документа.
- `find_manual_numbering`: Определяет ручную нумерацию в тексте.
- `int_to_roman`: Преобразует целое число в римскую цифру.

### `ooxml.py`

Содержит основной класс для обработки документа и преобразования его содержимого в HTML:

- `DocHandler`: Основной класс, который обрабатывает документ, извлекая параграфы и таблицы, и формирует их в иерархию.
- `table_extend`: Определяет, могут ли две таблицы быть объединены.
- `concat_tables`: Объединяет две таблицы.

## Пример использования
Экспорт в HTML

```python
import docx
from doc_parse import DocHandler, DocHTML

doc = docx.Document(path_to_docx_file)
handler = DocHandler(doc)
converter = DocHTML()

html_content, toc_links = converter.get_html(handler)
```

Экспорт в JSON

```python
import docx
from doc_parse import DocHandler, DocJSON

doc = docx.Document(path_to_docx_file)
handler = DocHandler(doc)
converter = DocJSON()

json_content = converter.get_html(handler)
```

Экспорт в HTML и JSON (используется один DocHandler, чтобы избежать повторной обработки документа)

```python
import docx
from doc_parse import DocHandler, DocHTML, DocJSON

doc = docx.Document(path_to_docx_file)
handler = DocHandler(doc)

# Convert to HTML
html_converter = DocHTML()
html_content, toc_links = html_converter.get_html(handler)

json_converter = DocJSON()
json_content = json_converter.get_json(handler)
```

## Ход процесса парсинга

1. **Импортирование документа**: Используется библиотека `docx` для загрузки документа.
2. **Инициализация обработчика**: Создается экземпляр класса `DocHandler`, который принимает документ и настройки.
3. **Обработка документа**: Вызывается метод `process` класса `DocHandler`, который:
   - Итерирует по содержимому документа.
   - Обрабатывает параграфы с помощью `ParHandler`.
   - Обрабатывает таблицы с помощью `TableHandler`.
4. **Формирование иерархии**: Для каждого параграфа и таблицы создается узел (`Node`) и добавляется в иерархию документа.
5. **Обработка нумерации и стилей**: Используется `NumberingDB.numerize` для обработки нумерации и стилей параграфов (более подробно описано ниже).
6. **Классификация текста**: При необходимости используется `BERTTextClassifier` для классификации текста.
7. **Формирование**: После обработки всего содержимого, документ готов для экспорта в HTML или JSON.

## Расчет нумерации с NumberingDB
### Инициализация класса `NumberingDB`

При создании экземпляра класса `NumberingDB`, выполняются следующие действия:

1. **Инициализация параметров**:
   - `doc`: DOCX документ для обработки.
   - `appendix_header_length`: Максимальная длина заголовка приложения.
   - `default_levels`: Количество уровней нумерации по умолчанию.
   - `default_font`: Размер шрифта по умолчанию.
   - `norm_numeration_model`: Путь к модели для классификации нумерованных заголовков.
   - `norm_heading_model`: Путь к модели для классификации ненумерованных заголовков.

2. **Парсинг XML нумерации**:
   - Извлекает XML-данные нумерации из DOCX документа.
   - Если парсинг не удался, устанавливает `self.num_xml` в пустой словарь.

3. **Инициализация дефолтных абстрактных уровней**:
   - Создает дефолтные уровни нумерации с помощью `self.init_default_abstract`.

4. **Получение и связывание уровней и абстрактных уровней**:
   - Извлекает ID тредов нумерации (`numId`) и абстрактные уровни нумерации из XML-данных.
   - Извлекает ID и метаданные стилей определенных в документе.
   - Связывает треды нумерации с абстрактными уровнями и стили с абстрактными уровнями.

5. **Инициализация инкремента нумерации**:
   - Создает словарь для отслеживания инкремента нумерации.

6. **Инициализация классификаторов**:
   - Загружает модели BERT для классификации потенциальных заголовков.

7. **Определение стоп-символов**:
   - Устанавливает список стоп-символов, которые могут влиять на обработку нумерации.

### Обработка параграфа в функции `NumberingDB.numerize`

Функция `NumberingDB.numerize` обрабатывает параграф, пытаясь определить и нормализовать его нумерацию. В этой функции происходит следующее:

1. **Обновление статистики размера шрифта**:
   - Добавляет размер шрифта параграфа в список `self.font_size`.

2. **Определение приоритета методов обработки**:
   - Устанавливает порядок применения методов для обработки нумерации:
     1. `numrize_by_meta`: Обрабатывает нумерацию по метаданным.
     2. `numrize_by_style`: Обрабатывает нумерацию по стилю.
     3. `numerize_by_text`: Определяет нумерацию в текстовом префиксе.
     4. `numerize_by_heading`: Обрабатывает нумерацию по заголовку.
     5. `numerize_by_appendix`: Определяет нумерацию для приложений.

3. **Применение методов обработки**:
   - Последовательно применяет каждый метод из списка `numerize_prioritet` к параграфу.
   - Если какой-либо метод успешно определил нумерацию (`par.node.num_prefix` не пустой), останавливает дальнейшую обработку.

4. **Возврат обработанного параграфа**:
   - Возвращает параграф с установленной нумерацией (если она была определена).

### Подробности методов обработки

- **`numrize_by_meta`**:
  - Извлекает `numId` и уровень нумерации из XML-данных параграфа.
  - Получает абстрактный ID нумерации.
  - Считает встроенную нумерацию для данного абстрактного ID и уровня.
  - Проверяет, является ли параграф заголовком и корректирует глубину нумерации.
  - Проверяет наличие стоп-символов в префиксе нумерации и начале текста.

- **`numrize_by_style`**:
  - Получает абстрактный ID нумерации по стилю параграфа.
  - Считает встроенную нумерацию для данного абстрактного ID и уровня.

- **`numerize_by_text`**:
  - Определяет нумерацию в текстовом префиксе параграфа.
  - Проверяет наличие стоп-символов в начале текста.
  - Проверяет, является ли параграф заголовком.
  - Проверяет результат классификатора - является ли текст параграфа нумерованным заголовком.

- **`numerize_by_heading`**:
  - Проверяет, является ли параграф заголовком.
  - Проверяет результат классификатора - является ли текст параграфа заголовком.

- **`numerize_by_appendix`**:
  - Определяет, является ли параграф заголовком приложения.

