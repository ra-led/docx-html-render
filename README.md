## Запуск полного стэка сервиса

1. Build the Docker image:
   ```bash
   docker-compose build
   ```

2. Run the Flask application:
   ```bash
   docker-compose up
   ```
   
Веб приложение будет доступно по адресу `http://localhost:5000`.

# Модуль для парсинга документов .docx

Этот модуль предназначен для парсинга документов формата .docx и преобразования их содержимого в HTML. В данном описании представлен обзор работы модуля, включая основные классы и функции, используемые для обработки документов.

## Структура модуля

Модуль состоит из следующих файлов:

- `__init__.py`
- `core.py`
- `ml.py`
- `numbering.py`
- `ooxml.py`

## Описание файлов

### `__init__.py`

Этот файл инициализирует модуль и импортирует основные классы из других файлов:

- `DocHandler` из `ooxml.py` для обработки документа.
- `DocHTML` из `export.py` для экспорта документа в HTML.

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

- **DocHandler**: Основной класс, который обрабатывает документ, извлекая параграфы и таблицы, и формирует их в иерархию.
- **table_extend**: Определяет, могут ли две таблицы быть объединены.
- **concat_tables**: Объединяет две таблицы.

## Ход процесса парсинга

1. **Импортирование документа**: Используется библиотека `docx` для загрузки документа.
2. **Инициализация обработчика**: Создается экземпляр класса `DocHandler`, который принимает документ и настройки.
3. **Обработка документа**: Вызывается метод `process` класса `DocHandler`, который:
   - Итерирует по содержимому документа.
   - Обрабатывает параграфы с помощью `ParHandler`.
   - Обрабатывает таблицы с помощью `TableHandler`.
4. **Формирование иерархии**: Для каждого параграфа и таблицы создается узел (`Node`) и добавляется в иерархию документа.
5. **Обработка нумерации и стилей**: Используется `NumberingDB` для обработки нумерации и стилей параграфов.
6. **Классификация текста**: При необходимости используется `BERTTextClassifier` для классификации текста.
7. **Формирование HTML**: После обработки всего содержимого, документ готов для экспорта в HTML.

## Пример использования
Этот пример демонстрирует, как загрузить документ и обработать его с помощью `DocHandler`. После вызова `process`, документ будет обработан, и его содержимое будет доступно для дальнейшей обработки или экспорта в HTML или JSON.

```python
import docx
from doc_parse.ooxml import DocHandler

# Загрузка документа
doc = docx.Document('path_to_docx')

# Инициализация обработчика
handler = DocHandler(doc)

# Обработка документа
handler.process()
```
