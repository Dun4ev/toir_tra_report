# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.7] - 2025-10-16
### Added
- Учет периодичности из `TZ_glob.xlsx` (колонка E) при подборе суффиксов индексных папок.
- Нормализация периодичности `М/Г` → `M/G` и их использование в названии итоговых каталогов.
- Сообщения об ошибках с детализацией по индексам, кодам Reserved и периодичности.

### Changed
- Документация: обновлён `README.md`, добавлено описание новой логики формирования папок.
- Расширены unit-тесты `tests/test_index_folder_builder.py` для проверки обработки периодичности.

## [2.6] - 2025-10-01
### Changed
- Актуализированы шаблоны формирования документов и настройки графического интерфейса.

### Fixed
- Исправлены ошибки группировки файлов с кириллицей в индексах.

## [2.5] - 2025-10-01
### Added
- Первый релиз ветки v2.x с обновлённым GUI и логикой подготовки отчётов ТОиР.
