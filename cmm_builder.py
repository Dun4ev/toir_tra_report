from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable

from openpyxl import load_workbook
from openpyxl.workbook.defined_name import DefinedName


StatusCallback = Callable[[str], None] | None
NormalizeKey = Callable[[str], str]
SUPPORTED_EXTENSIONS = (".docx", ".pdf")
DATE_FORMAT = "dd.mm.yyyy"
REPORT_NAME_RANGE = "ReportName"
CREATED_DATE_RANGE = "CreatedDate"
EXTRA_FIELD_RANGE = "ExtraField1"
RE_INDEX = re.compile(
    r"\b([IVXLCDM]+)\.(\d+)(?:\.(\d+))?(?:\.(\d+))?([A-Za-z\u0400-\u04FF])?\b",
    re.IGNORECASE,
)


@dataclass
class CommentSheetResult:
    """Результат пакетной генерации CMM."""

    created: list[Path]
    skipped_existing: list[Path]
    failed: list[tuple[Path, str]]


def _notify(callback: StatusCallback, message: str) -> None:
    """Отправляет сообщение о статусе, если колбэк задан."""
    if callback is not None:
        callback(message)


def _iter_candidate_files(base_dir: Path) -> Iterable[Path]:
    """Возвращает поток файлов с поддерживаемыми расширениями."""
    for ext in SUPPORTED_EXTENSIONS:
        yield from base_dir.rglob(f"*{ext}")


def extract_index_from_name(filename: str) -> str | None:
    """Извлекает индекс вида CT-DR из имени файла."""
    match = RE_INDEX.search(filename)
    if not match:
        return None
    roman, num1, num2, num3, suffix = match.groups()
    parts = [roman.upper(), num1]
    if num2:
        parts.append(num2)
    if num3:
        parts.append(num3)
    index = ".".join(parts)
    if suffix:
        index += suffix
    return index


def ensure_named_range(ws, wb, cell, name: str) -> None:
    """Создаёт именованный диапазон, если он отсутствует."""
    defined = {dn.name for dn in wb.defined_names.definedName}
    if name in defined:
        return
    wb.defined_names.add(
        DefinedName(name=name, attr_text=f"'{ws.title}'!{cell.coordinate}")
    )


def fill_basic_fields(workbook, report_name: str) -> None:
    """Заполняет базовые поля (имя отчёта и дату создания)."""
    worksheet = workbook.active
    defined = dict(workbook.defined_names.items())

    if REPORT_NAME_RANGE in defined:
        for sheet, coord in defined[REPORT_NAME_RANGE].destinations:
            target = workbook[sheet] if isinstance(sheet, str) else sheet
            target[coord].value = report_name
    else:
        worksheet["D1"].value = report_name
        ensure_named_range(worksheet, workbook, worksheet["D1"], REPORT_NAME_RANGE)

    today = datetime.now()
    if CREATED_DATE_RANGE in defined:
        for sheet, coord in defined[CREATED_DATE_RANGE].destinations:
            target = workbook[sheet] if isinstance(sheet, str) else sheet
            cell = target[coord]
            cell.value = today
            cell.number_format = DATE_FORMAT
    else:
        worksheet["D4"].value = today
        worksheet["D4"].number_format = DATE_FORMAT
        ensure_named_range(worksheet, workbook, worksheet["D4"], CREATED_DATE_RANGE)


def fill_extra_fields(
    workbook, report_name: str, tz_map: dict[str, str], normalize_key: NormalizeKey
) -> None:
    """Устанавливает описание из TZ в именованный диапазон ExtraField1."""
    worksheet = workbook.active
    index_code = extract_index_from_name(report_name)

    extra_value = "Отсутствует информация в TZ"
    if index_code:
        lookup_key = normalize_key(index_code)
        description = tz_map.get(lookup_key)
        if description:
            extra_value = description
        else:
            extra_value = f"Нет описания для {index_code}"
    else:
        extra_value = f"Индекс отсутствует ({report_name})"

    defined = dict(workbook.defined_names.items())
    if EXTRA_FIELD_RANGE in defined:
        for sheet, coord in defined[EXTRA_FIELD_RANGE].destinations:
            target = workbook[sheet] if isinstance(sheet, str) else sheet
            target[coord].value = extra_value
    else:
        worksheet["D6"].value = extra_value
        ensure_named_range(worksheet, workbook, worksheet["D6"], EXTRA_FIELD_RANGE)


def create_comment_sheet(
    report_path: Path,
    template_path: Path,
    tz_map: dict[str, str],
    normalize_key: NormalizeKey,
) -> Path:
    """Создаёт новый CMM рядом с исходным отчётом."""
    stem = report_path.stem
    output_path = report_path.with_name(f"{stem}_CMM.xlsx")
    if output_path.exists():
        raise FileExistsError(str(output_path))

    workbook = load_workbook(template_path)
    workbook.template = False
    fill_basic_fields(workbook, stem)
    fill_extra_fields(workbook, stem, tz_map, normalize_key)
    workbook.save(output_path)
    return output_path


def generate_comment_sheets(
    source_dir: Path,
    template_path: Path,
    tz_map: dict[str, str],
    normalize_key: NormalizeKey,
    status_callback: StatusCallback = None,
) -> CommentSheetResult:
    """Генерирует CMM для всех подходящих файлов в каталоге."""
    resolved_dir = source_dir.resolve()
    if not resolved_dir.exists() or not resolved_dir.is_dir():
        raise FileNotFoundError(f"Каталог не найден: {resolved_dir}")
    if not template_path.exists():
        raise FileNotFoundError(f"Шаблон не найден: {template_path}")

    _notify(status_callback, f"Поиск файлов для CMM в {resolved_dir}")
    candidates = sorted({p for p in _iter_candidate_files(resolved_dir)})
    created: list[Path] = []
    skipped: list[Path] = []
    failed: list[tuple[Path, str]] = []

    if not candidates:
        _notify(status_callback, "Нет файлов для обработки")
        return CommentSheetResult(created=[], skipped_existing=[], failed=[])

    for candidate in candidates:
        name_upper = candidate.name.upper()
        if not name_upper.startswith("CT-DR-"):
            continue

        _notify(status_callback, f"Обработка файла {candidate.name}")
        try:
            result_path = create_comment_sheet(
                candidate, template_path, tz_map, normalize_key
            )
        except FileExistsError:
            skipped.append(candidate)
            _notify(
                status_callback, f"Пропущен (уже существует CMM): {candidate.name}"
            )
            continue
        except Exception as error:  # noqa: BLE001
            failed.append((candidate, str(error)))
            _notify(status_callback, f"Ошибка: {candidate.name} - {error}")
            continue

        created.append(result_path)
        _notify(status_callback, f"Создан: {result_path.name}")

    return CommentSheetResult(created=created, skipped_existing=skipped, failed=failed)
