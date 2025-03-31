# app_functions.py

from docx import Document
import logging
from typing import Tuple, List, Optional

from utiles import insert_highlights
from taskes import *  # Импортирует все задачи
import pandas as pd  # Если требуется в дальнейшем

logger = logging.getLogger(__name__)

# Определение словаря задач
task2func = {
    '3. Курсив в докладной записке не используется': task3,
    '10. При указании номера документа или его пункта пробел не ставится': task10,
    '12. Формат кавычек - «...»': task12, 
    '13. Форматирование сокращений: "млн" и "млрд" без точек, "тыс." и "руб." с точками': task13,
    '14. Пробелы вокруг знака "/" удалены': task14,
    '15. Удаление пробела перед знаком "%"': task15,
    '19. Правильное использование тире и дефисов': task19,
    '23. Вместо слова \'Выявлено\' использовать \'Установлено\'': task23,
    '24. Вместо слова \'Сотрудник\' использовать \'Работник\'': task24,
    '25. Удаление множественных пробелов': task25
}

def apply_format(doc: Document, func, desc: str) -> Tuple[Document, List[str], List[Tuple[int, int]], List[Tuple[int, int]]]:
    """
    Применяет форматирующую функцию и возвращает обновленный документ, историю изменений,
    а также пустые списки для error_spans и fix_spans.
    """
    doc, history = func(doc)
    error_spans = []  # Пустой список, так как не используется
    fix_spans = []    # Пустой список, так как не используется
    return doc, history, error_spans, fix_spans

def apply_all_formats(doc: Document) -> Tuple[Document, List[str], List[Tuple[int, int]], List[Tuple[int, int]]]:
    """
    Применяет все форматирующие задачи из task2func.
    
    :param doc: Объект документа.
    :return: Кортеж (обновлённый документ, история изменений, error_spans, fix_spans).
    """
    history = []
    error_spans = []
    fix_spans = []
    for desc, func in task2func.items():
        logger.debug(f"Применение задачи: {desc}")
        doc, func_history, func_error_spans, func_fix_spans = apply_format(doc, func, desc)
        history.extend(func_history)
        error_spans.extend(func_error_spans)
        fix_spans.extend(func_fix_spans)
    return doc, history, error_spans, fix_spans

def apply_selected_formats(doc: Document, selected_tasks: List[str]) -> Tuple[Document, List[str], List[Tuple[int, int]], List[Tuple[int, int]]]:
    """
    Применяет выбранные форматирующие задачи.
    """
    history = []      # История изменений
    error_spans = []  # Пустой список
    fix_spans = []    # Пустой список

    for desc in selected_tasks:
        func = task2func.get(desc)
        if func:
            logger.debug(f"Применение задачи: {desc}")
            doc, func_history, func_error_spans, func_fix_spans = apply_format(doc, func, desc)
            history.extend(func_history)
            error_spans.extend(func_error_spans)
            fix_spans.extend(func_fix_spans)
        else:
            logger.warning(f"Форматирующая задача с описанием '{desc}' не найдена.")

    return doc, history, error_spans, fix_spans

def get_all_spans(doc: Document, selected_tasks: Optional[List[str]] = None) -> Tuple[List[str], List[Tuple[int, int]], List[Tuple[int, int]]]:
    """
    Получает все истории изменений и спаны ошибок/исправлений для выбранных задач.
    
    :param doc: Объект документа.
    :param selected_tasks: Список выбранных задач. Если None, применяются все задачи.
    :return: Кортеж (history, error_spans, fix_spans).
    """
    if selected_tasks is None:
        doc, history, error_spans, fix_spans = apply_all_formats(doc)
    else:
        doc, history, error_spans, fix_spans = apply_selected_formats(doc, selected_tasks)
    
    return history, error_spans, fix_spans