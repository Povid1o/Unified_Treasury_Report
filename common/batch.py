"""Запуск ETL по списку входных файлов: раздельные CSV (по умолчанию) или
один объединённый CSV (опция combine=True). Общий механизм для всех
"файловых" отчётов (ОВП, Структура баланса, ЧПД, NIM).
"""
from pathlib import Path
from typing import Callable, List

import pandas as pd

from common import ui


def run_batch(
    input_paths: List[Path],
    build_report: Callable[[Path], pd.DataFrame],
    save_report: Callable[[pd.DataFrame, Path], Path],
    default_output_path: Callable[[Path], Path],
    combine: bool,
    combined_output_path: Path,
) -> None:
    """Строит отчёт по каждому файлу из input_paths.

    combine=False (по умолчанию): каждый файл -> свой CSV (default_output_path).
    combine=True: все файлы объединяются в один DataFrame -> combined_output_path.
    """
    if not input_paths:
        raise ValueError("Не передано ни одного входного файла")

    if combine:
        frames = []
        for i, path in enumerate(input_paths, start=1):
            ui.console.print(f"[grey70]— [{i}/{len(input_paths)}] обрабатываю {path.name}...[/grey70]")
            frames.append(build_report(path))
        combined = pd.concat(frames, ignore_index=True)
        save_report(combined, combined_output_path)
        ui.success(
            f"Готово: {len(combined)} строк из {len(input_paths)} файлов "
            f"объединено в {combined_output_path}"
        )
        return

    for i, path in enumerate(input_paths, start=1):
        ui.console.print(f"[grey70]— [{i}/{len(input_paths)}] обрабатываю {path.name}...[/grey70]")
        df = build_report(path)
        output_path = default_output_path(path)
        save_report(df, output_path)
        ui.success(f"{path.name}: {len(df)} строк сохранено в {output_path}")
