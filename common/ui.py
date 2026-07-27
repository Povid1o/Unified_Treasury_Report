"""Оформление консоли: единый rich.Console, спиннер и стилизованные сообщения.

Console — общий на весь процесс: логирование (см. logging_utils) рендерится
через тот же объект, что и спиннер, поэтому лог-строки корректно "всплывают"
над крутящимся спиннером вместо того, чтобы ломать его отрисовку.
"""
from contextlib import contextmanager
from typing import Optional, Sequence

from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich import box

console = Console()

APP_TITLE = "Навигатор отчётов Казначейства"


def print_header() -> None:
    console.print(Panel.fit(f"[bold cyan]{APP_TITLE}[/bold cyan]", border_style="cyan", padding=(0, 2)))


def print_menu(reports: Sequence) -> None:
    table = Table(box=box.ROUNDED, show_header=True, header_style="bold cyan", border_style="grey50")
    table.add_column("#", justify="right", style="bold yellow", no_wrap=True)
    table.add_column("Отчёт", style="bold white", no_wrap=True)
    table.add_column("Описание", style="grey70")
    for i, r in enumerate(reports, start=1):
        table.add_row(str(i), r.title, r.description)
    table.add_row("0", "[red]Выход[/red]", "")
    console.print(table)


@contextmanager
def spinner(message: str):
    """Крутящийся индикатор загрузки. Лог-строки (через RichHandler на том же
    Console) корректно печатаются над ним, не ломая анимацию. Возвращает
    объект status — вызывающий код может status.update("новый текст") для
    отображения прогресса (например, "файл 2 из 5") без пересоздания спиннера.
    """
    with console.status(f"[bold cyan]{message}[/bold cyan]", spinner="dots") as status:
        yield status


def success(message: str) -> None:
    console.print(f"[bold green]✓[/bold green] {message}")


def error(message: str) -> None:
    console.print(f"[bold red]✗ Ошибка:[/bold red] {message}")


def warning(message: str) -> None:
    console.print(f"[yellow]![/yellow] {message}")


def cancelled(message: str = "Операция отменена пользователем.") -> None:
    console.print(f"[yellow]⏹  {message}[/yellow]")


def ask(prompt: str, default: Optional[str] = None) -> str:
    """Стилизованный input(). Ctrl+C всплывает как KeyboardInterrupt — вызывающий
    код должен его перехватывать (см. console.py), а не глотать здесь.
    """
    hint = f" [grey50]({default})[/grey50]" if default else ""
    value = console.input(f"[bold cyan]?[/bold cyan] {prompt}{hint}: ").strip()
    return value or (default or "")
