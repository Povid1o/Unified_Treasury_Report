"""Проверка, что код компилируется на САМОМ СТАРОМ поддерживаемом Python.

Зачем. На рабочих машинах стоит Python 3.9/3.10, а разрабатывается проект на
более новом. Синтаксис, появившийся позже, компилятор разработчика принимает
молча, и ошибка вылезает только у пользователя при запуске — причём сразу
фатальная, на импорте консоли. Именно так случилось с f-строкой, у которой
выражение перенесено на следующую строку: это разрешено только с Python 3.12
(PEP 701).

Проверять таким кодом нечего — нужен парсер С ГРАММАТИКОЙ НУЖНОЙ ВЕРСИИ:

* ast.parse(..., feature_version=(3, 9)) НЕ ловит такие f-строки — проверено;
* vermin тоже не ловит: он разбирает файл текущим интерпретатором;
* parso умеет грамматику конкретной версии и ловит.

Поэтому тест использует parso, если он установлен (pip install parso), и
пропускается, если нет. Полная гарантия — прогнать набор тестов настоящим
интерпретатором 3.9, см. README, раздел «Тесты».
"""
import sys
import unittest
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_DIR))

try:
    import parso
    HAS_PARSO = True
except ImportError:
    HAS_PARSO = False

# Самая старая версия, на которой проект должен запускаться.
MIN_PYTHON = "3.9"

SKIP_PARTS = {".venv", "venv", "__pycache__", "build", "dist"}


def project_sources():
    for path in sorted(BASE_DIR.rglob("*.py")):
        if not any(part in SKIP_PARTS for part in path.parts):
            yield path


@unittest.skipUnless(HAS_PARSO, "нужен пакет parso: pip install parso")
class SyntaxCompatibilityTests(unittest.TestCase):
    def test_every_source_file_parses_on_the_oldest_supported_python(self):
        grammar = parso.load_grammar(version=MIN_PYTHON)
        problems = []
        for path in project_sources():
            code = path.read_text(encoding="utf-8")
            for error in grammar.iter_errors(grammar.parse(code)):
                line = code.splitlines()[error.start_pos[0] - 1].strip()
                problems.append(
                    f"{path.relative_to(BASE_DIR)}:{error.start_pos[0]} "
                    f"{error.message}\n    {line[:100]}"
                )
        self.assertEqual(
            problems, [],
            f"код не компилируется на Python {MIN_PYTHON}:\n" + "\n".join(problems),
        )

    def test_the_checker_actually_catches_newer_syntax(self):
        """Тест бесполезен, если парсер молча принимает новый синтаксис."""
        too_new = "\n".join([
            'values = {"A": 1}',
            "x = f\"{', '.join(f'{v}' for v in values) or 'нет '",
            "      'продолжение'}\"",
        ])
        grammar = parso.load_grammar(version=MIN_PYTHON)
        self.assertTrue(
            list(grammar.iter_errors(grammar.parse(too_new))),
            "парсер не считает f-строку из PEP 701 ошибкой — проверка ничего не стоит",
        )


if __name__ == "__main__":
    unittest.main()
