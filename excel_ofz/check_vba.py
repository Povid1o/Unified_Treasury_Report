"""Статическая проверка VBA-модуля перед сборкой книги.

Компилятора VBA вне Windows нет, а ошибки вида «Only comments may appear after
End Sub» вылезают только при открытии книги в Excel. Здесь ловится то, что
можно поймать чтением текста:

* объявления уровня модуля после первой процедуры;
* несбалансированные блоки (Sub/Function/If/For/Do/With/Select Case/#If);
* `On Error GoTo` на несуществующую метку;
* `Exit Sub` внутри Function и наоборот;
* идентификаторы, заканчивающиеся подчёркиванием (VBA примет `_` за перенос).

Запуск: python excel_ofz/check_vba.py [путь к .bas]
"""
from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import List, NamedTuple

DEFAULT_BAS = Path(__file__).resolve().parent / "ofz_report.bas"

PROC_START = re.compile(
    r"^(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)",
    re.I)
PROC_END = re.compile(r"^End\s+(Sub|Function|Property)\b", re.I)
DECL_START = re.compile(r"^(?:Public|Private|Dim|Const|Global|Declare|Type|Enum|Option|Attribute)\b", re.I)
LABEL = re.compile(r"^(\w+):\s*$")
ON_ERROR_GOTO = re.compile(r"^On\s+Error\s+GoTo\s+(\w+)\s*$", re.I)


class Problem(NamedTuple):
    line: int
    text: str
    message: str


def _strip_strings_and_comments(line: str) -> str:
    """Убирает строковые литералы и комментарий, чтобы не путать разбор."""
    out = []
    in_string = False
    i = 0
    while i < len(line):
        ch = line[i]
        if in_string:
            if ch == '"':
                if i + 1 < len(line) and line[i + 1] == '"':
                    i += 2
                    continue
                in_string = False
            i += 1
            continue
        if ch == '"':
            in_string = True
            out.append('""')
            i += 1
            continue
        if ch == "'":
            break
        out.append(ch)
        i += 1
    text = "".join(out)
    return re.sub(r"^\s*Rem\b.*$", "", text, flags=re.I).strip()


def _logical_lines(raw_lines: List[str]) -> List[tuple[int, str]]:
    """Склеивает переносы ` _` в одну логическую строку, сохраняя её номер."""
    joined: List[tuple[int, str]] = []
    buffer = ""
    start = 0
    for n, raw in enumerate(raw_lines, 1):
        cleaned = _strip_strings_and_comments(raw)
        if not buffer:
            start = n
        if cleaned.endswith("_") and (len(cleaned) == 1 or cleaned[-2].isspace()):
            buffer += cleaned[:-1] + " "
            continue
        joined.append((start, (buffer + cleaned).strip()))
        buffer = ""
    if buffer:
        joined.append((start, buffer.strip()))
    return joined


def _split_statements(text: str) -> List[str]:
    """Разбивает `a: b` на отдельные операторы, не трогая метки."""
    if LABEL.match(text):
        return [text]
    return [part.strip() for part in text.split(":") if part.strip()]


def check(path: Path) -> List[Problem]:
    raw_lines = path.read_text(encoding="utf-8").splitlines()
    problems: List[Problem] = []

    proc_name = None
    proc_kind = None
    proc_line = 0
    seen_proc = False
    labels: set[str] = set()
    pending_goto: List[tuple[int, str]] = []
    blocks: List[tuple[int, str]] = []

    for lineno, text in _logical_lines(raw_lines):
        if not text:
            continue

        for stmt in _split_statements(text):
            m = PROC_START.match(stmt)
            if m and proc_name is None:
                proc_kind = m.group(1).split()[0].capitalize()
                proc_name = m.group(2)
                proc_line = lineno
                seen_proc = True
                labels = set()
                pending_goto = []
                continue

            if PROC_END.match(stmt):
                if proc_name is None:
                    problems.append(Problem(lineno, stmt, "End без открывающей процедуры"))
                else:
                    for goto_line, label in pending_goto:
                        if label not in labels:
                            problems.append(Problem(
                                goto_line, f"On Error GoTo {label}",
                                f"метки {label}: нет в процедуре {proc_name}"))
                    if blocks:
                        opened = ", ".join(f"{kind} (строка {ln})" for ln, kind in blocks)
                        problems.append(Problem(
                            lineno, stmt,
                            f"в {proc_name} не закрыты блоки: {opened}"))
                    blocks = []
                    proc_name = None
                continue

            if proc_name is None:
                # Уровень модуля: объявления допустимы только до первой процедуры.
                if DECL_START.match(stmt) and seen_proc:
                    problems.append(Problem(
                        lineno, stmt,
                        "объявление уровня модуля после процедуры — "
                        "перенесите в блок объявлений в начале модуля "
                        "(Excel: «Only comments may appear after End Sub»)"))
                continue

            label_match = LABEL.match(stmt)
            if label_match:
                labels.add(label_match.group(1))
                continue

            goto_match = ON_ERROR_GOTO.match(stmt)
            if goto_match and goto_match.group(1) != "0":
                pending_goto.append((lineno, goto_match.group(1)))

            if re.match(r"^Exit\s+(Sub|Function|Property)\b", stmt, re.I):
                kind = re.match(r"^Exit\s+(\w+)", stmt, re.I).group(1).capitalize()
                if proc_kind and kind != proc_kind:
                    problems.append(Problem(
                        lineno, stmt, f"Exit {kind} внутри {proc_kind} {proc_name}"))

            _track_blocks(stmt, lineno, blocks, problems)

        for word in re.findall(r"\b(\w*[A-Za-z0-9])_\b", text):
            problems.append(Problem(lineno, text, f"идентификатор {word}_ кончается подчёркиванием"))

    if proc_name is not None:
        problems.append(Problem(proc_line, proc_name, "процедура не закрыта End Sub/End Function"))

    return problems


def _track_blocks(stmt: str, lineno: int, blocks: List[tuple[int, str]],
                  problems: List[Problem]) -> None:
    """Считает открывающие и закрывающие блочные конструкции."""
    closers = {
        r"^End\s+If\b": "If",
        r"^End\s+With\b": "With",
        r"^End\s+Select\b": "Select",
        r"^Next\b": "For",
        r"^Loop\b": "Do",
        r"^Wend\b": "While",
        r"^#End\s+If\b": "#If",
    }
    for pattern, kind in closers.items():
        if re.match(pattern, stmt, re.I):
            if not blocks:
                problems.append(Problem(lineno, stmt, f"закрытие {kind} без открытия"))
            elif blocks[-1][1] != kind:
                problems.append(Problem(
                    lineno, stmt,
                    f"закрывается {kind}, а открыт {blocks[-1][1]} (строка {blocks[-1][0]})"))
                blocks.pop()
            else:
                blocks.pop()
            return

    # `If ... Then` без продолжения — блок; с продолжением — однострочник.
    if re.match(r"^If\b", stmt, re.I):
        if re.search(r"\bThen$", stmt, re.I):
            blocks.append((lineno, "If"))
        return
    if re.match(r"^#If\b", stmt, re.I):
        blocks.append((lineno, "#If"))
        return
    if re.match(r"^(For\s+Each\b|For\b)", stmt, re.I):
        blocks.append((lineno, "For"))
        return
    if re.match(r"^Do\b", stmt, re.I):
        blocks.append((lineno, "Do"))
        return
    if re.match(r"^While\b", stmt, re.I):
        blocks.append((lineno, "While"))
        return
    if re.match(r"^With\b", stmt, re.I):
        blocks.append((lineno, "With"))
        return
    if re.match(r"^Select\s+Case\b", stmt, re.I):
        blocks.append((lineno, "Select"))
        return


def main() -> int:
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_BAS
    problems = check(path)
    if not problems:
        print(f"{path.name}: замечаний нет")
        return 0
    print(f"{path.name}: найдено замечаний — {len(problems)}")
    for problem in problems:
        print(f"  строка {problem.line}: {problem.message}")
        print(f"    {problem.text[:110]}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
