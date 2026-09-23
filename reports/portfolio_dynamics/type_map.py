"""Справочник «портфель -> тип» в отдельном JSON-файле.

Зачем. Угадывание типа по коду портфеля (префиксы, подстроки) ошибается, а
поправить его можно было только правилами и исключениями в настройках —
длинными строками через запятую. Здесь разметка лежит списком, который
правится руками за минуту:

    {
      "TSS":      ["OFZ_PD", "OFZ_PK"],
      "AFS":      ["AFS_TR_RUR", "AFS_*"],
      "HTM":      ["OFZ_HTM"],
      "HTM_KUAP": ["HTM_KUAP_CORE"],
      "_не_размечены": {"NEW_CODE": "AFS?"}
    }

Ключ — тип, значение — коды портфелей. В коде можно ставить «*» и «?» (как
в маске файлов): точный код всегда важнее маски. Ключи, начинающиеся с «_»,
типами не считаются: в «_не_размечены» отчёт сам дописывает портфели,
которых в файле нет, вместе с угаданным типом — остаётся перенести код в
нужный список.

Файл — главный источник: его разметка перебивает и угадывание, и настройку
«Разметка отдельных портфелей». Перечень типов в отчёте — это ключи файла:
тип, которого в файле нет, в dim_portfolio не проставляется.

HTM_KUAP — подтип HTM: объём HTM везде считается вместе с КУАП (настройка
«Вложенность типов»), отдельно КУАП виден только там, где нужна
конкретизация, — в своей строке свода и в своём подлимите.
"""
import fnmatch
import json
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

BASE_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE_DIR))

import config  # noqa: E402
from reports.portfolio_dynamics.etl import (  # noqa: E402
    PortfolioDynamicsError, canonical_type, logger,
)

# Порядок ключей в новом файле: три типа и подтип HTM.
DEFAULT_TYPES = ("TSS", "AFS", "HTM", "HTM_KUAP")
UNMAPPED_KEY = "_не_размечены"
README_KEY = "_как_заполнять"
README = [
    "Ключ — тип портфеля, значение — список кодов портфелей (как в выгрузке).",
    "Можно маски: «AFS_*» — все коды, начинающиеся с AFS_. Точный код важнее маски.",
    "HTM_KUAP — подтип HTM: в объём и лимит HTM он входит сам, отдельно виден в своей строке.",
    "«_не_размечены» отчёт заполняет сам: код и угаданный тип. Перенесите код в нужный список.",
    "Портфель, указанный здесь, получает этот тип всегда — поверх правил и настроек.",
]


def normalize_type(name) -> str:
    """«htm-kuap» / «HTM KUAP» / «TTS» -> каноническое написание типа."""
    text = str(name).strip().upper().replace("-", "_").replace(" ", "_")
    return canonical_type(text)


def _is_pattern(code: str) -> bool:
    return any(ch in code for ch in "*?[")


@dataclass
class TypeMap:
    path: Optional[Path]
    types: List[str] = field(default_factory=lambda: list(DEFAULT_TYPES))
    exact: Dict[str, str] = field(default_factory=dict)
    patterns: List[Tuple[str, str]] = field(default_factory=list)
    unmapped: Dict[str, str] = field(default_factory=dict)

    def explicit(self, code) -> Optional[str]:
        """Тип, заданный в файле: точный код, затем самая длинная подходящая маска."""
        key = str(code).strip().upper()
        if key in self.exact:
            return self.exact[key]
        matches = [(p, t) for p, t in self.patterns if fnmatch.fnmatchcase(key, p)]
        if not matches:
            return None
        # Самая конкретная маска — с наибольшим числом обычных символов.
        return max(matches, key=lambda m: len(m[0].replace("*", "").replace("?", "")))[1]


def types_file() -> Path:
    return Path(config.PORTFOLIO_DYNAMICS_TYPES_FILE)


_CACHE: Dict[str, Tuple[float, TypeMap]] = {}


def forget_cache() -> None:
    _CACHE.clear()


def load(path: Optional[Path] = None) -> TypeMap:
    """Читает справочник. Файла нет — пустой справочник с четырьмя типами."""
    path = Path(path) if path is not None else types_file()
    if not path.is_file():
        return TypeMap(path=path)
    mtime = path.stat().st_mtime
    cached = _CACHE.get(str(path))
    if cached and cached[0] == mtime:
        return cached[1]
    parsed = _parse(path)
    _CACHE[str(path)] = (mtime, parsed)
    return parsed


def _parse(path: Path) -> TypeMap:
    try:
        raw = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        # Молча угадывать дальше нельзя: человек думает, что разметка
        # действует, а отчёт её не видит.
        raise PortfolioDynamicsError(
            f"Файл разметки типов {path} не читается: {exc.msg} (строка {exc.lineno}, "
            f"символ {exc.colno}). Частые причины — лишняя запятая после последнего "
            "кода в списке или кавычки «ёлочки» вместо \"прямых\"."
        ) from exc
    except OSError as exc:
        raise PortfolioDynamicsError(f"Файл разметки типов {path} не открылся: {exc}") from exc
    if not isinstance(raw, dict):
        raise PortfolioDynamicsError(
            f"Файл разметки типов {path}: ожидается объект {{\"ТИП\": [коды]}}, "
            f"а там {type(raw).__name__}."
        )

    result = TypeMap(path=path, types=[])
    seen: Dict[str, str] = {}
    conflicts = []
    for key, codes in raw.items():
        if str(key).startswith("_"):
            if key == UNMAPPED_KEY and isinstance(codes, dict):
                result.unmapped = {str(c).strip().upper(): str(t) for c, t in codes.items()}
            continue
        portfolio_type = normalize_type(key)
        if not portfolio_type:
            continue
        if portfolio_type not in result.types:
            result.types.append(portfolio_type)
        if codes is None:
            continue
        if isinstance(codes, str):
            codes = [codes]
        if not isinstance(codes, list):
            raise PortfolioDynamicsError(
                f"Файл разметки типов {path}: у типа {key} должен быть список кодов "
                f"в квадратных скобках, а там {type(codes).__name__}."
            )
        for code in codes:
            code = str(code).strip().upper()
            if not code:
                continue
            if code in seen and seen[code] != portfolio_type:
                conflicts.append(f"{code}: {seen[code]} и {portfolio_type}")
                # Из двух вариантов берём подтип: HTM_KUAP конкретнее HTM.
                if len(portfolio_type) <= len(seen[code]):
                    continue
            seen[code] = portfolio_type
            if _is_pattern(code):
                result.patterns = [(p, t) for p, t in result.patterns if p != code]
                result.patterns.append((code, portfolio_type))
            else:
                result.exact[code] = portfolio_type
    if conflicts:
        logger.warning(
            "Файл разметки типов: коды указаны у нескольких типов (%s) — взят более "
            "конкретный. Оставьте каждый код в одном списке.", "; ".join(conflicts),
        )
    if not result.types:
        result.types = list(DEFAULT_TYPES)
    return result


def registry() -> List[str]:
    """Типы, существующие в отчёте: ключи файла разметки (по умолчанию четыре)."""
    return list(load().types)


def explicit_type(code) -> Optional[str]:
    return load().explicit(code)


def _template() -> dict:
    data: dict = {README_KEY: README}
    for portfolio_type in DEFAULT_TYPES:
        data[portfolio_type] = []
    data[UNMAPPED_KEY] = {}
    return data


def _dump(data: dict) -> str:
    """JSON с одним кодом на строку — чтобы править и сравнивать было удобно."""
    return json.dumps(data, ensure_ascii=False, indent=2) + "\n"


def record_unmapped(guesses: Dict[str, Optional[str]], path: Optional[Path] = None) -> List[str]:
    """Дописывает в «_не_размечены» коды, которых в файле нет. Возвращает новые.

    Файла нет — создаёт его по шаблону: так справочник появляется сам, с
    полным списком портфелей, которые остаётся разнести по типам. Коды, уже
    перенесённые человеком в список типа, из «_не_размечены» убираются.
    """
    path = Path(path) if path is not None else types_file()
    current = load(path)
    if path.is_file():
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    else:
        data = _template()

    pending = data.get(UNMAPPED_KEY)
    if not isinstance(pending, dict):
        pending = {}
    before = dict(pending)

    for code in list(pending):
        if current.explicit(code) is not None:
            del pending[code]
    added = []
    for code, guessed in guesses.items():
        key = str(code).strip().upper()
        if not key or current.explicit(key) is not None:
            continue
        hint = (guessed + "?") if guessed else "?"
        if key not in pending:
            added.append(key)
        pending[key] = hint

    if pending == before and path.is_file():
        return []
    data.pop(UNMAPPED_KEY, None)
    data[UNMAPPED_KEY] = dict(sorted(pending.items()))
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(_dump(data), encoding="utf-8")
    except OSError as exc:
        logger.warning("Не удалось обновить файл разметки типов %s: %s", path, exc)
        return []
    forget_cache()
    return added


def unknown_types(types: Iterable) -> List[str]:
    """Типы из списка, которых нет в справочнике."""
    allowed = set(registry())
    return sorted({normalize_type(t) for t in types
                   if str(t or "").strip() and normalize_type(t) not in allowed})

