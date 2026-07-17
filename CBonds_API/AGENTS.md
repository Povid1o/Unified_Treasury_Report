# 🤖 Инструкция для ИИ-агентов по работе с API CBonds

Этот документ содержит критически важную информацию об особенностях работы с JSON API CBonds для данного проекта. Прочтите его перед тем, как писать код для интеграции или выгрузки данных.

## 1. Базовые принципы и Аутентификация

- **Базовый URL**: Всегда используйте `https://ws.cbonds.info/services/json` (избегайте `ws2.cbonds.info`). К эндпоинту всегда добавляется параметр языка: `/?lang=rus`.
- **Метод**: Используется **POST**.
- **Тело запроса (Payload)**: Все параметры передаются в JSON-теле, включая учетные данные:
  ```json
  {
    "auth": {"login": "ВАШ_ЛОГИН", "password": "ВАШ_ПАРОЛЬ"},
    "filters": [{"field": "...", "operator": "eq", "value": "..."}],
    "quantity": {"limit": 50, "offset": 0},
    "sorting": [{"field": "date", "order": "desc"}],
    "fields": [] 
  }
  ```

## 2. Доступные эндпоинты

Учетной записи в данном проекте доступны следующие методы:
- **Параметры облигаций**: `get_emissions`, `get_emission_default`, `get_emission_guarantors`, `get_flow_new` (купоны/амортизация), `get_offert`, `get_floater_coupons`.
- **Котировки**: `get_tradings_new` (Московская биржа).
- **Индексы и Макроэкономика**: `get_index_types`, `get_index_value_new`.

## 3. Критические особенности полей и фильтров (ВАЖНО!)

API CBonds имеет неконсистентный нейминг полей в разных базах данных. Ошибки здесь приводят к молчаливому возврату некорректных данных или падению с невнятной ошибкой сервера.

### 🔴 Ошибка фильтрации по ISIN
Разные эндпоинты используют разные имена полей для ISIN-кода при фильтрации:
- В `get_emissions` и `get_tradings_new` нужно использовать `"field": "isin_code"`.
- В `get_flow_new`, `get_offert`, `get_emission_default` нужно использовать `"field": "emission_isin_code"`.
*Если вы передадите `isin_code` в `get_flow_new`, фильтр будет проигнорирован, и API вернет самые старые записи из всей базы (например, купоны за 1994 год).*

### 🔴 Названия полей (fields)
Никогда не запрашивайте поле `"name"` в `get_emissions` (например, передавая `{"field": "name"}` в массив `fields`). В базе данных нет такого поля, и API ответит технической ошибкой `Undefined array key "name"`.
- Название облигации хранится в `document_rus` или `document_eng`.
- Название эмитента хранится в `emitent_name_rus`.

### 🔴 Значения (Values)
Имя поля, в котором хранится целевое значение, отличается:
- Для **индексов** (`get_index_value_new`): поле `value`.
- Для **купонов** (`get_flow_new`): поле `cupon_sum`.
- Для **амортизации** (`get_flow_new`): поле `redemtion`.

## 4. Обработка ошибок

Если API возвращает ошибку, статус-код может оставаться 200 OK. Ошибка приходит в виде словаря внутри JSON-ответа: `{"error": {"err_no": 900000, "err_str": "..."}}`.
Всегда проверяйте наличие ключа `"error"` в корне разобранного JSON-ответа.

---

## 5. Использование готовой библиотеки `CBondsAPI`

В корне проекта написан Python-модуль `cbonds_api_test.py`, содержащий класс `CBondsAPI`. **Вам следует использовать этот класс**, а не писать сырые requests-запросы с нуля, так как в классе уже учтены все вышеописанные баги и особенности фильтрации.

### Примеры использования `CBondsAPI`:

```python
from cbonds_api_test import CBondsAPI

# Инициализация (логин и пароль берутся из переменных окружения или аргументов по умолчанию)
api = CBondsAPI()

# 1. Получить общую информацию об облигации (использует get_emissions)
bond = api.get_bond_info("RU000A10DQA8")
bond_name = bond.get("document_rus")

# 2. Получить купоны (использует get_flow_new, внутри автоматически применен emission_isin_code)
cashflows = api.get_bond_cashflows("RU000A10DQA8")
for cf in cashflows:
    date = cf.get("date")
    coupon = cf.get("cupon_sum")
    amortization = cf.get("redemtion")

# 3. Получить котировки (использует get_tradings_new)
trades = api.get_market_trades("RU000A10DQA8", date_from="2024-01-01")

# 4. Получить макроэкономические индексы (использует get_index_value_new и словарь AVAILABLE_INDICES)
# Например, кривая доходности ОФЗ
yield_5y = api.get_yield_curve("5Y") 
# Например, ключевая ставка ЦБ
key_rate = api.get_key_rate()
# Например, инфляция
core_inflation = api.get_index_values("Core_Inflation_Monthly")
```

**Резюме для агента:** Прежде чем делать сырой запрос к CBonds, загляни в класс `CBondsAPI` — скорее всего нужный метод уже реализован с правильными параметрами. Если нужно реализовать новый метод (например, для `get_emission_guarantors`), скопируй подход из `get_bond_cashflows`, обратив внимание на суффиксы для фильтра по ISIN.
