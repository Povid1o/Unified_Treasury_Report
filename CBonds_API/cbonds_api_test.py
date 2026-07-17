import os
import requests
import json
from datetime import date
from typing import Optional, List, Dict, Union
from dotenv import load_dotenv

# Загружаем переменные окружения из файла .env (если он существует)
load_dotenv()


# ── Конфигурация ──────────────────────────────────────────────────────────────
BASE_URL = "https://ws.cbonds.info/services/json"
TIMEOUT = 30

# Доступные индексы (для метода get_index_values)
AVAILABLE_INDICES = {
    # RUONIA (Россия)
    "RUONIA_1M": 72063,
    "RUONIA_3M": 72065,
    "RUONIA_6M": 72067,
    "RUONIA_Index": 72061,
    "RUONIA_Rate": 1829,

    # Russia Government Bond Yield Curve
    "RUB_Yield_Curve_10Y": 24431,
    "RUB_Yield_Curve_15Y": 24433,
    "RUB_Yield_Curve_1M": 24411,
    "RUB_Yield_Curve_1W": 24407,
    "RUB_Yield_Curve_1Y": 24421,
    "RUB_Yield_Curve_20Y": 24435,
    "RUB_Yield_Curve_25Y": 24437,
    "RUB_Yield_Curve_2M": 24413,
    "RUB_Yield_Curve_2W": 24409,
    "RUB_Yield_Curve_2Y": 24423,
    "RUB_Yield_Curve_30Y": 72095,
    "RUB_Yield_Curve_3M": 24415,
    "RUB_Yield_Curve_3Y": 24425,
    "RUB_Yield_Curve_4M": 24417,
    "RUB_Yield_Curve_4Y": 72069,
    "RUB_Yield_Curve_5Y": 24427,
    "RUB_Yield_Curve_6M": 79369,
    "RUB_Yield_Curve_7Y": 24429,
    "RUB_Yield_Curve_8Y": 72071,
    "RUB_Yield_Curve_9M": 24419,
    "RUB_Yield_Curve_9Y": 72073,

    # Макроэкономика и Государственные показатели
    "Core_Inflation_Monthly": 51281, # Базовый уровень инфляции
    "Budget_Balance_Monthly": 51317, # Сальдо доходов бюджета
    "CB_Key_Rate": 21755,            # Ключевая ставка ЦБ РФ
    "Inflation_Rate_Yearly": 37639,  # Уровень инфляции в годовом выражении
    "RGBI": 1599,                    # Индекс гособлигаций RGBI

    # Курсы валют FOREX
    "CNY_RUB": 75093,
    "EUR_RUB": 72375,
    "GBP_RUB": 75049,
    "INR_RUB": 75473,
    "JPY_RUB": 75151,
    "USD_RUB": 40329,

    # Нефть
    "Brent": 624,
    "Urals": 1594,
}


class CBondsAPI:
    """Полный клиент для работы с API CBonds для всех доступных эндпоинтов."""
    
    def __init__(self, login: str = None, password: str = None):
        self.login = login or os.getenv("CBONDS_LOGIN")
        self.password = password or os.getenv("CBONDS_PASSWORD")
        
        if not self.login or not self.password:
            raise ValueError(
                "API credentials (login/password) are not set. "
                "Please set CBONDS_LOGIN and CBONDS_PASSWORD environment variables or pass them to CBondsAPI()."
            )
            
        self.auth = {"login": self.login, "password": self.password}

    def _post(self, operation: str, filters: list = None, fields: list = None,
              limit: int = 50, offset: int = 0, sorting: list = None) -> dict:
        """Базовый метод для отправки POST запроса."""
        url = f"{BASE_URL}/{operation}/?lang=rus"
        payload = {
            "auth": self.auth,
            "filters": filters or [],
            "quantity": {"limit": limit, "offset": offset},
            "sorting": sorting or [],
            "fields": fields or []
        }
        
        resp = requests.post(url, json=payload, timeout=TIMEOUT)
        
        if resp.status_code == 301:
            raise RuntimeError("HTTP 301: используется HTTP вместо HTTPS")
        if resp.status_code == 403:
            raise RuntimeError("HTTP 403: неверный логин/пароль или нет доступа к операции")
        if resp.status_code == 500:
            raise RuntimeError("HTTP 500: синтаксическая ошибка в запросе или нет доступа к операции")
        if resp.status_code == 504:
            raise RuntimeError("HTTP 504: таймаут сервера")
        resp.raise_for_status()

        data = resp.json()
        
        if isinstance(data, dict) and "error" in data:
            err = data["error"]
            err_msg = err.get("err_str", str(err)) if isinstance(err, dict) else str(err)
            raise RuntimeError(f"API error: {err_msg}")
            
        return data

    # =====================================================================
    # 1. ПАРАМЕТРЫ ОБЛИГАЦИЙ (ЭМИССИИ)
    # =====================================================================

    def get_bond_info(self, isin: str) -> Dict:
        """
        Получить параметры конкретной облигации по ISIN.
        Использует эндпоинт get_emissions.
        """
        data = self._post(
            operation="get_emissions",
            filters=[{"field": "isin_code", "operator": "eq", "value": isin}],
            limit=1
        )
        items = data.get("items", [])
        return items[0] if items else {}

    def search_bonds(self, isin_prefix: str = "RU000A", currency: str = "RUB", limit: int = 50) -> List[Dict]:
        """
        Поиск облигаций по префиксу ISIN и валюте.
        Например, для ОФЗ isin_prefix="RU000A", currency="RUB".
        """
        filters = []
        if isin_prefix:
            filters.append({"field": "isin_code", "operator": "bw", "value": isin_prefix})
        if currency:
            filters.append({"field": "currency_name", "operator": "eq", "value": currency})
            
        data = self._post(
            operation="get_emissions",
            filters=filters,
            limit=limit,
            sorting=[{"field": "id", "order": "asc"}]
        )
        return data.get("items", [])

    def get_bond_cashflows(self, isin: str, limit: int = 100) -> List[Dict]:
        """
        Получить денежные потоки (купоны и амортизация) по облигации (get_flow_new).
        """
        data = self._post(
            operation="get_flow_new",
            filters=[{"field": "emission_isin_code", "operator": "eq", "value": isin}],
            limit=limit,
            sorting=[{"field": "date", "order": "asc"}]
        )
        return data.get("items", [])

    def get_bond_offers(self, isin: str, limit: int = 50) -> List[Dict]:
        """Получить оферты (put/call опционы) по облигации (get_offert)."""
        data = self._post(
            operation="get_offert",
            filters=[{"field": "emission_isin_code", "operator": "eq", "value": isin}],
            limit=limit
        )
        return data.get("items", [])

    def get_bond_defaults(self, isin: str, limit: int = 50) -> List[Dict]:
        """Получить информацию о дефолтах по облигации (get_emission_default)."""
        data = self._post(
            operation="get_emission_default",
            filters=[{"field": "emission_isin_code", "operator": "eq", "value": isin}],
            limit=limit
        )
        return data.get("items", [])


    # =====================================================================
    # 2. КОТИРОВКИ ОБЛИГАЦИЙ
    # =====================================================================

    def get_market_trades(self, isin: str, date_from: Optional[str] = None, 
                          date_to: Optional[str] = None, limit: int = 100,
                          offset: int = 0) -> List[Dict]:
        """
        Получить котировки и сделки с Московской биржи (get_tradings_new).
        :param isin: ISIN облигации.
        :param date_from: Начальная дата (YYYY-MM-DD).
        :param date_to: Конечная дата (YYYY-MM-DD).
        :param limit: Лимит количества записей (по умолчанию 100).
        :param offset: Смещение для пагинации (по умолчанию 0).
        """
        filters = [{"field": "isin_code", "operator": "eq", "value": isin}]
        if date_from:
            filters.append({"field": "date", "operator": "ge", "value": date_from})
        if date_to:
            filters.append({"field": "date", "operator": "le", "value": date_to})
            
        data = self._post(
            operation="get_tradings_new",
            filters=filters,
            limit=limit,
            offset=offset,
            sorting=[{"field": "date", "order": "desc"}]
        )
        return data.get("items", [])

    def get_realtime_trades(self, isin: str, limit: int = 100, offset: int = 0) -> List[Dict]:
        """
        Получить реалтайм котировки/сделки по облигации (get_tradings_realtime).
        ВНИМАНИЕ: Для доступа к этому эндпоинту требуется продвинутая подписка.
        """
        filters = [{"field": "isin_code", "operator": "eq", "value": isin}]
        data = self._post(
            operation="get_tradings_realtime",
            filters=filters,
            limit=limit,
            offset=offset
        )
        return data.get("items", [])

    def get_realtime_stocks_trades(self, isin: str, limit: int = 100, offset: int = 0) -> List[Dict]:
        """
        Получить реалтайм котировки по акциям (get_tradings_stocks_full_realtime_new).
        ВНИМАНИЕ: Для доступа к этому эндпоинту требуется продвинутая подписка.
        """
        filters = [{"field": "isin_code", "operator": "eq", "value": isin}]
        data = self._post(
            operation="get_tradings_stocks_full_realtime_new",
            filters=filters,
            limit=limit,
            offset=offset
        )
        return data.get("items", [])

    # =====================================================================
    # 3. ИНДЕКСЫ И МАКРОЭКОНОМИКА
    # =====================================================================

    def get_index_values(self, index_id: Union[int, str], 
                         date_from: Optional[str] = None, 
                         date_to: Optional[str] = None, 
                         limit: int = 50,
                         offset: int = 0) -> List[Dict]:
        """
        Получает значения индекса по его ID или текстовому ключу.
        :param limit: Лимит количества записей (по умолчанию 50).
        :param offset: Смещение для пагинации (по умолчанию 0).
        """
        if isinstance(index_id, str) and index_id in AVAILABLE_INDICES:
            index_id = AVAILABLE_INDICES[index_id]
            
        filters = [{"field": "type_id", "operator": "eq", "value": str(index_id)}]
        if date_from:
            filters.append({"field": "date", "operator": "ge", "value": date_from})
        if date_to:
            filters.append({"field": "date", "operator": "le", "value": date_to})

        data = self._post(
            operation="get_index_value_new", 
            filters=filters, 
            limit=limit,
            sorting=[{"field": "date", "order": "desc"}]
        )
        return data.get("items", [])

    def get_ruonia_rate(self, limit: int = 5) -> List[Dict]:
        return self.get_index_values("RUONIA_Rate", limit=limit)

    def get_key_rate(self, limit: int = 5) -> List[Dict]:
        return self.get_index_values("CB_Key_Rate", limit=limit)

    def get_forex_rate(self, currency_pair: str, limit: int = 5) -> List[Dict]:
        if currency_pair not in AVAILABLE_INDICES:
            raise ValueError(f"Валютная пара {currency_pair} не найдена.")
        return self.get_index_values(currency_pair, limit=limit)

    def get_yield_curve(self, maturity: str, limit: int = 5) -> List[Dict]:
        key = f"RUB_Yield_Curve_{maturity}"
        if key not in AVAILABLE_INDICES:
            raise ValueError(f"Срок {maturity} не найден. Доступные: 1W, 1M, 1Y, 5Y, 10Y и т.д.")
        return self.get_index_values(key, limit=limit)


# =====================================================================
# ТЕСТЫ / ПРИМЕРЫ ИСПОЛЬЗОВАНИЯ
# =====================================================================

def run_tests():
    api = CBondsAPI()
    print("Подключение к CBonds API инициализировано.\n")
    
    try:
        # 1. Тест получения информации по облигации (по ISIN)
        test_isin = "RU000A0JWSQ7" # Пример ISIN
        print(f"--- Тест: Параметры облигации {test_isin} ---")
        bond_info = api.get_bond_info(test_isin)
        if bond_info:
            print(f"Название: {bond_info.get('document_rus')}")
            print(f"Эмитент: {bond_info.get('emitent_name_rus')}")
            print(f"Дата погашения: {bond_info.get('maturity_date')}")
            print(f"Ставка купона: {bond_info.get('emission_coupon_rate')}%")
        else:
            print("Облигация не найдена.")
        print()
            
        # 2. Тест получения купонов по облигации
        print(f"--- Тест: График выплат (купоны) {test_isin} ---")
        flows = api.get_bond_cashflows(test_isin, limit=3)
        if flows:
            for flow in flows:
                print(f"Дата выплаты: {flow.get('date')}, Купон: {flow.get('cupon_sum')}, Амортизация: {flow.get('redemtion')}")
        else:
            print("Купоны не найдены.")
        print()
        
        # 3. Тест получения индексов
        print("--- Тест: Ставка RUONIA ---")
        ruonia = api.get_ruonia_rate(limit=3)
        for r in ruonia:
             print(f"Дата: {r.get('date')}, Значение: {r.get('value')}")

    except Exception as e:
        print(f"Ошибка при выполнении запроса: {e}")

if __name__ == "__main__":
    run_tests()
