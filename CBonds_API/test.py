import requests
import pandas as pd
from cbonds_api_test import CBondsAPI

api = CBondsAPI()
ISIN = "RU000A0JS3W6"

# 1. Получаем общую информацию об облигации
bond = api.get_bond_info(ISIN)
print(bond['document_rus'])   # Название выпуска
print(bond['maturity_date'])  # Дата погашения
# --- 1. Получаем валюту и объемы ---
print(f"Валюта: {bond.get('currency_name')}")
print(f"Объем в обращении: {bond.get('outstanding_volume')}")
print(f"Объем анонсированный: {bond.get('announced_volume_new')}")
print(f"Объем размещенный: {bond.get('placed_volume_new')}")

# --- 2. Получаем G-спред и дюрацию (например, за последние дни) ---
trades = api.get_market_trades(ISIN, limit=1) # Берем последний торговый день
if trades:
    last_trade = trades[0]
    print(f"\nДанные торгов на {last_trade.get('date')}:")
    print(f"G-спред: {last_trade.get('g_spread')}")
    print(f"Дюрация (dur): {last_trade.get('dur')}")
    print(f"Модиф. дюрация (dur_mod): {last_trade.get('dur_mod')}")

# 3. Получаем купоны (денежные потоки)
coupons = api.get_bond_cashflows(ISIN)
for coupon in coupons[:10]: # Выводим только первые 10 выплат для компактности
    print(f"Дата: {coupon.get('date')}, Купон: {coupon.get('cupon_sum')}, Амортизация: {coupon.get('redemtion')}")


print("\nbond.keys():\n", bond.keys())
print("\nlast_trade.keys():\n",last_trade.keys())