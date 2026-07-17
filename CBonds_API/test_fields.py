import requests, json, os
from dotenv import load_dotenv

load_dotenv()
auth = {'login': os.getenv('CBONDS_LOGIN'), 'password': os.getenv('CBONDS_PASSWORD')}

url = 'https://ws.cbonds.info/services/json/get_tradings_new/?lang=rus'
payload = {
    'auth': auth,
    'filters': [{'field': 'isin_code', 'operator': 'eq', 'value': 'RU000A10DQA8'}],
    'quantity': {'limit': 1, 'offset': 0},
    'sorting': [{'field': 'date', 'order': 'desc'}],
    'fields': []
}
resp = requests.post(url, json=payload, timeout=15)
data = resp.json()
if 'items' in data and len(data['items']) > 0:
    print('Fields in get_tradings_new:', list(data['items'][0].keys()))
else:
    print('No items found or error:', data)
