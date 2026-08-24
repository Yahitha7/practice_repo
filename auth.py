import requests
from config import CLIENT_ID, CLIENT_SECRET, TOKEN_URL

def get_token():

    payload = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET
    }

    response = requests.post(TOKEN_URL, json=payload)
    response.raise_for_status()

    return response.json()["access_token"]
