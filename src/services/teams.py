import os
from typing import Dict
from azure.identity import UsernamePasswordCredential
from dotenv import load_dotenv

load_dotenv()

def get_teams_headers() -> Dict[str, str]:
    client_id = os.getenv("MS_CLIENT_ID", "")
    username = os.getenv("MS_USERNAME", "")
    password = os.getenv("MS_PASSWORD", "")
    base_url = os.getenv("MS_GRAPH_BASE_URL") or "https://graph.microsoft.com"

    if not client_id or not username or not password:
        raise ValueError("MS_CLIENT_ID, MS_USERNAME ou MS_PASSWORD não definidos no .env")

    credential = UsernamePasswordCredential(
        client_id=client_id,
        username=username,
        password=password,
    )

    scope = f"{base_url}/.default"
    token = credential.get_token(scope).token

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    return headers