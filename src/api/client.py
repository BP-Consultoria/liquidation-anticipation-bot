import httpx

from config import settings


class ApiClient:
    def __init__(self) -> None:
        self._client = httpx.AsyncClient(
            base_url=settings.api_base_url,
            headers={"Authorization": f"Bearer {settings.api_key}"},
            timeout=settings.api_timeout,
        )

    async def get(self, path: str, **kwargs) -> httpx.Response:
        return await self._client.get(path, **kwargs)

    async def post(self, path: str, **kwargs) -> httpx.Response:
        return await self._client.post(path, **kwargs)

    async def close(self) -> None:
        await self._client.aclose()
