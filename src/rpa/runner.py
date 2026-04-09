from api.client import ApiClient


async def run() -> None:
    client = ApiClient()
    try:
        # TODO: implement RPA workflow steps here
        pass
    finally:
        await client.close()
