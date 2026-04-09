import asyncio

from config import settings
from rpa.runner import run


def main() -> None:
    asyncio.run(run())


if __name__ == "__main__":
    main()
