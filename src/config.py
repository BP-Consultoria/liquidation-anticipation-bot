from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    model_config = {"env_file": ".env", "env_file_encoding": "utf-8"}

    app_name: str = "liquidation-anticipation-bot"
    debug: bool = False

    # API settings
    api_base_url: str = ""
    api_key: str = ""
    api_timeout: int = 30

    # Database settings (anticipation_db)
    db_driver: str = "mssql+pyodbc"
    db_host: str = "localhost"
    db_port: int = 1433
    db_name: str = ""
    db_user: str = ""
    db_password: str = ""
    db_odbc_driver: str = "ODBC Driver 17 for SQL Server"

    # Database settings (securitização)
    sec_host: str = ""
    sec_port: int = 1433
    sec_name: str = ""
    sec_user: str = ""
    sec_password: str = ""
    sec_odbc_driver: str = "ODBC Driver 17 for SQL Server"

    @property
    def database_url(self) -> str:
        if self.db_driver.startswith("mssql"):
            params = (
                f"DRIVER={{{self.db_odbc_driver}}};"
                f"SERVER={self.db_host},{self.db_port};"
                f"DATABASE={self.db_name};"
                f"UID={self.db_user};"
                f"PWD={self.db_password}"
            )
            return f"{self.db_driver}:///?odbc_connect={params}"
        return (
            f"{self.db_driver}://{self.db_user}:{self.db_password}"
            f"@{self.db_host}:{self.db_port}/{self.db_name}"
        )


settings = Settings()
