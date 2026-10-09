from functools import lru_cache
from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", extra="ignore")

    database_url: str = Field(
        default="postgresql+psycopg://pesoctrl:pesoctrl@localhost:15432/pesoctrl",
        alias="DATABASE_URL",
    )
    cors_origins: str = Field(default="http://localhost:5174,http://127.0.0.1:5174", alias="CORS_ORIGINS")
    adb_path: str = Field(default="", alias="ADB_PATH")

    @property
    def cors_origin_list(self) -> list[str]:
        return [origin.strip() for origin in self.cors_origins.split(",") if origin.strip()]


@lru_cache
def get_settings() -> Settings:
    return Settings()
