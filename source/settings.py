from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    server_name: str
    project_name: str
    token_auth: str
    model_config = SettingsConfigDict(env_file='../resources/credits.env')


def get_settings():
    return Settings()

