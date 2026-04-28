"""Environment-driven configuration for the audit service."""
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", extra="ignore")

    # Shared HS256 secret with dispatch — must match on both Railway projects
    audit_service_jwt_secret: str

    # Operational limits (see contract §4.1, §9)
    max_processing_seconds: int = 120
    max_concurrent_audits: int = 3

    # Observability
    log_level: str = "INFO"

    # Service metadata (returned by /v1/health)
    service_version: str = "1.0.0"


settings = Settings()
