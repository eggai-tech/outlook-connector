"""Layered service configuration.

Per `the spec <docs/DESIGN.md#configuration>`:

- **Structural config** lives in a YAML file (path from ``CONFIG_FILE``,
  default ``config.yaml``): mailbox list, poll interval, bus connection, log level.

- **No secrets in the config file. They can only come from env.**

Validation is strict so startup can fail quickly.
"""

from functools import lru_cache
import os
from datetime import datetime
from typing import Literal

from pydantic import BaseModel, Field
from pydantic_settings import (
    BaseSettings,
    PydanticBaseSettingsSource,
    SettingsConfigDict,
    YamlConfigSettingsSource,
)

CONFIG_FILE_ENV = "CONFIG_FILE"
DEFAULT_CONFIG_FILE = "config.yaml"

# these keys can only come from the environment
# stop if any of them is present in the config fail
_ENV_ONLY_KEYS = [
    "azure_tenant_id",
    "azure_client_id",
    "azure_client_secret",
]


class BusConfig(BaseModel):
    """EggAI bus connection settings."""

    transport: Literal["kafka", "redis", "inmemory"] = "kafka"
    # Broker URL / bootstrap servers. None falls back to the transport default.
    broker_url: str | None = None
    channel: str = "emails"


class Settings(BaseSettings):
    """The connector's validated configuration."""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_nested_delimiter="__",
        extra="ignore",
    )

    mailbox: str = Field(min_length=1)
    poll_interval_seconds: float = Field(default=60.0, gt=0)
    # Optional manual cursor seed (ISO 8601). When unset, each mailbox cursor
    # starts at process-start "now" so only mail received after startup is
    # bridged. Set this to backfill from a known point in time instead.
    initial_cursor: datetime | None = None
    log_level: str = "INFO"
    bus: BusConfig = Field(default_factory=BusConfig)

    # Not allowed in config file. Env only.
    azure_tenant_id: str
    azure_client_id: str
    azure_client_secret: str

    @classmethod
    def settings_customise_sources(
        cls,
        settings_cls: type[BaseSettings],
        init_settings: PydanticBaseSettingsSource,
        env_settings: PydanticBaseSettingsSource,
        dotenv_settings: PydanticBaseSettingsSource,
        file_secret_settings: PydanticBaseSettingsSource,
    ) -> tuple[PydanticBaseSettingsSource, ...]:
        """
        Configure where we get settings from: environment and config.yaml
        """
        yaml_file = os.getenv(CONFIG_FILE_ENV, DEFAULT_CONFIG_FILE)
        yaml_source = YamlConfigSettingsSource(settings_cls, yaml_file=yaml_file)
        return (init_settings, env_settings, dotenv_settings, yaml_source)


def _reject_env_only_keys(yaml_file: str) -> None:
    """Fail loudly if the config file carries Azure credentials.

    Without this the keys would be *silently* ignored (environment beats the
    file), so a stale ``azure:`` block would look like it was still in effect.
    """
    document = YamlConfigSettingsSource(Settings, yaml_file=yaml_file)()
    present = [key for key in _ENV_ONLY_KEYS if key in document]
    if present:
        raise ValueError(
            f"{yaml_file} is not allowed to contain {', '.join(present)}. Env only."
        )


@lru_cache
def get_settings() -> Settings:
    """Load and validate configuration; raises on any invalid/missing value."""
    yaml_file = os.getenv(CONFIG_FILE_ENV, DEFAULT_CONFIG_FILE)
    if not os.path.exists(yaml_file):
        raise FileNotFoundError(f"Config file not found: {yaml_file}")
    _reject_env_only_keys(yaml_file)
    return Settings()
