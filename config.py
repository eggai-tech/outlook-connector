"""Layered service configuration.

Per `the spec <docs/implementation.md#configuration>`:

- **Structural config** lives in a YAML file (path from ``CONFIG_FILE``,
  default ``config.yaml``): mailbox list, poll interval, bus connection, Azure
  ``tenant_id``/``client_id``, log level.
- **Secrets** come from environment variables only — ``client_secret`` is never
  read from the YAML file. The :class:`_StructuralYamlSource` actively rejects a
  ``client_secret`` written into the file so the secret cannot leak there.

Validation is strict so startup can *fail fast*: a missing secret, a missing
structural field, an empty mailbox list, or an unknown transport all raise.
"""

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


class SecretInConfigError(Exception):
    """Raised when a secret (``client_secret``) is found in the YAML config file."""


class BusConfig(BaseModel):
    """EggAI bus connection settings."""

    transport: Literal["kafka", "redis", "inmemory"] = "kafka"
    # Broker URL / bootstrap servers. None falls back to the transport default.
    broker_url: str | None = None
    channel: str = "emails"


class AzureConfig(BaseModel):
    """Azure app-registration identifiers (the secret is supplied via env)."""

    tenant_id: str
    client_id: str


class _StructuralYamlSource(YamlConfigSettingsSource):
    """A YAML source that refuses to carry the ``client_secret``."""

    def __call__(self) -> dict:
        data = super().__call__()
        if "client_secret" in data:
            raise SecretInConfigError(
                "client_secret must not appear in the config file; "
                "supply it via the client_secret environment variable."
            )
        return data


class Settings(BaseSettings):
    """The connector's validated configuration."""

    model_config = SettingsConfigDict(
        env_nested_delimiter="__",
        extra="forbid",
    )

    mailboxes: list[str] = Field(min_length=1)
    poll_interval_seconds: float = Field(default=60.0, gt=0)
    # Optional manual cursor seed (ISO 8601). When unset, each mailbox cursor
    # starts at process-start "now" so only mail received after startup is
    # bridged. Set this to backfill from a known point in time instead.
    initial_cursor: datetime | None = None
    log_level: str = "INFO"
    bus: BusConfig = Field(default_factory=BusConfig)
    azure: AzureConfig

    # Secret — environment variable only, never the YAML file.
    client_secret: str

    @classmethod
    def settings_customise_sources(
        cls,
        settings_cls: type[BaseSettings],
        init_settings: PydanticBaseSettingsSource,
        env_settings: PydanticBaseSettingsSource,
        dotenv_settings: PydanticBaseSettingsSource,
        file_secret_settings: PydanticBaseSettingsSource,
    ) -> tuple[PydanticBaseSettingsSource, ...]:
        yaml_file = os.getenv(CONFIG_FILE_ENV, DEFAULT_CONFIG_FILE)
        yaml_source = _StructuralYamlSource(settings_cls, yaml_file=yaml_file)
        # Precedence (first wins): explicit init args, then environment (the
        # secret), then the structural YAML file.
        return (init_settings, env_settings, yaml_source)


def load_settings() -> Settings:
    """Load and validate configuration; raises on any invalid/missing value."""
    yaml_file = os.getenv(CONFIG_FILE_ENV, DEFAULT_CONFIG_FILE)
    if not os.path.exists(yaml_file):
        raise FileNotFoundError(f"Config file not found: {yaml_file}")
    return Settings()
