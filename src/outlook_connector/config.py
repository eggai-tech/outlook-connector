"""Layered service configuration.

Per `the spec <docs/DESIGN.md#configuration>`:

- **Structural config** lives in a YAML file (path from ``CONFIG_FILE``,
  default ``config.yaml``): mailbox list, poll interval, bus connection, log level.

- **No secrets in the config file. They can only come from env.**

Validation is strict so startup can fail quickly.
"""

from functools import lru_cache
import os
from datetime import UTC, datetime
from typing import Literal

from pydantic import BaseModel, Field, field_validator
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
    # Redis only: approximate cap on the channel's stream length, applied as
    # XADD MAXLEN ~. Redis streams keep every entry in RAM, acked or not, so
    # without this the stream grows until the server OOMs. Trimming is by
    # count, oldest first, regardless of whether consumers have read them.
    # None = unbounded. Ignored by kafka and inmemory.
    max_len: int | None = Field(default=None, gt=0)


class Settings(BaseSettings):
    """The connector's validated configuration."""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_nested_delimiter="__",
        extra="ignore",
    )

    mailbox: str = Field(min_length=1)
    # Folder to poll (well-known Graph name like "inbox", or a display name).
    source_folder: str = Field(default="inbox", min_length=1)
    poll_interval_seconds: float = Field(default=60.0, gt=0)
    # At most this many messages fetched + published per poll cycle, oldest
    # first; a backlog (or a restart's re-emit of the whole folder) drains
    # across cycles instead of in one unbounded burst of full-message fetches.
    # null = no bound — only sensible for small, drained folders.
    batch_max_messages: int | None = Field(default=100, gt=0)
    # Attachments larger than this are published as metadata only (body=None).
    # Mind your broker's message-size limit (kafka defaults to ~1MB).
    # None = no cap.
    max_attachment_bytes: int | None = Field(default=8 * 1024 * 1024, gt=0)
    # When true the connector makes one extra Graph call per message
    # (``/messages/{id}/$value``) and publishes the full RFC 822 message —
    # headers, bodies and every attachment base64-encoded again — as
    # ``email.mime_content``. Attachments stay on the wire as well, so the
    # event roughly doubles in size. NOT bounded by ``max_attachment_bytes``:
    # mind the broker's message-size limit (kafka defaults to ~1 MiB).
    include_mime_content: bool = False
    # Optional lower bound (ISO 8601): mail received before this instant is
    # never listed or published. By default the whole source folder is
    # bridged — every rescan re-emits anything still present, and consumers
    # dedupe. Set this to keep an old, full folder from being backfilled.
    ignore_received_before: datetime | None = None
    # Port for the HTTP health/status endpoint (GET /health), bound on all
    # interfaces. null disables the endpoint entirely.
    health_port: int | None = Field(default=8000, gt=0, le=65535)

    @field_validator("ignore_received_before")
    @classmethod
    def _bound_must_be_aware(cls, value: datetime | None) -> datetime | None:
        """Coerce a naive timestamp to UTC.

        Graph timestamps are timezone-aware; a naive bound would compare
        wrongly (or raise) downstream. A bare ISO string like
        "2026-06-26T08:00:00" is treated as UTC.
        """
        if value is not None and value.tzinfo is None:
            return value.replace(tzinfo=UTC)
        return value
    log_level: str = "INFO"
    bus: BusConfig = Field(default_factory=BusConfig)

    # Not allowed in config file. Env only. min_length rejects the empty
    # strings a compose file with `${AZURE_...:-}` defaults would pass in.
    azure_tenant_id: str = Field(min_length=1)
    azure_client_id: str = Field(min_length=1)
    azure_client_secret: str = Field(min_length=1)

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
