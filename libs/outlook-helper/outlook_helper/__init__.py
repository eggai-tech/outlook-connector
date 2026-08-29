"""outlook-helper: work with M365 email through the Microsoft Graph API."""

from outlook_helper.attachments import Attachment
from outlook_helper.auth import (
    ClientSecretCredential,
    Credential,
    DeviceCodeCredential,
)
from outlook_helper.client import OutlookClient
from outlook_helper.config import (
    AppOnlyConfig,
    CertificateConfig,
    DelegatedConfig,
)
from outlook_helper.exceptions import GraphError
from outlook_helper.schemas import (
    InternetMessageHeader,
    OutlookAttachment,
    OutlookAttachmentMeta,
    OutlookBody,
    EmailAddress,
    OutlookFolder,
    OutlookMessage,
)

__all__ = [
    "GraphError",
    "DelegatedConfig",
    "AppOnlyConfig",
    "CertificateConfig",
    "Credential",
    "ClientSecretCredential",
    "DeviceCodeCredential",
    "OutlookMessage",
    "EmailAddress",
    "OutlookBody",
    "InternetMessageHeader",
    "OutlookAttachment",
    "OutlookAttachmentMeta",
    "OutlookFolder",
    "Attachment",
    "OutlookClient",
]
