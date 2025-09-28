"""이메일 파이프라인 패키지 | Email pipeline package."""

from .base import LogiBaseModel
from .cli import app as cli_app
from .config import SupabaseSettings
from .models import (
    Currency,
    EmailAttachment,
    EmailMessageRecord,
    EmbeddingPayload,
    OntologySnapshot,
)
from .ontology import OntologyBuilder
from .parser import EmailParser
from .repository import SupabaseEmailRepository
from .service import EmailIngestionService, OpenAIEmbeddingProvider

__version__ = "3.4.0"
__all__ = [
    "LogiBaseModel",
    "SupabaseSettings",
    "Currency",
    "EmailAttachment",
    "EmailMessageRecord",
    "EmbeddingPayload",
    "OntologySnapshot",
    "OntologyBuilder",
    "EmailParser",
    "SupabaseEmailRepository",
    "EmailIngestionService",
    "OpenAIEmbeddingProvider",
    "cli_app",
]
