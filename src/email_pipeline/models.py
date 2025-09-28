"""이메일 파이프라인 데이터 모델 | Data models for email pipeline."""

from __future__ import annotations

from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Any, Dict, List, Optional

from .base import LogiBaseModel


class Currency(str, Enum):
    """통화 코드 열거형 | Currency enumeration."""

    AED = "AED"
    USD = "USD"
    EUR = "EUR"
    GBP = "GBP"
    CNY = "CNY"
    KRW = "KRW"


class EmailAttachment(LogiBaseModel):
    """이메일 첨부 데이터 모델 | Email attachment data model."""

    filename: str
    content_id: Optional[str]
    content_type: Optional[str]
    size_bytes: int
    checksum: str
    storage_path: Path


class EmailMessageRecord(LogiBaseModel):
    """이메일 메타데이터 및 본문 모델 | Email metadata and body model."""

    entry_id: Optional[str]
    message_id: str
    subject: str
    from_address: str
    to_addresses: List[str]
    cc_addresses: List[str]
    bcc_addresses: List[str]
    received_at: datetime
    body_text: str
    body_html: Optional[str]
    categories: List[str]
    headers: Dict[str, str]
    importance: Optional[str]
    has_attachments: bool
    attachments: List[EmailAttachment]
    source_path: Path
    ontology_snapshot: Optional[Dict[str, Any]] = None

    def as_row(self) -> Dict[str, Any]:
        """Supabase 행 변환 수행 | Convert record into Supabase row."""

        return {
            "entry_id": self.entry_id,
            "message_id": self.message_id,
            "subject": self.subject,
            "from_address": self.from_address,
            "to_addresses": self.to_addresses,
            "cc_addresses": self.cc_addresses,
            "bcc_addresses": self.bcc_addresses,
            "received_at": self.received_at.isoformat(),
            "body_text": self.body_text,
            "body_html": self.body_html,
            "categories": self.categories,
            "headers": self.headers,
            "importance": self.importance,
            "has_attachments": self.has_attachments,
            "source_path": str(self.source_path),
        }


class EmbeddingPayload(LogiBaseModel):
    """임베딩 데이터 적재 모델 | Embedding data load model."""

    message_id: str
    chunk_id: str
    vector: List[float]
    metadata: Dict[str, Any]


class OntologySnapshot(LogiBaseModel):
    """온톨로지 스냅샷 데이터 | Ontology snapshot data."""

    message_id: str
    snapshot: Dict[str, Any]
