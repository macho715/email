"""Supabase 연동 저장소 | Supabase integration repository."""

from __future__ import annotations

from typing import Any, Iterable, Optional, Protocol, Sequence

from .config import SupabaseSettings
from .models import (
    EmailAttachment,
    EmailMessageRecord,
    EmbeddingPayload,
    OntologySnapshot,
)


class SupabaseTableProtocol(Protocol):
    """Supabase 테이블 프로토콜 | Supabase table protocol."""

    def upsert(self, data: Any) -> Any:
        """업서트 실행 | Execute upsert."""


class SupabaseClientProtocol(Protocol):
    """Supabase 클라이언트 프로토콜 | Supabase client protocol."""

    def table(self, name: str) -> SupabaseTableProtocol:
        """테이블 핸들을 반환 | Return table handle."""


class VecsCollectionProtocol(Protocol):
    """Vecs 컬렉션 프로토콜 | Vecs collection protocol."""

    def upsert(
        self, records: Sequence[tuple[str, list[float], dict[str, Any]]]
    ) -> None:
        """벡터를 업서트 | Upsert vectors."""


class SupabaseEmailRepository:
    """Supabase 이메일 저장소 | Supabase email repository."""

    def __init__(
        self,
        settings: SupabaseSettings,
        supabase_client: SupabaseClientProtocol,
        embedding_collection: Optional[VecsCollectionProtocol] = None,
    ) -> None:
        """저장소를 초기화합니다 | Initialize repository."""

        self.settings = settings
        self._supabase = supabase_client
        self._collection = embedding_collection

    def upsert_email(self, record: EmailMessageRecord) -> None:
        """이메일 메타데이터 업서트 | Upsert email metadata."""

        self._table(self.settings.emails_table).upsert(record.as_row())

    def upsert_attachments(self, record: EmailMessageRecord) -> None:
        """첨부 메타데이터 업서트 | Upsert attachment metadata."""

        if not record.attachments:
            return
        rows = [
            self._attachment_row(record.message_id, attachment)
            for attachment in record.attachments
        ]
        self._table(self.settings.attachments_table).upsert(rows)

    def upsert_ontology(self, snapshot: OntologySnapshot) -> None:
        """온톨로지 스냅샷 업서트 | Upsert ontology snapshot."""

        payload = {"message_id": snapshot.message_id, "snapshot": snapshot.snapshot}
        self._table(self.settings.ontology_table).upsert(payload)

    def upsert_embeddings(self, embeddings: Iterable[EmbeddingPayload]) -> None:
        """임베딩을 업서트합니다 | Upsert embeddings."""

        if self._collection is None:
            raise RuntimeError("Vecs collection is not configured for embeddings")
        records: list[tuple[str, list[float], dict[str, Any]]] = []
        for payload in embeddings:
            metadata = dict(payload.metadata)
            metadata.update(
                {"message_id": payload.message_id, "chunk_id": payload.chunk_id}
            )
            records.append((payload.chunk_id, payload.vector, metadata))
        if records:
            self._collection.upsert(records)

    def _table(self, name: str) -> SupabaseTableProtocol:
        """테이블 핸들을 제공합니다 | Provide table handle."""

        return self._supabase.table(name)

    @staticmethod
    def _attachment_row(message_id: str, attachment: EmailAttachment) -> dict[str, Any]:
        """첨부 행을 구성합니다 | Build attachment row."""

        return {
            "message_id": message_id,
            "filename": attachment.filename,
            "content_id": attachment.content_id,
            "content_type": attachment.content_type,
            "size_bytes": attachment.size_bytes,
            "checksum": attachment.checksum,
            "storage_path": str(attachment.storage_path),
        }
