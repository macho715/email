"""이메일 파이프라인 서비스 | Email pipeline service."""

from __future__ import annotations

import math
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Protocol

from .config import SupabaseSettings
from .models import EmailMessageRecord, EmbeddingPayload, OntologySnapshot
from .parser import EmailParser
from .repository import SupabaseEmailRepository


class EmbeddingProviderProtocol(Protocol):
    """임베딩 제공자 프로토콜 | Embedding provider protocol."""

    def embed(self, text: str, metadata: Dict[str, str]) -> List[float]:
        """텍스트 임베딩 수행 | Embed text."""


class OpenAIEmbeddingProvider:
    """OpenAI 임베딩 제공자 | OpenAI embedding provider."""

    def __init__(self, api_key: str, model: str = "text-embedding-3-small") -> None:
        """OpenAI 임베딩 제공자를 초기화 | Initialize OpenAI embedding provider."""

        from openai import OpenAI

        self.model = model
        self._client = OpenAI(api_key=api_key)

    def embed(self, text: str, metadata: Dict[str, str]) -> List[float]:
        """OpenAI 임베딩 호출 | Call OpenAI embedding."""

        response = self._client.embeddings.create(model=self.model, input=text)
        return list(response.data[0].embedding)


class EmailIngestionService:
    """이메일 수집 서비스 | Email ingestion service."""

    def __init__(
        self,
        settings: SupabaseSettings,
        repository: SupabaseEmailRepository,
        parser: EmailParser,
        embedding_provider: Optional[EmbeddingProviderProtocol] = None,
    ) -> None:
        """서비스를 초기화합니다 | Initialize service."""

        self.settings = settings
        self.repository = repository
        self.parser = parser
        self.embedding_provider = embedding_provider

    def ingest(
        self, source_path: Path, entry_id: Optional[str] = None
    ) -> EmailMessageRecord:
        """이메일 파일을 적재 | Ingest email file."""

        record = self.parser.parse(source_path, entry_id=entry_id)
        self.repository.upsert_email(record)
        self.repository.upsert_attachments(record)
        ontology_snapshot = OntologySnapshot(
            message_id=record.message_id, snapshot=record.ontology_snapshot or {}
        )
        self.repository.upsert_ontology(ontology_snapshot)
        if self.embedding_provider:
            embeddings = list(self._build_embeddings(record))
            if embeddings:
                self.repository.upsert_embeddings(embeddings)
        return record

    def _build_embeddings(self, record: EmailMessageRecord) -> List[EmbeddingPayload]:
        """임베딩 페이로드 생성 | Build embedding payloads."""

        if not self.embedding_provider:
            return []
        payloads: List[EmbeddingPayload] = []
        chunks = list(self._chunk_body(record.body_text))
        for index, chunk in enumerate(chunks):
            metadata = {"message_id": record.message_id, "chunk_index": str(index)}
            vector = self.embedding_provider.embed(chunk, metadata)
            payloads.append(
                EmbeddingPayload(
                    message_id=record.message_id,
                    chunk_id=f"{record.message_id}-chunk-{index}",
                    vector=vector,
                    metadata={"type": "body", "subject": record.subject},
                )
            )
        subject_vector = self.embedding_provider.embed(
            record.subject, {"message_id": record.message_id, "chunk_index": "subject"}
        )
        payloads.append(
            EmbeddingPayload(
                message_id=record.message_id,
                chunk_id=f"{record.message_id}-subject",
                vector=subject_vector,
                metadata={"type": "subject"},
            )
        )
        return payloads

    @staticmethod
    def _chunk_body(body: str, max_tokens: int = 800) -> Iterable[str]:
        """본문을 청크로 나눕니다 | Split body into chunks."""

        if not body:
            return []
        length = len(body)
        segments = math.ceil(length / max_tokens)
        return (
            body[index * max_tokens : (index + 1) * max_tokens]
            for index in range(segments)
        )
