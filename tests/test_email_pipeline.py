"""Email pipeline unit tests."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, cast

sys.path.append(str(Path(__file__).resolve().parents[1] / "src"))

from email_pipeline import (
    EmailIngestionService,
    EmailParser,
    OntologyBuilder,
    SupabaseEmailRepository,
    SupabaseSettings,
)
from email_pipeline.models import EmbeddingPayload, OntologySnapshot
from email_pipeline.parser import MessageProtocol


class FakeAttachment:
    """Test attachment mock."""

    def __init__(self, name: str, data: bytes, cid: Optional[str] = None) -> None:
        """Initialize attachment mock."""

        self.longFilename: Optional[str] = name
        self.shortFilename: Optional[str] = name
        self.data: bytes = data
        self._cid = cid
        self._mime = "text/plain"

    @property
    def cid(self) -> Optional[str]:
        """Return CID."""

        return self._cid

    @property
    def mimeType(self) -> Optional[str]:
        """Return MIME type."""

        return self._mime


class FakeMessage:
    """Test message mock."""

    def __init__(self, attachments: Iterable[FakeAttachment]) -> None:
        """Initialize message mock."""

        self.subject: str = "HVDC status report"
        self.sender: str = "planner@example.com"
        self.to: Optional[str] = "ops@example.com;manager@example.com"
        self.cc: Optional[str] = "lead@example.com"
        self.bcc: Optional[str] = ""
        self.date: str = "Fri, 12 Jul 2025 09:30:00 +0000"
        self.message_id: Optional[str] = "<sample@example.com>"
        self.header: str = (
            "Message-ID: <sample@example.com>\n"
            "In-Reply-To: <parent@example.com>\n"
            "References: <root@example.com> <parent@example.com>\n"
            "Thread-Topic: Logistics update"
        )
        self.body: str = "Body text for logistics update"
        self.htmlBody: Optional[str] = "<p>Body text for logistics update</p>"
        self.importance: Optional[str] = "Normal"
        self.categories: Optional[str] = "Logistics;HVDC"
        self.attachments: Sequence[FakeAttachment] = tuple(attachments)


class StubTable:
    """Supabase table mock."""

    def __init__(self) -> None:
        """Initialize table mock."""

        self.rows: List[Any] = []

    def upsert(self, data: Any) -> Any:
        """Record upsert call."""

        if isinstance(data, list):
            self.rows.extend(data)
        else:
            self.rows.append(data)
        return {"data": data}


class StubSupabaseClient:
    """Supabase client mock."""

    def __init__(self) -> None:
        """Initialize client mock."""

        self.tables: Dict[str, StubTable] = {}

    def table(self, name: str) -> StubTable:
        """Provide table access."""

        if name not in self.tables:
            self.tables[name] = StubTable()
        return self.tables[name]


class StubCollection:
    """Vecs collection mock."""

    def __init__(self) -> None:
        """Initialize collection mock."""

        self.records: List[tuple[str, list[float], Dict[str, Any]]] = []

    def upsert(
        self, records: Iterable[tuple[str, list[float], Dict[str, Any]]]
    ) -> None:
        """Record upsert call."""

        self.records.extend(records)


class StubEmbeddingProvider:
    """Embedding provider mock."""

    def embed(self, text: str, metadata: Dict[str, str]) -> List[float]:
        """Simulate embedding."""

        base = float(len(text))
        return [round(base / 100, 2), round(base / 200, 2), round(base / 300, 2)]


def build_parser(tmp_path: Path) -> EmailParser:
    """Build parser for tests."""

    attachments = [FakeAttachment("report.txt", b"sample data")]
    message = FakeMessage(attachments)

    def loader(_: Path) -> MessageProtocol:
        return cast(MessageProtocol, message)

    return EmailParser(tmp_path, message_loader=loader)


def test_email_parser_extracts_metadata(tmp_path: Path) -> None:
    """Parser extracts metadata."""

    parser = build_parser(tmp_path)
    record = parser.parse(tmp_path / "mail.msg")
    assert record.subject == "HVDC status report"
    assert record.from_address == "planner@example.com"
    assert record.to_addresses == ["ops@example.com", "manager@example.com"]
    assert record.cc_addresses == ["lead@example.com"]
    assert record.has_attachments is True
    assert record.attachments[0].filename == "report.txt"
    assert record.attachments[0].storage_path.exists()


def test_email_parser_handles_duplicate_attachment_names(tmp_path: Path) -> None:
    """Parser generates unique filenames when attachments collide."""

    duplicate_attachments = [
        FakeAttachment("report.txt", b"first copy"),
        FakeAttachment("report.txt", b"second copy"),
    ]
    message = FakeMessage(duplicate_attachments)

    def loader(_: Path) -> MessageProtocol:
        return cast(MessageProtocol, message)

    parser = EmailParser(tmp_path, message_loader=loader)
    record = parser.parse(tmp_path / "mail.msg")

    filenames = [attachment.filename for attachment in record.attachments]
    assert filenames == ["report.txt", "report-1.txt"]
    storage_paths = [attachment.storage_path for attachment in record.attachments]
    assert [path.name for path in storage_paths] == filenames
    assert storage_paths[0] != storage_paths[1]
    assert all(path.exists() for path in storage_paths)
    assert storage_paths[0].read_bytes() == duplicate_attachments[0].data
    assert storage_paths[1].read_bytes() == duplicate_attachments[1].data


def test_ontology_builder_generates_json_ld(tmp_path: Path) -> None:
    """Ontology builder generates JSON-LD."""

    parser = build_parser(tmp_path)
    record = parser.parse(tmp_path / "mail.msg")
    builder = OntologyBuilder()
    snapshot = builder.build(record)
    assert snapshot["@type"] == "EmailMessage"
    assert snapshot["sender"]["email"] == "planner@example.com"
    assert snapshot["messageAttachment"][0]["name"] == "report.txt"


def test_repository_persists_entities(tmp_path: Path) -> None:
    """Repository persists data."""

    parser = build_parser(tmp_path)
    record = parser.parse(tmp_path / "mail.msg")
    settings = SupabaseSettings(
        url="https://example.supabase.co",
        service_role_key="service",
        anon_key="anon",
    )
    client = StubSupabaseClient()
    collection = StubCollection()
    repository = SupabaseEmailRepository(
        settings, client, embedding_collection=collection
    )
    repository.upsert_email(record)
    repository.upsert_attachments(record)
    repository.upsert_ontology(
        OntologySnapshot(
            message_id=record.message_id, snapshot=record.ontology_snapshot or {}
        )
    )
    payload = EmbeddingPayload(
        message_id=record.message_id,
        chunk_id="chunk-1",
        vector=[0.1, 0.2, 0.3],
        metadata={"type": "body"},
    )
    repository.upsert_embeddings([payload])
    assert client.tables[settings.emails_table].rows
    assert client.tables[settings.attachments_table].rows
    assert client.tables[settings.ontology_table].rows
    assert collection.records


def test_ingestion_service_runs_full_pipeline(tmp_path: Path) -> None:
    """Ingestion service runs pipeline."""

    parser = build_parser(tmp_path)
    settings = SupabaseSettings(
        url="https://example.supabase.co",
        service_role_key="service",
        anon_key="anon",
    )
    client = StubSupabaseClient()
    collection = StubCollection()
    repository = SupabaseEmailRepository(
        settings, client, embedding_collection=collection
    )
    provider = StubEmbeddingProvider()
    service = EmailIngestionService(
        settings, repository, parser, embedding_provider=provider
    )
    record = service.ingest(tmp_path / "mail.msg", entry_id="ENTRY-1")
    assert record.entry_id == "ENTRY-1"
    assert client.tables[settings.emails_table].rows
    assert collection.records


def test_supabase_settings_from_env(monkeypatch: Any) -> None:
    """Load settings from environment."""

    monkeypatch.setenv("SUPABASE_URL", "https://example.supabase.co")
    monkeypatch.setenv("SUPABASE_SERVICE_ROLE_KEY", "service-key")
    monkeypatch.setenv("SUPABASE_ANON_KEY", "anon-key")
    monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
    settings = SupabaseSettings.from_env()
    assert settings.url == "https://example.supabase.co"
    assert settings.service_role_key == "service-key"
    assert settings.anon_key == "anon-key"
    assert settings.openai_api_key == "sk-test"

