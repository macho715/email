"""이메일 파일 파서 구현 | Email file parser implementation."""

from __future__ import annotations

import email
import hashlib
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Protocol, Sequence, cast

from .models import EmailAttachment, EmailMessageRecord
from .ontology import OntologyBuilder


class AttachmentProtocol(Protocol):
    """첨부 객체 프로토콜 | Attachment object protocol."""

    longFilename: Optional[str]
    shortFilename: Optional[str]
    data: bytes

    @property
    def cid(self) -> Optional[str]:
        """콘텐츠 ID 반환 | Return content id."""

    @property
    def mimeType(self) -> Optional[str]:
        """MIME 타입 반환 | Return MIME type."""


class MessageProtocol(Protocol):
    """extract-msg 메시지 프로토콜 | Protocol for extract-msg message."""

    subject: str
    sender: str
    to: Optional[str]
    cc: Optional[str]
    bcc: Optional[str]
    date: str
    message_id: Optional[str]
    header: str
    body: str
    htmlBody: Optional[str]
    importance: Optional[str]
    categories: Optional[str]
    attachments: Sequence[AttachmentProtocol]


class EmailParser:
    """이메일 파일을 구조화 데이터로 파싱 | Parse email files into structured data."""

    def __init__(
        self,
        attachment_dir: Path,
        ontology_builder: Optional[OntologyBuilder] = None,
        message_loader: Optional[Callable[[Path], MessageProtocol]] = None,
    ) -> None:
        """파서를 초기화합니다 | Initialize parser."""

        self.attachment_dir = attachment_dir
        self.attachment_dir.mkdir(parents=True, exist_ok=True)
        self._ontology_builder = ontology_builder or OntologyBuilder()
        if message_loader is None:
            self._message_loader: Callable[[Path], MessageProtocol] = (
                self._default_loader
            )
        else:
            self._message_loader = message_loader

    def parse(
        self, source_path: Path, entry_id: Optional[str] = None
    ) -> EmailMessageRecord:
        """이메일 파일을 파싱합니다 | Parse email file."""

        message = self._message_loader(source_path)
        message_id = message.message_id or self._build_message_id(source_path)
        received_at = self._parse_date(message.date)
        attachments = list(self._process_attachments(message.attachments, message_id))
        record = EmailMessageRecord(
            entry_id=entry_id,
            message_id=message_id,
            subject=message.subject,
            from_address=message.sender,
            to_addresses=self._parse_addresses(message.to),
            cc_addresses=self._parse_addresses(message.cc),
            bcc_addresses=self._parse_addresses(message.bcc),
            received_at=received_at,
            body_text=message.body,
            body_html=getattr(message, "htmlBody", None),
            categories=self._parse_categories(message.categories),
            headers=self._parse_headers(message.header),
            importance=getattr(message, "importance", None),
            has_attachments=bool(attachments),
            attachments=attachments,
            source_path=source_path,
        )
        record.ontology_snapshot = self._ontology_builder.build(record)
        return record

    def _process_attachments(
        self, attachments: Iterable[AttachmentProtocol], message_id: str
    ) -> Iterable[EmailAttachment]:
        """첨부 파일을 저장하고 메타데이터 생성 | Store attachments and build metadata."""

        for index, attachment in enumerate(attachments):
            filename = (
                attachment.longFilename
                or attachment.shortFilename
                or f"attachment-{index + 1}"
            )
            safe_name = self._sanitize_filename(filename)
            safe_message_id = self._sanitize_filename(message_id)
            target_dir = self.attachment_dir / safe_message_id
            target_dir.mkdir(parents=True, exist_ok=True)
            resolved_name, target_path = self._resolve_attachment_path(target_dir, safe_name)
            target_path.write_bytes(attachment.data)
            checksum = hashlib.sha256(attachment.data).hexdigest()
            size_bytes = len(attachment.data)
            yield EmailAttachment(
                filename=resolved_name,
                content_id=getattr(attachment, "cid", None),
                content_type=getattr(attachment, "mimeType", None),
                size_bytes=size_bytes,
                checksum=checksum,
                storage_path=target_path,
            )


    @staticmethod
    def _resolve_attachment_path(target_dir: Path, filename: str) -> tuple[str, Path]:
        """Generate a collision-free filename inside target_dir."""

        candidate_name = filename or "attachment"
        candidate_path = target_dir / candidate_name
        if not candidate_path.exists():
            return candidate_name, candidate_path
        stem = Path(candidate_name).stem
        suffix = Path(candidate_name).suffix
        counter = 1
        while True:
            next_name = f"{stem}-{counter}{suffix}"
            next_path = target_dir / next_name
            if not next_path.exists():
                return next_name, next_path
            counter += 1

    @staticmethod
    def _sanitize_filename(name: str) -> str:
        """파일명을 안전하게 정규화 | Sanitize filenames safely."""

        return (
            "".join(ch for ch in name if ch.isalnum() or ch in {".", "_", "-"})
            or "attachment"
        )

    @staticmethod
    def _parse_addresses(raw_value: Optional[str]) -> List[str]:
        """주소 문자열을 리스트로 변환 | Convert address string into list."""

        if not raw_value:
            return []
        separators = [";", ","]
        addresses = [raw_value]
        for separator in separators:
            addresses = sum((item.split(separator) for item in addresses), [])
        return [addr.strip() for addr in addresses if addr.strip()]

    @staticmethod
    def _parse_categories(raw_value: Optional[str]) -> List[str]:
        """카테고리 필드 파싱 | Parse category field."""

        if not raw_value:
            return []
        return [value.strip() for value in raw_value.split(";") if value.strip()]

    @staticmethod
    def _parse_headers(raw_header: str) -> Dict[str, str]:
        """헤더 블록을 딕셔너리로 변환 | Convert header block into dict."""

        message = email.message_from_string(raw_header)
        return {key: value for key, value in message.items()}

    @staticmethod
    def _parse_date(raw_date: str) -> datetime:
        """날짜 문자열을 datetime으로 변환 | Convert date string into datetime."""

        try:
            parsed = parsedate_to_datetime(raw_date)
            if parsed is not None:
                return parsed
        except (TypeError, ValueError):
            pass
        return datetime.now(tz=timezone.utc)

    @staticmethod
    def _build_message_id(source_path: Path) -> str:
        """파일 경로 기반 메시지 ID 생성 | Build message id from path."""

        digest = hashlib.sha1(str(source_path).encode("utf-8")).hexdigest()
        return f"local-{digest}@hvdc.local"

    @staticmethod
    def _default_loader(source_path: Path) -> MessageProtocol:
        """기본 extract-msg 로더 | Default extract-msg loader."""

        try:
            from extract_msg import Message  # pyright: ignore[reportMissingImports]
        except ImportError as exc:  # pragma: no cover
            raise RuntimeError(
                "extract-msg library is required for parsing .msg files"
            ) from exc
        return cast(MessageProtocol, Message(str(source_path)))