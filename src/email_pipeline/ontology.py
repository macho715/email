"""이메일 온톨로지 스냅샷 빌더 | Email ontology snapshot builder."""

from __future__ import annotations

from typing import Any, Dict, List

from .models import EmailAttachment, EmailMessageRecord


class OntologyBuilder:
    """이메일 온톨로지 생성기 | Email ontology generator."""

    def build(self, record: EmailMessageRecord) -> Dict[str, Any]:
        """JSON-LD 온톨로지 스냅샷 생성 | Build JSON-LD ontology snapshot."""

        attachments = [self._build_attachment_node(item) for item in record.attachments]
        header_properties = self._build_header_properties(record.headers)
        snapshot: Dict[str, Any] = {
            "@context": "https://schema.org",
            "@type": "EmailMessage",
            "@id": record.message_id,
            "identifier": record.message_id,
            "dateReceived": record.received_at.isoformat(),
            "headline": record.subject,
            "sender": {
                "@type": "ContactPoint",
                "email": record.from_address,
            },
            "recipient": [
                {"@type": "ContactPoint", "email": address}
                for address in record.to_addresses
            ],
            "about": record.categories,
            "messageAttachment": attachments,
            "isPartOf": record.headers.get("References"),
            "inReplyTo": record.headers.get("In-Reply-To"),
            "additionalProperty": header_properties,
        }
        if record.body_html:
            snapshot["articleBody"] = record.body_html
        else:
            snapshot["articleBody"] = record.body_text
        return snapshot

    @staticmethod
    def _build_attachment_node(attachment: EmailAttachment) -> Dict[str, Any]:
        """첨부 메타데이터 노드 생성 | Build attachment metadata node."""

        return {
            "@type": "DigitalDocument",
            "name": attachment.filename,
            "encodingFormat": attachment.content_type,
            "contentSize": attachment.size_bytes,
            "identifier": attachment.checksum,
            "contentUrl": str(attachment.storage_path),
        }

    @staticmethod
    def _build_header_properties(headers: Dict[str, str]) -> List[Dict[str, Any]]:
        """헤더를 속성 목록으로 변환 | Convert headers to property list."""

        interesting_headers = [
            "Message-ID",
            "In-Reply-To",
            "References",
            "Thread-Index",
            "Thread-Topic",
        ]
        properties: List[Dict[str, Any]] = []
        for name in interesting_headers:
            value = headers.get(name)
            if value:
                properties.append(
                    {
                        "@type": "PropertyValue",
                        "name": name,
                        "value": value,
                    }
                )
        return properties
