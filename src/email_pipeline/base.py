"""이메일 파이프라인 기본 모델 | Email pipeline base models."""

from __future__ import annotations

from datetime import datetime
from typing import Any, Dict, Optional
from uuid import UUID, uuid4

from pydantic import BaseModel, ConfigDict, Field, field_serializer


class LogiBaseModel(BaseModel):
    """MACHO-GPT v3.4-mini 기본 모델 | Base model for MACHO-GPT v3.4-mini."""

    model_config = ConfigDict(
        use_enum_values=True,
        validate_assignment=True,
        arbitrary_types_allowed=True,
        str_strip_whitespace=True,
    )

    id: UUID = Field(default_factory=uuid4)
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime = Field(default_factory=datetime.utcnow)
    confidence: float = Field(default=0.90, ge=0.0, le=1.0)
    mode: str = Field(default="PRIME", description="Current containment mode")
    metadata: Dict[str, Any] = Field(default_factory=dict)

    @field_serializer('created_at', 'updated_at')
    def serialize_datetime(self, value: datetime) -> str:
        """날짜시간 직렬화 | Serialize datetime."""
        return value.isoformat()

    @field_serializer('id')
    def serialize_uuid(self, value: UUID) -> str:
        """UUID 직렬화 | Serialize UUID."""
        return str(value)

    def update_timestamp(self) -> None:
        """업데이트 시간 갱신 | Update timestamp."""
        self.updated_at = datetime.utcnow()

    def set_confidence(self, confidence: float) -> None:
        """신뢰도 설정 | Set confidence level."""
        if not 0.0 <= confidence <= 1.0:
            raise ValueError("Confidence must be between 0.0 and 1.0")
        self.confidence = confidence

    def add_metadata(self, key: str, value: Any) -> None:
        """메타데이터 추가 | Add metadata."""
        self.metadata[key] = value
        self.update_timestamp()

    def get_metadata(self, key: str, default: Any = None) -> Any:
        """메타데이터 조회 | Get metadata."""
        return self.metadata.get(key, default)
