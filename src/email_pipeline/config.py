"""이메일 파이프라인 설정 | Email pipeline settings."""

from __future__ import annotations

import os
from typing import Optional

from dotenv import load_dotenv

from .base import LogiBaseModel


class SupabaseSettings(LogiBaseModel):
    """Supabase 연결 설정 | Supabase connection settings."""

    url: str
    service_role_key: str
    anon_key: Optional[str]
    db_schema: str = "public"
    emails_table: str = "emails"
    attachments_table: str = "email_attachments"
    ontology_table: str = "email_ontology"
    embeddings_table: str = "email_embeddings"
    vector_collection: str = "email_embeddings"
    openai_api_key: Optional[str] = None

    @classmethod
    def from_env(cls) -> "SupabaseSettings":
        """환경 변수에서 설정 로드 | Load settings from environment."""

        load_dotenv()
        url = os.getenv("SUPABASE_URL")
        service_key = os.getenv("SUPABASE_SERVICE_ROLE_KEY")
        anon_key = os.getenv("SUPABASE_ANON_KEY")
        openai_key = os.getenv("OPENAI_API_KEY")
        if not url or not service_key:
            raise RuntimeError(
                "Supabase environment variables are not fully configured"
            )
        return cls(
            url=url,
            service_role_key=service_key,
            anon_key=anon_key,
            openai_api_key=openai_key,
        )
