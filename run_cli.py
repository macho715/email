#!/usr/bin/env python3
"""이메일 파이프라인 CLI 실행 스크립트 | Email pipeline CLI runner."""

import sys
from pathlib import Path

# src 디렉토리를 Python 경로에 추가
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

# CLI 앱 실행
from email_pipeline.cli import app

if __name__ == "__main__":
    app()
