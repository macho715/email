# HVDC 이메일 처리 파이프라인 | Email Processing Pipeline

HVDC 프로젝트를 위한 고급 이메일 지식 파이프라인입니다. Outlook VBA, Python extract-msg, Supabase, 그리고 OpenAI 임베딩을 통한 완전한 이메일 처리 시스템을 제공합니다.

[![GitHub](https://img.shields.io/badge/GitHub-Repository-blue?style=flat-square&logo=github)](https://github.com/macho715/email)
[![Python](https://img.shields.io/badge/Python-3.11+-green?style=flat-square&logo=python)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-yellow?style=flat-square)](LICENSE)

## 🚀 빠른 시작 | Quick Start

### 1. 설치 | Installation

```bash
# 저장소 클론
git clone https://github.com/macho715/email.git
cd email

# 자동 설치 실행
python install.py
```

### 2. 환경 설정 | Environment Setup

`.env` 파일을 편집하여 Supabase 설정을 구성하세요:

```env
SUPABASE_URL=https://your-project.supabase.co
SUPABASE_SERVICE_ROLE_KEY=your-service-role-key
SUPABASE_ANON_KEY=your-anon-key
OPENAI_API_KEY=sk-your-openai-key
```

### 3. 사용 예시 | Usage Example

```python
from email_pipeline import EmailIngestionService, SupabaseSettings
from pathlib import Path

# 설정 로드
settings = SupabaseSettings.from_env()

# 이메일 처리
service = EmailIngestionService(settings, repository, parser, embedding_provider)
record = service.ingest(Path("email.msg"))
```

## 📁 프로젝트 구조 | Project Structure

```
email_processing_pipeline/
├── src/email_pipeline/          # 핵심 패키지
│   ├── __init__.py             # 패키지 초기화
│   ├── base.py                 # 기본 모델 클래스
│   ├── config.py               # Supabase 설정
│   ├── models.py               # 데이터 모델
│   ├── ontology.py             # JSON-LD 온톨로지 빌더
│   ├── parser.py               # 이메일 파서
│   ├── repository.py           # Supabase 저장소
│   └── service.py              # 이메일 수집 서비스
├── tests/                      # 단위 테스트
├── docs/                       # 문서
├── config.yaml                 # 설정 파일
├── requirements.txt            # 의존성
├── pyproject.toml             # 프로젝트 설정
└── install.py                 # 설치 스크립트
```

## 🔧 주요 기능 | Key Features

### 1. 이메일 캡처 | Email Capture
- **Outlook VBA**: `Application_NewMailEx` 이벤트 기반 실시간 캡처
- **구조화된 저장**: 일자별 폴더에 `.msg` 파일 저장
- **첨부 파일 관리**: EntryID 기반 첨부 파일 저장

### 2. 이메일 파싱 | Email Parsing
- **extract-msg**: Outlook `.msg` 파일 완전 지원
- **메타데이터 추출**: 제목, 발신자, 수신자, 헤더 등
- **첨부 파일 처리**: 자동 저장 및 메타데이터 생성
- **RFC 5322 준수**: 표준 이메일 형식 지원

### 3. 데이터 저장 | Data Persistence
- **Supabase**: PostgreSQL 기반 클라우드 데이터베이스
- **pgvector**: 벡터 임베딩 저장 및 검색
- **RLS 보안**: Row Level Security로 데이터 보호
- **구조화된 테이블**: emails, attachments, ontology, embeddings

### 4. 임베딩 및 검색 | Embedding & Search
- **OpenAI 임베딩**: `text-embedding-3-small` 모델 사용
- **청크 분할**: 800자 이하로 본문 분할
- **벡터 검색**: 유사도 기반 검색 지원
- **메타데이터 필터링**: SQL 필터와 벡터 검색 결합

### 5. 온톨로지 | Ontology
- **JSON-LD**: schema.org/EmailMessage 표준 준수
- **구조화된 메타데이터**: 발신자, 수신자, 첨부 파일 등
- **검색 최적화**: 의미론적 검색을 위한 구조화

## 🧪 테스트 | Testing

### TDD 개발 방법론 | TDD Development Methodology

이 프로젝트는 **Kent Beck의 TDD**와 **Tidy First** 원칙을 따릅니다:

1. **RED**: 실패하는 테스트 작성
2. **GREEN**: 테스트 통과를 위한 최소 구현
3. **REFACTOR**: 구조 개선 (행위 변경 없이)

```bash
# plan.md 기반 TDD 개발
# 1. 다음 미표시 테스트 확인
# 2. RED → GREEN → REFACTOR 사이클 실행

# 전체 테스트 실행
pytest tests/ -v

# 커버리지 포함
pytest tests/ --cov=src/email_pipeline --cov-report=html

# 특정 테스트 실행
pytest tests/test_email_pipeline.py::test_email_parser_extracts_metadata -v

# TDD 사이클 실행 (plan.md 기반)
# /go 명령어로 다음 미표시 테스트부터 시작
```

### 테스트 구조 | Test Structure

- **단위 테스트**: 개별 모듈 테스트
- **통합 테스트**: 모듈 간 연동 테스트
- **보안 테스트**: PII/NDA/컴플라이언스 테스트
- **성능 테스트**: 대용량 처리 벤치마크

## 📊 성능 지표 | Performance Metrics

- **캡처율**: ≥99.5% (Outlook VBA 기반)
- **임베딩 지연**: ≤0.30초/문서
- **파싱 성공률**: ≥95% (extract-msg 기반)
- **저장 성공률**: ≥99% (Supabase 기반)

## 🔒 보안 | Security

- **환경 변수**: 민감한 정보는 `.env` 파일에 저장
- **RLS**: Supabase Row Level Security 적용
- **암호화**: 전송 및 저장 시 암호화
- **접근 제어**: 서비스 롤 기반 권한 관리

## 💻 시스템 환경 | System Environment

### 개발 환경 사양 | Development Environment Specifications

**운영체제**: Windows 11 Enterprise
**프로세서**: 11th Gen Intel(R) Core(TM) i7-1165G7 @ 2.80GHz
**Outlook**: Office LTSC Professional Plus 2021, 버전 2108 (빌드 14334.20296)

### 개발 도구 | Development Tools

- **IDE**: Cursor (주 개발 환경)
- **AI 지원**: ChatGPT (Edge 브라우저 사이드바에서만 사용 가능)
- **버전 관리**: Git (GitHub 접속 불가 - 로컬 개발 환경)
- **브라우저**: Microsoft Edge

### 환경 제약사항 | Environment Constraints

- **GitHub 접속 불가**: 네트워크 제약으로 인한 GitHub 접속 불가
- **AI 도구 제한**: ChatGPT는 Edge 브라우저 사이드바에서만 사용 가능
- **로컬 개발**: 모든 개발 및 테스트는 로컬 환경에서 수행

### 호환성 정보 | Compatibility Information

- ✅ **Windows 11 Enterprise**: 완전 지원
- ✅ **Office LTSC 2021**: VBA 매크로 지원
- ✅ **Cursor IDE**: 통합 개발 환경 지원
- ✅ **Python 3.11+**: 모든 기능 지원
- ⚠️ **GitHub**: 접속 불가 (로컬 개발만 가능)

## 🚀 배포 | Deployment

### 로컬 개발 | Local Development
```bash
python install.py
python -m email_pipeline.ingest sample.msg
```

### 프로덕션 | Production
```bash
# Docker 사용
docker build -t hvdc-email-pipeline .
docker run -e SUPABASE_URL=... hvdc-email-pipeline

# 또는 직접 실행
python -m email_pipeline.service
```

## 📚 문서 | Documentation

- [영어 가이드](docs/en/email_pipeline.md)
- [한국어 가이드](docs/kr/email_pipeline.md)
- [API 문서](docs/api/)
- [설정 가이드](docs/configuration.md)

## 🤝 기여 | Contributing

### 로컬 개발 환경 | Local Development Environment

현재 환경은 GitHub 접속이 불가능한 로컬 개발 환경입니다:

1. **로컬 저장소에서 직접 작업**
2. **Git 커밋을 통한 버전 관리**
3. **TDD 방법론 준수** (RED → GREEN → REFACTOR)
4. **테스트 추가 및 실행**
5. **로컬 백업 및 문서화**

### 개발 워크플로우 | Development Workflow

```bash
# 로컬 개발 워크플로우
git add .
git commit -m "[BEHAVIORAL] Add new feature"
git commit -m "[STRUCTURAL] Refactor code structure"

# 테스트 실행
pytest tests/ -v --cov=src/email_pipeline

# CLI 테스트
python run_cli.py test
python run_cli.py status
```

### 환경별 개발 가이드 | Environment-Specific Development Guide

- **Cursor IDE**: 주 개발 환경으로 모든 코딩 작업 수행
- **ChatGPT (Edge 사이드바)**: 코드 리뷰 및 최적화 제안용
- **로컬 Git**: 버전 관리 및 백업용
- **Windows 11 Enterprise**: 모든 기능 테스트 환경

## 📄 라이선스 | License

MIT License - 자세한 내용은 [LICENSE](LICENSE) 파일을 참조하세요.

## 🆘 지원 | Support

### 로컬 환경 지원 | Local Environment Support

- **문서**: 프로젝트 문서 및 plan.md 참조
- **CLI 도움말**: `python run_cli.py --help`
- **테스트 실행**: `python run_cli.py test`
- **시스템 상태**: `python run_cli.py status`
- **로그 확인**: `logs/` 디렉토리 내 로그 파일 참조

### 문제 해결 | Troubleshooting

#### 일반적인 문제 | Common Issues

1. **환경 변수 설정 오류**
   ```bash
   # .env 파일 확인
   python run_cli.py status
   ```

2. **Outlook VBA 매크로 오류**
   - Outlook 보안 설정 확인
   - VBA 코드 재컴파일 (F5)
   - 매크로 보안 설정 변경

3. **Python 의존성 문제**
   ```bash
   # 의존성 재설치
   pip install -r requirements.txt
   ```

4. **Supabase 연결 오류**
   - 네트워크 연결 확인
   - API 키 유효성 검증
   - 환경 변수 재설정

### 개발 도구 활용 | Development Tools Usage

- **Cursor IDE**: 모든 코드 편집 및 디버깅
- **ChatGPT (Edge 사이드바)**: 코드 리뷰 및 최적화
- **Windows 11 Enterprise**: 통합 테스트 환경
- **Office LTSC 2021**: Outlook VBA 통합 테스트

---

**HVDC Project v3.4-mini** | Samsung C&T Logistics & ADNOC·DSV Partnership
