#!/usr/bin/env python3
"""캡처된 이메일 파일들을 일괄 처리하는 스크립트 | Batch process captured email files."""

import sys
from pathlib import Path
import time
from datetime import datetime

# src 디렉토리를 Python 경로에 추가
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

from email_pipeline import EmailIngestionService, SupabaseSettings
from email_pipeline.parser import EmailParser
from email_pipeline.repository import SupabaseEmailRepository
from email_pipeline.service import OpenAIEmbeddingProvider


def process_captured_emails(capture_folder: str = r"C:\HVDC\EmailCapture"):
    """캡처된 이메일 파일들을 처리합니다."""
    
    capture_path = Path(capture_folder)
    if not capture_path.exists():
        print(f"❌ 캡처 폴더를 찾을 수 없습니다: {capture_folder}")
        return
    
    # .msg 파일 찾기
    msg_files = list(capture_path.glob("*.msg"))
    if not msg_files:
        print(f"📭 처리할 이메일 파일이 없습니다: {capture_folder}")
        return
    
    print(f"📧 {len(msg_files)}개의 이메일 파일을 발견했습니다.")
    
    # 설정 로드 (환경 변수에서)
    try:
        settings = SupabaseSettings.from_env()
        print("✅ Supabase 설정을 로드했습니다.")
    except Exception as e:
        print(f"⚠️ Supabase 설정 오류: {e}")
        print("📝 모의 모드로 실행합니다.")
        settings = None
    
    # 출력 디렉토리 설정
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)
    attachment_dir = output_dir / "attachments"
    
    # 파서 초기화
    parser = EmailParser(attachment_dir)
    
    # 저장소 초기화 (모의 객체 사용)
    if settings:
        try:
            from supabase import create_client
            supabase = create_client(settings.url, settings.service_role_key)
            repository = SupabaseEmailRepository(settings, supabase)
            print("✅ Supabase 저장소를 초기화했습니다.")
        except Exception as e:
            print(f"⚠️ Supabase 연결 실패: {e}")
            print("📝 모의 모드로 실행합니다.")
            settings = None
    
    if not settings:
        # 모의 저장소 사용
        from tests.test_email_pipeline import StubSupabaseClient, StubCollection
        client = StubSupabaseClient()
        collection = StubCollection()
        repository = SupabaseEmailRepository(
            SupabaseSettings(url="mock", service_role_key="mock"),
            client,
            collection
        )
        print("📝 모의 저장소를 사용합니다.")
    
    # 임베딩 제공자 초기화
    embedding_provider = None
    if settings and settings.openai_api_key:
        try:
            embedding_provider = OpenAIEmbeddingProvider(settings.openai_api_key)
            print("✅ OpenAI 임베딩 제공자를 초기화했습니다.")
        except Exception as e:
            print(f"⚠️ OpenAI 임베딩 초기화 실패: {e}")
    
    # 서비스 초기화
    service = EmailIngestionService(
        settings or SupabaseSettings(url="mock", service_role_key="mock"),
        repository,
        parser,
        embedding_provider
    )
    
    # 이메일 처리
    processed_count = 0
    error_count = 0
    
    print(f"\n🚀 이메일 처리를 시작합니다...")
    print("=" * 50)
    
    for i, msg_file in enumerate(msg_files, 1):
        try:
            print(f"[{i}/{len(msg_files)}] 처리 중: {msg_file.name}")
            
            # 이메일 처리
            record = service.ingest(msg_file)
            
            print(f"  ✅ 성공: {record.subject[:50]}...")
            print(f"  📧 발신자: {record.from_address}")
            print(f"  📎 첨부파일: {len(record.attachments)}개")
            print(f"  🎯 신뢰도: {record.confidence:.2%}")
            
            processed_count += 1
            
        except Exception as e:
            print(f"  ❌ 오류: {e}")
            error_count += 1
        
        # 진행률 표시
        if i % 5 == 0 or i == len(msg_files):
            progress = (i / len(msg_files)) * 100
            print(f"  📊 진행률: {progress:.1f}% ({i}/{len(msg_files)})")
        
        # 잠시 대기 (시스템 부하 방지)
        time.sleep(0.1)
    
    print("\n" + "=" * 50)
    print(f"🎉 처리 완료!")
    print(f"  ✅ 성공: {processed_count}개")
    print(f"  ❌ 실패: {error_count}개")
    print(f"  📊 성공률: {(processed_count/(processed_count+error_count)*100):.1f}%")
    
    # 결과 요약
    if processed_count > 0:
        print(f"\n📁 출력 파일 위치:")
        print(f"  📧 이메일: {output_dir}")
        print(f"  📎 첨부파일: {attachment_dir}")
        
        if settings:
            print(f"\n🗄️ 데이터베이스 저장 완료:")
            print(f"  🔗 Supabase: {settings.url}")
            print(f"  📊 테이블: {settings.emails_table}")


def main():
    """메인 함수"""
    print("🔧 HVDC 이메일 처리 파이프라인")
    print("=" * 40)
    
    # 명령행 인수 처리
    capture_folder = r"C:\HVDC\EmailCapture"
    if len(sys.argv) > 1:
        capture_folder = sys.argv[1]
    
    print(f"📁 캡처 폴더: {capture_folder}")
    print(f"⏰ 시작 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    try:
        process_captured_emails(capture_folder)
    except KeyboardInterrupt:
        print("\n⏹️ 사용자에 의해 중단되었습니다.")
    except Exception as e:
        print(f"\n💥 예상치 못한 오류: {e}")
        sys.exit(1)
    
    print(f"\n⏰ 완료 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")


if __name__ == "__main__":
    main()

