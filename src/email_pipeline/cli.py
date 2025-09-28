"""이메일 파이프라인 CLI 인터페이스 | Email pipeline CLI interface."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Optional

import typer
from rich.console import Console
from rich.table import Table

from .config import SupabaseSettings
from .service import EmailIngestionService, OpenAIEmbeddingProvider
from .parser import EmailParser
from .repository import SupabaseEmailRepository

app = typer.Typer(
    name="email-pipeline",
    help="HVDC 이메일 처리 파이프라인 | HVDC Email Processing Pipeline",
    rich_markup_mode="rich",
)
console = Console()


@app.command()
def status():
    """시스템 상태 확인 | Check system status."""
    console.print("🔍 [bold blue]HVDC Email Pipeline Status[/bold blue]")
    
    # 환경 변수 확인
    try:
        settings = SupabaseSettings.from_env()
        console.print("✅ [green]Supabase 설정 로드됨[/green]")
        console.print(f"   URL: {settings.url}")
        console.print(f"   테이블: {settings.emails_table}")
    except Exception as e:
        console.print(f"❌ [red]Supabase 설정 오류: {e}[/red]")
        return
    
    # 의존성 확인
    dependencies = [
        ("extract-msg", "이메일 파싱"),
        ("supabase", "데이터베이스"),
        ("openai", "임베딩"),
        ("vecs", "벡터 검색"),
    ]
    
    table = Table(title="의존성 상태 | Dependencies Status")
    table.add_column("패키지", style="cyan")
    table.add_column("용도", style="magenta")
    table.add_column("상태", style="green")
    
    for package, purpose in dependencies:
        try:
            __import__(package.replace("-", "_"))
            status = "✅ 설치됨"
        except ImportError:
            status = "❌ 미설치"
        table.add_row(package, purpose, status)
    
    console.print(table)


@app.command()
def test():
    """테스트 실행 | Run tests."""
    console.print("🧪 [bold blue]테스트 실행 중...[/bold blue]")
    
    import subprocess
    result = subprocess.run([
        sys.executable, "-m", "pytest", "tests/", "-v", "--tb=short"
    ], capture_output=True, text=True)
    
    if result.returncode == 0:
        console.print("✅ [green]모든 테스트 통과![/green]")
    else:
        console.print("❌ [red]테스트 실패:[/red]")
        console.print(result.stdout)
        console.print(result.stderr)


@app.command()
def ingest(
    email_file: Path = typer.Argument(..., help="처리할 이메일 파일 (.msg)"),
    entry_id: Optional[str] = typer.Option(None, "--entry-id", help="Entry ID"),
    output_dir: Path = typer.Option(Path("output"), "--output", help="출력 디렉토리"),
    use_embeddings: bool = typer.Option(True, "--embeddings/--no-embeddings", help="임베딩 생성 여부"),
):
    """이메일 파일 처리 | Process email file."""
    console.print(f"📧 [bold blue]이메일 처리 중: {email_file}[/bold blue]")
    
    if not email_file.exists():
        console.print(f"❌ [red]파일을 찾을 수 없습니다: {email_file}[/red]")
        raise typer.Exit(1)
    
    try:
        # 설정 로드
        settings = SupabaseSettings.from_env()
        
        # 출력 디렉토리 생성
        output_dir.mkdir(parents=True, exist_ok=True)
        attachment_dir = output_dir / "attachments"
        
        # 파서 초기화
        parser = EmailParser(attachment_dir)
        
        # 저장소 초기화 (모의 객체 사용)
        from .tests.test_email_pipeline import StubSupabaseClient, StubCollection
        client = StubSupabaseClient()
        collection = StubCollection()
        repository = SupabaseEmailRepository(settings, client, collection)
        
        # 임베딩 제공자 초기화
        embedding_provider = None
        if use_embeddings and settings.openai_api_key:
            embedding_provider = OpenAIEmbeddingProvider(settings.openai_api_key)
        
        # 서비스 초기화
        service = EmailIngestionService(settings, repository, parser, embedding_provider)
        
        # 이메일 처리
        record = service.ingest(email_file, entry_id=entry_id)
        
        # 결과 출력
        console.print("✅ [green]이메일 처리 완료![/green]")
        console.print(f"   메시지 ID: {record.message_id}")
        console.print(f"   제목: {record.subject}")
        console.print(f"   발신자: {record.from_address}")
        console.print(f"   첨부 파일: {len(record.attachments)}개")
        console.print(f"   신뢰도: {record.confidence:.2%}")
        
    except Exception as e:
        console.print(f"❌ [red]처리 중 오류 발생: {e}[/red]")
        raise typer.Exit(1)


@app.command()
def setup():
    """초기 설정 | Initial setup."""
    console.print("🔧 [bold blue]HVDC Email Pipeline 초기 설정[/bold blue]")
    
    # .env 파일 확인
    env_file = Path(".env")
    if not env_file.exists():
        console.print("📝 [yellow].env 파일을 생성합니다...[/yellow]")
        env_example = Path("env_example.txt")
        if env_example.exists():
            env_file.write_text(env_example.read_text())
            console.print("✅ [green].env 파일이 생성되었습니다.[/green]")
            console.print("   [yellow]실제 값으로 업데이트해주세요.[/yellow]")
        else:
            console.print("❌ [red]env_example.txt 파일을 찾을 수 없습니다.[/red]")
    else:
        console.print("✅ [green].env 파일이 이미 존재합니다.[/green]")
    
    # 테스트 실행
    console.print("\n🧪 [blue]테스트를 실행합니다...[/blue]")
    test()


@app.command()
def demo():
    """데모 실행 | Run demo."""
    console.print("🎯 [bold blue]HVDC Email Pipeline 데모[/bold blue]")
    
    # 샘플 데이터 생성
    sample_data = {
        "entry_id": "DEMO-001",
        "message_id": "<demo@hvdc.local>",
        "subject": "HVDC 프로젝트 상태 보고서",
        "from_address": "planner@hvdc.local",
        "to_addresses": ["ops@hvdc.local", "manager@hvdc.local"],
        "cc_addresses": ["lead@hvdc.local"],
        "received_at": "2025-01-26T10:30:00+00:00",
        "body_text": "HVDC 프로젝트의 현재 상태를 보고드립니다.",
        "categories": ["Logistics", "HVDC", "Status"],
        "has_attachments": True,
        "confidence": 0.95
    }
    
    table = Table(title="샘플 이메일 데이터 | Sample Email Data")
    table.add_column("필드", style="cyan")
    table.add_column("값", style="green")
    
    for key, value in sample_data.items():
        if isinstance(value, list):
            value = ", ".join(value)
        table.add_row(key, str(value))
    
    console.print(table)
    console.print("\n💡 [blue]실제 이메일 파일을 처리하려면:[/blue]")
    console.print("   [green]python -m email_pipeline.cli ingest sample.msg[/green]")


if __name__ == "__main__":
    app()
