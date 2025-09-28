#!/usr/bin/env python3
"""이메일 파이프라인 설치 스크립트 | Email pipeline installation script."""

import subprocess
import sys
from pathlib import Path


def install_dependencies():
    """의존성 설치 | Install dependencies."""
    print("📦 Installing email pipeline dependencies...")
    
    try:
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", "-r", "requirements.txt"
        ])
        print("✅ Dependencies installed successfully!")
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to install dependencies: {e}")
        return False
    
    return True


def setup_environment():
    """환경 설정 | Setup environment."""
    print("🔧 Setting up environment...")
    
    env_file = Path(".env")
    if not env_file.exists():
        print("📝 Creating .env file from template...")
        with open("env_example.txt", "r") as src:
            content = src.read()
        with open(".env", "w") as dst:
            dst.write(content)
        print("✅ .env file created! Please update with your actual values.")
    else:
        print("✅ .env file already exists.")
    
    return True


def run_tests():
    """테스트 실행 | Run tests."""
    print("🧪 Running tests...")
    
    try:
        subprocess.check_call([
            sys.executable, "-m", "pytest", "tests/", "-v"
        ])
        print("✅ All tests passed!")
    except subprocess.CalledProcessError as e:
        print(f"❌ Tests failed: {e}")
        return False
    
    return True


def main():
    """메인 설치 함수 | Main installation function."""
    print("🚀 HVDC Email Processing Pipeline Installation")
    print("=" * 50)
    
    # 의존성 설치
    if not install_dependencies():
        sys.exit(1)
    
    # 환경 설정
    if not setup_environment():
        sys.exit(1)
    
    # 테스트 실행
    if not run_tests():
        print("⚠️  Tests failed, but installation completed.")
    
    print("\n🎉 Installation completed successfully!")
    print("\n📋 Next steps:")
    print("1. Update .env file with your Supabase credentials")
    print("2. Configure your email settings in config.yaml")
    print("3. Run: python -m email_pipeline.ingest <email_file.msg>")


if __name__ == "__main__":
    main()
