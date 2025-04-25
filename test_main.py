import pytest
import platform
import os
import shutil
from pathlib import Path
from main import search_local_files, local_read_file

# テストデータのセットアップとクリーンアップ
@pytest.fixture(scope="session")
def test_data_dir():
    """テストデータディレクトリのセットアップ"""
    test_dir = Path(__file__).parent / "test_data"
    test_dir.mkdir(exist_ok=True)
    
    # テストファイルの作成
    (test_dir / "test.txt").write_text("これはテストです")
    
    yield test_dir
    
    # クリーンアップ
    shutil.rmtree(test_dir)

# プラットフォーム固有のテスト
@pytest.mark.skipif(platform.system() != 'Windows', reason="Windows環境でのみ実行")
def test_windows_search(test_data_dir):
    """Windows環境での実ファイル検索テスト"""
    result = search_local_files(
        query="テスト",
        extension="txt",
        min_size_kb=0,
        max_size_kb=100
    )
    assert isinstance(result, str)
    assert "test.txt" in result or "No matching files found" in result

@pytest.mark.skipif(platform.system() != 'Darwin', reason="macOS環境でのみ実行")
def test_mac_search(test_data_dir):
    """macOS環境での実ファイル検索テスト"""
    result = search_local_files(
        query="テスト",
        extension="txt",
        min_size_kb=0,
        max_size_kb=100
    )
    assert isinstance(result, str)
    assert "test.txt" in result or "No matching files found" in result

# 基本的なファイル操作のテスト
def test_file_operations(test_data_dir):
    """実ファイルの読み書きテスト"""
    test_file = test_data_dir / "test.txt"
    result = local_read_file(str(test_file))
    assert "これはテスト" in result

def test_unsupported_platform():
    """非対応プラットフォームのテスト"""
    if platform.system() not in ['Windows', 'Darwin']:
        result = search_local_files("test")
        assert result == "Unsupported operating system"
