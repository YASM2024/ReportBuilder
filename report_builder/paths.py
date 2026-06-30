"""アプリケーションルートとユーザーデータパスの解決。

exe 化（PyInstaller onefile）後は config / import / export / templates を
実行ファイルと同じ階層に置き、カレントディレクトリではなく exe 基準で解決する。
ソースコードは _MEIPASS に展開されるが、ユーザーデータは exe 隣を参照する。
"""

from __future__ import annotations

import sys
from pathlib import Path


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False)) or hasattr(sys, "_MEIPASS")


def get_bundle_root() -> Path:
    """バンドルされたソースのルート（開発時はプロジェクトルート）。"""
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        return Path(meipass)
    return Path(__file__).resolve().parent.parent


def get_app_root() -> Path:
    """ユーザーデータの基準ディレクトリ（プロジェクトルート、または exe のあるフォルダ）。"""
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return get_bundle_root()


def resolve_app_path(value: str | Path) -> Path:
    """app_root 基準でパスを解決する（絶対パスはそのまま）。"""
    path = Path(value)
    if path.is_absolute():
        return path
    return (get_app_root() / path).resolve()
