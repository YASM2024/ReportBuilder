#!/usr/bin/env python3
"""CSV → Excel レポート生成ツール。"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from report_builder.builder import build_reports


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="設定ファイルに基づき、CSVをExcelレポートに成型します。",
    )
    parser.add_argument(
        "-c",
        "--config",
        required=True,
        help="レポート設定YAMLのパス",
    )
    parser.add_argument(
        "--csv",
        help="処理するCSVファイル（省略時は設定の入力フォルダ内を一括処理）",
    )
    args = parser.parse_args(argv)

    try:
        results = build_reports(args.config, args.csv)
    except (FileNotFoundError, ValueError, KeyError, IndexError) as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 1

    for result in results:
        print(f"生成完了: {result.output_path}")
        print(f"  入力: {result.csv_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
