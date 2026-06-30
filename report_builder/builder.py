"""レポート生成のオーケストレーション。"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .config_loader import ReportConfig, load_config
from .excel_writer import build_output_path, find_csv_files, write_report
from .paths import resolve_app_path


@dataclass
class BuildResult:
    csv_path: Path
    output_path: Path


def build_reports(config_path: str | Path, csv_file: str | Path | None = None) -> list[BuildResult]:
    config = load_config(config_path)

    if csv_file:
        targets = [resolve_app_path(csv_file)]
    else:
        targets = find_csv_files(config)

    results: list[BuildResult] = []
    for index, csv_path in enumerate(targets):
        if not csv_path.exists():
            raise FileNotFoundError(f"CSVファイルが見つかりません: {csv_path}")

        if csv_file or len(targets) == 1:
            output_path = build_output_path(config, csv_path)
        else:
            stem = csv_path.stem
            if config.output.include_timestamp:
                from datetime import datetime

                timestamp = datetime.now().strftime(config.output.timestamp_format)
                filename = f"{config.output.base_name}_{stem}_{timestamp}.xlsx"
            else:
                filename = f"{config.output.base_name}_{stem}.xlsx"
            output_path = config.excel_output_dir / filename

        saved = write_report(config, csv_path, output_path)
        results.append(BuildResult(csv_path=csv_path, output_path=saved))

    return results
