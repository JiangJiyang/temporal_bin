from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import main
from merge import extract_tasks, merge_results


def process_tasks(
    request: Any,
    config_path: str | Path = Path(__file__).resolve().parent / "config.json",
    override_test_mode: bool | None = None,
) -> dict[str, Any]:
    tasks = extract_tasks(request)
    results = [
        main.process_request(task, config_path=config_path, persist_result=False, override_test_mode=override_test_mode)
        for task in tasks
    ]
    merged = merge_results(results)
    if len(results) == 1:
        single = dict(results[0])
        single["task_count"] = 1
        return single
    return merged


def run_tasks(request: Any, config_path: str = "./config.json", override_test_mode: bool | None = None) -> dict[str, Any]:
    return process_tasks(request, config_path=config_path, override_test_mode=override_test_mode)


def main_cli() -> int:
    import argparse
    import sys

    parser = argparse.ArgumentParser(description="统一 task 入口")
    parser.add_argument("--config", default=str(Path(__file__).resolve().parent / "config.json"))
    parser.add_argument("--request-json", help="task 请求 JSON 字符串，支持单 task / tasks 包装 / task 数组")
    parser.add_argument("--request-file", help="task 请求 JSON 文件，支持单 task / tasks 包装 / task 数组")
    parser.add_argument("--stdin-json", action="store_true", help="从标准输入读取 task 请求 JSON")
    parser.add_argument("--test-mode", choices=("true", "false"))
    args = parser.parse_args()

    override_test_mode = None if args.test_mode is None else args.test_mode == "true"

    content = ""
    if args.request_json:
        content = args.request_json
    elif args.request_file:
        content = Path(args.request_file).read_text(encoding="utf-8")
    elif args.stdin_json or not sys.stdin.isatty():
        content = sys.stdin.read().strip()

    if not content:
        parser.error("未提供 task 请求 JSON")

    result = process_tasks(json.loads(content), config_path=args.config, override_test_mode=override_test_mode)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    status = str(result.get("status", ""))
    return 0 if status in {"success", "partial_success"} else 1


if __name__ == "__main__":
    raise SystemExit(main_cli())
