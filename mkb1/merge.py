from __future__ import annotations

import json
from pathlib import Path
from typing import Any

INPUT_FIELDS = ("zipName", "folderName", "fileName", "fileFetchPath")


def normalize_single_payload(payload: dict[str, Any]) -> dict[str, list[str]]:
    if not isinstance(payload, dict):
        raise ValueError("task must be an object")
    normalized: dict[str, list[str]] = {}
    missing = [field for field in INPUT_FIELDS if field not in payload]
    if missing:
        raise ValueError(f"task missing required fields: {', '.join(missing)}")
    for field in INPUT_FIELDS:
        value = payload.get(field, [])
        if not isinstance(value, list):
            raise ValueError(f"{field} must be a list")
        normalized[field] = ["" if item is None else str(item) for item in value]
    lengths = {len(normalized[field]) for field in INPUT_FIELDS}
    if len(lengths) != 1:
        raise ValueError("zipName, folderName, fileName, and fileFetchPath must have the same length")
    if not lengths or max(lengths) == 0:
        raise ValueError("task input batch is empty")
    return normalized


def extract_tasks(request: Any) -> list[dict[str, list[str]]]:
    if isinstance(request, list):
        tasks = request
    elif isinstance(request, dict) and "tasks" in request:
        tasks = request.get("tasks") or []
    else:
        tasks = [request]
    normalized_tasks: list[dict[str, list[str]]] = []
    for index, task in enumerate(tasks):
        try:
            normalized_tasks.append(normalize_single_payload(task))
        except Exception as exc:
            raise ValueError(f"task #{index + 1} is invalid: {exc}") from exc
    if not normalized_tasks:
        raise ValueError("tasks is empty")
    return normalized_tasks


def merge_results(results: list[dict[str, Any]]) -> dict[str, Any]:
    success_count = sum(1 for item in results if item.get("status") == "success")
    fail_count = len(results) - success_count
    return {
        "status": "success" if fail_count == 0 else "partial_success" if success_count else "fail",
        "task_count": len(results),
        "success_count": success_count,
        "fail_count": fail_count,
        "results": results,
    }


def load_request(source: str) -> Any:
    path = Path(source)
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return json.loads(source)
