from __future__ import annotations

import argparse
import csv
import json
import re
import uuid
from dataclasses import dataclass
from datetime import datetime
from functools import lru_cache
from pathlib import Path
from typing import Any

from openai import OpenAI

from layout_trace import build_trace_results, format_trace_markdown
from pipeline_extensions import ingest_and_route
from settings import (
    ConversionSettings,
    ModelSettings,
    OcrSettings,
    RoutingSettings,
    load_app_settings,
)


@dataclass(frozen=True)
class BranchConfig:
    label: str
    match_keywords: list[str]
    model_key: str
    field_map: dict[str, str]
    output_file_name: str


@dataclass(frozen=True)
class PromptConfig:
    system_prompt: str
    instructions: list[str]
    user_prompt_template: str


@dataclass(frozen=True)
class WorkflowConfig:
    csv_headers: list[str]
    branches: list[BranchConfig]


@dataclass(frozen=True)
class RuntimeConfig:
    workdir: Path
    temp_dir: Path
    output_dir: Path
    default_question: str
    prompt_config_file: Path
    workflow_config_file: Path
    prompt_config: PromptConfig
    workflow_config: WorkflowConfig
    qwen: ModelSettings
    deepseek: ModelSettings
    routing: RoutingSettings
    ocr: OcrSettings
    conversion: ConversionSettings
    save_intermediate: bool


def load_prompt_config(prompt_config_file: str | Path) -> PromptConfig:
    path = Path(prompt_config_file).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Prompt config not found: {path}")
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("Prompt config must be a JSON object.")

    system_prompt = str(payload.get("system_prompt", "")).strip()
    user_prompt_template = str(payload.get("user_prompt_template", "")).strip()
    instructions_value = payload.get("instructions", [])
    if not system_prompt:
        raise ValueError("Prompt config is missing system_prompt.")
    if not user_prompt_template:
        raise ValueError("Prompt config is missing user_prompt_template.")
    if not isinstance(instructions_value, list) or not all(isinstance(item, str) for item in instructions_value):
        raise ValueError("Prompt config instructions must be a list of strings.")

    return PromptConfig(
        system_prompt=system_prompt,
        instructions=[item.strip() for item in instructions_value if item.strip()],
        user_prompt_template=user_prompt_template,
    )


def load_workflow_config(workflow_config_file: str | Path) -> WorkflowConfig:
    path = Path(workflow_config_file).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Workflow config not found: {path}")
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("Workflow config must be a JSON object.")

    csv_headers_value = payload.get("csv_headers", [])
    branches_value = payload.get("branches", [])
    if not isinstance(csv_headers_value, list) or not all(isinstance(item, str) for item in csv_headers_value):
        raise ValueError("workflow.csv_headers must be a list of strings.")
    if not isinstance(branches_value, list) or not branches_value:
        raise ValueError("workflow.branches must be a non-empty list.")

    branches: list[BranchConfig] = []
    for item in branches_value:
        if not isinstance(item, dict):
            raise ValueError("Each branch must be a JSON object.")
        label = str(item.get("label", "")).strip()
        model_key = str(item.get("model_key", "")).strip()
        output_file_name = str(item.get("output_file_name", "")).strip()
        field_map = item.get("field_map", {})
        match_keywords = item.get("match_keywords", [])

        if not label:
            raise ValueError("Branch label is required.")
        if not model_key:
            raise ValueError(f"Branch {label} is missing model_key.")
        if not output_file_name:
            raise ValueError(f"Branch {label} is missing output_file_name.")
        if not isinstance(field_map, dict) or not all(isinstance(k, str) and isinstance(v, str) for k, v in field_map.items()):
            raise ValueError(f"Branch {label} field_map must be a string-to-string map.")
        if not isinstance(match_keywords, list) or not all(isinstance(keyword, str) for keyword in match_keywords):
            raise ValueError(f"Branch {label} match_keywords must be a list of strings.")

        branches.append(
            BranchConfig(
                label=label,
                match_keywords=[keyword.strip() for keyword in match_keywords if keyword.strip()],
                model_key=model_key,
                field_map={str(k).strip(): str(v).strip() for k, v in field_map.items()},
                output_file_name=output_file_name,
            )
        )

    return WorkflowConfig(
        csv_headers=[item.strip() for item in csv_headers_value if item.strip()],
        branches=branches,
    )


def load_runtime_config(config_file: str | Path | None = None) -> RuntimeConfig:
    app_settings = load_app_settings(config_file)
    return RuntimeConfig(
        workdir=app_settings.workdir,
        temp_dir=app_settings.temp_dir,
        output_dir=app_settings.output_dir,
        default_question=app_settings.default_question,
        prompt_config_file=app_settings.prompt_config_file,
        workflow_config_file=app_settings.workflow_config_file,
        prompt_config=load_prompt_config(app_settings.prompt_config_file),
        workflow_config=load_workflow_config(app_settings.workflow_config_file),
        qwen=app_settings.qwen,
        deepseek=app_settings.deepseek,
        routing=app_settings.routing,
        ocr=app_settings.ocr,
        conversion=app_settings.conversion,
        save_intermediate=app_settings.save_intermediate,
    )


@lru_cache(maxsize=4)
def get_runtime_config(config_file: str | Path | None = None) -> RuntimeConfig:
    return load_runtime_config(config_file)


def merge_field_content(field_map: dict[str, str], section_map: dict[str, str]) -> dict[str, str]:
    return {field: section_map[title] for field, title in field_map.items() if title and title in section_map}


def clean_json_string(text: str) -> str:
    cleaned = text.strip()
    if "<回答>" in cleaned and "</回答>" in cleaned:
        cleaned = cleaned.split("<回答>", 1)[1].split("</回答>", 1)[0].strip()
    cleaned = cleaned.strip("`").strip()
    if cleaned.lower().startswith("json"):
        cleaned = cleaned[4:].lstrip()
    try:
        json.loads(cleaned)
        return cleaned
    except json.JSONDecodeError:
        match = re.search(r"(\{.*\}|\[.*\])", cleaned, re.S)
        if not match:
            raise
        candidate = match.group(1).strip()
        json.loads(candidate)
        return candidate


def parse_model_json(text: str) -> dict[str, Any]:
    payload = json.loads(clean_json_string(text))
    if not isinstance(payload, dict):
        raise ValueError("Model output is not a JSON object.")
    return payload


def _date_like(text_value: str) -> bool:
    return any(
        re.match(pattern, text_value.strip())
        for pattern in (
            r"^\d{4}-\d{1,2}-\d{1,2}$",
            r"^\d{4}/\d{1,2}/\d{1,2}$",
            r"^\d{4}\.\d{1,2}\.\d{1,2}$",
            r"^\d{4}年\d{1,2}月\d{1,2}日$",
        )
    )


def write_csv(data: dict[str, Any], output_file: Path, headers: list[str]) -> Path:
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with output_file.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        row: list[str] = []
        for header in headers:
            value = data.get(header, "")
            if value is None:
                row.append(" ")
            elif isinstance(value, (list, dict)):
                row.append(json.dumps(value, ensure_ascii=False))
            else:
                text_value = str(value)
                row.append("'" + text_value if _date_like(text_value) else text_value)
        writer.writerow(row)
    return output_file.resolve()


def select_branch(question: str, branches: list[BranchConfig]) -> BranchConfig:
    lowered_question = question.lower()
    for branch in branches:
        if any(keyword.lower() in lowered_question for keyword in branch.match_keywords):
            return branch
    return branches[0]


def select_branch_by_label(label: str, branches: list[BranchConfig]) -> BranchConfig:
    for branch in branches:
        if branch.label == label:
            return branch
    raise ValueError(f"Unsupported branch label: {label}")


def build_prompt(question: str, branch: BranchConfig, document_markdown: str, prompt_config: PromptConfig) -> str:
    rules = {
        "question": question,
        "branch": branch.label,
        "field_title_map": branch.field_map,
        "instructions": prompt_config.instructions,
    }
    return prompt_config.user_prompt_template.format(
        rules_json=json.dumps(rules, ensure_ascii=False, indent=2),
        document_markdown=document_markdown,
    )


def create_client(settings: ModelSettings) -> OpenAI:
    if not settings.api_key:
        raise ValueError(f"{settings.name} is missing API key.")
    if not settings.base_url:
        raise ValueError(f"{settings.name} is missing base_url.")
    base_url = settings.base_url.strip()
    if not base_url.startswith(("http://", "https://")):
        base_url = f"https://{base_url}"
    return OpenAI(api_key=settings.api_key, base_url=base_url, timeout=settings.timeout)


def call_model(settings: ModelSettings, prompt: str, system_prompt: str) -> str:
    client = create_client(settings)
    response = client.chat.completions.create(
        model=settings.name,
        temperature=0,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt},
        ],
    )
    return (response.choices[0].message.content or "").strip()


def save_intermediate(
    config: RuntimeConfig,
    branch: BranchConfig,
    question: str,
    markdown: str,
    raw_output: str,
    trace: str,
    route_payload: dict[str, Any],
) -> Path:
    run_dir = config.output_dir / datetime.now().strftime("%Y%m%d_%H%M%S") / uuid.uuid4().hex[:8]
    run_dir.mkdir(parents=True, exist_ok=True)
    (run_dir / "question.txt").write_text(question, encoding="utf-8")
    (run_dir / "branch.txt").write_text(branch.label, encoding="utf-8")
    (run_dir / "route.json").write_text(json.dumps(route_payload, ensure_ascii=False, indent=2), encoding="utf-8")
    (run_dir / "document.md").write_text(markdown, encoding="utf-8")
    (run_dir / "raw_output.txt").write_text(raw_output, encoding="utf-8")
    (run_dir / "trace_output.md").write_text(trace, encoding="utf-8")
    return run_dir


def run_xiangmuku_kexingxing(
    question: str | None = None,
    files: list[str | Path] | None = None,
    config_file: str | Path | None = None,
    config: RuntimeConfig | None = None,
) -> dict[str, Any]:
    runtime_config = config or load_runtime_config(config_file)
    normalized_question = (question or runtime_config.default_question or "").strip()
    if not normalized_question:
        normalized_question = "请根据上传文档自动识别类型并抽取字段"

    normalized_files = [Path(item).expanduser().resolve() for item in (files or [])]
    if not normalized_files:
        raise ValueError("Missing uploaded files.")
    missing = [str(path) for path in normalized_files if not path.exists()]
    if missing:
        raise FileNotFoundError(f"These files do not exist: {missing}")

    prompt_config = runtime_config.prompt_config
    workflow_config = runtime_config.workflow_config

    pipeline_result = ingest_and_route(
        files=normalized_files,
        question=normalized_question,
        temp_dir=runtime_config.temp_dir,
        qwen=runtime_config.qwen,
        deepseek=runtime_config.deepseek,
        routing=runtime_config.routing,
        ocr=runtime_config.ocr,
        conversion=runtime_config.conversion,
        branches=workflow_config.branches,
    )
    branch = select_branch_by_label(pipeline_result.selected_branch_label, workflow_config.branches)
    prompt = build_prompt(normalized_question, branch, pipeline_result.combined_markdown, prompt_config)
    model_settings = runtime_config.deepseek if branch.model_key == "deepseek" else runtime_config.qwen
    raw_output = call_model(model_settings, prompt, prompt_config.system_prompt)
    parsed = parse_model_json(raw_output)
    trace_source = merge_field_content(branch.field_map, pipeline_result.section_map)
    trace_results = build_trace_results(parsed, trace_source, pipeline_result.layout_index)
    answer_markdown = format_trace_markdown(trace_results)
    csv_path = write_csv(parsed, runtime_config.output_dir / branch.output_file_name, workflow_config.csv_headers)

    route_payload = {
        "selected_branch": branch.label,
        "route_confidence": pipeline_result.route_confidence,
        "route_reason": pipeline_result.route_reason,
        "file_summaries": pipeline_result.file_summaries,
    }
    run_dir = (
        save_intermediate(
            runtime_config,
            branch,
            normalized_question,
            pipeline_result.combined_markdown,
            raw_output,
            answer_markdown,
            route_payload,
        )
        if runtime_config.save_intermediate
        else None
    )

    return {
        "question": normalized_question,
        "branch": branch.label,
        "selected_branch": branch.label,
        "route_confidence": pipeline_result.route_confidence,
        "route_reason": pipeline_result.route_reason,
        "used_model": model_settings.name,
        "files": [str(path) for path in normalized_files],
        "csv_path": str(csv_path),
        "run_dir": str(run_dir) if run_dir else None,
        "raw_model_output": raw_output,
        "parsed_json": parsed,
        "answer_markdown": answer_markdown,
    }


def _main() -> int:
    parser = argparse.ArgumentParser(description="Run local project-library feasibility workflow.")
    parser.add_argument("question", nargs="?")
    parser.add_argument("--file", action="append", dest="files", default=[])
    parser.add_argument("--config", default=None)
    args = parser.parse_args()
    result = run_xiangmuku_kexingxing(args.question, args.files, args.config)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
