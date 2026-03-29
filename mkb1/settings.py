from __future__ import annotations

import os
import tomllib
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_APP_CONFIG_FILE = PROJECT_ROOT / "config" / "app.toml"


def _load_dotenv(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}

    values: dict[str, str] = {}
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[7:].lstrip()
        if "=" not in line:
            continue
        name, value = line.split("=", 1)
        name = name.strip()
        if not name:
            continue
        values[name] = value.strip().strip('"').strip("'")
    return values


DOTENV_VALUES = _load_dotenv(PROJECT_ROOT / ".env")


def _bool_value(value: Any, default: bool = False) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "on"}


def _setting(name: str, default: Any = "") -> Any:
    return os.getenv(name, DOTENV_VALUES.get(name, default))


def _resolve_path(root_dir: Path, raw_path: str | Path | None, default: str | Path) -> Path:
    candidate = Path(raw_path or default).expanduser()
    if not candidate.is_absolute():
        candidate = (root_dir / candidate).resolve()
    return candidate


def _resolve_existing_path(root_dir: Path, raw_path: str | Path | None, default: str | Path) -> Path:
    default_path = _resolve_path(root_dir, default, default)
    candidate = _resolve_path(root_dir, raw_path, default)
    if candidate.exists() or not default_path.exists():
        return candidate
    return default_path


def _resolve_workflow_name(workflow_base_dir: Path, raw_name: str | None, fallback_name: str) -> str:
    candidate = str(raw_name or "").strip()
    fallback = fallback_name.strip()
    if candidate and (workflow_base_dir / candidate).exists():
        return candidate
    if fallback and (workflow_base_dir / fallback).exists():
        return fallback
    return candidate or fallback


def _dict_section(value: Any) -> dict[str, Any]:
    return value if isinstance(value, dict) else {}


@dataclass(frozen=True)
class ModelSettings:
    provider: str
    name: str
    api_key: str
    base_url: str
    timeout: int


@dataclass(frozen=True)
class RoutingSettings:
    model_key: str
    model_name: str | None
    temperature: float
    max_chars: int
    system_prompt: str
    fallback_to_keyword: bool


@dataclass(frozen=True)
class OcrSettings:
    enabled: bool
    provider: str
    language: str
    dpi: int


@dataclass(frozen=True)
class ConversionSettings:
    soffice_command: str
    use_office_com_on_windows: bool


@dataclass(frozen=True)
class AppSettings:
    env: str
    debug: bool
    root_dir: Path
    config_file: Path
    workdir: Path
    temp_dir: Path
    output_dir: Path
    default_workflow: str
    default_question: str
    default_input_file: str | None
    prompt_config_file: Path
    workflow_config_file: Path
    qwen: ModelSettings
    deepseek: ModelSettings
    routing: RoutingSettings
    ocr: OcrSettings
    conversion: ConversionSettings
    log_level: str
    save_intermediate: bool


def _load_model_settings(section: dict[str, Any], prefix: str) -> ModelSettings:
    return ModelSettings(
        provider=str(_setting(f"{prefix}_PROVIDER", section.get("provider", ""))).strip(),
        name=str(_setting(f"{prefix}_MODEL", section.get("model", ""))).strip(),
        api_key=str(_setting(f"{prefix}_API_KEY", section.get("api_key", ""))).strip(),
        base_url=str(_setting(f"{prefix}_BASE_URL", section.get("base_url", ""))).strip(),
        timeout=int(_setting(f"{prefix}_TIMEOUT", section.get("timeout", 300))),
    )


def load_app_settings(config_file: str | Path | None = None) -> AppSettings:
    config_path = _resolve_path(
        PROJECT_ROOT,
        config_file or _setting("APP_CONFIG_FILE", str(DEFAULT_APP_CONFIG_FILE)),
        DEFAULT_APP_CONFIG_FILE,
    )
    payload = tomllib.loads(config_path.read_text(encoding="utf-8"))

    app_section = _dict_section(payload.get("app"))
    path_section = _dict_section(payload.get("paths"))
    workflow_section = _dict_section(payload.get("workflow"))
    model_section = _dict_section(payload.get("models"))
    qwen_section = _dict_section(model_section.get("qwen"))
    deepseek_section = _dict_section(model_section.get("deepseek"))
    routing_section = _dict_section(payload.get("routing"))
    ocr_section = _dict_section(payload.get("ocr"))
    conversion_section = _dict_section(payload.get("conversion"))

    workdir = _resolve_path(PROJECT_ROOT, _setting("WORKDIR", path_section.get("workdir", "local_data")), "local_data")
    temp_dir = _resolve_path(PROJECT_ROOT, _setting("TEMP_DIR", path_section.get("temp_dir", workdir / "tmp")), workdir / "tmp")
    output_dir = _resolve_path(PROJECT_ROOT, _setting("OUTPUT_DIR", path_section.get("output_dir", workdir / "output")), workdir / "output")

    workflow_base_dir = _resolve_path(PROJECT_ROOT, workflow_section.get("base_dir", "config/workflows"), "config/workflows")
    configured_workflow = str(app_section.get("default_workflow", "xiangmuku_kexingxing")).strip() or "xiangmuku_kexingxing"
    default_workflow = _resolve_workflow_name(
        workflow_base_dir,
        str(_setting("DEFAULT_WORKFLOW", configured_workflow)).strip(),
        configured_workflow,
    )
    prompt_default = workflow_base_dir / default_workflow / str(workflow_section.get("prompt_file", "prompt.json"))
    workflow_default = workflow_base_dir / default_workflow / str(workflow_section.get("workflow_file", "workflow.json"))

    prompt_config_file = _resolve_existing_path(
        PROJECT_ROOT,
        _setting("PROMPT_CONFIG_FILE", str(prompt_default)),
        prompt_default,
    )
    workflow_config_file = _resolve_existing_path(
        PROJECT_ROOT,
        _setting("WORKFLOW_CONFIG_FILE", str(workflow_default)),
        workflow_default,
    )

    workdir.mkdir(parents=True, exist_ok=True)
    temp_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    qwen = _load_model_settings(qwen_section, "QWEN")
    deepseek = _load_model_settings(deepseek_section, "DEEPSEEK")

    return AppSettings(
        env=str(_setting("APP_ENV", app_section.get("env", "local"))).strip(),
        debug=_bool_value(_setting("APP_DEBUG", app_section.get("debug", False))),
        root_dir=PROJECT_ROOT,
        config_file=config_path,
        workdir=workdir,
        temp_dir=temp_dir,
        output_dir=output_dir,
        default_workflow=default_workflow,
        default_question=str(_setting("DEFAULT_QUESTION", app_section.get("default_question", ""))).strip(),
        default_input_file=str(_setting("DEFAULT_INPUT_FILE", app_section.get("default_input_file", ""))).strip() or None,
        prompt_config_file=prompt_config_file,
        workflow_config_file=workflow_config_file,
        qwen=qwen,
        deepseek=deepseek,
        routing=RoutingSettings(
            model_key=str(_setting("ROUTING_MODEL_KEY", routing_section.get("model_key", "qwen"))).strip(),
            model_name=str(_setting("ROUTING_MODEL", routing_section.get("model", ""))).strip() or None,
            temperature=float(_setting("ROUTING_TEMPERATURE", routing_section.get("temperature", 0))),
            max_chars=int(_setting("ROUTING_MAX_CHARS", routing_section.get("max_chars", 6000))),
            system_prompt=str(_setting("ROUTING_SYSTEM_PROMPT", routing_section.get("system_prompt", "You route uploaded documents to one workflow branch and only return JSON."))).strip(),
            fallback_to_keyword=_bool_value(_setting("ROUTING_FALLBACK_TO_KEYWORD", routing_section.get("fallback_to_keyword", True))),
        ),
        ocr=OcrSettings(
            enabled=_bool_value(_setting("OCR_ENABLED", ocr_section.get("enabled", False))),
            provider=str(_setting("OCR_PROVIDER", ocr_section.get("provider", "rapidocr"))).strip(),
            language=str(_setting("OCR_LANGUAGE", ocr_section.get("language", "ch"))).strip(),
            dpi=int(_setting("OCR_DPI", ocr_section.get("dpi", 220))),
        ),
        conversion=ConversionSettings(
            soffice_command=str(_setting("SOFFICE_COMMAND", conversion_section.get("soffice_command", "soffice"))).strip(),
            use_office_com_on_windows=_bool_value(
                _setting("USE_OFFICE_COM_ON_WINDOWS", conversion_section.get("use_office_com_on_windows", True))
            ),
        ),
        log_level=str(_setting("LOG_LEVEL", app_section.get("log_level", "INFO"))).strip(),
        save_intermediate=_bool_value(_setting("SAVE_INTERMEDIATE", app_section.get("save_intermediate", True))),
    )


@lru_cache(maxsize=4)
def get_app_settings(config_file: str | Path | None = None) -> AppSettings:
    return load_app_settings(config_file)
