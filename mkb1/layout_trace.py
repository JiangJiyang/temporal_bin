from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class Rect:
    x0: float
    y0: float
    x1: float
    y1: float


@dataclass(frozen=True)
class TraceResult:
    field_name: str
    value: str
    page: int | None
    page_reason: str | None
    quote: str
    rects: list[Rect]


def _normalize_text(text: str) -> str:
    cleaned = text.replace("|", " ").replace("#", " ").replace("<br>", " ")
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip().lower()


def _tokenize(text: str) -> list[str]:
    return re.findall(r"[\u4e00-\u9fff]+|[A-Za-z0-9_]+", _normalize_text(text))


def _score_match(candidate: str, line_text: str) -> int:
    normalized_candidate = _normalize_text(candidate)
    normalized_line = _normalize_text(line_text)
    if not normalized_candidate or not normalized_line:
        return 0
    if normalized_candidate in normalized_line:
        return len(normalized_candidate) + 50
    if normalized_line in normalized_candidate:
        return len(normalized_line) + 40
    candidate_tokens = set(_tokenize(normalized_candidate))
    line_tokens = set(_tokenize(normalized_line))
    if not candidate_tokens or not line_tokens:
        return 0
    return len(candidate_tokens & line_tokens) * 10


def _candidate_fragments(value: Any, source_excerpt: str) -> list[str]:
    fragments: list[str] = []
    if source_excerpt:
        for line in source_excerpt.splitlines():
            stripped = line.strip()
            if len(_normalize_text(stripped)) >= 4:
                fragments.append(stripped)
    if value is not None:
        stripped_value = str(value).strip()
        if stripped_value:
            fragments.append(stripped_value)
    seen: set[str] = set()
    result: list[str] = []
    for item in fragments:
        normalized = _normalize_text(item)
        if normalized and normalized not in seen:
            seen.add(normalized)
            result.append(item)
    return result[:8]


def _rect_from_payload(payload: dict[str, float]) -> Rect:
    return Rect(
        x0=float(payload.get("x0", 0.0)),
        y0=float(payload.get("y0", 0.0)),
        x1=float(payload.get("x1", 1.0)),
        y1=float(payload.get("y1", 1.0)),
    )


def _resolve_quote_and_rects(value: Any, source_excerpt: str, layout_index: dict[str, Any]) -> tuple[int | None, str, list[Rect], str | None]:
    lines = list(layout_index.get("lines", []))
    if not lines:
        return None, source_excerpt.strip(), [], "未生成版面定位数据"

    candidates = _candidate_fragments(value, source_excerpt)
    if not candidates:
        return None, source_excerpt.strip() or str(value).strip(), [], "缺少可定位文本"

    best_line: dict[str, Any] | None = None
    best_score = -1
    for candidate in candidates:
        for line in lines:
            score = _score_match(candidate, str(line.get("text", "")))
            if score > best_score:
                best_score = score
                best_line = line
    if best_line is None or best_score <= 0:
        if source_excerpt.strip():
            reason = "未在版面文本中匹配到来源片段"
        else:
            reason = "缺少来源章节片段，无法定位页码"
        return None, source_excerpt.strip() or str(value).strip(), [], reason

    page = int(best_line.get("page"))
    same_page = [line for line in lines if int(line.get("page")) == page]

    matched_lines = []
    for line in same_page:
        line_score = max(_score_match(candidate, str(line.get("text", ""))) for candidate in candidates) if candidates else 0
        if line_score > 0:
            matched_lines.append((line_score, line))
    matched_lines.sort(key=lambda item: (-item[0], item[1]["rect"]["y0"], item[1]["rect"]["x0"]))
    selected = [item[1] for item in matched_lines[:4]] or [best_line]

    quote = "\n".join(str(item.get("text", "")).strip() for item in selected if str(item.get("text", "")).strip()).strip()
    rects = [_rect_from_payload(item.get("rect", {})) for item in selected]
    return page, quote or (source_excerpt.strip() or str(value).strip()), rects, None


def build_trace_results(parsed: dict[str, Any], source_map: dict[str, str], layout_index: dict[str, Any]) -> list[TraceResult]:
    results: list[TraceResult] = []
    for field_name, value in parsed.items():
        if value is None or value == "":
            continue
        source_excerpt = source_map.get(field_name, "")
        page, quote, rects, page_reason = _resolve_quote_and_rects(value, source_excerpt, layout_index)
        results.append(
            TraceResult(
                field_name=field_name,
                value=str(value),
                page=page,
                page_reason=page_reason,
                quote=quote,
                rects=rects,
            )
        )
    return results


def format_trace_markdown(trace_results: list[TraceResult]) -> str:
    lines: list[str] = []
    for item in trace_results:
        lines.append(f"### {item.field_name}")
        lines.append(item.value)
        lines.append("")
        lines.append(f"page: {item.page if item.page is not None else (item.page_reason or '未定位页码')}")
        lines.append("")
        lines.append("quote:")
        lines.append(item.quote or "")
        lines.append("")
        lines.append("position.rects:")
        if item.rects:
            for rect in item.rects:
                lines.append(f"- x0: {rect.x0:.6f}")
                lines.append(f"  y0: {rect.y0:.6f}")
                lines.append(f"  x1: {rect.x1:.6f}")
                lines.append(f"  y1: {rect.y1:.6f}")
        else:
            lines.append("- []")
        lines.append("")
    return "\n".join(lines).strip()
