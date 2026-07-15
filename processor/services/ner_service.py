from __future__ import annotations

import os
import re
import threading
from dataclasses import dataclass
from typing import Callable


DEFAULT_MODEL = "fastino/gliner2-multi-v1"
CYRILLIC_PATTERN = re.compile(r"[А-Яа-яЁёӘәҒғҚқҢңӨөҰұҮүҺһІі]")
LATIN_FULL_NAME_PATTERN = re.compile(r"[A-Za-z'’-]+(?:\s+[A-Za-z'’-]+)+")
LABELED_FIO_PATTERN = re.compile(
    r"\bФИО(?:\s+[А-Яа-яЁёӘәҒғҚқҢңӨөҰұҮүҺһІіA-Za-z]+)?[\s*:=—-]+([^;,]+)",
    flags=re.IGNORECASE,
)


@dataclass(frozen=True)
class NameExtraction:
    names: tuple[str, ...]
    reason: str = ""


_model = None
_model_name = ""
_model_load_lock = threading.Lock()
_inference_lock = threading.Lock()


class PersonNameExtractor:
    def __init__(self, model_name: str | None = None, batch_size: int | None = None) -> None:
        self.model_name = model_name or os.getenv("NER_MODEL", DEFAULT_MODEL)
        self.batch_size = batch_size or int(os.getenv("NER_BATCH_SIZE", "8"))

    def extract_many(
        self,
        texts: list[str],
        should_cancel: Callable[[], bool] | None = None,
    ) -> list[NameExtraction]:
        if not texts:
            return []
        model = self._get_model()
        output: list[NameExtraction] = []
        chunk_size = max(self.batch_size * 4, 8)
        with _inference_lock:
            for start in range(0, len(texts), chunk_size):
                if should_cancel and should_cancel():
                    raise InterruptedError("Операция отменена во время NER")
                chunk = [str(text or "") for text in texts[start : start + chunk_size]]
                results = model.batch_extract_entities(
                    chunk,
                    {"person": "Full name"},
                    include_confidence=True,
                    include_spans=True,
                    batch_size=self.batch_size,
                )
                output.extend(
                    _normalize_result(text, result)
                    for text, result in zip(chunk, results, strict=True)
                )
        return output

    def _get_model(self):
        global _model, _model_name
        with _model_load_lock:
            if _model is None or _model_name != self.model_name:
                from gliner2 import GLiNER2

                _model = GLiNER2.from_pretrained(self.model_name)
                _model_name = self.model_name
        return _model


def _normalize_result(source_text: str, result: dict) -> NameExtraction:
    entities = result.get("entities", {}).get("person", [])
    candidates: list[tuple[int, int, str]] = []
    for entity in entities:
        raw_name = entity.get("text", "") if isinstance(entity, dict) else str(entity)
        name = _clean_name(raw_name)
        if not _valid_name(name):
            continue
        start = entity.get("start", source_text.find(raw_name)) if isinstance(entity, dict) else source_text.find(raw_name)
        end = entity.get("end", start + len(raw_name)) if isinstance(entity, dict) else start + len(raw_name)
        candidates.append((start, end, name))

    candidates.sort(key=lambda item: (item[0], item[1]))
    merged: list[tuple[int, int, str]] = []
    for start, end, name in candidates:
        if merged and start >= merged[-1][1] and not source_text[merged[-1][1] : start].strip():
            old_start, _, old_name = merged[-1]
            merged[-1] = (old_start, end, f"{old_name} {name}")
        else:
            merged.append((start, end, name))

    names = [name for _, _, name in merged]
    labeled_match = LABELED_FIO_PATTERN.search(source_text)
    if labeled_match:
        labeled_name = _clean_name(labeled_match.group(1))
        if _valid_name(labeled_name):
            names = [name for name in names if name.casefold() not in labeled_name.casefold()]
            names.append(labeled_name)

    unique_names: list[str] = []
    for name in names:
        if name.casefold() not in {item.casefold() for item in unique_names}:
            unique_names.append(name)
    reason = "" if unique_names else "GLiNER 2 не обнаружил допустимое ФИО"
    return NameExtraction(tuple(unique_names), reason)


def _clean_name(value: str) -> str:
    return " ".join(str(value or "").split()).strip(" ,;:.*-\t\n")


def _valid_name(name: str) -> bool:
    if not name or any(character.isdigit() for character in name):
        return False
    return bool(CYRILLIC_PATTERN.search(name) or LATIN_FULL_NAME_PATTERN.fullmatch(name))
