from __future__ import annotations

import re
from functools import lru_cache
from dataclasses import dataclass
from typing import Any

import pymorphy3


NAME_LETTERS_PATTERN = re.compile(r"[A-Za-zА-Яа-яЁёӘәҒғҚқҢңӨөҰұҮүҺһІі]")
NAME_TOKEN_PATTERN = re.compile(r"[A-Za-zА-Яа-яЁёӘәҒғҚқҢңӨөҰұҮүҺһІі]+")


@dataclass(frozen=True)
class NameMatchEvidence:
    criterion: str
    strength: int


def normalize_person_token(value: Any) -> str:
    return str(value or "").casefold().replace("ё", "е")


def person_name_tokens(value: Any) -> tuple[str, ...]:
    return tuple(
        normalize_person_token(token)
        for token in NAME_TOKEN_PATTERN.findall(str(value or ""))
        if NAME_LETTERS_PATTERN.search(token)
    )


def name_lookup_keys(value: Any) -> set[str]:
    keys: set[str] = set()
    for token in person_name_tokens(value):
        if len(token) < 2:
            continue
        keys.add(token)
        if len(token) >= 5:
            keys.add(f"^{token[:5]}")
    return keys


@lru_cache(maxsize=1)
def _morph_analyzer() -> pymorphy3.MorphAnalyzer:
    return pymorphy3.MorphAnalyzer()


@lru_cache(maxsize=20_000)
def _nominative_person_token(token: str, preserve_feminine_name: bool) -> str:
    if len(token) <= 1 or not re.fullmatch(NAME_TOKEN_PATTERN, token):
        return token
    if re.search(r"[әғқңөұүһі]", token) or token.endswith(
        ("ұлы", "улы", "қызы", "кызы")
    ):
        return token
    parses = _morph_analyzer().parse(token)
    name_parses = [
        parsed
        for parsed in parses
        if parsed.tag.grammemes & {"Name", "Surn", "Patr"}
        and "sing" in parsed.tag.grammemes
    ]
    # A nominative interpretation can rank below a more common Russian word
    # interpretation, especially for Kazakh and feminine names. In that case
    # the original token is already safe and must not be changed.
    if any("nomn" in parsed.tag.grammemes for parsed in name_parses):
        return token
    if preserve_feminine_name and token.endswith(("а", "я", "е")) and any(
        "Name" in parsed.tag.grammemes for parsed in name_parses
    ):
        return token
    for parsed in name_parses:
        nominative = parsed.inflect({"nomn"})
        if nominative:
            return normalize_person_token(nominative.word)
    return token


def nominative_person_name_tokens(value: Any) -> tuple[str, ...]:
    """Return name tokens with Russian grammatical endings removed."""
    tokens = person_name_tokens(value)
    has_male_patronymic = any(
        re.search(r"(?:вич|вича|вичу|вичем|виче)$", token)
        for token in tokens
    )
    has_female_marker = any(
        re.search(r"(?:вна|вны|вной|чна|чны|чной|ова|ева|ина|ская)$", token)
        for token in tokens
    )
    preserve_feminine_name = has_female_marker and not has_male_patronymic
    return tuple(
        _nominative_person_token(token, preserve_feminine_name)
        for token in tokens
    )


def nominative_name_lookup_keys(value: Any) -> set[str]:
    keys: set[str] = set()
    for token in nominative_person_name_tokens(value):
        if len(token) < 2:
            continue
        keys.add(token)
        if len(token) >= 5:
            keys.add(f"^{token[:5]}")
    return keys


def match_person_name(reference_name: Any, detected_name: Any) -> NameMatchEvidence | None:
    reference = person_name_tokens(reference_name)
    detected = person_name_tokens(detected_name)
    if len(reference) < 2 or len(detected) < 2:
        return None

    tokenizations = [detected]
    expanded = _expand_compact_initials(detected, reference)
    if expanded != detected:
        tokenizations.append(expanded)

    matches = [
        evidence
        for tokens in tokenizations
        if (evidence := _match_tokens(reference, tokens)) is not None
    ]
    if not matches:
        return None
    return max(matches, key=lambda item: item.strength)


def match_person_name_nominative(
    reference_name: Any,
    detected_name: Any,
) -> NameMatchEvidence | None:
    """Retry a name match after converting both names to dictionary forms."""
    reference = nominative_person_name_tokens(reference_name)
    detected = nominative_person_name_tokens(detected_name)
    if len(reference) < 2 or len(detected) < 2:
        return None

    tokenizations = [detected]
    expanded = _expand_compact_initials(detected, reference)
    if expanded != detected:
        tokenizations.append(expanded)

    matches = [
        evidence
        for tokens in tokenizations
        if (evidence := _match_tokens(reference, tokens)) is not None
    ]
    if not matches:
        return None
    return max(matches, key=lambda item: item.strength)


def _expand_compact_initials(
    detected: tuple[str, ...],
    reference: tuple[str, ...],
) -> tuple[str, ...]:
    if len(reference) < 3 or not any(len(token) >= 4 for token in detected):
        return detected
    expanded: list[str] = []
    changed = False
    for token in detected:
        if 1 < len(token) <= 3 and all(
            any(reference_token.startswith(character) for reference_token in reference)
            for character in token
        ):
            expanded.extend(token)
            changed = True
        else:
            expanded.append(token)
    return tuple(expanded) if changed else detected


def _match_tokens(
    reference: tuple[str, ...],
    detected: tuple[str, ...],
) -> NameMatchEvidence | None:
    full_tokens = [token for token in detected if len(token) > 1]
    initial_tokens = [token for token in detected if len(token) == 1]

    # A surname alone, surname + one initial, and a fully abbreviated name are
    # deliberately too weak. Two full elements or one full element plus two
    # initials are the minimum accepted forms.
    if not full_tokens:
        return None
    if len(full_tokens) < 2 and len(initial_tokens) < 2:
        return None

    matched = _find_distinct_token_assignment(reference, detected)
    if not matched:
        return None

    if len(full_tokens) >= 3 and not initial_tokens:
        return NameMatchEvidence("Полное ФИО GLiNER", 4)
    if len(full_tokens) >= 2 and initial_tokens:
        return NameMatchEvidence("Два элемента ФИО + инициалы GLiNER", 3)
    if len(full_tokens) >= 2:
        return NameMatchEvidence("Два элемента ФИО GLiNER", 2)
    return NameMatchEvidence("Элемент ФИО + два инициала GLiNER", 2)


def _find_distinct_token_assignment(
    reference: tuple[str, ...],
    detected: tuple[str, ...],
) -> bool:
    options: list[list[int]] = []
    for detected_token in detected:
        matching_indexes = [
            index
            for index, reference_token in enumerate(reference)
            if _person_tokens_match(reference_token, detected_token)
        ]
        if not matching_indexes:
            return False
        options.append(matching_indexes)

    options.sort(key=len)

    def assign(position: int, used: set[int]) -> bool:
        if position == len(options):
            return True
        for index in options[position]:
            if index in used:
                continue
            used.add(index)
            if assign(position + 1, used):
                return True
            used.remove(index)
        return False

    return assign(0, set())


def _person_tokens_match(reference: str, detected: str) -> bool:
    if len(detected) == 1:
        return reference.startswith(detected)
    if reference == detected:
        return True
    if min(len(reference), len(detected)) < 5 or abs(len(reference) - len(detected)) > 3:
        return False
    common = 0
    for left, right in zip(reference, detected):
        if left != right:
            break
        common += 1
    # Allows common Russian grammatical endings only when almost the whole
    # token agrees; this is intentionally stricter than general fuzzy search.
    return common >= 3 and common >= min(len(reference), len(detected)) - 2
