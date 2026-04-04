from __future__ import annotations

import difflib
import re
from dataclasses import dataclass, field
from enum import Enum
from typing import Any

from docx import Document
from docx.shared import Pt


class ChangeType(str, Enum):
    INSERTION = "insertion"
    DELETION = "deletion"
    MODIFICATION = "modification"
    FORMATTING = "formatting"


@dataclass
class RunInfo:
    text: str
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    font_name: str | None = None
    font_size_pt: float | None = None
    font_color_rgb: tuple[int, int, int] | None = None


@dataclass
class ParagraphInfo:
    text: str
    location_desc: str = "Body"
    runs: list[RunInfo] = field(default_factory=list)
    alignment: str | None = None
    space_before_pt: float | None = None
    space_after_pt: float | None = None
    first_line_indent_pt: float | None = None
    left_indent_pt: float | None = None
    style_name: str | None = None


@dataclass
class WordToken:
    text: str
    is_whitespace: bool = False


@dataclass
class DiffSegment:
    text: str
    type: str  # "equal", "insert", "delete"


@dataclass
class FormattingRange:
    start: int
    end: int
    changes: list[str] = field(default_factory=list)


@dataclass
class ParagraphDiff:
    type: str  # "equal", "insert", "delete", "modify", "formatting"
    original_info: ParagraphInfo | None = None
    changed_info: ParagraphInfo | None = None
    location_desc: str = ""
    segments: list[DiffSegment] = field(default_factory=list)
    formatting_changes: list[str] = field(default_factory=list)
    has_run_formatting_changes: bool = False
    run_formatting_ranges: list[FormattingRange] = field(default_factory=list)
    original_para_index: int | None = None
    changed_para_index: int | None = None


@dataclass
class Change:
    type: ChangeType
    paragraph_index: int
    location_desc: str
    original_text: str
    new_text: str
    formatting_detail: str = ""


def _normalize_bool(val):
    if val is None or val is False:
        return False
    return True


def _extract_runs(paragraph) -> list[RunInfo]:
    runs = []
    for run in paragraph.runs:
        color_rgb = None
        if run.font.color and run.font.color.type is not None:
            if run.font.color.rgb is not None:
                rgb = run.font.color.rgb
                color_rgb = (rgb[0], rgb[1], rgb[2])

        size_pt = None
        if run.font.size is not None:
            size_pt = run.font.size.pt

        underline_val = _normalize_bool(run.underline)

        runs.append(
            RunInfo(
                text=run.text,
                bold=_normalize_bool(run.bold),
                italic=_normalize_bool(run.italic),
                underline=underline_val,
                font_name=run.font.name,
                font_size_pt=size_pt,
                font_color_rgb=color_rgb,
            )
        )
    return runs


def _extract_alignment(paragraph) -> str | None:
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    align = paragraph.alignment
    if align is None:
        return None
    mapping = {
        WD_ALIGN_PARAGRAPH.LEFT: "left",
        WD_ALIGN_PARAGRAPH.CENTER: "center",
        WD_ALIGN_PARAGRAPH.RIGHT: "right",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    }
    return mapping.get(align)


def _extract_style_name(paragraph) -> str | None:
    style = paragraph.style
    if style:
        return style.name
    return None


def _extract_paragraph_info(paragraph, location_desc: str) -> ParagraphInfo:
    pf = paragraph.paragraph_format

    space_before = None
    space_after = None
    first_indent = None
    left_indent = None

    if pf.space_before is not None:
        space_before = pf.space_before.pt
    if pf.space_after is not None:
        space_after = pf.space_after.pt
    if pf.first_line_indent is not None:
        first_indent = pf.first_line_indent.pt
    if pf.left_indent is not None:
        left_indent = pf.left_indent.pt

    return ParagraphInfo(
        text=paragraph.text,
        location_desc=location_desc,
        runs=_extract_runs(paragraph),
        alignment=_extract_alignment(paragraph),
        space_before_pt=space_before,
        space_after_pt=space_after,
        first_line_indent_pt=first_indent,
        left_indent_pt=left_indent,
        style_name=_extract_style_name(paragraph),
    )


def extract_paragraph_infos(doc: Document) -> list[ParagraphInfo]:
    from docx_redline.doc_walk import list_redline_paragraphs

    return [
        _extract_paragraph_info(p, loc) for p, loc in list_redline_paragraphs(doc)
    ]


def extract_paragraphs(doc: Document) -> list[ParagraphInfo]:
    """Backward-compatible alias for :func:`extract_paragraph_infos`."""
    return extract_paragraph_infos(doc)


def _tokenize(text: str) -> list[WordToken]:
    tokens = []
    pattern = re.compile(r"(\s+)")
    parts = pattern.split(text)
    for part in parts:
        if not part:
            continue
        tokens.append(WordToken(text=part, is_whitespace=bool(pattern.fullmatch(part))))
    return tokens


def _tokenize_words(text: str) -> list[str]:
    tokens = []
    pattern = re.compile(r"(\S+|\s+)")
    for m in pattern.finditer(text):
        tokens.append(m.group())
    return tokens


def _word_diff(original_text: str, changed_text: str) -> list[DiffSegment]:
    if original_text == changed_text:
        if original_text:
            return [DiffSegment(text=original_text, type="equal")]
        return []

    orig_words = _tokenize_words(original_text)
    new_words = _tokenize_words(changed_text)

    sm = difflib.SequenceMatcher(None, orig_words, new_words, autojunk=False)
    segments = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            segments.append(DiffSegment(text="".join(orig_words[i1:i2]), type="equal"))
        elif tag == "replace":
            del_text = "".join(orig_words[i1:i2])
            ins_text = "".join(new_words[j1:j2])
            if del_text.strip():
                segments.append(DiffSegment(text=del_text, type="delete"))
            if ins_text.strip():
                segments.append(DiffSegment(text=ins_text, type="insert"))
        elif tag == "delete":
            segments.append(DiffSegment(text="".join(orig_words[i1:i2]), type="delete"))
        elif tag == "insert":
            segments.append(DiffSegment(text="".join(new_words[j1:j2]), type="insert"))

    return segments


def _run_format_tuple(run: RunInfo) -> tuple:
    return (
        run.bold,
        run.italic,
        run.underline,
        run.font_name,
        run.font_size_pt,
        run.font_color_rgb,
    )


def _build_char_format_map(runs: list[RunInfo]) -> list[tuple]:
    fmt_map = []
    for run in runs:
        fmt = _run_format_tuple(run)
        for _ in run.text:
            fmt_map.append(fmt)
    return fmt_map


def _compare_formatting(
    orig: ParagraphInfo, changed: ParagraphInfo
) -> tuple[list[str], bool, list[FormattingRange]]:
    changes = []
    has_run_changes = False
    run_ranges: list[FormattingRange] = []

    field_labels = [
        "bold",
        "italic",
        "underline",
        "font name",
        "font size",
        "font color",
    ]

    min_len = min(len(orig.text), len(changed.text))
    orig_map = _build_char_format_map(orig.runs)
    changed_map = _build_char_format_map(changed.runs)

    i = 0
    while i < min_len:
        o_fmt = orig_map[i] if i < len(orig_map) else None
        c_fmt = changed_map[i] if i < len(changed_map) else None
        if o_fmt != c_fmt:
            start = i
            differing_fields = set()
            while i < min_len:
                o_f = orig_map[i] if i < len(orig_map) else None
                c_f = changed_map[i] if i < len(changed_map) else None
                if o_f == c_f:
                    break
                for fi, (ov, cv) in enumerate(
                    zip(
                        (o_f if o_f else (False, False, False, None, None, None)),
                        (c_f if c_f else (False, False, False, None, None, None)),
                    )
                ):
                    if ov != cv:
                        differing_fields.add(field_labels[fi])
                i += 1
            field_list = sorted(differing_fields)
            run_ranges.append(FormattingRange(start=start, end=i, changes=field_list))
            has_run_changes = True
            changes.append(f"Chars {start}-{i}: {', '.join(field_list)} changed")
        else:
            i += 1

    para_fields = [
        ("alignment", orig.alignment, changed.alignment),
        (
            "space before",
            _pt_str(orig.space_before_pt),
            _pt_str(changed.space_before_pt),
        ),
        ("space after", _pt_str(orig.space_after_pt), _pt_str(changed.space_after_pt)),
        (
            "first line indent",
            _pt_str(orig.first_line_indent_pt),
            _pt_str(changed.first_line_indent_pt),
        ),
        ("left indent", _pt_str(orig.left_indent_pt), _pt_str(changed.left_indent_pt)),
        ("style", orig.style_name, changed.style_name),
    ]

    for field_name, o_val, c_val in para_fields:
        if o_val != c_val:
            changes.append(
                f"Paragraph {field_name}: {_format_val(o_val)} \u2192 {_format_val(c_val)}"
            )

    return changes, has_run_changes, run_ranges


def _format_val(val: Any) -> str:
    if val is None:
        return "(inherited)"
    return str(val)


def _pt_str(val: float | None) -> str | None:
    if val is None:
        return None
    return f"{val:.1f}pt"


def align_paragraphs(
    orig_paras: list[ParagraphInfo],
    changed_paras: list[ParagraphInfo],
) -> list[ParagraphDiff]:
    orig_texts = [p.text for p in orig_paras]
    changed_texts = [p.text for p in changed_paras]

    sm = difflib.SequenceMatcher(None, orig_texts, changed_texts, autojunk=False)
    diffs = []

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                oi = i1 + k
                ci = j1 + k
                fmt_changes, has_run_changes, run_ranges = _compare_formatting(
                    orig_paras[oi], changed_paras[ci]
                )
                if fmt_changes:
                    diffs.append(
                        ParagraphDiff(
                            type="formatting",
                            original_info=orig_paras[oi],
                            changed_info=changed_paras[ci],
                            location_desc=orig_paras[oi].location_desc,
                            formatting_changes=fmt_changes,
                            has_run_formatting_changes=has_run_changes,
                            run_formatting_ranges=run_ranges,
                            original_para_index=oi,
                            changed_para_index=ci,
                            segments=[
                                DiffSegment(text=orig_paras[oi].text, type="equal")
                            ],
                        )
                    )
                else:
                    diffs.append(
                        ParagraphDiff(
                            type="equal",
                            original_info=orig_paras[oi],
                            changed_info=changed_paras[ci],
                            location_desc=orig_paras[oi].location_desc,
                            segments=[
                                DiffSegment(text=orig_paras[oi].text, type="equal")
                            ],
                            original_para_index=oi,
                            changed_para_index=ci,
                        )
                    )

        elif tag == "replace":
            orig_chunk = orig_paras[i1:i2]
            changed_chunk = changed_paras[j1:j2]

            orig_texts_chunk = [p.text for p in orig_chunk]
            changed_texts_chunk = [p.text for p in changed_chunk]

            sub_sm = difflib.SequenceMatcher(
                None, orig_texts_chunk, changed_texts_chunk, autojunk=False
            )

            used_changed = set()
            for sub_tag, si1, si2, sj1, sj2 in sub_sm.get_opcodes():
                if sub_tag == "equal":
                    for k in range(si2 - si1):
                        oi = i1 + si1 + k
                        ci = j1 + sj1 + k
                        fmt_changes, has_run_changes, run_ranges = _compare_formatting(
                            orig_paras[oi], changed_paras[ci]
                        )
                        if fmt_changes:
                            diffs.append(
                                ParagraphDiff(
                                    type="formatting",
                                    original_info=orig_paras[oi],
                                    changed_info=changed_paras[ci],
                                    location_desc=orig_paras[oi].location_desc,
                                    formatting_changes=fmt_changes,
                                    has_run_formatting_changes=has_run_changes,
                                    run_formatting_ranges=run_ranges,
                                    original_para_index=oi,
                                    changed_para_index=ci,
                                    segments=[
                                        DiffSegment(
                                            text=orig_paras[oi].text, type="equal"
                                        )
                                    ],
                                )
                            )
                        else:
                            diffs.append(
                                ParagraphDiff(
                                    type="equal",
                                    original_info=orig_paras[oi],
                                    changed_info=changed_paras[ci],
                                    location_desc=orig_paras[oi].location_desc,
                                    segments=[
                                        DiffSegment(
                                            text=orig_paras[oi].text, type="equal"
                                        )
                                    ],
                                    original_para_index=oi,
                                    changed_para_index=ci,
                                )
                            )
                        used_changed.add(ci)
                elif sub_tag == "replace":
                    for k in range(si2 - si1):
                        oi = i1 + si1 + k
                        diffs.append(
                            ParagraphDiff(
                                type="modify",
                                original_info=orig_paras[oi],
                                changed_info=None,
                                location_desc=orig_paras[oi].location_desc,
                                original_para_index=oi,
                            )
                        )
                elif sub_tag == "delete":
                    for k in range(si2 - si1):
                        oi = i1 + si1 + k
                        diffs.append(
                            ParagraphDiff(
                                type="delete",
                                original_info=orig_paras[oi],
                                location_desc=orig_paras[oi].location_desc,
                                original_para_index=oi,
                            )
                        )
                elif sub_tag == "insert":
                    for k in range(sj2 - sj1):
                        ci = j1 + sj1 + k
                        diffs.append(
                            ParagraphDiff(
                                type="insert",
                                changed_info=changed_paras[ci],
                                location_desc=changed_paras[ci].location_desc,
                                changed_para_index=ci,
                            )
                        )
                        used_changed.add(ci)

            for ci in range(j1, j2):
                if ci not in used_changed:
                    diffs.append(
                        ParagraphDiff(
                            type="insert",
                            changed_info=changed_paras[ci],
                            location_desc=changed_paras[ci].location_desc,
                            changed_para_index=ci,
                        )
                    )

        elif tag == "delete":
            for k in range(i2 - i1):
                oi = i1 + k
                diffs.append(
                    ParagraphDiff(
                        type="delete",
                        original_info=orig_paras[oi],
                        location_desc=orig_paras[oi].location_desc,
                        original_para_index=oi,
                    )
                )

        elif tag == "insert":
            for k in range(j2 - j1):
                ci = j1 + k
                diffs.append(
                    ParagraphDiff(
                        type="insert",
                        changed_info=changed_paras[ci],
                        location_desc=changed_paras[ci].location_desc,
                        changed_para_index=ci,
                    )
                )

    return diffs


def _norm_para_text(s: str) -> str:
    return " ".join(s.split())


def _build_norm_text_index(changed_paras: list[ParagraphInfo]) -> dict[str, list[int]]:
    idx_by_norm: dict[str, list[int]] = {}
    for i, p in enumerate(changed_paras):
        idx_by_norm.setdefault(_norm_para_text(p.text), []).append(i)
    return idx_by_norm


def _modify_match_candidates(
    orig_para: ParagraphInfo,
    changed_paras: list[ParagraphInfo],
    consumed_changed: set[int],
    idx_by_norm: dict[str, list[int]],
) -> list[int]:
    """Narrow candidate changed paragraphs before running SequenceMatcher."""
    all_free = [i for i in range(len(changed_paras)) if i not in consumed_changed]
    if not all_free:
        return []

    orig_norm = _norm_para_text(orig_para.text)
    exact = [i for i in idx_by_norm.get(orig_norm, []) if i not in consumed_changed]
    if exact:
        return exact

    olen = max(len(orig_para.text), 1)
    narrowed = [
        i
        for i in all_free
        if 0.25 <= max(len(changed_paras[i].text), 1) / olen <= 4.0
    ]
    return narrowed if narrowed else all_free


def resolve_modifications(
    diffs: list[ParagraphDiff],
    orig_paras: list[ParagraphInfo],
    changed_paras: list[ParagraphInfo],
) -> list[ParagraphDiff]:
    result = []
    consumed_changed: set[int] = set()
    idx_by_norm = _build_norm_text_index(changed_paras)

    for d in diffs:
        if d.type == "modify":
            orig_para = orig_paras[d.original_para_index]
            best_match_idx = None
            best_ratio = 0.3

            search = _modify_match_candidates(
                orig_para, changed_paras, consumed_changed, idx_by_norm
            )
            for ci in search:
                ratio = difflib.SequenceMatcher(
                    None, orig_para.text, changed_paras[ci].text
                ).ratio()
                if ratio > best_ratio:
                    best_ratio = ratio
                    best_match_idx = ci

            if best_match_idx is not None:
                consumed_changed.add(best_match_idx)
                changed_para = changed_paras[best_match_idx]
                segments = _word_diff(orig_para.text, changed_para.text)
                fmt_changes, has_run_changes, run_ranges = _compare_formatting(
                    orig_para, changed_para
                )

                if (
                    fmt_changes
                    and segments
                    and all(s.type == "equal" for s in segments)
                ):
                    d.type = "formatting"
                    d.changed_info = changed_para
                    d.changed_para_index = best_match_idx
                    d.formatting_changes = fmt_changes
                    d.has_run_formatting_changes = has_run_changes
                    d.run_formatting_ranges = run_ranges
                    d.segments = segments
                else:
                    d.changed_info = changed_para
                    d.changed_para_index = best_match_idx
                    d.segments = segments
                    d.formatting_changes = fmt_changes
                    d.has_run_formatting_changes = has_run_changes
                    d.run_formatting_ranges = run_ranges
            else:
                d.type = "delete"

        result.append(d)

    result = [
        d
        for d in result
        if not (
            d.type == "insert"
            and d.changed_para_index is not None
            and d.changed_para_index in consumed_changed
        )
    ]

    return result


def _build_change_list(diffs: list[ParagraphDiff]) -> list[Change]:
    changes = []
    para_counter = 0

    for d in diffs:
        para_counter += 1
        loc = d.location_desc or f"Block {para_counter}"

        if d.type == "insert":
            changes.append(
                Change(
                    type=ChangeType.INSERTION,
                    paragraph_index=para_counter,
                    location_desc=loc,
                    original_text="",
                    new_text=d.changed_info.text if d.changed_info else "",
                )
            )

        elif d.type == "delete":
            changes.append(
                Change(
                    type=ChangeType.DELETION,
                    paragraph_index=para_counter,
                    location_desc=loc,
                    original_text=d.original_info.text if d.original_info else "",
                    new_text="",
                )
            )

        elif d.type == "modify":
            del_parts = [s.text for s in d.segments if s.type == "delete"]
            ins_parts = [s.text for s in d.segments if s.type == "insert"]
            detail = ""
            if del_parts and ins_parts:
                detail = f'Replaced "{_truncate("".join(del_parts))}" with "{_truncate("".join(ins_parts))}"'
            elif del_parts:
                detail = f'Deleted "{_truncate("".join(del_parts))}"'
            elif ins_parts:
                detail = f'Inserted "{_truncate("".join(ins_parts))}"'

            changes.append(
                Change(
                    type=ChangeType.MODIFICATION,
                    paragraph_index=para_counter,
                    location_desc=loc,
                    original_text=d.original_info.text if d.original_info else "",
                    new_text=d.changed_info.text if d.changed_info else "",
                    formatting_detail=detail,
                )
            )

        elif d.type == "formatting":
            changes.append(
                Change(
                    type=ChangeType.FORMATTING,
                    paragraph_index=para_counter,
                    location_desc=loc,
                    original_text=d.original_info.text if d.original_info else "",
                    new_text=d.changed_info.text if d.changed_info else "",
                    formatting_detail="; ".join(d.formatting_changes),
                )
            )

    return changes


def _truncate(text: str, max_len: int = 60) -> str:
    text = " ".join(text.split())
    if len(text) > max_len:
        return text[: max_len - 3] + "..."
    return text


def compare_documents(
    original_path: str, changed_path: str
) -> tuple[list[ParagraphDiff], list[Change]]:
    orig_doc = Document(original_path)
    changed_doc = Document(changed_path)

    orig_paras = extract_paragraph_infos(orig_doc)
    changed_paras = extract_paragraph_infos(changed_doc)

    diffs = align_paragraphs(orig_paras, changed_paras)
    diffs = resolve_modifications(diffs, orig_paras, changed_paras)
    changes = _build_change_list(diffs)

    return diffs, changes
