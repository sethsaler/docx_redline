"""Walk document body (including nested tables) and unique header/footer stories."""

from __future__ import annotations

from collections.abc import Iterator

from docx.document import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


def iter_paragraphs_in_container(container) -> Iterator[Paragraph]:
    """Paragraphs in document order, recursing into table cells."""
    for block in container.iter_inner_content():
        if isinstance(block, Paragraph):
            yield block
        elif isinstance(block, Table):
            for row in block.rows:
                for cell in row.cells:
                    yield from iter_paragraphs_in_container(cell)


def iter_redline_paragraphs(
    doc: Document,
) -> Iterator[tuple[Paragraph, str]]:
    """
    (paragraph, location_label) for every paragraph compared in the redline.

    Main body comes first (including table cells), then each distinct header/footer
    part once (skipping sections linked to previous and duplicate part definitions).
    """
    for i, p in enumerate(iter_paragraphs_in_container(doc), start=1):
        yield p, f"Body ¶{i}"

    seen_part_ids: set[int] = set()
    for si, section in enumerate(doc.sections):
        for label, hf in (
            ("Header", section.header),
            ("First-page header", section.first_page_header),
            ("Even-page header", section.even_page_header),
            ("Footer", section.footer),
            ("First-page footer", section.first_page_footer),
            ("Even-page footer", section.even_page_footer),
        ):
            if si > 0 and hf.is_linked_to_previous:
                continue
            part = hf.part
            pid = id(part)
            if pid in seen_part_ids:
                continue
            seen_part_ids.add(pid)
            for j, p in enumerate(iter_paragraphs_in_container(hf), start=1):
                yield p, f"Section {si + 1} {label} ¶{j}"


def list_redline_paragraphs(doc: Document) -> list[tuple[Paragraph, str]]:
    return list(iter_redline_paragraphs(doc))
