from __future__ import annotations

import os


def normalize_user_path(path: str) -> str:
    """Strip whitespace, expand ``~``, and resolve to an absolute path."""
    path = path.strip()
    if not path:
        raise ValueError("Path is empty.")
    if path.startswith("~"):
        path = os.path.expanduser(path)
    return os.path.abspath(path)


def redlines_desktop_dir() -> str:
    """Return ``~/Desktop/Redlines``, creating the directory if it does not exist."""
    d = os.path.join(os.path.expanduser("~/Desktop"), "Redlines")
    os.makedirs(d, exist_ok=True)
    return d


def ensure_parent_dir(path: str) -> None:
    """Create the parent directory of ``path`` if needed (no-op for bare filenames)."""
    parent = os.path.dirname(os.path.abspath(path))
    if parent:
        os.makedirs(parent, exist_ok=True)


def validate_docx_input_path(path: str) -> None:
    """
    Ensure ``path`` refers to an existing .docx file.
    ``path`` should already be normalized (see ``normalize_user_path``).
    """
    if not os.path.isfile(path):
        raise ValueError(f"File not found: {path}")
    if not path.lower().endswith(".docx"):
        raise ValueError("File must be a .docx file.")


def default_output_path(
    original_path: str,
    changed_path: str,
    cwd: str | None = None,
    *,
    track_changes: bool = False,
) -> str:
    """
    Default output path: ``~/Desktop/Redlines/<original_stem>_redlines.docx``.

    ``changed_path`` and ``cwd`` are ignored but kept for backward-compatible
    call sites; the filename uses only the original document's base name.
    """
    _ = (changed_path, cwd)
    base = redlines_desktop_dir()
    orig_base = os.path.splitext(os.path.basename(original_path))[0]
    suffix = "_tracked" if track_changes else ""
    return os.path.join(base, f"{orig_base}_redlines{suffix}.docx")
