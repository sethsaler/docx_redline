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
    """Default output filename next to ``cwd`` (or current directory)."""
    base = cwd if cwd is not None else os.getcwd()
    orig_base = os.path.splitext(os.path.basename(original_path))[0]
    changed_base = os.path.splitext(os.path.basename(changed_path))[0]
    suffix = "_tracked" if track_changes else ""
    return os.path.join(base, f"redline_{orig_base}_vs_{changed_base}{suffix}.docx")
