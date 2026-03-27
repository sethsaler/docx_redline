from __future__ import annotations

import sys

from docx_redline.formatter import generate_redline
from docx_redline.paths import default_output_path, normalize_user_path, validate_docx_input_path


def _prompt_file(label: str) -> str:
    while True:
        path = input(f"{label}: ").strip()
        if not path:
            print("  Please enter a path.")
            continue
        try:
            normalized = normalize_user_path(path)
            validate_docx_input_path(normalized)
        except ValueError as e:
            print(f"  {e}")
            continue
        return normalized


def main():
    print("=" * 50)
    print("  DOCX Redline Comparison Tool")
    print("=" * 50)
    print()

    original = _prompt_file("Original file path (the before version)")
    changed = _prompt_file("Changed file path  (the after version)")

    default_output = default_output_path(original, changed)

    print()
    output = input(f"Output file path [{default_output}]: ").strip()
    if not output:
        output = default_output

    try:
        output = normalize_user_path(output)
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    print()
    print(f"Comparing:")
    print(f"  Original: {original}")
    print(f"  Modified: {changed}")
    print()

    try:
        generate_redline(original, changed, output)
    except Exception as e:
        print(f"Error generating redline: {e}", file=sys.stderr)
        print()
        print("Press Enter to exit.")
        input()
        sys.exit(1)

    print(f"Redline saved to: {output}")


if __name__ == "__main__":
    main()
