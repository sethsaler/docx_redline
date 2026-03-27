from __future__ import annotations

import os
import sys

from docx_redline.formatter import generate_redline


def _prompt_file(label: str) -> str:
    while True:
        path = input(f"{label}: ").strip()
        if not path:
            print("  Please enter a path.")
            continue
        home = os.path.expanduser("~")
        if path.startswith("~"):
            path = home + path[1:]
        if not os.path.isfile(path):
            print(f"  File not found: {path}")
            continue
        if not path.lower().endswith(".docx"):
            print("  File must be a .docx file.")
            continue
        return os.path.abspath(path)


def main():
    print("=" * 50)
    print("  DOCX Redline Comparison Tool")
    print("=" * 50)
    print()

    original = _prompt_file("Original file path (the before version)")
    changed = _prompt_file("Changed file path  (the after version)")

    orig_base = os.path.splitext(os.path.basename(original))[0]
    changed_base = os.path.splitext(os.path.basename(changed))[0]
    default_output = os.path.join(
        os.getcwd(), f"redline_{orig_base}_vs_{changed_base}.docx"
    )

    print()
    output = input(f"Output file path [{default_output}]: ").strip()
    if not output:
        output = default_output

    home = os.path.expanduser("~")
    if output.startswith("~"):
        output = home + output[1:]
    output = os.path.abspath(output)

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
