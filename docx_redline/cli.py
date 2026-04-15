from __future__ import annotations

import argparse
import os
import sys
from typing import List, Optional

from docx.exceptions import InvalidXmlError, PythonDocxError

from docx_redline.formatter import generate_redline
from docx_redline.paths import default_output_path, normalize_user_path


def _package_version() -> str:
    try:
        from importlib.metadata import PackageNotFoundError, version

        return version("docx-redline")
    except (ImportError, PackageNotFoundError):
        return "unknown"


def main(argv: Optional[List[str]] = None):
    parser = argparse.ArgumentParser(
        prog="docx-redline",
        description="Generate a redlined comparison of two .docx files.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {_package_version()}",
    )
    parser.add_argument("original", help="Path to the original .docx file")
    parser.add_argument("changed", help="Path to the modified .docx file")
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help=(
            "Path for the output redlined .docx file "
            "(default: ~/Desktop/Redlines/<original_stem>_redlines.docx)"
        ),
    )
    parser.add_argument(
        "-q",
        "--quiet",
        action="store_true",
        help="Print only errors and the final output path.",
    )
    parser.add_argument(
        "-f",
        "--force",
        action="store_true",
        help="Overwrite the output file if it already exists.",
    )
    parser.add_argument(
        "--mode",
        choices=("styled", "track_changes"),
        default="styled",
        help=(
            "styled: merge diff with red underline/strike and a change report (default). "
            "track_changes: apply edits as Word revision markup (w:ins/w:del) on the original."
        ),
    )

    args = parser.parse_args(argv)

    if not os.path.isfile(args.original):
        print(f"Error: original file not found: {args.original}", file=sys.stderr)
        sys.exit(1)
    if not os.path.isfile(args.changed):
        print(f"Error: changed file not found: {args.changed}", file=sys.stderr)
        sys.exit(1)

    if not args.original.lower().endswith(".docx"):
        print("Error: original file must be a .docx file", file=sys.stderr)
        sys.exit(1)
    if not args.changed.lower().endswith(".docx"):
        print("Error: changed file must be a .docx file", file=sys.stderr)
        sys.exit(1)

    if args.output:
        output_path = normalize_user_path(args.output)
    else:
        output_path = default_output_path(
            args.original,
            args.changed,
            track_changes=args.mode == "track_changes",
        )

    if os.path.exists(output_path) and not args.force:
        print(
            f"Error: output file already exists: {output_path} (use --force to overwrite)",
            file=sys.stderr,
        )
        sys.exit(1)

    if not args.quiet:
        print("Comparing:")
        print(f"  Original: {args.original}")
        print(f"  Modified: {args.changed}")
        print()

    try:
        generate_redline(
            args.original, args.changed, output_path, output_mode=args.mode
        )
    except (InvalidXmlError, PythonDocxError) as e:
        print(
            f"Error: could not read Word document (invalid or corrupt .docx): {e}",
            file=sys.stderr,
        )
        sys.exit(1)
    except PermissionError as e:
        print(f"Error: permission denied: {e}", file=sys.stderr)
        sys.exit(1)
    except OSError as e:
        print(f"Error: could not write output: {e}", file=sys.stderr)
        sys.exit(1)
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    print(f"Redline saved to: {output_path}")


if __name__ == "__main__":
    main()
