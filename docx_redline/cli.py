from __future__ import annotations

import argparse
import os
import sys
from typing import List, Optional

from docx_redline.formatter import generate_redline


def main(argv: Optional[List[str]] = None):
    parser = argparse.ArgumentParser(
        prog="docx-redline",
        description="Generate a redlined comparison of two .docx files.",
    )
    parser.add_argument("original", help="Path to the original .docx file")
    parser.add_argument("changed", help="Path to the modified .docx file")
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Path for the output redlined .docx file (default: redline_<original>_vs_<changed>.docx)",
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
        output_path = args.output
    else:
        orig_base = os.path.splitext(os.path.basename(args.original))[0]
        changed_base = os.path.splitext(os.path.basename(args.changed))[0]
        output_path = f"redline_{orig_base}_vs_{changed_base}.docx"

    print(f"Comparing:")
    print(f"  Original: {args.original}")
    print(f"  Modified: {args.changed}")
    print()

    try:
        generate_redline(args.original, args.changed, output_path)
    except Exception as e:
        print(f"Error generating redline: {e}", file=sys.stderr)
        sys.exit(1)

    print(f"Redline saved to: {output_path}")


if __name__ == "__main__":
    main()
