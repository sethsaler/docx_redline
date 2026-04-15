from __future__ import annotations

import os
import sys

from docx_redline.formatter import generate_redline
from docx_redline.paths import (
    default_output_path,
    ensure_parent_dir,
    normalize_user_path,
    validate_docx_input_path,
)


def _run_gui_window() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    def _browse_docx(var: tk.StringVar) -> None:
        path = filedialog.askopenfilename(
            title="Choose a Word document",
            filetypes=(
                ("Word documents", "*.docx"),
                ("All files", "*.*"),
            ),
        )
        if path:
            var.set(path)

    def _browse_save(
        var: tk.StringVar,
        original: str | None,
        changed: str | None,
        track_changes: bool,
    ) -> None:
        initial = var.get().strip() or None
        if not initial and original and changed:
            try:
                initial = default_output_path(
                    original, changed, track_changes=track_changes
                )
            except Exception:
                initial = None
        path = filedialog.asksaveasfilename(
            title="Save redlined document as",
            defaultextension=".docx",
            filetypes=(("Word documents", "*.docx"), ("All files", "*.*")),
            initialfile=os.path.basename(initial) if initial else None,
            initialdir=os.path.dirname(initial) if initial and os.path.dirname(initial) else None,
        )
        if path:
            var.set(path)

    root = tk.Tk()
    root.title("DOCX Redline Comparison")
    root.minsize(520, 280)

    pad = {"padx": 10, "pady": 6}
    original_var = tk.StringVar()
    changed_var = tk.StringVar()
    output_var = tk.StringVar()
    output_mode_var = tk.StringVar(value="styled")

    frm = ttk.Frame(root, padding=12)
    frm.grid(row=0, column=0, sticky="nsew")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    frm.columnconfigure(1, weight=1)

    ttk.Label(frm, text="Original (.docx)").grid(row=0, column=0, sticky="w", **pad)
    ttk.Entry(frm, textvariable=original_var, width=52).grid(row=0, column=1, sticky="ew", **pad)
    ttk.Button(frm, text="Browse…", command=lambda: _browse_docx(original_var)).grid(
        row=0, column=2, **pad
    )

    ttk.Label(frm, text="Modified (.docx)").grid(row=1, column=0, sticky="w", **pad)
    ttk.Entry(frm, textvariable=changed_var, width=52).grid(row=1, column=1, sticky="ew", **pad)
    ttk.Button(frm, text="Browse…", command=lambda: _browse_docx(changed_var)).grid(
        row=1, column=2, **pad
    )

    ttk.Label(frm, text="Output (.docx)").grid(row=2, column=0, sticky="w", **pad)
    ttk.Entry(frm, textvariable=output_var, width=52).grid(row=2, column=1, sticky="ew", **pad)
    ttk.Button(
        frm,
        text="Browse…",
        command=lambda: _browse_save(
            output_var,
            original_var.get().strip() or None,
            changed_var.get().strip() or None,
            output_mode_var.get() == "track_changes",
        ),
    ).grid(row=2, column=2, **pad)

    mode_frm = ttk.LabelFrame(frm, text="Output mode", padding=(8, 6))
    mode_frm.grid(row=3, column=0, columnspan=3, sticky="ew", **pad)
    ttk.Radiobutton(
        mode_frm,
        text="Styled redline (underline / strike + change report)",
        variable=output_mode_var,
        value="styled",
    ).grid(row=0, column=0, sticky="w")
    ttk.Radiobutton(
        mode_frm,
        text="Track changes (Word revision markup on the original)",
        variable=output_mode_var,
        value="track_changes",
    ).grid(row=1, column=0, sticky="w")

    status_var = tk.StringVar(value="Choose two documents, then click Generate redline.")
    ttk.Label(frm, textvariable=status_var, wraplength=480).grid(
        row=4, column=0, columnspan=3, sticky="w", **pad
    )

    def on_generate() -> None:
        o_raw = original_var.get().strip()
        c_raw = changed_var.get().strip()
        out_raw = output_var.get().strip()

        if not o_raw or not c_raw:
            messagebox.showwarning(
                "Missing files",
                "Please choose both the original and modified .docx files.",
            )
            return

        try:
            original = normalize_user_path(o_raw)
            changed = normalize_user_path(c_raw)
            validate_docx_input_path(original)
            validate_docx_input_path(changed)
        except ValueError as e:
            messagebox.showerror("Invalid input", str(e))
            return

        track_changes = output_mode_var.get() == "track_changes"
        if not out_raw:
            out_raw = default_output_path(
                original, changed, track_changes=track_changes
            )
            output_var.set(out_raw)

        try:
            output = normalize_user_path(out_raw)
        except ValueError as e:
            messagebox.showerror("Invalid output path", str(e))
            return

        try:
            ensure_parent_dir(output)
        except OSError as e:
            messagebox.showerror(
                "Invalid output path",
                f"Could not create output folder:\n{e}",
            )
            return

        status_var.set("Generating redline…")
        root.update_idletasks()

        try:
            generate_redline(
                original,
                changed,
                output,
                output_mode="track_changes" if track_changes else "styled",
            )
        except Exception as e:
            status_var.set("Ready.")
            messagebox.showerror("Error", f"Could not generate redline:\n{e}")
            return

        status_var.set(f"Saved: {output}")
        messagebox.showinfo("Done", f"Redline saved to:\n{output}")

    btn_row = ttk.Frame(frm)
    btn_row.grid(row=5, column=0, columnspan=3, sticky="e", **pad)
    ttk.Button(btn_row, text="Generate redline", command=on_generate).pack(side=tk.RIGHT)

    root.bind("<Return>", lambda e: on_generate())
    root.mainloop()


def main() -> None:
    try:
        import tkinter as tk
    except ImportError:
        print(
            "The graphical file picker requires Tk (tkinter). Falling back to terminal prompts.\n"
            "For the Browse window: use Python from python.org (macOS) with Tcl/Tk, "
            "or `brew install python-tk`, or on Linux install `python3-tk`, then recreate `.venv` with setup.sh.",
            file=sys.stderr,
        )
        from docx_redline.cli_interactive import main as interactive_main

        interactive_main()
        return

    try:
        _run_gui_window()
    except tk.TclError as e:
        print(
            "Could not open the graphical window. Is Tk installed correctly?\n"
            "You can use terminal prompts instead:\n"
            "  .venv/bin/python -m docx_redline.cli_interactive\n"
            f"Details: {e}",
            file=sys.stderr,
        )
        sys.exit(1)


if __name__ == "__main__":
    main()
