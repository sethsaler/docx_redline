from __future__ import annotations

import os
import sys
from typing import Callable

from docx_redline.formatter import generate_redline
from docx_redline.paths import default_output_path, normalize_user_path, validate_docx_input_path


def _run_gui_window() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    try:
        from tkinterdnd2 import DND_FILES, TkinterDnD
    except ImportError:
        DND_FILES = None  # type: ignore[assignment,misc]
        TkinterDnD = None  # type: ignore[assignment,misc]

    last_open_dir: list[str] = [""]

    def _remember_dir_for_path(path: str) -> None:
        d = os.path.dirname(os.path.abspath(path))
        if os.path.isdir(d):
            last_open_dir[0] = d

    def _open_dialog_initialdir(current_field: str) -> str | None:
        for candidate in (
            current_field and os.path.dirname(os.path.abspath(current_field.strip())),
            last_open_dir[0],
            os.path.expanduser("~"),
        ):
            if candidate and os.path.isdir(candidate):
                return candidate
        return None

    def _browse_docx(var: tk.StringVar) -> None:
        path = filedialog.askopenfilename(
            title="Choose a Word document",
            filetypes=(
                ("Word documents", "*.docx"),
                ("All files", "*.*"),
            ),
            initialdir=_open_dialog_initialdir(var.get()),
        )
        if path:
            var.set(path)
            _remember_dir_for_path(path)

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
        initial_dir = None
        initial_file = None
        if initial:
            initial_dir = os.path.dirname(initial) or None
            initial_file = os.path.basename(initial)
        if not initial_dir or not os.path.isdir(initial_dir):
            initial_dir = _open_dialog_initialdir(var.get())
        path = filedialog.asksaveasfilename(
            title="Save redlined document as",
            defaultextension=".docx",
            filetypes=(("Word documents", "*.docx"), ("All files", "*.*")),
            initialfile=initial_file,
            initialdir=initial_dir,
        )
        if path:
            var.set(path)
            _remember_dir_for_path(path)

    root: tk.Tk
    if TkinterDnD is not None:
        try:
            root = TkinterDnD.Tk()
        except Exception:
            root = tk.Tk()
    else:
        root = tk.Tk()

    root.title("DOCX Redline Comparison")
    root.minsize(600, 300)

    pad = {"padx": 10, "pady": 6}
    small_pad = {"padx": (0, 4), "pady": 6}
    original_var = tk.StringVar()
    changed_var = tk.StringVar()
    output_var = tk.StringVar()
    output_mode_var = tk.StringVar(value="styled")

    frm = ttk.Frame(root, padding=12)
    frm.grid(row=0, column=0, sticky="nsew")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    frm.columnconfigure(1, weight=1)

    def _clear_var(var: tk.StringVar) -> None:
        var.set("")

    def _file_row(
        row: int,
        label: str,
        var: tk.StringVar,
        browse_cmd: Callable[[], None],
        *,
        allow_drop: bool = True,
    ) -> None:
        ttk.Label(frm, text=label).grid(row=row, column=0, sticky="w", **pad)
        ent = ttk.Entry(frm, textvariable=var, width=48)
        ent.grid(row=row, column=1, sticky="ew", **pad)
        actions = ttk.Frame(frm)
        actions.grid(row=row, column=2, sticky="e", **pad)
        ttk.Button(actions, text="Clear", width=7, command=lambda: _clear_var(var)).pack(
            side=tk.LEFT, **small_pad
        )
        ttk.Button(actions, text="Browse…", command=browse_cmd).pack(side=tk.LEFT)

        if allow_drop and DND_FILES is not None and hasattr(ent, "drop_target_register"):
            def on_drop(event: tk.Event, v: tk.StringVar = var) -> None:  # type: ignore[name-defined]
                try:
                    paths = root.tk.splitlist(event.data)  # type: ignore[attr-defined]
                except tk.TclError:
                    return
                if not paths:
                    return
                p = os.path.normpath(paths[0])
                if not p.lower().endswith(".docx"):
                    messagebox.showwarning(
                        "Not a Word file",
                        "Please drop a .docx file.",
                    )
                    return
                v.set(p)
                _remember_dir_for_path(p)

            ent.drop_target_register(DND_FILES)
            ent.dnd_bind("<<Drop>>", on_drop)

    _file_row(0, "Original (.docx)", original_var, lambda: _browse_docx(original_var))
    _file_row(1, "Modified (.docx)", changed_var, lambda: _browse_docx(changed_var))

    def swap_original_modified() -> None:
        o, c = original_var.get(), changed_var.get()
        original_var.set(c)
        changed_var.set(o)

    swap_frm = ttk.Frame(frm)
    swap_frm.grid(row=2, column=1, columnspan=2, sticky="e", **pad)
    ttk.Button(
        swap_frm,
        text="Swap original ↔ modified",
        command=swap_original_modified,
    ).pack(side=tk.RIGHT)

    def browse_output() -> None:
        _browse_save(
            output_var,
            original_var.get().strip() or None,
            changed_var.get().strip() or None,
            output_mode_var.get() == "track_changes",
        )

    _file_row(3, "Output (.docx)", output_var, browse_output, allow_drop=False)

    mode_frm = ttk.LabelFrame(frm, text="Output mode", padding=(8, 6))
    mode_frm.grid(row=4, column=0, columnspan=3, sticky="ew", **pad)
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
    status_lbl = ttk.Label(frm, textvariable=status_var, wraplength=480)
    status_lbl.grid(row=5, column=0, columnspan=3, sticky="w", **pad)

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

        parent = os.path.dirname(output)
        if parent and not os.path.isdir(parent):
            messagebox.showerror(
                "Invalid output path",
                f"Output folder does not exist:\n{parent}",
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

        _remember_dir_for_path(output)
        status_var.set(f"Saved: {output}")
        messagebox.showinfo("Done", f"Redline saved to:\n{output}")

    btn_row = ttk.Frame(frm)
    btn_row.grid(row=6, column=0, columnspan=3, sticky="e", **pad)
    ttk.Button(btn_row, text="Generate redline", command=on_generate).pack(side=tk.RIGHT)

    def sync_wraplength(_event: tk.Event | None = None) -> None:
        try:
            w = frm.winfo_width()
        except tk.TclError:
            return
        # Leave margin for outer padding and label column.
        status_lbl.configure(wraplength=max(280, w - 48))

    frm.bind("<Configure>", sync_wraplength)
    root.after_idle(sync_wraplength)

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
