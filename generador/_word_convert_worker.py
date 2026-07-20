# -*- coding: utf-8 -*-
"""Worker aislado: convierte .doc -> .docx con Word COM (proceso aparte)."""
from __future__ import annotations

import sys
from pathlib import Path


def main(src: str, dst: str) -> int:
    import pythoncom
    import win32com.client  # type: ignore

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(
            str(Path(src).resolve()),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        dst_abs = str(Path(dst).resolve())
        try:
            doc.SaveAs2(dst_abs, FileFormat=16)
        except Exception:
            doc.SaveAs(dst_abs, FileFormat=16)
        doc.Close(False)
        return 0
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("usage: _word_convert_worker.py src.doc dst.docx", file=sys.stderr)
        sys.exit(2)
    try:
        raise SystemExit(main(sys.argv[1], sys.argv[2]))
    except Exception as e:
        print(f"{type(e).__name__}: {e}", file=sys.stderr)
        raise SystemExit(1)
