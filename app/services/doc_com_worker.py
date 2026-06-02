"""Isolated subprocess worker for .doc COM extraction."""

from __future__ import annotations

import argparse
import json
import subprocess
from pathlib import Path


def _kill_stale_office_processes() -> None:
    # Stale Office/WPS background processes can cause COM startup hangs.
    process_names = ["WINWORD.EXE", "wps.exe"]
    for name in process_names:
        try:
            subprocess.run(
                ["taskkill", "/F", "/IM", name],
                capture_output=True,
                text=True,
                check=False,
                timeout=10,
            )
        except Exception:
            pass


def _extract_doc_text_via_com(doc_path: Path, prog_ids: list[str]) -> str:
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    pythoncom.CoInitialize()
    app = None
    doc = None
    last_error: Exception | None = None
    try:
        _kill_stale_office_processes()
        for prog_id in prog_ids:
            try:
                # DispatchEx creates a new instance instead of attaching to
                # potentially blocked existing Office/WPS instance.
                app = win32com.client.DispatchEx(prog_id)
                break
            except Exception as exc:  # noqa: PERF203
                last_error = exc
                app = None
        if app is None:
            raise RuntimeError(f"No COM server available. Last error: {last_error}")

        app.Visible = False
        # Prevent modal prompts that can block automation silently.
        try:
            app.DisplayAlerts = 0
        except Exception:
            pass
        doc = app.Documents.Open(
            str(doc_path.resolve()),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            NoEncodingDialog=True,
        )
        return doc.Content.Text or ""
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _probe_com_prog_id(prog_id: str) -> None:
    import winreg

    # Registration-only probe via registry; avoids COM server launch and
    # works even on environments where pythoncom lacks CLSID helpers.
    paths = [
        prog_id,
        f"{prog_id}\\CLSID",
        rf"Wow6432Node\Classes\{prog_id}",
        rf"Wow6432Node\Classes\{prog_id}\CLSID",
    ]
    last_error: Exception | None = None
    for sub_key in paths:
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, sub_key):
                return
        except Exception as exc:
            last_error = exc
    raise RuntimeError(
        f"ProgID not registered in HKCR: {prog_id}. Last error: {last_error}"
    )


def run_doc_com_worker_cli(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--doc-com-worker", action="store_true")
    parser.add_argument("--doc-path", required=True)
    parser.add_argument("--prog-ids", required=True)
    args = parser.parse_args(argv)

    doc_path = Path(args.doc_path)
    prog_ids = [part.strip() for part in args.prog_ids.split(",") if part.strip()]
    if not prog_ids:
        payload = {"ok": False, "error": "No COM ProgIDs provided."}
        print(json.dumps(payload))
        return 2

    try:
        text = _extract_doc_text_via_com(doc_path, prog_ids)
        print(json.dumps({"ok": True, "text": text}))
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}))
        return 1


def run_probe_com_worker_cli(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--probe-com-worker", action="store_true")
    parser.add_argument("--prog-id", required=True)
    args = parser.parse_args(argv)

    prog_id = args.prog_id.strip()
    if not prog_id:
        print(json.dumps({"ok": False, "error": "Empty COM ProgID"}))
        return 2
    try:
        _probe_com_prog_id(prog_id)
        print(json.dumps({"ok": True, "detail": "available"}))
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}))
        return 1

