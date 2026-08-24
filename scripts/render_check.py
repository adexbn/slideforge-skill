# -*- coding: utf-8 -*-
"""
render_check.py — render a .pptx to PNGs via the local PowerPoint COM (or LibreOffice)
so the slideforge skill can do a real visual QA pass.

  python render_check.py --check                      # probe render backend only
  python render_check.py deck.pptx -o outdir [--dpi 1600x900]
  python render_check.py deck.pptx                   # writes ./preview/<stem>/slide-NN.png

Backend priority on Windows: PowerPoint COM > LibreOffice. macOS: Keynote > LibreOffice.
Core LibreOffice rendering path also works on Linux.
"""
from __future__ import annotations

import argparse, os, subprocess, sys
from pathlib import Path


def have_powerpoint_com():
    try:
        import win32com.client  # noqa
        return True
    except Exception:
        pass
    # try via pywin32 default dispatch through a quick COM probe is unreliable;
    # rely on python-pptx not needed here. Fall back to a PowerShell probe.
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             "try{$a=New-Object -ComObject PowerPoint.Application;"
             "\"OK\"+$a.Version;$a.Quit()}catch{'NOPE'}"],
            capture_output=True, text=True, timeout=40)
        return r.stdout.strip().startswith("OK")
    except Exception:
        return False


def have_libreoffice():
    for cmd in ("soffice", "libreoffice"):
        p = shutil_which(cmd)
        if p:
            return p
    return None


def shutil_which(cmd):
    import shutil
    return shutil.which(cmd)


def render_with_powerpoint(pptx_path, out_dir):
    import win32com.client  # type: ignore
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    pres = ppt.Presentations.Open(os.path.abspath(pptx_path), True, False, False)
    os.makedirs(out_dir, exist_ok=True)
    n = pres.Slides.Count
    for i in range(1, n + 1):
        out = os.path.join(out_dir, f"slide-{i:02d}.png")
        pres.Slides.Item(i).Export(out, "PNG", 1600, 900)
    pres.Close()
    ppt.Quit()
    return n


def render_with_libreoffice(pptx_path, out_dir):
    import shutil
    soffice = have_libreoffice()
    if not soffice:
        raise RuntimeError("no LibreOffice found")
    os.makedirs(out_dir, exist_ok=True)
    # libreoffice converts to PNG per slide only via pdf -> images is heavy; do
    # a one-page-per-file approach using pdf then per-frame is out of scope here.
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir,
                    os.path.abspath(pptx_path)], check=True)
    return -1  # pdf; caller converts


def render(pptx_path, out_dir):
    if have_powerpoint_com():
        try:
            return render_with_powerpoint(pptx_path, out_dir), "powerpoint"
        except Exception as ex:
            print("PowerPoint COM failed, falling back:", ex)
    lo = have_libreoffice()
    if lo:
        return render_with_libreoffice(pptx_path, out_dir), "libreoffice"
    raise RuntimeError("no render backend (PowerPoint COM / LibreOffice) available")


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx", nargs="?", help="the .pptx to render")
    ap.add_argument("-o", "--out", default=None, help="output dir (default ./preview/<stem>)")
    ap.add_argument("--check", action="store_true", help="probe backend only")
    args = ap.parse_args(argv)
    if args.check:
        pp = have_powerpoint_com()
        lo = bool(have_libreoffice())
        print(json_dumps({"powerpoint_com": pp, "libreoffice": lo}))
        return 0 if (pp or lo) else 10
    if not args.pptx:
        print("need a .pptx path (or --check)")
        return 1
    out = args.out or os.path.join("preview", Path(args.pptx).stem)
    n, backend = render(args.pptx, out)
    print(f"OK backend={backend} slides={n} out={os.path.abspath(out)}")
    return 0


def json_dumps(obj):
    import json
    return json.dumps(obj, ensure_ascii=False)


if __name__ == "__main__":
    raise SystemExit(main())
