#!/usr/bin/env python3
"""Run the VBA Python-subset interpreter inside headless LibreOffice (VBA
compatibility mode) and compare its output with CPython.

usage: /usr/bin/python3 run_tests.py [-v] [-k substring] [--python /path/to/python3]

Needs LibreOffice and its Python-UNO bridge (python3-uno). cases/x01_bigint.py
needs VBA's Decimal type, which LibreOffice lacks, so it only runs in Excel.
A case with a matching .out file is compared with that file instead of CPython.
For tests that end
with an uncaught exception, only the final "ExcType: message" line of the
traceback is compared.
"""
import argparse, glob, os, re, shutil, socket, subprocess, sys, tempfile, time

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.environ.get("PYVBA_SRC", os.path.join(HERE, "..", "src"))

import uno  # noqa: E402  (provided by LibreOffice)
from com.sun.star.script import ModuleInfo  # noqa: E402

HARNESS = """Option VBASupport 1
Option Explicit
Function RunPy(ByVal src As String) As String
    On Error GoTo EH
    RunPy = PyRunProgram(src)
    Exit Function
EH:
    RunPy = "!!VBA ERROR " & Err.Number & ": " & Err.Description
End Function
"""

def vba_source(path):
    text = open(path, encoding="ascii").read().replace("\r\n", "\n")
    out, in_header = [], path.endswith(".cls")
    for line in text.split("\n"):
        if in_header:
            if line.startswith("END"):
                in_header = False
            continue
        if line.startswith("Attribute VB_"):
            continue
        out.append(line)
    opts = ["Option VBASupport 1"]
    if path.endswith(".cls"):
        opts.append("Option ClassModule")
    return "\n".join(opts + out)

def free_port():
    s = socket.socket(); s.bind(("127.0.0.1", 0)); p = s.getsockname()[1]; s.close(); return p

def start_office():
    port = free_port()
    profile = tempfile.mkdtemp(prefix="lo-profile-")
    proc = subprocess.Popen(
        ["soffice", "--headless", "--invisible", "--norestore", "--nologo",
         "-env:UserInstallation=file://" + profile,
         "--accept=socket,host=127.0.0.1,port=%d;urp;" % port],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(120):
        try:
            ctx = resolver.resolve("uno:socket,host=127.0.0.1,port=%d;urp;StarOffice.ComponentContext" % port)
            return proc, ctx, profile
        except Exception:
            time.sleep(0.5)
    raise RuntimeError("could not connect to LibreOffice")

def load_modules(doc):
    libs = doc.BasicLibraries
    libs.VBACompatibilityMode = True
    lib = libs.getByName("Standard")
    for path in sorted(glob.glob(os.path.join(SRC, "*.cls"))) + sorted(glob.glob(os.path.join(SRC, "*.bas"))):
        name = os.path.splitext(os.path.basename(path))[0]
        if path.endswith(".cls"):
            info = ModuleInfo()
            info.ModuleType = uno.getConstantByName("com.sun.star.script.ModuleType.CLASS")
            lib.insertModuleInfo(name, info)
        lib.insertByName(name, vba_source(path))
    lib.insertByName("TestHarness", HARNESS)
    return doc.getScriptProvider().getScript(
        "vnd.sun.star.script:Standard.TestHarness.RunPy?language=Basic&location=document")

def expected_output(pyexe, path):
    out_file = os.path.splitext(path)[0] + ".out"
    if os.path.exists(out_file):
        return open(out_file, encoding="utf-8").read()
    r = subprocess.run([pyexe, "-I", path], capture_output=True, text=True, timeout=60, cwd=os.path.dirname(path))
    out = r.stdout
    if r.returncode != 0:
        lines = [l for l in r.stderr.strip().split("\n") if l.strip()]
        out += lines[-1] + "\n" if lines else ""
    return out

def normalize_actual(text):
    lines = text.split("\n")
    keep = []
    for l in lines:
        if l.startswith("Traceback (most recent call last):") or l.startswith("  File \"<python>\"") or l.startswith("  ... ("):
            continue
        keep.append(l)
    return "\n".join(keep)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-v", action="store_true")
    ap.add_argument("-k", default="")
    ap.add_argument("--python", default=shutil.which("python3.12") or "/usr/local/bin/python3")
    ap.add_argument("files", nargs="*")
    a = ap.parse_args()
    files = a.files or sorted(glob.glob(os.path.join(HERE, "cases", "t*.py")) +
                              glob.glob(os.path.join(HERE, "cases", "x02*.py")))
    files = [os.path.abspath(f) for f in files if a.k in os.path.basename(f)]
    proc, ctx, profile = start_office()
    failures = 0
    try:
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
        run = load_modules(doc)
        for f in files:
            src = open(f, encoding="utf-8").read()
            exp = expected_output(a.python, f)
            t0 = time.time()
            got = run.invoke((src,), (), ())[0]
            dt = time.time() - t0
            ok = normalize_actual(got) == exp
            name = os.path.basename(f)
            print("%-4s %-32s %6.2fs" % ("ok" if ok else "FAIL", name, dt))
            if not ok:
                failures += 1
                import difflib
                for line in difflib.unified_diff(exp.split("\n"), normalize_actual(got).split("\n"),
                                                 "cpython", "vba", lineterm="", n=2):
                    print("     " + line)
                if a.v:
                    print("----- raw vba output -----\n" + got)
        doc.close(True)
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass
        try:
            proc.wait(timeout=20)
        except Exception:
            proc.kill()
        shutil.rmtree(profile, ignore_errors=True)
    print("%d/%d passed" % (len(files) - failures, len(files)))
    sys.exit(1 if failures else 0)

if __name__ == "__main__":
    main()
