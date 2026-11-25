#!/usr/bin/env python3
"""Run MiniLisp inside headless LibreOffice (VBA compatibility mode) and
compare the output of every cases/*.scm with cases/*.out.

A case whose first line is ";; echo" is run like a REPL (the value of every
top-level expression is printed). A case whose first line is ";; each" runs
every non-empty line on its own, in a fresh interpreter, REPL style - handy
for error messages, because an error stops the rest of a program.
After the cases, a few checks drive the Excel entry points (RunSheet, the
=LISP() worksheet function, cell / set-cell!).

usage: /usr/bin/python3 run_tests.py [-v] [-k substring]
Needs LibreOffice and its Python-UNO bridge (python3-uno).
"""
import argparse, difflib, glob, os, shutil, socket, subprocess, sys, tempfile, time

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.environ.get("MINILISP_SRC", os.path.join(HERE, "..", "src"))

import uno  # noqa: E402  (provided by LibreOffice)
from com.sun.star.script import ModuleInfo  # noqa: E402

HARNESS = """Option VBASupport 1
Option Explicit
Function RunTest(ByVal src As String, ByVal echo As Boolean) As String
    On Error GoTo EH
    RunTest = LispRun(src, echo)
    Exit Function
EH:
    RunTest = "!!VBA ERROR " & Err.Number & ": " & Err.Description
End Function
Function CallLisp(ByVal src As String) As Variant
    CallLisp = LISP(src)
End Function
Sub CallRunSheet()
    RunSheet
End Sub
"""

def read_source(path):
    """VBA sources are GBK (what the VBA editor on Chinese Windows expects); accept UTF-8 too."""
    raw = open(path, "rb").read()
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("cp936")

def vba_source(path):
    text = read_source(path).replace("\r\n", "\n")
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
    sp = doc.getScriptProvider()
    get = lambda n: sp.getScript("vnd.sun.star.script:Standard.TestHarness.%s?language=Basic&location=document" % n)
    return get("RunTest"), get("CallLisp"), get("CallRunSheet")

def run_case(run, src):
    first = src.split("\n", 1)[0].strip()
    if first == ";; each":
        out = []
        for line in src.split("\n")[1:]:
            if line.strip() and not line.strip().startswith(";"):
                out.append(run.invoke((line, True), (), ())[0])
        return "".join(out)
    return run.invoke((src, first == ";; echo"), (), ())[0]

def report(name, ok, exp, got, dt, verbose):
    print("%-4s %-28s %6.2fs" % ("ok" if ok else "FAIL", name, dt))
    if not ok:
        exp, got = str(exp), str(got)
        for line in difflib.unified_diff(exp.split("\n"), got.split("\n"), "expected", "vba", lineterm="", n=2):
            print("     " + line)
        if verbose:
            print("----- raw vba output -----\n" + got)

def excel_checks(doc, run, call_lisp, call_runsheet):
    """(name, expected, got) for the Excel entry points."""
    sheet = doc.Sheets.getByIndex(0)
    cell = lambda a: sheet.getCellRangeByName(a)
    results = []
    # cell / set-cell!
    cell("D1").setValue(41)
    cell("D2").setString("Ada")
    got = run.invoke(('(display (+ (cell "D1") 1)) (display (cell "D2")) (display (cell "D3"))'
                      ' (set-cell! "E1" (* 6 7)) (set-cell! "E2" "hi") (set-cell! "E3" (list 1 2))', False), (), ())[0]
    results.append(("cell", "42Ada#f", got))
    results.append(("set-cell!", "42|hi|(1 2)",
                    "%g|%s|%s" % (cell("E1").getValue(), cell("E2").getString(), cell("E3").getString())))
    # =LISP() worksheet function
    got = [call_lisp.invoke((s,), (), ())[0] for s in
           ["(+ 1 2)", "(define (sq x) (* x x)) (sq 12)", '(string-append "a" "b")', "(list 1 (quote x))",
            "(> 1 2)", '(display "out")', "(car '())"]]
    results.append(("LISP()", [3, 144, "ab", "(1 x)", False, "out", "Error: car: empty list"], got))
    # RunSheet: column A holds code, results go to column B; definitions carry over
    rows = ["(define (sq x) (* x x))", "(sq 9)", "", '(display "hi") (newline) (+ 1 2)', "(oops)", "(sq 3)"]
    for i, code in enumerate(rows):
        if code:
            cell("A%d" % (i + 1)).setString(code)
    call_runsheet.invoke((), (), ())
    got = [cell("B%d" % (i + 1)).getString() for i in range(len(rows))]
    results.append(("RunSheet", repr(["", "81", "", "hi\n3", "Error: unbound variable: oops", "9"]), repr(got)))
    return results

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-v", action="store_true")
    ap.add_argument("-k", default="")
    a = ap.parse_args()
    files = sorted(f for f in glob.glob(os.path.join(HERE, "cases", "*.scm")) if a.k in os.path.basename(f))
    proc, ctx, profile = start_office()
    total = failures = 0
    try:
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
        run, call_lisp, call_runsheet = load_modules(doc)
        for f in files:
            src = open(f, encoding="utf-8").read()
            exp = open(os.path.splitext(f)[0] + ".out", encoding="utf-8").read()
            t0 = time.time()
            got = run_case(run, src)
            ok = got == exp
            total += 1; failures += not ok
            report(os.path.basename(f), ok, exp, got, time.time() - t0, a.v)
        if a.k in "excel":
            for name, exp, got in excel_checks(doc, run, call_lisp, call_runsheet):
                ok = exp == got
                total += 1; failures += not ok
                report("excel: " + name, ok, exp, got, 0, a.v)
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
    print("%d/%d passed" % (total - failures, total))
    sys.exit(1 if failures else 0)

if __name__ == "__main__":
    main()
