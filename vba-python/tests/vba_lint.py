#!/usr/bin/env python3
"""Static checks for VBA compile errors that LibreOffice's Basic does not
report but Excel's VBA compiler does:
  * undeclared identifiers (Option Explicit)
  * ByRef argument type mismatch (typed variable passed to a ByRef param of another type)
  * wrong number of arguments to known procedures
  * Sub called with parentheses around several arguments
  * Exit Function inside a Sub (and vice versa), duplicate declarations
  * lines longer than 1023 characters
usage: python3 vba_lint.py ../src/*.bas ../src/*.cls
"""
import re, sys, os

VBA_WORDS = set("""
and as boolean byref byte byval call case const currency date decimal declare dim do double each else elseif empty
end enum erase error exit explicit false for friend function get global gosub goto if implements in integer is let
like long loop lset me mod new next not nothing null object on option optional or paramarray preserve private
property public raise redim rem resume return rset select set single static step stop string sub then to true type
typeof until variant wend while with withevents xor longlong longptr compare binary text base
abs array asc ascw atn cbool cbyte ccur cdate cdbl cdec chr chr$ chrw chrw$ cint clng csng cstr cvar cverr date$ dateserial
doevents environ eof err exp filter fix format format$ hex hex$ iif instr instrrev int isarray isdate isempty iserror
ismissing isnull isnumeric isobject join lbound lcase lcase$ left left$ len log ltrim ltrim$ mid mid$ now oct oct$ replace
right right$ rnd round rtrim rtrim$ sgn sin space space$ split sqr str str$ strcomp strconv string$ tan timer trim trim$
typename ubound ucase ucase$ val vartype weekday year month day hour minute second randomize msgbox inputbox
vbempty vbnull vbinteger vblong vbsingle vbdouble vbcurrency vbdate vbstring vbobject vberror vbboolean vbvariant
vbdecimal vbbyte vbarray vbcrlf vbcr vblf vbtab vbnullstring vbformfeed vbbinarycompare vbtextcompare vbobjecterror
vbokonly vbinformation vbexclamation vbcritical vbquestion vbyesno vbnewline rgb application thisworkbook activesheet
selection cos
""".split())

def strip_comment(line):
    out, inq = [], False
    for i, ch in enumerate(line):
        if ch == '"':
            inq = not inq
        if ch == "'" and not inq:
            break
        out.append(ch)
    return "".join(out).rstrip()

def logical_lines(text):
    lines = text.replace("\r\n", "\n").split("\n")
    res, buf, start = [], "", 0
    for n, raw in enumerate(lines, 1):
        l = strip_comment(raw)
        if not buf:
            start = n
        if l.endswith(" _"):
            buf += l[:-2] + " "
            continue
        buf += l
        res.append((start, buf, raw))
        buf = ""
    return res

def blank_strings(s):
    return re.sub(r'"(?:[^"]|"")*"', '""', s)

TYPE_RE = r"(?:As\s+(?:New\s+)?([A-Za-z_][A-Za-z0-9_]*))?"
PROC_RE = re.compile(r"^\s*(Public\s+|Private\s+|Friend\s+)?(Static\s+)?(Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+([A-Za-z_][A-Za-z0-9_]*)\s*\((.*)\)\s*(?:As\s+([A-Za-z_][A-Za-z0-9_]*)(\(\))?)?\s*$", re.I)

def split_top(s, sep=","):
    parts, depth, cur, inq = [], 0, "", False
    for ch in s:
        if ch == '"':
            inq = not inq
        if not inq:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == sep and depth == 0:
                parts.append(cur); cur = ""; continue
        cur += ch
    if cur.strip() or parts:
        parts.append(cur)
    return [p.strip() for p in parts]

def parse_params(s):
    params = []
    for p in split_top(s):
        if not p:
            continue
        m = re.match(r"^(Optional\s+)?(ByVal\s+|ByRef\s+)?(ParamArray\s+)?([A-Za-z_][A-Za-z0-9_]*)(\(\))?\s*(?:As\s+([A-Za-z_][A-Za-z0-9_]*))?(\s*=.*)?$", p, re.I)
        if not m:
            raise ValueError("bad param: " + p)
        opt, mode, pa, name, arr, typ, dflt = m.groups()
        params.append(dict(name=name, optional=bool(opt) or bool(pa), byval=(mode or "").strip().lower() == "byval",
                           paramarray=bool(pa), array=bool(arr), type=(typ or "Variant").lower()))
    return params

class Module:
    def __init__(self, path):
        self.path = path
        self.name = os.path.splitext(os.path.basename(path))[0]
        self.is_class = path.endswith(".cls")
        self.lines = logical_lines(open(path, encoding="ascii").read())
        self.module_vars = {}     # name -> type
        self.public = {}
        self.procs = []           # (name, kind, params, rettype, start_idx, end_idx, public)
        self.parse()

    def parse(self):
        i, n = 0, len(self.lines)
        in_header = self.is_class
        while i < n:
            ln, l, raw = self.lines[i]
            s = l.strip()
            if in_header:
                if s.startswith("END"):
                    in_header = False
                i += 1; continue
            m = PROC_RE.match(l)
            if m:
                vis, _, kind, name, params, ret, retarr = m.groups()
                kind = kind.split()[0].lower()
                j = i + 1
                endre = re.compile(r"^\s*End\s+(Sub|Function|Property)\s*$", re.I)
                while j < n and not endre.match(self.lines[j][1]):
                    j += 1
                pub = not (vis and vis.strip().lower() == "private")
                self.procs.append(dict(name=name, kind=kind, params=parse_params(params), ret=(ret or "Variant").lower(),
                                       start=i, end=j, public=pub, module=self))
                i = j + 1; continue
            m = re.match(r"^\s*(Public|Private|Dim|Global)\s+(Const\s+)?(.*)$", l, re.I)
            if m and not s.lower().startswith(("public sub", "private sub", "public function", "private function")):
                vis, const, rest = m.groups()
                for decl in split_top(rest):
                    dm = re.match(r"^([A-Za-z_][A-Za-z0-9_]*)(\(.*?\))?\s*(?:As\s+(?:New\s+)?([A-Za-z_][A-Za-z0-9_]*))?", decl, re.I)
                    if dm:
                        typ = (dm.group(3) or "Variant").lower()
                        self.module_vars[dm.group(1).lower()] = typ
                        if vis.lower() in ("public", "global") and not self.is_class:
                            self.public[dm.group(1).lower()] = typ
            i += 1

def main(paths):
    mods = [Module(p) for p in paths]
    errors = []
    procs = {}        # lower name -> proc (public procs of std modules, plus all procs by module)
    classes = {m.name.lower(): m for m in mods if m.is_class}
    global_vars = {}
    for m in mods:
        if not m.is_class:
            global_vars.update(m.public)
            for p in m.procs:
                if p["public"]:
                    if p["name"].lower() in procs:
                        errors.append("%s: duplicate public procedure %s" % (m.name, p["name"]))
                    procs[p["name"].lower()] = p
    class_members = {}
    for c in classes.values():
        mem = dict(c.module_vars)
        for p in c.procs:
            mem[p["name"].lower()] = p
        class_members[c.name.lower()] = mem

    ident = re.compile(r"(?<![\w.#&])([A-Za-z_][A-Za-z0-9_]*[$]?)")
    for m in mods:
        for ln, l, raw in m.lines:
            if len(raw) > 1023:
                errors.append("%s:%d: line longer than 1023 characters" % (m.name, ln))
        private_procs = {p["name"].lower(): p for p in m.procs}
        for p in m.procs:
            local = {}
            for prm in p["params"]:
                if prm["name"].lower() in local:
                    errors.append("%s.%s: duplicate parameter %s" % (m.name, p["name"], prm["name"]))
                local[prm["name"].lower()] = prm["type"] + ("()" if prm["array"] else "")
            labels = set()
            body = m.lines[p["start"] + 1:p["end"]]
            for ln, l, raw in body:
                s = blank_strings(l).strip()
                lm = re.match(r"^([A-Za-z_][A-Za-z0-9_]*):(\s|$)", s)
                if lm and not re.match(r"^(Case|Else)\b", s, re.I):
                    labels.add(lm.group(1).lower())
                dm = re.match(r"^(Dim|Static|Const)\s+(.*)$", s, re.I)
                if dm:
                    for decl in split_top(dm.group(2)):
                        d2 = re.match(r"^([A-Za-z_][A-Za-z0-9_]*)(\(.*?\))?\s*(?:As\s+(?:New\s+)?([A-Za-z_][A-Za-z0-9_]*))?", decl, re.I)
                        if d2:
                            nm = d2.group(1).lower()
                            if nm in local:
                                errors.append("%s.%s:%d: duplicate declaration %s" % (m.name, p["name"], ln, nm))
                            local[nm] = (d2.group(3) or "Variant").lower() + ("()" if d2.group(2) else "")
            def var_type(nm):
                nm = nm.lower()
                if nm in local: return local[nm]
                if nm in m.module_vars: return m.module_vars[nm]
                if nm in global_vars: return global_vars[nm]
                return None
            def known(nm):
                k = nm.lower().rstrip("$")
                return (k in local or k in labels or k in m.module_vars or k in global_vars or k in procs
                        or k in private_procs or k in classes or k in VBA_WORDS or nm.lower() in VBA_WORDS
                        or (m.is_class and k in class_members.get(m.name.lower(), {})) or k == p["name"].lower())
            def lookup_proc(nm):
                k = nm.lower()
                return private_procs.get(k) or procs.get(k)
            for ln, l, raw in body:
                s = blank_strings(l).strip()
                if not s:
                    continue
                if re.match(r"^Exit\s+Function", s, re.I) and p["kind"] != "function":
                    errors.append("%s.%s:%d: Exit Function in a %s" % (m.name, p["name"], ln, p["kind"]))
                if re.match(r"^Exit\s+Sub", s, re.I) and p["kind"] != "sub":
                    errors.append("%s.%s:%d: Exit Sub in a %s" % (m.name, p["name"], ln, p["kind"]))
                code = re.sub(r"\b[A-Za-z_]\w*:=", "", s)   # named arguments
                for tok in ident.finditer(code):
                    nm = tok.group(1)
                    if not known(nm):
                        errors.append("%s.%s:%d: undeclared identifier '%s'" % (m.name, p["name"], ln, nm))
                # statements split by ':' (outside strings) for call checks
                for stmt in split_top(re.sub(r"^\s*(Else)?If .*? Then\s+", "", s), ":"):
                    stmt = re.sub(r"^\s*Else\s+", "", stmt.strip())
                    check_call_stmt(stmt, m, p, lookup_proc, var_type, errors, ln)
                # function calls inside expressions: Name(args)
                for fm in re.finditer(r"(?<![\w.])([A-Za-z_]\w*)\s*\(", code):
                    pr = lookup_proc(fm.group(1))
                    if pr and pr["kind"] == "function":
                        args = extract_paren(code, fm.end() - 1)
                        if args is not None:
                            check_args(pr, split_top(args) if args.strip() else [], m, p, var_type, errors, ln)
    for e in errors:
        print(e)
    print("%d problem(s)" % len(errors))
    return 1 if errors else 0

def extract_paren(s, i):
    depth, inq = 0, False
    for j in range(i, len(s)):
        ch = s[j]
        if ch == '"': inq = not inq
        if inq: continue
        if ch == "(": depth += 1
        elif ch == ")":
            depth -= 1
            if depth == 0:
                return s[i + 1:j]
    return None

def check_call_stmt(stmt, m, p, lookup_proc, var_type, errors, ln):
    if not stmt:
        return
    km = re.match(r"^(Call\s+)?([A-Za-z_]\w*)(.*)$", stmt, re.I)
    if not km:
        return
    is_call, name, rest = km.groups()
    if re.match(r"^\s*(=|\.|\()", rest) and not (rest.strip().startswith("(") and not re.search(r"\)\s*=", rest)):
        return
    pr = lookup_proc(name)
    if not pr:
        return
    rest = rest.strip()
    if is_call:
        args = extract_paren(rest, 0) if rest.startswith("(") else ""
        args = split_top(args) if args and args.strip() else []
    else:
        if rest.startswith("(") and extract_paren(rest, 0) is not None and rest.endswith(")") and len(split_top(extract_paren(rest, 0))) > 1:
            errors.append("%s.%s:%d: Sub call with parentheses around several arguments: %s" % (m.name, p["name"], ln, stmt))
            return
        args = split_top(rest) if rest else []
    check_args(pr, args, m, p, var_type, errors, ln)

def check_args(pr, args, m, p, var_type, errors, ln):
    params = pr["params"]
    has_pa = any(x["paramarray"] for x in params)
    required = sum(1 for x in params if not x["optional"])
    if len(args) < required or (not has_pa and len(args) > len(params)):
        errors.append("%s.%s:%d: %s expects %d..%d args, got %d" % (m.name, p["name"], ln, pr["name"], required, len(params), len(args)))
        return
    for a, prm in zip(args, params):
        if prm["paramarray"]:
            break
        if prm["byval"] or prm["type"] == "variant" and not prm["array"]:
            continue
        am = re.match(r"^([A-Za-z_]\w*)$", a)
        if not am:
            continue
        t = var_type(am.group(1))
        if t is None:
            continue
        want = prm["type"] + ("()" if prm["array"] else "")
        if t != want:
            errors.append("%s.%s:%d: ByRef argument type mismatch: %s (%s) passed to %s %s As %s" % (
                m.name, p["name"], ln, am.group(1), t, pr["name"], prm["name"], want))

if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
