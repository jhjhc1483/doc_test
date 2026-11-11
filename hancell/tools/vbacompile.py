"""진짜 VBA가 '컴파일' 단계에서 거부하는 실수를 찾습니다.

LibreOffice 는 줄을 실행할 때만 검사하므로, 한 번도 실행되지 않은 줄의 실수는 지나칩니다.
한셀(VBA)은 모듈 전체를 먼저 검사하고 하나라도 틀리면 그 모듈을 통째로 실행하지 않습니다.

- 선언하지 않은 이름 (Option Explicit)
- For Each 의 변수가 Variant/Object 가 아님
- 다른 모듈의 Private 프로시저·변수를 씀
- 반환값을 받지 않는 호출에 괄호와 쉼표 (Foo(a, b))
"""
import re
import sys

KEYWORDS = set('''
if then else elseif end sub function for each in to step next do loop while until wend select case is not and or xor
mod dim as new set let exit call on error resume goto with optional byval byref public private const true false nothing
empty null me redim preserve option explicit boolean long integer double string variant object date single byte currency
like static type attribute eqv imp
'''.split())

BUILTINS = set('''
len mid mid$ left left$ right right$ instr instrrev replace split join trim trim$ ltrim rtrim ucase ucase$ lcase lcase$
chr chr$ chrw chrw$ asc ascw val cstr clng cint cdbl cdate cbool csng cbyte ccur isnumeric isarray isempty isnull iserror
isobject isdate typename vartype ubound lbound array int fix round abs sgn sqr format format$ now time year month
day weekday hour minute second dateserial timeserial dateadd datediff msgbox inputbox rgb iif choose environ environ$
space string doevents timer err debug application activesheet activeworkbook thisworkbook activewindow selection
worksheets workbooks sheets range cells rows columns collection vblf vbcr vbcrlf vbtab vbyesno vbyes vbno vbquestion
vbinformation vbstring vbdate vbboolean vbnullstring vbexclamation vbcritical vbokonly vbokcancel vbcancel print
createobject getobject strconv filelen dir kill
'''.split())

PROC = re.compile(r'^\s*(Public |Private )?(Sub|Function) (\w+)\s*\((.*)\)(\s+As\s+\w+)?\s*$', re.I)
END_PROC = re.compile(r'^\s*End (Sub|Function)\b', re.I)
IDENT = re.compile(r'(?<![\w.$])([A-Za-z_]\w*\$?)')


def strip(line):
    out = re.sub(r'"(?:[^"]|"")*"', '""', line)
    q = out.find("'")
    if q >= 0:
        out = out[:q]
    return out


def joined_lines(text):
    lines = text.replace('\r\n', '\n').split('\n')
    out, buf, start = [], '', 0
    for i, l in enumerate(lines):
        if not buf:
            start = i + 1
        s = strip(l)
        if s.rstrip().endswith(' _'):
            buf += s.rstrip()[:-1] + ' '
            continue
        out.append((start, buf + s))
        buf = ''
    return out


def decl_names(rest):
    rest = re.sub(r'\([^()]*\)', '', rest)
    names = []
    for part in rest.split(','):
        part = re.sub(r'^\s*(WithEvents|Preserve)\s+', '', part.strip(), flags=re.I)
        m = re.match(r'(\w+)(\s+As\s+(New\s+)?(\w+))?', part, re.I)
        if m:
            names.append((m.group(1), (m.group(4) or 'Variant')))
    return names


def params_of(sig):
    out = []
    for part in sig.split(','):
        part = re.sub(r'\b(Optional|ByVal|ByRef|ParamArray)\b', '', part, flags=re.I).strip()
        part = part.split('=')[0]
        m = re.match(r'(\w+)(\(\))?(\s+As\s+(\w+))?', part.strip(), re.I)
        if m:
            out.append((m.group(1), m.group(4) or 'Variant'))
    return out


def params_full(sig):
    """[(name, type, byval, isarray)]"""
    out = []
    for part in sig.split(','):
        if not part.strip():
            continue
        byval = bool(re.search(r'\bByVal\b', part, re.I))
        q = re.sub(r'\b(Optional|ByVal|ByRef|ParamArray)\b', '', part, flags=re.I).strip().split('=')[0]
        m = re.match(r'(\w+)(\s*\(\))?(\s+As\s+(\w+))?', q.strip(), re.I)
        if m:
            out.append((m.group(1), (m.group(4) or 'Variant').lower(), byval, bool(m.group(2))))
    return out


def split_args(s):
    args, depth, cur = [], 0, ''
    for ch in s:
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
        if ch == ',' and depth == 0:
            args.append(cur.strip())
            cur = ''
        else:
            cur += ch
    if cur.strip() or args:
        args.append(cur.strip())
    return args


def scan(files):
    mods = {}
    for fn in files:
        lines = joined_lines(open(fn, encoding='utf-8').read())
        mod = {'file': fn, 'vars': {}, 'procs': {}, 'body': []}
        cur = None
        for no, l in lines:
            pm = PROC.match(l)
            if pm:
                vis = (pm.group(1) or 'Public ').strip().lower()
                cur = {'name': pm.group(3), 'vis': vis, 'params': params_of(pm.group(4)), 'lines': [], 'no': no,
                       'sig': params_full(pm.group(4)), 'kind': pm.group(2).lower()}
                mod['procs'][pm.group(3).lower()] = cur
                continue
            if cur is not None:
                if END_PROC.match(l):
                    cur = None
                else:
                    cur['lines'].append((no, l))
                continue
            m = re.match(r'^\s*(Public|Private|Dim|Global)\s+(Const\s+)?(.*)$', l, re.I)
            if m and not re.match(r'^\s*(Public|Private)\s+(Sub|Function)', l, re.I):
                vis = 'private' if m.group(1).lower() in ('private', 'dim') else 'public'
                for nm, _t in decl_names(re.sub(r'=.*$', '', m.group(3))):
                    mod['vars'][nm.lower()] = vis
            m = re.match(r'^\s*Const\s+(\w+)', l, re.I)
            if m:
                mod['vars'][m.group(1).lower()] = 'private'
        mods[fn] = mod
    return mods


def check(files):
    mods = scan(files)
    errors = []
    public_names = {}
    for fn, m in mods.items():
        for k, v in m['vars'].items():
            if v == 'public':
                public_names[k] = fn
        for k, p in m['procs'].items():
            if p['vis'] == 'public':
                public_names[k] = fn
    sigs = {}
    for fn, m in mods.items():
        for k, p in m['procs'].items():
            sigs[k] = p['sig']
    for fn, m in mods.items():
        own = set(m['vars']) | set(m['procs'])
        for pname, p in m['procs'].items():
            local = {}
            for nm, t in p['params']:
                local[nm.lower()] = t
            labels = set()
            arrays = set(nm.lower() for nm, t, bv, isarr in p['sig'] if isarr)
            for no, l in p['lines']:
                for st in [x.strip() for x in re.split(r':(?!=)', l)]:
                    dm = re.match(r'^(Dim|Static|ReDim(\s+Preserve)?|Const)\s+(.*)$', st, re.I)
                    if dm:
                        rest = dm.group(3)
                        if dm.group(1).lower() == 'const':
                            rest = re.sub(r'=.*$', '', rest)
                        for part in rest.split(','):
                            mm = re.match(r'\s*(\w+)\s*\(', part)
                            if mm:
                                arrays.add(mm.group(1).lower())
                        for nm, t in decl_names(rest):
                            local.setdefault(nm.lower(), t)
                    if re.match(r'^Exit\s+(Sub|Function)\b', st, re.I):
                        kind = re.match(r'^Exit\s+(Sub|Function)', st, re.I).group(1).lower()
                        if kind != p['kind']:
                            errors.append('%s:%d: %s 안에서 Exit %s' % (fn, no, p['kind'], kind))
                    lm = re.match(r'^(\w+):$', l.strip())
                    if lm:
                        labels.add(lm.group(1).lower())
            for no, l in p['lines']:
                code = l
                # For Each control variable
                for fm in re.finditer(r'\bFor\s+Each\s+(\w+)\s+In\b', code, re.I):
                    v = fm.group(1).lower()
                    t = local.get(v)
                    if t is None and v in m['vars']:
                        t = 'Variant'
                    if t and t.lower() not in ('variant', 'object'):
                        errors.append('%s:%d: For Each 변수 %s 는 %s — Variant 여야 합니다' % (fn, no, fm.group(1), t))
                # call with parentheses and commas as a statement
                for st in [x.strip() for x in re.split(r':(?!=)', code)]:
                    cm = re.match(r'^(\w+)\s*\((.*)\)$', st)
                    if cm and re.match(r'^\w+\s*\([^)]*\)\s*\.', st):
                        cm = None   # obj(i).Method ...: not a call with parentheses
                    if cm and '=' not in st.split(')')[-1] and cm.group(1).lower() not in KEYWORDS:
                        depth, comma = 0, False
                        for ch in cm.group(2):
                            if ch == '(':
                                depth += 1
                            elif ch == ')':
                                depth -= 1
                            elif ch == ',' and depth == 0:
                                comma = True
                        if comma and not re.match(r'^\w+\s*\(.*\)\s*=', st):
                            errors.append('%s:%d: 값을 받지 않는 호출에 괄호: %s' % (fn, no, st[:60]))
                # ByRef argument types
                for st in [x.strip() for x in re.split(r':(?!=)', code)]:
                    for cm in re.finditer(r'(?<![\w.])(\w+)(\s*\(|\s+)(?=\S)', st):
                        callee = cm.group(1).lower()
                        if callee not in sigs or callee in KEYWORDS:
                            continue
                        rest = st[cm.end():] if cm.group(2).strip() == '' else st[cm.end():]
                        if cm.group(2).strip() == '(':
                            depth, j = 1, 0
                            while j < len(rest) and depth:
                                if rest[j] == '(':
                                    depth += 1
                                elif rest[j] == ')':
                                    depth -= 1
                                j += 1
                            argstr = rest[:j - 1]
                        else:
                            if not st.lower().startswith(callee):
                                continue
                            argstr = rest
                        args = split_args(argstr)
                        for (pn, pt, byval, isarr), a in zip(sigs[callee], args):
                            if byval or not re.fullmatch(r'\w+', a or ''):
                                continue
                            ak = a.lower()
                            at = local.get(ak)
                            if at is None:
                                continue
                            at = at.lower()
                            if isarr and ak not in arrays and at != 'variant':
                                errors.append('%s:%d: %s 에 배열이 아닌 %s 를 넘김' % (fn, no, callee, a))
                            elif isarr and ak not in arrays:
                                errors.append('%s:%d: %s 의 배열 인자 %s 에 Variant %s 를 넘김 (형식 불일치)' % (fn, no, callee, pn, a))
                            elif not isarr and pt != 'variant' and at != pt and ak not in arrays:
                                errors.append('%s:%d: ByRef 인자 형식 불일치: %s(%s As %s) 에 %s As %s' % (fn, no, callee, pn, pt, a, at))
                # identifiers
                if re.match(r'^\s*\w+:\s*$', code):
                    continue
                for im in IDENT.finditer(code):
                    name = im.group(1)
                    k = name.lower()
                    after = code[im.end():im.end() + 3]
                    if k in KEYWORDS or k in BUILTINS or k.rstrip('$') in BUILTINS:
                        continue
                    if k in local or k in labels or k == pname:
                        continue
                    if k in own:
                        continue
                    if k in public_names:
                        continue
                    # named argument  Foo:=
                    if after.startswith(':='):
                        continue
                    other = [f for f, mm in mods.items() if f != fn and (k in mm['procs'] or k in mm['vars'])]
                    if other:
                        errors.append('%s:%d: 다른 모듈의 Private 이름 %s (%s)' % (fn, no, name, other[0]))
                    else:
                        errors.append('%s:%d: 선언하지 않은 이름 %s' % (fn, no, name))
    # one message per (file, name)
    seen, out = set(), []
    for e in errors:
        key = re.sub(r':\d+:', ':', e)
        if key not in seen:
            seen.add(key)
            out.append(e)
    return out


if __name__ == '__main__':
    errs = check(sys.argv[1:])
    for e in errs:
        print(e)
    sys.exit(1 if errs else 0)
