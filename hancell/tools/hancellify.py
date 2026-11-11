"""읽기 쉬운 원본(src/*.bas)을 한셀 2020 매크로 편집기에 붙여넣을 모듈로 바꿉니다.

한셀 구문 점검(hancell/한셀구문점검.txt)에서 알게 된 것:
  - 됨: 줄 이어 쓰기, Do While/Exit Do, 모듈 Const·변수, 배열을 돌려주는 Function,
        Choose·Format$·Val·$함수, Optional 인자, ReDim·Preserve, 배열 인자, Select Case,
        With, On Error GoTo, 2차원 배열, GoTo
  - 한셀이 꺼짐: Collection (For Each, 빠진 인자와 함께 씀)
  - 큰 모듈은 붙여넣어도 아무것도 실행되지 않았음

그래서 붙여넣기 파일은 확인된 문법만 쓰도록 바꿉니다.
  1. 줄 이어 쓰기( _)를 한 줄로 합칩니다.
  2. 주석을 뺍니다.
  3. Const 이름을 값으로 바꿉니다 (모듈 사이 상수를 쓰지 않음).
  4. ':' 로 이은 문장은 한 줄에 하나씩 씁니다. 한 줄 If 에 문장이 여럿이면 블록 If 로 바꿉니다.
  5. 프로시저를 '앞 모듈만 부르는' 순서로 작은 모듈 여러 개에 나눠 담습니다.
     그래서 모듈을 하나씩 붙여넣고 그 모듈의 확인 매크로(CheckN)로 바로 확인할 수 있습니다.
  6. Collection, For Each, New, 이름 붙인 인자(:=), 빠진 인자가 남아 있으면 멈춥니다.
"""
import re

LIMIT = 22000          # 모듈 하나의 글자 수 한도 (붙여넣기 창 32767자 제한보다 넉넉히 작게)

KEYWORDS = {
    'if', 'then', 'else', 'elseif', 'end', 'for', 'next', 'to', 'step', 'do', 'loop', 'while', 'until', 'wend',
    'select', 'case', 'with', 'dim', 'redim', 'preserve', 'as', 'byval', 'byref', 'optional', 'public', 'private',
    'sub', 'function', 'exit', 'call', 'set', 'let', 'and', 'or', 'not', 'xor', 'mod', 'is', 'like', 'new',
    'nothing', 'true', 'false', 'empty', 'null', 'on', 'error', 'resume', 'goto', 'const', 'string', 'long',
    'double', 'integer', 'boolean', 'variant', 'object', 'date', 'single', 'byte', 'currency', 'me', 'each',
    'in', 'option', 'explicit', 'static', 'type', 'enum', 'property', 'get', 'friend', 'global',
}

PROC_START = re.compile(r'^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function)\s+(\w+)', re.I)
PROC_END = re.compile(r'^\s*End\s+(Sub|Function)\s*$', re.I)
CONST_DECL = re.compile(r'^\s*(?:(?:Public|Private|Global)\s+)?Const\s+(\w+)(?:\s+As\s+\w+)?\s*=\s*(.+?)\s*$', re.I)
IDENT = re.compile(r'[A-Za-z_][A-Za-z0-9_]*')


class BuildError(Exception):
    pass


def mask(code):
    """문자열 안의 글자를 \\0 으로 가린 같은 길이의 문자열 (따옴표는 남김)"""
    out = []
    ins = False
    i = 0
    n = len(code)
    while i < n:
        ch = code[i]
        if ins:
            if ch == '"':
                if i + 1 < n and code[i + 1] == '"':
                    out.append('\0\0')
                    i += 2
                    continue
                ins = False
                out.append('"')
            else:
                out.append('\0')
        else:
            if ch == '"':
                ins = True
            out.append(ch)
        i += 1
    if ins:
        raise BuildError('따옴표가 닫히지 않음: ' + code)
    return ''.join(out)


def strip_comment(line):
    ins = False
    for i, ch in enumerate(line):
        if ch == '"':
            ins = not ins          # "" 안의 따옴표 두 개는 두 번 뒤집혀 그대로
        elif ch == "'" and not ins:
            line = line[:i]
            break
    else:
        if ins:
            raise BuildError('따옴표가 닫히지 않음: ' + line)
    if re.match(r'^\s*Rem\b', line, re.I):
        return ''
    return line.rstrip()


def join_continuations(lines):
    out = []
    buf = None
    for ln in lines:
        code = strip_comment(ln)
        if buf is not None:
            code = buf + ' ' + code.lstrip()
            buf = None
        if re.search(r'(^|\s)_$', code):
            buf = code[:-1].rstrip()
            continue
        out.append(code)
    if buf is not None:
        raise BuildError('줄 이어 쓰기로 끝남')
    return out


def top_level_splits(code, sep_re):
    """문자열과 괄호 밖에서 sep_re 가 맞는 위치들"""
    m = mask(code)
    depth = 0
    pos = []
    i = 0
    while i < len(m):
        ch = m[i]
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
        elif depth == 0:
            mm = sep_re.match(m, i)
            if mm:
                pos.append((i, mm.end()))
                i = mm.end()
                continue
        i += 1
    return pos


COLON = re.compile(r':(?!=)')
ELSE_KW = re.compile(r'(?<![A-Za-z0-9_])Else(?![A-Za-z0-9_])', re.I)


def split_colons(code):
    parts = []
    last = 0
    for a, b in top_level_splits(code, COLON):
        parts.append(code[last:a])
        last = b
    parts.append(code[last:])
    return [p.strip() for p in parts if p.strip()]


def is_label(code):
    m = re.match(r'^\s*([A-Za-z_]\w*):\s*$', code)
    return bool(m) and m.group(1).lower() not in KEYWORDS


def expand_statement(code, indent):
    """한 줄(주석 없음)을 확인된 문법의 여러 줄로"""
    if not code.strip():
        return []
    if is_label(code):
        return [code.strip()]
    m = re.match(r'^\s*([A-Za-z_]\w*):(?!=)\s*(\S.*)$', code)
    if m and m.group(1).lower() not in KEYWORDS:
        return [m.group(1) + ':'] + expand_statement(m.group(2), indent)
    mk = mask(code)
    mi = re.match(r'^\s*If\b', mk, re.I)
    if mi:
        tpos = [p for p in top_level_splits(code, re.compile(r'(?<![A-Za-z0-9_])Then(?![A-Za-z0-9_])', re.I))]
        if not tpos:
            raise BuildError('Then 이 없는 If: ' + code)
        a, b = tpos[0]
        rest = code[b:].strip()
        if rest:
            head = code[:b].strip()
            elses = top_level_splits(rest, ELSE_KW)
            if elses:
                ea, eb = elses[0]
                then_part, else_part = rest[:ea].strip(), rest[eb:].strip()
            else:
                then_part, else_part = rest, None
            then_stmts = split_colons(then_part)
            if len(then_stmts) == 1 and else_part is None and not re.match(r'^\s*If\b', then_stmts[0], re.I):
                return [indent + head + ' ' + then_stmts[0]]
            out = [indent + head]
            for s in then_stmts:
                out += expand_statement(s, indent + '    ')
            if else_part is not None:
                out.append(indent + 'Else')
                for s in split_colons(else_part):
                    out += expand_statement(s, indent + '    ')
            out.append(indent + 'End If')
            return out
    stmts = split_colons(code)
    if len(stmts) == 1:
        return [indent + stmts[0]]
    out = []
    first = stmts[0]
    out += expand_statement(first, indent)
    inner = indent
    if re.match(r'^(Case|Else)\b', first, re.I):
        inner = indent + '    '
    for s in stmts[1:]:
        out += expand_statement(s, inner)
    return out


def indent_of(line):
    return line[:len(line) - len(line.lstrip())]


class Proc:
    def __init__(self, name, kind, vis, lines, module):
        self.name = name
        self.kind = kind
        self.vis = vis
        self.lines = lines
        self.module = module
        self.calls = set()

    def size(self):
        return sum(len(l) + 2 for l in self.lines) + 2


def parse_module(name, text):
    lines = text.replace('\r\n', '\n').split('\n')
    lines = [l for l in lines if not l.startswith('Attribute VB_')]
    lines = join_continuations(lines)
    consts = {}
    decls = []
    procs = []
    cur = None
    for ln in lines:
        if cur is None:
            if not ln.strip():
                continue
            m = PROC_START.match(ln)
            if m:
                cur = Proc(m.group(3), m.group(2), (m.group(1) or 'Public').capitalize(), [ln.strip()], name)
                continue
            if re.match(r'^\s*Option\s', ln, re.I):
                continue
            mc = CONST_DECL.match(ln)
            if mc:
                if ',' in [c for c in mask(mc.group(2)) if c not in '\0']:
                    raise BuildError('%s: Const 한 줄에 하나만: %s' % (name, ln))
                consts[mc.group(1)] = mc.group(2).strip()
                continue
            if re.match(r'^\s*(Public|Private|Dim|Global)\s', ln, re.I):
                if re.match(r'^\s*(Public|Private)\s+(Type|Enum|Declare)\b', ln, re.I):
                    raise BuildError('%s: Type/Enum/Declare 는 쓰지 않음: %s' % (name, ln))
                decls.append(re.sub(r'^\s*(Public|Private|Dim|Global)\s+', 'Public ', ln.strip(), flags=re.I))
                continue
            raise BuildError('%s: 모듈 맨 위에 알 수 없는 줄: %s' % (name, ln))
        else:
            if PROC_END.match(ln):
                cur.lines.append(ln.strip())
                procs.append(cur)
                cur = None
                continue
            if not ln.strip():
                continue
            if re.match(r'^\s*Const\s', ln, re.I):
                raise BuildError('%s.%s: 프로시저 안 Const 는 쓰지 않음' % (name, cur.name))
            cur.lines.append(ln)
    if cur is not None:
        raise BuildError('%s: %s 가 끝나지 않음' % (name, cur.name))
    return consts, decls, procs


def replace_idents(code, mapping):
    """문자열 밖, '.' 뒤가 아닌 이름을 mapping(소문자 키)으로 바꿈"""
    mk = mask(code)
    out = []
    last = 0
    for m in IDENT.finditer(mk):
        key = m.group(0).lower()
        if key in mapping:
            a = m.start()
            if a > 0 and mk[a - 1] == '.':
                continue
            if a > 0 and (mk[a - 1].isalnum() or mk[a - 1] == '_'):
                continue
            out.append(code[last:a])
            out.append(mapping[key])
            last = m.end()
    out.append(code[last:])
    return ''.join(out)


def resolve_consts(consts):
    low = {k.lower(): v for k, v in consts.items()}
    for _ in range(10):
        changed = False
        for k, v in list(low.items()):
            nv = replace_idents(v, low)
            if nv != v:
                low[k] = nv
                changed = True
        if not changed:
            break
    return low


def local_names(proc):
    """매개변수와 Dim 한 이름 (소문자)"""
    names = set()
    head = mask(proc.lines[0])
    m = re.search(r'\((.*)\)', head)
    if m:
        for part in m.group(1).split(','):
            part = re.sub(r'^\s*(Optional\s+)?(ByVal\s+|ByRef\s+)?(ParamArray\s+)?', '', part, flags=re.I)
            mm = IDENT.match(part)
            if mm:
                names.add(mm.group(0).lower())
    for ln in proc.lines[1:]:
        mk = mask(ln)
        for st in split_colons(mk):
            md = re.match(r'^\s*(Dim|ReDim(\s+Preserve)?)\s+(.*)$', st, re.I)
            if md:
                body = re.sub(r'\([^()]*\)', '', md.group(3))
                for part in body.split(','):
                    mm = IDENT.match(part.strip())
                    if mm:
                        names.add(mm.group(0).lower())
    return names


FORBIDDEN = [
    (re.compile(r'\bCollection\b', re.I), 'Collection (한셀이 꺼짐)'),
    (re.compile(r'\bFor\s+Each\b', re.I), 'For Each'),
    (re.compile(r'\bNew\b', re.I), 'New'),
    (re.compile(r':='), '이름 붙인 인자 (:=)'),
    (re.compile(r'\(\s*,|,\s*,|,\s*\)'), '빠진 인자'),
    (re.compile(r'\.Value\s*=\s*Array\(', re.I), 'Range.Value = Array(...)'),
    (re.compile(r'(^|\s)_$'), '줄 이어 쓰기'),
]


def check_forbidden(lines, where):
    errs = []
    for ln in lines:
        mk = mask(ln)
        mk = re.sub(r'\0+', '', mk)
        for rx, what in FORBIDDEN:
            if rx.search(mk):
                errs.append('%s: %s 가 남아 있음: %s' % (where, what, ln.strip()))
        if len(ln) > 1000:
            errs.append('%s: 줄이 너무 김(%d자)' % (where, len(ln)))
    return errs


def tarjan(nodes, edges):
    index = {}
    low = {}
    stack = []
    on = set()
    sccs = []
    counter = [0]

    def strong(v):
        # 반복문으로 (재귀 깊이 제한을 피함)
        work = [(v, iter(sorted(edges[v], key=lambda x: order[x])))]
        index[v] = low[v] = counter[0]
        counter[0] += 1
        stack.append(v)
        on.add(v)
        while work:
            node, it = work[-1]
            advanced = False
            for w in it:
                if w not in index:
                    index[w] = low[w] = counter[0]
                    counter[0] += 1
                    stack.append(w)
                    on.add(w)
                    work.append((w, iter(sorted(edges[w], key=lambda x: order[x]))))
                    advanced = True
                    break
                elif w in on:
                    low[node] = min(low[node], index[w])
            if advanced:
                continue
            work.pop()
            if work:
                low[work[-1][0]] = min(low[work[-1][0]], low[node])
            if low[node] == index[node]:
                comp = []
                while True:
                    w = stack.pop()
                    on.discard(w)
                    comp.append(w)
                    if w == node:
                        break
                sccs.append(comp)
    order = {n: i for i, n in enumerate(nodes)}
    for v in nodes:
        if v not in index:
            strong(v)
    return sccs        # 부르는 쪽보다 불리는 쪽이 먼저 나옴


def build_modules(sources, title, drop=()):
    """sources: [(모듈 이름, 원본 글)] -> [모듈 글] (붙여넣기 순서)"""
    consts = {}
    decls = []
    procs = []
    for name, text in sources:
        c, d, p = parse_module(name, text)
        for k in c:
            if k.lower() in {x.lower() for x in consts}:
                raise BuildError('Const 이름이 겹침: ' + k)
        consts.update(c)
        decls += d
        procs += [x for x in p if x.name not in drop]
    cmap = resolve_consts(consts)

    # 이름 겹침 검사
    seen = {}
    for p in procs:
        k = p.name.lower()
        if k in seen:
            raise BuildError('프로시저 이름이 겹침: %s (%s, %s)' % (p.name, seen[k], p.module))
        seen[k] = p.module
    gnames = set()
    for d in decls:
        body = re.sub(r'\([^()]*\)', '', mask(d)[len('Public '):])
        for part in body.split(','):
            mm = IDENT.match(part.strip())
            if mm:
                n = mm.group(0).lower()
                if n in gnames or n in seen:
                    raise BuildError('모듈 변수 이름이 겹침: ' + mm.group(0))
                gnames.add(n)
    for p in procs:
        clash = local_names(p) & set(cmap)
        if clash:
            raise BuildError('%s.%s: 지역 이름이 Const 와 같음: %s' % (p.module, p.name, ', '.join(sorted(clash))))

    # 문장 바꾸기
    for p in procs:
        out = [replace_idents(p.lines[0], cmap)]
        for ln in p.lines[1:-1]:
            code = replace_idents(ln, cmap)
            out += expand_statement(code, indent_of(code))
        out.append(p.lines[-1])
        p.lines = out
    decls = [replace_idents(d, cmap) for d in decls]

    # 부르는 관계
    byname = {p.name.lower(): p for p in procs}
    for p in procs:
        for ln in p.lines:
            mk = mask(ln)
            for m in IDENT.finditer(mk):
                k = m.group(0).lower()
                if k in byname and k != p.name.lower():
                    a = m.start()
                    if a > 0 and mk[a - 1] == '.':
                        continue
                    p.calls.add(k)
    nodes = [p.name.lower() for p in procs]
    edges = {p.name.lower(): p.calls for p in procs}
    sccs = tarjan(nodes, edges)

    # 작은 모듈로 나누기
    head_size = sum(len(d) + 2 for d in decls) + 400
    mods = [[]]
    size = head_size
    for comp in sccs:
        comp = sorted(comp, key=nodes.index)
        csize = sum(byname[n].size() for n in comp)
        if mods[-1] and size + csize > LIMIT:
            mods.append([])
            size = 400
        mods[-1] += [byname[n] for n in comp]
        size += csize
    where = {}
    for i, m in enumerate(mods):
        for p in m:
            where[p.name.lower()] = i
    # 다른 모듈에서 부르는 Private -> Public
    for p in procs:
        for k in p.calls:
            q = byname[k]
            if where[q.name.lower()] > where[p.name.lower()]:
                raise BuildError('순서 오류: %s 가 뒤 모듈의 %s 를 부름' % (p.name, q.name))
            if where[q.name.lower()] != where[p.name.lower()] and q.vis == 'Private':
                q.vis = 'Public'
                q.lines[0] = re.sub(r'^Private\s+', 'Public ', q.lines[0], flags=re.I)

    n = len(mods)
    texts = []
    errs = []
    for i, m in enumerate(mods, 1):
        out = ["' %s - 모듈 %d / %d" % (title, i, n),
               "' 한셀 매크로 편집기에 이 파일 전체를 붙여넣고, 맨 아래 Check%d 를 실행해 확인하세요." % i,
               "' (원본: hancell/src, 만드는 도구: hancell/tools/build.py)",
               'Option Explicit', '']
        if i == 1:
            out += decls + ['']
        for p in m:
            out += p.lines + ['']
        nxt = ('다음: %d번 파일을 새 모듈에 붙여넣으세요.' % (i + 1)) if i < n else '모든 모듈을 넣었습니다. 이제 Setup 을 실행하세요.'
        # CheckN: 이 모듈을 읽었는지 + 앞 모듈들의 함수를 부를 수 있는지
        out += ['Public Function ModOk%d() As Long' % i,
                '    ModOk%d = %d' % (i, i),
                'End Function', '',
                'Public Sub Check%d()' % i,
                '    Dim n As Long, msg As String',
                '    msg = "%d번 모듈을 읽었습니다 (전체 %d개)."' % (i, n),
                '    On Error Resume Next',
                '    n = ' + ' + '.join('ModOk%d()' % k for k in range(1, i + 1)),
                '    If Err.Number <> 0 Or n <> %d Then' % (i * (i + 1) // 2),
                '        msg = msg & vbLf & "그런데 앞 모듈을 부르지 못했습니다: " & Err.Description',
                '    Else',
                '        msg = msg & vbLf & "%s"' % nxt,
                '    End If',
                '    On Error GoTo 0',
                '    MsgBox msg, 0, "확인"',
                'End Sub', '',
                "' ==== 모듈 %d 끝 ====" % i]
        errs += check_forbidden(out, '모듈 %d' % i)
        text = '\r\n'.join(out) + '\r\n'
        if len(text) > LIMIT + 3000:
            errs.append('모듈 %d 이 너무 큼 (%d자)' % (i, len(text)))
        texts.append(text)
    if errs:
        raise BuildError('\n'.join(errs))
    return texts
