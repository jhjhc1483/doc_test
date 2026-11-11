"""VBA 원본(src/*.bas)의 흔한 실수를 미리 찾습니다.

- 이름은 대소문자를 가리지 않으므로 rbText 와 RBText 는 같은 이름 (겹치면 컴파일 오류)
- 예약어(Any, Line 등)를 변수 이름으로 쓴 경우
- 모듈 사이에 Public 이름이 겹치는 경우
- If/For 안에서 Dim 한 변수 (LibreOffice 는 Dim 을 실행할 때 만들어서 블록 밖에서 쓰면 오류)
"""
import re
import sys

RESERVED = {'any', 'line', 'get', 'let', 'set', 'type', 'error', 'me', 'global', 'step', 'empty', 'null',
            'date', 'string', 'val', 'name', 'input', 'print', 'open', 'close', 'width', 'lock', 'unlock',
            'mid', 'left', 'right', 'len', 'resume', 'stop', 'time', 'erase', 'seek', 'object',
            'wend', 'text', 'local', 'random', 'read', 'write', 'output', 'binary', 'append', 'access',
            'compare', 'base', 'explicit', 'shared', 'lib', 'alias', 'property', 'until', 'while', 'xor',
            'eqv', 'imp', 'like', 'is', 'mod', 'option', 'rem', 'return', 'select', 'static', 'to', 'with'}

DECL = re.compile(r'^\s*(Dim|Private|Public|Const|Static|ReDim)\b(.*)$', re.I)
PROC = re.compile(r'^\s*(Public |Private )?(Sub|Function) (\w+)\s*\((.*?)\)', re.I)


def strip_strings(line):
    return re.sub(r'"(?:[^"]|"")*"', '""', line).split("'")[0]


def names_in_decl(rest):
    rest = re.sub(r'\(.*?\)', '', rest)
    out = []
    for part in rest.split(','):
        part = part.strip()
        part = re.sub(r'^(WithEvents|Const|Preserve)\s+', '', part, flags=re.I)
        m = re.match(r'(\w+)', part)
        if m:
            out.append(m.group(1))
    return out


def params(sig):
    out = []
    for part in sig.split(','):
        part = re.sub(r'\b(Optional|ByVal|ByRef|ParamArray)\b', '', part, flags=re.I).strip()
        m = re.match(r'(\w+)', part)
        if m:
            out.append(m.group(1))
    return out


def lint(files):
    errors = []
    public = {}
    for fn in files:
        lines = open(fn, encoding='utf-8').read().replace('\r\n', '\n').split('\n')
        # join continuation lines
        joined, buf, starts = [], '', []
        for i, l in enumerate(lines):
            if not buf:
                starts.append(i + 1)
            if l.rstrip().endswith(' _'):
                buf += l.rstrip()[:-1] + ' '
                continue
            joined.append(buf + l)
            buf = ''
        module = {}
        proc = None
        local = {}
        depth = 0
        for ln, raw in zip(starts, joined):
            l = strip_strings(raw)
            pm = PROC.match(l)
            if pm:
                proc = pm.group(3)
                local = {}
                depth = 0
                key = proc.lower()
                if key in module and module[key] != proc:
                    errors.append('%s:%d: "%s" 가 "%s" 와 겹칩니다(대소문자 무시)' % (fn, ln, proc, module[key]))
                module[key] = proc
                if (pm.group(1) or 'Public ').strip().lower() == 'public':
                    if key in public and public[key][0] != fn:
                        errors.append('%s:%d: Public "%s" 가 %s 에도 있습니다' % (fn, ln, proc, public[key][0]))
                    public[key] = (fn, proc)
                for p in params(pm.group(4)):
                    if p.lower() in local:
                        errors.append('%s:%d: 인자 "%s" 가 겹칩니다' % (fn, ln, p))
                    if p.lower() in RESERVED:
                        errors.append('%s:%d: 예약어 "%s" 를 인자 이름으로 씀' % (fn, ln, p))
                    local[p.lower()] = p
                continue
            if re.match(r'^\s*End (Sub|Function)\b', l, re.I):
                proc = None
                continue
            if re.match(r'^\s*If .+ Then\s+\S.*\sElse\s', l, re.I):
                errors.append('%s:%d: 한 줄 If ... Then ... Else 는 LibreOffice 가 못 읽습니다 — 여러 줄로 쓰세요' % (fn, ln))
            stmts = [x.strip() for x in l.split(':')]
            for st in stmts:
                if re.match(r'^(If .* Then|For |Do\b|Select Case|With )', st, re.I) and not re.search(r'\bThen\s+\S', st, re.I):
                    depth += 1
                if re.match(r'^(End If|Next\b|Loop\b|End Select|End With)', st, re.I):
                    depth = max(0, depth - 1)
                dm = DECL.match(st)
                if not dm or re.match(r'^\s*(Private|Public)\s+(Sub|Function)', st, re.I):
                    continue
                kw = dm.group(1).lower()
                if kw == 'redim':
                    continue
                rest = dm.group(2)
                rest = re.sub(r'^\s*Const\b', '', rest, flags=re.I)
                for nm in names_in_decl(rest):
                    k = nm.lower()
                    if k in RESERVED:
                        errors.append('%s:%d: 예약어 "%s" 를 이름으로 씀' % (fn, ln, nm))
                    if proc:
                        if k in local:
                            errors.append('%s:%d: "%s" 가 "%s" 와 겹칩니다(대소문자 무시)' % (fn, ln, nm, local[k]))
                        if k == proc.lower():
                            errors.append('%s:%d: 변수 "%s" 가 프로시저 이름과 같습니다' % (fn, ln, nm))
                        local[k] = nm
                        if kw == 'dim' and depth > 0:
                            errors.append('%s:%d: 블록 안의 Dim "%s" — 프로시저 맨 위로 옮기세요' % (fn, ln, nm))
                    else:
                        if k in module and module[k] != nm:
                            errors.append('%s:%d: "%s" 가 "%s" 와 겹칩니다(대소문자 무시)' % (fn, ln, nm, module[k]))
                        elif k in module:
                            errors.append('%s:%d: "%s" 가 두 번 선언됨' % (fn, ln, nm))
                        module[k] = nm
    return errors


if __name__ == '__main__':
    errs = lint(sys.argv[1:])
    for e in errs:
        print(e)
    sys.exit(1 if errs else 0)
