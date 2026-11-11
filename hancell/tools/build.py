"""한셀용 매크로 파일(.bas)과 매크로가 들어간 통합 문서(.xlsm)를 만듭니다.

src/*.bas 는 읽기 쉽게 한글을 그대로 쓴 원본입니다.
VBA 편집기는 한글(유니코드)을 파일에 제대로 담지 못할 때가 있어서,
문자열 속 한글은 U("\\uXXXX") 꼴로 바꿔 ASCII만 남긴 .bas 를 만듭니다.

usage: python3 build.py            # 두 화면 모두
       python3 build.py exec       # 임원 화면만
"""
import os
import re
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)              # hancell/
SRC = os.path.join(ROOT, 'src')
OUT_BAS = os.path.join(ROOT, '매크로')

OUT_PASTE = os.path.join(ROOT, '붙여넣기')

APPS = {
    'exec': {'modules': ['HcCore', 'HcExec', 'HcExecIO', 'HcExecSetup'], 'book': '임원일일예정.xlsm', 'paste': '임원'},
    'daily': {'modules': ['HcCore', 'HcDaily', 'HcDailyMd', 'HcDailyIO', 'HcDailySetup'], 'book': '부서별일일예정.xlsm',
              'paste': '부서별'},
}


def enc_str(body):
    """VBA 문자열 내용(따옴표 안, "" 는 그대로) → \\uXXXX 꼴"""
    out = []
    for ch in body:
        o = ord(ch)
        if ch == '\\':
            out.append('\\u005C')
        elif o < 128:
            out.append(ch)
        elif o > 0xFFFF:
            o -= 0x10000
            out.append('\\u%04X\\u%04X' % (0xD800 + (o >> 10), 0xDC00 + (o & 0x3FF)))
        else:
            out.append('\\u%04X' % o)
    return ''.join(out)


def convert_line(line, fname, no):
    res = []
    i = 0
    n = len(line)
    while i < n:
        ch = line[i]
        if ch == '"':
            j = i + 1
            while j < n:
                if line[j] == '"':
                    if j + 1 < n and line[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            if j >= n:
                raise SystemExit('%s:%d: 따옴표가 닫히지 않음' % (fname, no))
            body = line[i + 1:j]
            if any(ord(c) >= 128 for c in body):
                res.append('U("' + enc_str(body) + '")')
            else:
                res.append(line[i:j + 1])
            i = j + 1
            continue
        if ch == "'":
            comment = line[i:]
            if any(ord(c) >= 128 for c in comment):
                comment = "'"          # 한글 주석은 뺀다 (원본 src 에 남아 있음)
            res.append(comment)
            break
        res.append(ch)
        i += 1
    out = ''.join(res).rstrip()
    if re.match(r'^\s*(Public\s+|Private\s+)?Const\b', out) and 'U("' in out:
        raise SystemExit('%s:%d: Const 에는 한글 문자열을 쓸 수 없습니다' % (fname, no))
    if any(ord(c) >= 128 for c in out):
        raise SystemExit('%s:%d: 문자열 밖에 한글이 있습니다: %s' % (fname, no, out))
    if len(out) > 1000:
        raise SystemExit('%s:%d: 줄이 너무 깁니다(%d자). 문자열을 나누세요.' % (fname, no, len(out)))
    return out


def convert_file(name):
    path = os.path.join(SRC, name + '.bas')
    lines = open(path, encoding='utf-8').read().replace('\r\n', '\n').split('\n')
    out = [convert_line(ln, name, k + 1) for k, ln in enumerate(lines)]
    while out and not out[-1]:
        out.pop()
    text = '\r\n'.join(out) + '\r\n'
    os.makedirs(OUT_BAS, exist_ok=True)
    dst = os.path.join(OUT_BAS, name + '.bas')
    with open(dst, 'w', encoding='ascii', newline='') as f:
        f.write(text)
    return dst, len(text)


def write_paste(app, cfg):
    """한셀 매크로 편집기에 붙여넣을 파일: Attribute 줄을 뺀 모듈, 순서대로 번호"""
    d = os.path.join(OUT_PASTE, cfg['paste'])
    os.makedirs(d, exist_ok=True)
    for f in os.listdir(d):
        if f.endswith('.txt'):
            os.remove(os.path.join(d, f))
    for k, m in enumerate(cfg['modules'], 1):
        src = open(os.path.join(OUT_BAS, m + '.bas'), encoding='ascii', newline='').read()
        lines = [ln for ln in src.replace('\r\n', '\n').split('\n') if not ln.startswith('Attribute VB_')]
        text = '\r\n'.join(lines)
        # 빈 파일이 나가지 않도록 확인
        if len(text) < len(src) * 0.95 or 'End Sub' not in text and 'End Function' not in text:
            raise SystemExit('붙여넣기 파일이 비었거나 너무 짧습니다: ' + m)
        with open(os.path.join(d, '%d_%s.txt' % (k, m)), 'w', encoding='ascii', newline='') as fh:
            fh.write(text)
    print('paste files:', os.path.relpath(d, ROOT))


def blank_xlsx(path):
    import openpyxl
    wb = openpyxl.Workbook()
    wb.active.title = 'Sheet1'
    wb.save(path)


def build(app):
    cfg = APPS[app]
    sys.path.insert(0, HERE)
    import vbalint
    errs = vbalint.lint([os.path.join(SRC, m + '.bas') for m in cfg['modules']])
    if errs:
        print('\n'.join(errs))
        raise SystemExit('VBA 검사에서 문제가 나왔습니다')
    bas = []
    for m in cfg['modules']:
        dst, size = convert_file(m)
        print('%s: %d bytes' % (os.path.relpath(dst, ROOT), size))
        if size > 60000:
            print('  경고: 모듈이 60KB를 넘습니다. 나누는 것이 좋습니다.')
        bas.append(dst)
    write_paste(app, cfg)
    tmp = os.path.join(tempfile.gettempdir(), 'hc_blank.xlsx')
    blank_xlsx(tmp)
    # LibreOffice로 만든 .xlsm 은 한셀이 매크로를 읽지 못하므로 시험용으로만 만든다
    book = os.path.join(tempfile.gettempdir(), cfg['book'])
    cmd = [sys.executable, os.path.join(HERE, 'lo_run.py'), '--in', tmp]
    for b in bas:
        cmd += ['--bas', b]
    cmd += ['--run', cfg['modules'][-1] + '.BuildRelease', '--save', book]
    if os.environ.get('HC_PDF'):
        cmd += ['--pdf', os.environ['HC_PDF'], '--pdfsheet', 'A4']
    print(' '.join(cmd))
    r = subprocess.run(cmd)
    if r.returncode:
        raise SystemExit('LibreOffice 실행 실패')


if __name__ == '__main__':
    apps = sys.argv[1:] or ['exec', 'daily']
    for a in apps:
        build(a)
