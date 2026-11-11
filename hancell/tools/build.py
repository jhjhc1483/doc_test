"""한셀에 붙여넣을 매크로 모듈을 만들고 LibreOffice 에서 그대로 시험합니다.

src/*.bas 는 읽기 쉬운 원본입니다. 한셀 2020 은 VBA 일부(Collection 등)에서 꺼지고
큰 모듈은 읽지 못해서, tools/hancellify.py 가 확인된 문법만 쓰는 작은 모듈 여러 개로 바꿉니다.
결과: hancell/붙여넣기/<화면>/모듈1.txt, 모듈2.txt, ... (붙여넣는 순서)

usage: python3 build.py            # 임원 화면
       python3 build.py exec daily # 두 화면
"""
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)              # hancell/
SRC = os.path.join(ROOT, 'src')
OUT_PASTE = os.path.join(ROOT, '붙여넣기')

APPS = {
    'exec': {'modules': ['HcCore', 'HcExec', 'HcExecIO', 'HcExecSetup'], 'test': 'HcExecTest',
             'title': '임원 일일 예정사항', 'book': '임원일일예정.xlsm', 'paste': '임원'},
    'daily': {'modules': ['HcCore', 'HcDaily', 'HcDailyMd', 'HcDailyIO', 'HcDailySetup'], 'test': 'HcDailyTest',
              'title': '부서별 일일 예정사항', 'book': '부서별일일예정.xlsm', 'paste': '부서별'},
}

# 원본에만 있는 확인용 프로시저 (붙여넣기 모듈에는 CheckN 이 새로 붙음)
DROP = {'CheckHcCore', 'CheckHcExec', 'CheckHcExecIO', 'CheckHcExecSetup',
        'CheckHcDaily', 'CheckHcDailyMd', 'CheckHcDailyIO', 'CheckHcDailySetup'}


def blank_xlsx(path):
    import openpyxl
    wb = openpyxl.Workbook()
    wb.active.title = 'Sheet1'
    wb.save(path)


def build(app):
    cfg = APPS[app]
    sys.path.insert(0, HERE)
    import vbalint
    import vbacompile
    import hancellify
    src_files = [os.path.join(SRC, m + '.bas') for m in cfg['modules']]
    test_file = os.path.join(SRC, cfg['test'] + '.bas')
    errs = vbalint.lint(src_files)
    errs += vbacompile.check(src_files + [test_file])
    if errs:
        print('\n'.join(errs))
        raise SystemExit('VBA 검사에서 문제가 나왔습니다 (원본)')

    sources = [(m, open(os.path.join(SRC, m + '.bas'), encoding='utf-8').read()) for m in cfg['modules']]
    try:
        texts = hancellify.build_modules(sources, cfg['title'], drop=DROP)
    except hancellify.BuildError as e:
        raise SystemExit('한셀용으로 바꾸지 못했습니다:\n%s' % e)

    # 붙여넣기 파일
    d = os.path.join(OUT_PASTE, cfg['paste'])
    if os.path.isdir(d):
        shutil.rmtree(d)
    os.makedirs(d)
    for k, t in enumerate(texts, 1):
        with open(os.path.join(d, '모듈%d.txt' % k), 'w', encoding='utf-8', newline='') as fh:
            fh.write(t)
        print('%s/모듈%d.txt: %d자, %d줄' % (os.path.relpath(d, ROOT), k, len(t), t.count('\n')))

    # 붙여넣기 파일 그대로 다시 검사하고 LibreOffice 에서 실행
    tmpd = tempfile.mkdtemp(prefix='hc_paste_')
    bas = []
    for k, t in enumerate(texts, 1):
        p = os.path.join(tmpd, 'Mod%d.bas' % k)
        with open(p, 'w', encoding='utf-8', newline='') as fh:
            fh.write(t)
        bas.append(p)
    errs = vbalint.lint(bas)
    errs += vbacompile.check(bas + [test_file])
    if errs:
        print('\n'.join(errs))
        raise SystemExit('VBA 검사에서 문제가 나왔습니다 (붙여넣기 파일)')
    tmp = os.path.join(tmpd, 'hc_blank.xlsx')
    blank_xlsx(tmp)
    book = os.path.join(tempfile.gettempdir(), cfg['book'])
    cmd = [sys.executable, os.path.join(HERE, 'lo_run.py'), '--in', tmp]
    for b in bas + [test_file]:
        cmd += ['--bas', b]
    cmd += ['--run', cfg['test'] + '.BuildRelease']
    for extra in os.environ.get('HC_RUN', '').split():
        cmd += ['--run', cfg['test'] + '.' + extra]
    cmd += ['--save', book]
    if os.environ.get('HC_PDF'):
        cmd += ['--pdf', os.environ['HC_PDF'], '--pdfsheet', 'A4']
    print(' '.join(cmd))
    r = subprocess.run(cmd)
    if r.returncode:
        raise SystemExit('LibreOffice 실행 실패')


if __name__ == '__main__':
    apps = sys.argv[1:] or ['exec']
    for a in apps:
        build(a)
