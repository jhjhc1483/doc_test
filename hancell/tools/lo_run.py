"""LibreOffice(headless)로 VBA 매크로를 시험해 보는 도구.

한셀이 없는 환경에서 매크로가 돌아가는지, A4 결과가 어떻게 나오는지 미리 보려고 씁니다.
(LibreOffice의 VBA 호환 기능을 씁니다. 한셀과 100% 같지는 않습니다.)

usage:
  python3 lo_run.py --in book.xlsx --bas HcCore.bas --bas HcExec.bas --run BuildRelease \
                    [--save out.xlsm] [--pdf out.pdf --pdfsheet A4]
"""
import argparse
import os
import subprocess
import sys
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue


def pv(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def sync_asian_fonts(doc):
    """LibreOffice keeps a separate font for Korean (Asian) text; VBA Font.* only sets the Western one.
    Hancell has one font per cell, so copy Western -> Asian everywhere before saving or printing."""
    from com.sun.star.sheet.CellFlags import VALUE, STRING, FORMULA, DATETIME
    flags = VALUE | STRING | FORMULA | DATETIME
    for i in range(doc.Sheets.Count):
        sh = doc.Sheets.getByIndex(i)
        cells = sh.queryContentCells(flags)
        en = cells.Cells.createEnumeration()
        while en.hasMoreElements():
            c = en.nextElement()
            for src, dst in (('CharFontName', 'CharFontNameAsian'), ('CharHeight', 'CharHeightAsian'),
                             ('CharWeight', 'CharWeightAsian'), ('CharPosture', 'CharPostureAsian')):
                try:
                    c.setPropertyValue(dst, c.getPropertyValue(src))
                except Exception:  # noqa: BLE001
                    pass
            try:
                pe = c.getText().createEnumeration()
                while pe.hasMoreElements():
                    para = pe.nextElement()
                    te = para.createEnumeration()
                    while te.hasMoreElements():
                        por = te.nextElement()
                        for src, dst in (('CharHeight', 'CharHeightAsian'), ('CharWeight', 'CharWeightAsian'),
                                         ('CharFontName', 'CharFontNameAsian')):
                            try:
                                por.setPropertyValue(dst, por.getPropertyValue(src))
                            except Exception:  # noqa: BLE001
                                pass
            except Exception:  # noqa: BLE001
                pass


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--in', dest='inp', required=True)
    ap.add_argument('--bas', action='append', default=[])
    ap.add_argument('--run', action='append', default=[])
    ap.add_argument('--pdf')
    ap.add_argument('--pdfsheet')
    ap.add_argument('--save')
    a = ap.parse_args()

    pipe = 'hcpipe%d' % os.getpid()
    prof = os.path.join(tempfile.gettempdir(), 'hc_lo_profile')
    proc = subprocess.Popen(['soffice', '--headless', '--invisible', '--nologo', '--norestore',
                             '-env:UserInstallation=file://' + prof,
                             '--accept=pipe,name=%s;urp;' % pipe],
                            stdout=subprocess.DEVNULL,
                            stderr=open(os.environ['HC_LO_LOG'], 'w') if os.environ.get('HC_LO_LOG') else subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    ctx = None
    for _ in range(120):
        try:
            ctx = resolver.resolve('uno:pipe,name=%s;urp;StarOffice.ComponentContext' % pipe)
            break
        except Exception:
            time.sleep(0.5)
    if ctx is None:
        print('could not connect to soffice')
        proc.kill()
        return 1
    desktop = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
    url = uno.systemPathToFileUrl(os.path.abspath(a.inp))
    doc = desktop.loadComponentFromURL(url, '_blank', 0, (pv('Hidden', False), pv('MacroExecutionMode', 4)))
    rc = 0
    try:
        libs = doc.BasicLibraries
        mods = []
        if a.bas:
            libs.VBACompatibilityMode = True
            libname = 'VBAProject'
            try:
                libs.ProjectName = libname
            except Exception as e:  # noqa: BLE001
                print('project name:', e)
            if not libs.hasByName(libname):
                libs.createLibrary(libname)
            lib = libs.getByName(libname)
            from com.sun.star.script import ModuleInfo
            from com.sun.star.script.ModuleType import NORMAL
            for f in a.bas:
                src = open(f, encoding='utf-8').read().replace('\r\n', '\n')
                lines = [ln for ln in src.split('\n') if not ln.startswith('Attribute VB_')]
                src = 'Option VBASupport 1\n' + '\n'.join(lines)
                name = os.path.splitext(os.path.basename(f))[0]
                mods.append(name)
                if lib.hasByName(name):
                    lib.replaceByName(name, src)
                else:
                    lib.insertByName(name, src)
                mi = ModuleInfo()
                mi.ModuleType = NORMAL
                if not lib.hasModuleInfo(name):
                    lib.insertModuleInfo(name, mi)
        if a.bas:
            shim = (
                'Option VBASupport 1\n'
                'Public Sub Activate()\n    gHostLO = True\nEnd Sub\n'
                'Public Sub LoShimFont(ByVal sh As String, ByVal addr As String, ByVal nm As String, ByVal size As Double, ByVal bold As Boolean)\n'
                '    Dim o As Object\n'
                '    Set o = ThisComponent.Sheets.getByName(sh).getCellRangeByName(addr)\n'
                '    o.CharHeightAsian = size\n'
                '    o.CharFontNameAsian = nm\n'
                '    If bold Then\n        o.CharWeightAsian = 150\n    Else\n        o.CharWeightAsian = 100\n    End If\n'
                'End Sub\n')
            lib.insertByName('LoShim', shim)
            mi = ModuleInfo()
            mi.ModuleType = NORMAL
            lib.insertModuleInfo('LoShim', mi)
            try:
                doc.getScriptProvider().getScript(
                    'vnd.sun.star.script:VBAProject.LoShim.Activate?language=Basic&location=document').invoke((), (), ())
            except Exception as e:  # noqa: BLE001
                print('shim activate:', e)
        for m in a.run:
            target = m if '.' in m else mods[-1] + '.' + m
            script_url = 'vnd.sun.star.script:VBAProject.%s?language=Basic&location=document' % target
            t0 = time.time()
            try:
                res = doc.getScriptProvider().getScript(script_url).invoke((), (), ())
                print('run %s: ok %.1fs %s' % (target, time.time() - t0, res[0] if res and res[0] else ''))
            except Exception as e:  # noqa: BLE001
                print('run %s: ERROR %s' % (target, e))
                rc = 2
        sync_asian_fonts(doc)
        if a.bas:
            lib = doc.BasicLibraries.getByName('VBAProject')
            if lib.hasByName('LoShim'):
                lib.removeByName('LoShim')
        if a.save:
            out = uno.systemPathToFileUrl(os.path.abspath(a.save))
            flt = 'Calc MS Excel 2007 VBA XML' if a.save.endswith('.xlsm') else 'Calc MS Excel 2007 XML'
            doc.storeToURL(out, (pv('FilterName', flt),))
            print('saved', a.save)
        if a.pdf:
            out = uno.systemPathToFileUrl(os.path.abspath(a.pdf))
            fd = ()
            if a.pdfsheet:
                sh = doc.Sheets.getByName(a.pdfsheet)
                fd = (pv('Selection', sh),)
            doc.storeToURL(out, (pv('FilterName', 'calc_pdf_Export'),
                                 pv('FilterData', uno.Any('[]com.sun.star.beans.PropertyValue', fd))))
            print('pdf', a.pdf)
    finally:
        try:
            doc.close(True)
        except Exception:  # noqa: BLE001
            pass
        try:
            desktop.terminate()
        except Exception:  # noqa: BLE001
            pass
        time.sleep(0.5)
        proc.kill()
    return rc


if __name__ == '__main__':
    sys.exit(main())
