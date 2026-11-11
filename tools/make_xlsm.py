"""엑셀용·한셀용 매크로 통합 문서(.xlsm) 만들기 — 외부 라이브러리 없이 파이썬만으로.

    python3 tools/make_xlsm.py

  excel/임원일일예정_대시보드.xlsm   ← excel/임원일일예정_대시보드.bas (모듈 ExecDashboard)
  hancell/임원일일예정_한셀.xlsm     ← hancell/붙여넣기/임원/모듈1.txt ~ 모듈7.txt (Module1 ~ Module7)

VBA 코드를 한국어 Windows 코드 페이지(CP949)로 넣습니다. (LibreOffice로 .xlsm을 저장하면
매크로 속 한글이 깨지기 때문에 직접 만듭니다.) 코드는 원본 글만 넣고 컴파일된 부분은 넣지 않아,
처음 열 때 프로그램이 원본에서 다시 컴파일합니다.

형식 근거: [MS-OVBA](VBA 프로젝트·압축), [MS-CFB](복합 파일), ECMA-376(.xlsm 패키지).
"""
import os
import struct
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CODEPAGE = 949
ENC = 'cp949'

# CP949에 없는 글자 (한셀 붙여넣기 모듈의 문자열 몇 곳) → .xlsm 안에서만 바꿈
CP949_FIX = [('"-–○●◦·•▶▷□■"', '"-○●·▶▷□■"'), ('"*＊-·•"', '"*＊-·"'), ('"-–·•"', '"-·"'),
             ('—', '-'), ('›', '>')]


# ---------------------------------------------------------------------------
# [MS-OVBA] 2.4.1 압축
# ---------------------------------------------------------------------------
def _compress_chunk(chunk):
    out = bytearray()
    n, i, table = len(chunk), 0, {}

    def remember(p):
        if p + 3 <= n:
            table.setdefault(bytes(chunk[p:p + 3]), []).append(p)

    while i < n:
        flag_at = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if i >= n:
                break
            best_len = best_off = 0
            if i > 0 and i + 3 <= n:
                bit_count = max((i - 1).bit_length(), 4)
                max_len = (0xFFFF >> bit_count) + 3
                for c in reversed(table.get(bytes(chunk[i:i + 3]), [])[-48:]):
                    ln = 3
                    while i + ln < n and ln < max_len and chunk[c + ln] == chunk[i + ln]:
                        ln += 1
                    if ln > best_len:
                        best_len, best_off = ln, i - c
                        if ln == max_len:
                            break
            if best_len >= 3:
                bit_count = max((i - 1).bit_length(), 4)
                out += struct.pack('<H', ((best_off - 1) << (16 - bit_count)) | (best_len - 3))
                flags |= 1 << bit
                for k in range(best_len):
                    remember(i + k)
                i += best_len
            else:
                out.append(chunk[i])
                remember(i)
                i += 1
        out[flag_at] = flags
    return bytes(out)


def ovba_compress(data):
    out = bytearray([1])
    for pos in range(0, len(data), 4096):
        chunk = data[pos:pos + 4096]
        comp = _compress_chunk(chunk)
        if len(chunk) == 4096 and len(comp) > 4096:          # 압축이 손해면 그대로
            out += struct.pack('<H', 0x3000 | 4095) + chunk
        else:
            assert len(comp) + 2 <= 4098, 'chunk too large'
            out += struct.pack('<H', 0xB000 | (len(comp) + 2 - 3)) + comp
    return bytes(out)


# ---------------------------------------------------------------------------
# [MS-OVBA] 2.4.3 PROJECT 스트림의 CMG·DPB·GC 값 (보호 없음·암호 없음·보이기)
# ---------------------------------------------------------------------------
def ovba_encrypt(data, project_id, seed):
    version, proj_key = 2, sum(project_id.encode(ENC)) & 0xFF
    out = bytearray([seed, seed ^ version, seed ^ proj_key])
    unenc1, enc1, enc2 = proj_key, seed ^ proj_key, seed ^ version
    payload = [0] * ((seed & 6) // 2) + list(struct.pack('<I', len(data))) + list(data)
    for b in payload:
        e = b ^ ((enc2 + unenc1) & 0xFF)
        out.append(e)
        enc2, enc1, unenc1 = enc1, e, b
    return out.hex().upper()


# ---------------------------------------------------------------------------
# VBA 프로젝트 (vbaProject.bin 의 스트림들)
# ---------------------------------------------------------------------------
DOC_BASE = {'ThisWorkbook': '0{00020819-0000-0000-C000-000000000046}',
            'Sheet1': '0{00020820-0000-0000-C000-000000000046}'}
PROJECT_ID = '{5B3A1C2E-8D4F-4E61-9A7B-1C2D3E4F5A6B}'


def _rec(rid, payload):
    return struct.pack('<HI', rid, len(payload)) + payload


def vba_streams(modules):
    """modules: [(이름, 코드)] 보통 모듈. ThisWorkbook·Sheet1 문서 모듈은 함께 넣음."""
    docs = [(nm, 'Attribute VB_Name = "%s"\r\nAttribute VB_Base = "%s"\r\nAttribute VB_GlobalNameSpace = False\r\n'
                 'Attribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\n'
                 'Attribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n' % (nm, base))
            for nm, base in DOC_BASE.items()]
    std = []
    for nm, code in modules:
        lines = [ln for ln in code.replace('\r\n', '\n').split('\n') if not ln.startswith('Attribute VB_')]
        std.append((nm, 'Attribute VB_Name = "%s"\r\n' % nm + '\r\n'.join(lines).rstrip('\r\n') + '\r\n'))
    allmods = [(n, c, 0x22) for n, c in docs] + [(n, c, 0x21) for n, c in std]

    # dir 스트림
    e = lambda s: s.encode(ENC)
    u = lambda s: s.encode('utf-16-le')
    d = bytearray()
    d += _rec(0x01, struct.pack('<I', 1))                 # SYSKIND: Win32
    d += _rec(0x02, struct.pack('<I', 0x409))             # LCID
    d += _rec(0x14, struct.pack('<I', 0x409))             # LCIDINVOKE
    d += _rec(0x03, struct.pack('<H', CODEPAGE))          # CODEPAGE
    d += _rec(0x04, e('VBAProject'))                      # NAME
    d += _rec(0x05, b'') + _rec(0x40, b'')                # DOCSTRING (+유니코드)
    d += _rec(0x06, b'') + _rec(0x3D, b'')                # HELPFILEPATH 1·2
    d += _rec(0x07, struct.pack('<I', 0))                 # HELPCONTEXT
    d += _rec(0x08, struct.pack('<I', 0))                 # LIBFLAGS
    d += struct.pack('<HIIH', 0x09, 4, 1, 0)              # VERSION
    d += _rec(0x0C, b'') + _rec(0x3C, b'')                # CONSTANTS (+유니코드)
    libid = '*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation'
    d += _rec(0x16, e('stdole')) + _rec(0x3E, u('stdole'))                    # REFERENCENAME
    d += _rec(0x0D, struct.pack('<I', len(e(libid))) + e(libid) + struct.pack('<IH', 0, 0))   # REFERENCEREGISTERED
    d += struct.pack('<HIH', 0x0F, 2, len(allmods))       # MODULES count
    d += struct.pack('<HIH', 0x13, 2, 0xFFFF)             # PROJECTCOOKIE
    for nm, _code, kind in allmods:
        d += _rec(0x19, e(nm)) + _rec(0x47, u(nm))        # MODULENAME, MODULENAMEUNICODE
        d += _rec(0x1A, e(nm)) + _rec(0x32, u(nm))        # MODULESTREAMNAME
        d += _rec(0x1C, b'') + _rec(0x48, b'')            # MODULEDOCSTRING
        d += _rec(0x31, struct.pack('<I', 0))             # MODULEOFFSET: 코드가 맨 앞 (컴파일된 부분 없음)
        d += _rec(0x1E, struct.pack('<I', 0))             # MODULEHELPCONTEXT
        d += struct.pack('<HIH', 0x2C, 2, 0xFFFF)         # MODULECOOKIE
        d += struct.pack('<HI', kind, 0)                  # MODULETYPE (0x21 보통, 0x22 문서)
        d += struct.pack('<HI', 0x2B, 0)                  # 모듈 끝
    d += struct.pack('<HI', 0x10, 0)                      # dir 끝

    proj = ['ID="%s"' % PROJECT_ID]
    proj += ['Document=%s/&H00000000' % n for n, _c, k in allmods if k == 0x22]
    proj += ['Module=%s' % n for n, _c, k in allmods if k == 0x21]
    proj += ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
             'CMG="%s"' % ovba_encrypt(b'\x00\x00\x00\x00', PROJECT_ID, 0x37),
             'DPB="%s"' % ovba_encrypt(b'\x00', PROJECT_ID, 0x52),
             'GC="%s"' % ovba_encrypt(b'\xff', PROJECT_ID, 0x6A),
             '', '[Host Extender Info]', '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000', '',
             '[Workspace]'] + ['%s=0, 0, 0, 0, C' % n for n, _c, _k in allmods] + ['']
    wm = b''.join(e(n) + b'\x00' + u(n) + b'\x00\x00' for n, _c, _k in allmods) + b'\x00\x00'

    streams = {'PROJECT': '\r\n'.join(proj).encode(ENC), 'PROJECTwm': wm,
               'VBA/_VBA_PROJECT': bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00]),
               'VBA/dir': ovba_compress(bytes(d))}
    for nm, code, _k in allmods:
        streams['VBA/' + nm] = ovba_compress(code.encode(ENC))
    return streams


# ---------------------------------------------------------------------------
# [MS-CFB] 복합 파일 (버전 3, 512바이트 섹터, 4096바이트 미만은 미니 스트림)
# ---------------------------------------------------------------------------
FREE, ENDCHAIN, FATSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF


def _cfb_key(name):
    return (len(name), name.upper())


def cfb_write(streams):
    # 디렉터리 나무 만들기
    nodes = [{'name': 'Root Entry', 'type': 5, 'kids': {}, 'data': b''}]

    def node_for(path):
        cur = 0
        for part in path:
            kids = nodes[cur]['kids']
            if part not in kids:
                nodes.append({'name': part, 'type': 1, 'kids': {}, 'data': b''})
                kids[part] = len(nodes) - 1
            cur = kids[part]
        return cur

    for path, data in streams.items():
        parts = path.split('/')
        parent = node_for(parts[:-1])
        nodes.append({'name': parts[-1], 'type': 2, 'kids': {}, 'data': data})
        nodes[parent]['kids'][parts[-1]] = len(nodes) - 1

    # 형제들은 균형 잡힌 이진 나무 (가장 깊은 층만 빨강 → 올바른 레드블랙 나무)
    for nd in nodes:
        nd.update(left=NOSTREAM, right=NOSTREAM, child=NOSTREAM, color=1)

    def build(ids, depth, out_depths):
        if not ids:
            return NOSTREAM
        m = len(ids) // 2
        me = ids[m]
        nodes[me]['depth'] = depth
        out_depths.append(me)
        nodes[me]['left'] = build(ids[:m], depth + 1, out_depths)
        nodes[me]['right'] = build(ids[m + 1:], depth + 1, out_depths)
        return me

    for nd in nodes:
        if nd['kids']:
            ids = sorted(nd['kids'].values(), key=lambda k: _cfb_key(nodes[k]['name']))
            seen = []
            nd['child'] = build(ids, 0, seen)
            full = (len(ids) + 1) & len(ids) == 0              # 꽉 찬 나무면 모두 검정
            maxd = max(nodes[k]['depth'] for k in seen)
            for k in seen:
                nodes[k]['color'] = 0 if (not full and nodes[k]['depth'] == maxd and maxd > 0) else 1

    # 미니 스트림 / 일반 스트림 나누기
    sectors = []                                          # 일반 섹터 내용(512바이트씩)
    fat = []

    def alloc(data, mark=None):
        if not data:
            return ENDCHAIN
        cnt = (len(data) + 511) // 512
        start = len(sectors)
        for k in range(cnt):
            sectors.append(data[k * 512:(k + 1) * 512].ljust(512, b'\x00'))
            fat.append(mark if mark is not None else (start + k + 1 if k < cnt - 1 else ENDCHAIN))
        return start

    mini = bytearray()
    minifat = []
    for nd in nodes:
        if nd['type'] != 2:
            continue
        data = nd['data']
        if len(data) < 4096:
            if data:
                cnt = (len(data) + 63) // 64
                start = len(minifat)
                mini += data.ljust(cnt * 64, b'\x00')
                minifat += [start + k + 1 if k < cnt - 1 else ENDCHAIN for k in range(cnt)]
                nd['start'] = start
            else:
                nd['start'] = ENDCHAIN
        else:
            nd['start'] = alloc(data)
        nd['size'] = len(data)
    nodes[0]['start'] = alloc(bytes(mini)) if mini else ENDCHAIN
    nodes[0]['size'] = len(mini)
    minifat_bytes = b''.join(struct.pack('<I', v) for v in minifat)
    if minifat_bytes:
        minifat_bytes += b'\xff\xff\xff\xff' * ((-len(minifat)) % 128)
    first_minifat = alloc(minifat_bytes) if minifat_bytes else ENDCHAIN
    n_minifat = len(minifat_bytes) // 512

    def dir_entry(nd):
        nm = nd['name'].encode('utf-16-le') + b'\x00\x00'
        return (nm.ljust(64, b'\x00') + struct.pack('<HBBIII', len(nm), nd['type'], nd['color'],
                                                     nd['left'], nd['right'], nd['child'])
                + b'\x00' * 16 + struct.pack('<I', 0) + b'\x00' * 16
                + struct.pack('<IQ', nd.get('start', 0) if nd['type'] != 1 else 0, nd.get('size', 0)))
    dirs = b''.join(dir_entry(nd) for nd in nodes)
    empty = b'\x00' * 64 + struct.pack('<HBBIII', 0, 0, 0, NOSTREAM, NOSTREAM, NOSTREAM) + b'\x00' * 48
    dirs += empty * ((-len(nodes)) % 4)
    first_dir = alloc(dirs)

    # FAT 섹터 (자기 자신도 FAT에 FATSECT 로)
    n_fat = 1
    while (len(fat) + n_fat + 127) // 128 > n_fat:
        n_fat += 1
    fat_start = len(sectors)
    fat += [FATSECT] * n_fat
    fat += [FREE] * (n_fat * 128 - len(fat))
    fat_bytes = b''.join(struct.pack('<I', v) for v in fat)
    for k in range(n_fat):
        sectors.append(fat_bytes[k * 512:(k + 1) * 512])
    assert n_fat <= 109

    difat = [fat_start + k for k in range(n_fat)] + [FREE] * (109 - n_fat)
    header = (b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1' + b'\x00' * 16
              + struct.pack('<HHHHH', 0x3E, 3, 0xFFFE, 9, 6) + b'\x00' * 6
              + struct.pack('<IIIIIIIII', 0, n_fat, first_dir, 0, 4096, first_minifat, n_minifat, ENDCHAIN, 0)
              + b''.join(struct.pack('<I', v) for v in difat))
    assert len(header) == 512
    return header + b''.join(sectors)


# ---------------------------------------------------------------------------
# .xlsm 패키지 (시트 하나: '시작' 안내)
# ---------------------------------------------------------------------------
def _x(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def xlsm_bytes(sheet_name, rows, vba_bin):
    NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    files = {
        '[Content_Types].xml': head + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>',
        '_rels/.rels': head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/></Relationships>' % R,
        'xl/workbook.xml': head + '<workbook xmlns="%s" xmlns:r="%s"><workbookPr codeName="ThisWorkbook"/>'
        '<bookViews><workbookView/></bookViews><sheets><sheet name="%s" sheetId="1" r:id="rId1"/></sheets></workbook>'
        % (NS, R, _x(sheet_name)),
        'xl/_rels/workbook.xml.rels': head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="%s/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="%s/styles" Target="styles.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>'
        '</Relationships>' % (R, R),
        'xl/styles.xml': head + '<styleSheet xmlns="%s">'
        '<fonts count="3"><font><sz val="11"/><name val="맑은 고딕"/><family val="3"/><charset val="129"/></font>'
        '<font><b/><sz val="14"/><name val="맑은 고딕"/><family val="3"/><charset val="129"/></font>'
        '<font><b/><sz val="11"/><name val="맑은 고딕"/><family val="3"/><charset val="129"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>'
        '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>' % NS,
    }
    sd = ''
    for r, (style, text) in enumerate(rows, 1):
        if text:
            sd += '<row r="%d"><c r="A%d" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c></row>' % (
                r, r, style, _x(text))
    files['xl/worksheets/sheet1.xml'] = (
        head + '<worksheet xmlns="%s" xmlns:r="%s"><sheetPr codeName="Sheet1"/>'
        '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="16.5"/>'
        '<cols><col min="1" max="1" width="120" customWidth="1" style="3"/></cols><sheetData>%s</sheetData>'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>' % (NS, R, sd))
    import io
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        fixed = (2026, 1, 1, 0, 0, 0)                     # 다시 만들어도 같은 파일이 나오게
        for name in ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels',
                     'xl/styles.xml', 'xl/worksheets/sheet1.xml']:
            z.writestr(zipfile.ZipInfo(name, fixed), files[name].encode('utf-8'), zipfile.ZIP_DEFLATED)
        z.writestr(zipfile.ZipInfo('xl/vbaProject.bin', fixed), vba_bin, zipfile.ZIP_DEFLATED)
    return buf.getvalue()


# ---------------------------------------------------------------------------
def make_excel():
    code = open(os.path.join(ROOT, 'excel', '임원일일예정_대시보드.bas'), 'rb').read().decode(ENC)
    vba = cfb_write(vba_streams([('ExecDashboard', code)]))
    rows = [
        (1, '임원 일일 예정사항 — 엑셀 매크로 판'), (0, ''),
        (2, '처음 한 번'),
        (0, '1. 위쪽 노란 줄의 [콘텐츠 사용]을 누릅니다.'),
        (0, '   빨간 줄 "Microsoft에서 이 파일의 매크로 실행을 차단했습니다"가 보이면: 엑셀을 닫고, 탐색기에서 이 파일을 오른쪽 단추 → 속성 → 아래쪽 [차단 해제]에 체크 → 확인 뒤 다시 엽니다.'),
        (0, '2. Alt+F8 → Setup → [실행]. 대시보드·설정·임원별 시트와 예시 일정 8명이 만들어집니다.'),
        (0, '3. Ctrl+S 로 저장합니다(.xlsm 그대로). 다음부터는 이 파일을 열어 씁니다. 이 "시작" 시트는 지워도 됩니다.'),
        (0, ''),
        (2, '매일'),
        (0, '임원별 시트에 일정을 쓰고 대시보드의 [새로 고침](Ctrl+Shift+R) → [인쇄 미리보기] 또는 [PDF로 저장]. 자세한 내용은 저장소 README의 "엑셀(VBA) 대시보드" 절.'),
    ]
    out = os.path.join(ROOT, 'excel', '임원일일예정_대시보드.xlsm')
    open(out, 'wb').write(xlsm_bytes('시작', rows, vba))
    return out


def make_hancell():
    mods = []
    for k in range(1, 8):
        code = open(os.path.join(ROOT, 'hancell', '붙여넣기', '임원', '모듈%d.txt' % k), encoding='utf-8').read()
        for a, b in CP949_FIX:
            code = code.replace(a, b)
        mods.append(('Module%d' % k, code))
    vba = cfb_write(vba_streams(mods))
    rows = [
        (1, '임원 일일 예정사항 — 한셀 매크로 판'), (0, ''),
        (0, '이 파일에는 매크로 모듈 7개(Module1~Module7)가 들어 있습니다. 한셀에서 아직 열어 보지 못한 파일입니다.'), (0, ''),
        (2, '처음 한 번'),
        (0, '1. 열 때 매크로를 쓸지 물으면 [사용]을 고릅니다.'),
        (0, '2. 매크로 목록(도구 → 매크로)에서 Setup → 실행. 입력·A4·임원 시트 8개·맨밑표·서식 시트와 단추가 만들어집니다.'),
        (0, '3. [다른 이름으로 저장] → 매크로 사용 문서(.xlsm). 다음부터는 이 파일을 열어 씁니다.'), (0, ''),
        (2, '매크로 목록에 Setup이 없거나 실행되지 않으면'),
        (0, '한셀이 이 파일에 든 매크로를 읽지 못한 것입니다. 저장소 README "한셀 버전 — 처음 한 번 만들기"처럼 hancell/붙여넣기/임원/모듈1.txt ~ 모듈7.txt 를 새 문서에 차례로 붙여넣고 Setup 을 실행한 뒤 .xlsm 으로 저장하세요.'),
    ]
    out = os.path.join(ROOT, 'hancell', '임원일일예정_한셀.xlsm')
    open(out, 'wb').write(xlsm_bytes('시작', rows, vba))
    return out


if __name__ == '__main__':
    for f in (make_excel(), make_hancell()):
        print('%s: %d bytes' % (os.path.relpath(f, ROOT), os.path.getsize(f)))
