// 배포 파일 만들기: src/의 두 화면(평문 원본)을 묶어 → 루트의 일일예정사항.html (화면·코드를 암호화한 파일 하나)
//
//   node tools/seal.mjs            (처음 비밀번호 1234)
//   node tools/seal.mjs 비밀번호
//
// - src/shell.html: 비밀번호 화면 + 위쪽 메뉴(부서별 / 임원) + 칸(iframe)을 띄우는 틀
// - src/daily_schedule.html, src/executive_schedule.html: 두 화면. 각자 열어서 따로 시험할 수도 있습니다(암호화 안 됨).
//   두 파일에서 꾸밈(head의 style)·'앱 화면' 표시 사이 HTML·app-code 안의 코드를 꺼내 칸 하나짜리 문서로 만들고,
//   두 문서를 묶어 AES-256-GCM으로 잠급니다 (키: 비밀번호 + 무작위 salt로 PBKDF2-SHA256 60만 번).
//   형식은 틀의 잠금 스크립트와 같습니다: ENC1.<salt>.<반복 횟수>.<iv>.<암호문+인증값> (Base64)
// - 평문으로 남는 것: 탭 제목, 비밀번호 화면, 위쪽 메뉴 이름. 두 화면의 꾸밈(CSS)·HTML·코드는 모두 암호문 안에 있습니다.
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const SHELL = 'shell.html', OUT = '일일예정사항.html';
const APPS = { daily: 'daily_schedule.html', exec: 'executive_schedule.html' };
const ITER = 600000;
const MARK_A = '<!-- ▼ 앱 화면', MARK_B = '<!-- ▲ 앱 화면 끝 -->';
const CODE_OPEN = '<script type="text/plain" id="app-code">', CLOSE = '</script>';
const ST_EMPTY = '<script id="embedded-state" type="application/json">null</script>';
// 메모장으로 열었을 때 보이면 안 되는 글 (예시 부서·임원 이름, 코드, 꾸밈)
const MUST_HIDE = ['기획조정과', '총무과', '김민수', '이정훈', 'function sampleState', 'function renderPreview', '주간 업무보고', '.doc-title', '.topbar'];

const pw = process.argv[2] || '1234';
const { subtle } = globalThis.crypto;
const b64 = (u8) => Buffer.from(u8).toString('base64');

async function seal(text) {
  const salt = globalThis.crypto.getRandomValues(new Uint8Array(16));
  const iv = globalThis.crypto.getRandomValues(new Uint8Array(12));
  const base = await subtle.importKey('raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveKey']);
  const key = await subtle.deriveKey({ name: 'PBKDF2', salt, iterations: ITER, hash: 'SHA-256' }, base,
    { name: 'AES-GCM', length: 256 }, false, ['encrypt']);
  const ct = new Uint8Array(await subtle.encrypt({ name: 'AES-GCM', iv }, key, new TextEncoder().encode(text)));
  return ['ENC1', b64(salt), ITER, b64(iv), b64(ct)].join('.');
}

// 칸(iframe) 안에 들어갈 화면 하나짜리 문서: 원본의 head(잠금 화면 꾸밈 빼고) + 앱 화면 HTML + 코드.
// 코드 앞에 틀의 열쇠·저장 도구를 이어 주는 짧은 스크립트를 넣습니다.
function appDoc(id, f) {
  const src = readFileSync(join(ROOT, 'src', f), 'utf8');
  const h0 = src.indexOf('<head>'), h1 = src.indexOf('</head>');
  const a = src.indexOf(MARK_A), b = src.indexOf(MARK_B);
  const cs = src.indexOf(CODE_OPEN), ce = cs < 0 ? -1 : src.indexOf(CLOSE, cs + CODE_OPEN.length);
  if (h0 < 0 || h1 < h0 || a < h1 || b < a || cs < b || ce < 0) throw new Error(f + ': head, 앱 화면 표시나 app-code를 찾지 못했습니다.');
  const head = src.slice(h0 + 6, h1).replace(/<style id="lock-style">[\s\S]*?<\/style>\s*/, '');
  if (!head.includes(ST_EMPTY)) throw new Error(f + ': head에 빈 embedded-state가 없습니다.');
  const markup = src.slice(src.indexOf('-->', a) + 3, b);
  const code = src.slice(cs + CODE_OPEN.length, ce);
  const bridge = [
    '/* 합친 파일의 칸 안: 틀(위 창)의 열쇠·비밀번호 바꾸기·파일 만들기를 같이 씀 */',
    '(function () {',
    '  var P = window.parent;',
    '  window.__docLock = P.__docLock; window.__lock = P.__lock; window.__host = P.__hostFor(' + JSON.stringify(id) + ');',
    '  /* Ctrl+P: 이 화면의 [인쇄 / PDF] (브라우저 기본 인쇄는 창 전체를 찍음) */',
    '  document.addEventListener("keydown", function (e) {',
    '    if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && (e.code === "KeyP" || e.key === "p" || e.key === "P")) {',
    '      e.preventDefault();',
    '      var b = document.getElementById("btnPrint");',
    '      if (b) b.click(); else window.print();',
    '    }',
    '  });',
    '})();',
  ].join('\n');
  // 인쇄할 때 칸(iframe) 안 문서는 스크롤하지 않음: A4 시트가 칸보다 조금(0.5mm) 커도 스크롤바가 같이 찍히지 않게
  const framePrint = '<style id="frame-print">@media print { html, body { overflow: hidden !important; scrollbar-width: none; } ' +
    '::-webkit-scrollbar { display: none; width: 0; height: 0; } }</style>\n';
  return '<!DOCTYPE html>\n<html lang="ko">\n<head>' + head + framePrint + '</head>\n<body>' + markup +
    '<script>\n' + bridge + '\n</script>\n<script>' + code + '</script>\n</body>\n</html>\n';
}

const shell = readFileSync(join(ROOT, 'src', SHELL), 'utf8');
const cs = shell.indexOf(CODE_OPEN), ce = cs < 0 ? -1 : shell.indexOf(CLOSE, cs + CODE_OPEN.length);
if (cs < 0 || ce < 0) throw new Error(SHELL + ': app-code를 찾지 못했습니다.');
const apps = {};
for (const [id, f] of Object.entries(APPS)) apps[id] = appDoc(id, f);
const text = JSON.stringify({ v: 2, apps });
const env = await seal(text);
const out = shell.slice(0, cs + CODE_OPEN.length) + '\n' + env + '\n' + shell.slice(ce);
const leak = MUST_HIDE.filter((w) => out.includes(w));
if (leak.length) throw new Error(OUT + ': 평문으로 남은 글 — ' + leak.join(', '));
writeFileSync(join(ROOT, OUT), out);
console.log(OUT + ': ' + Math.round(Buffer.byteLength(out) / 1024) + 'KB (두 화면 ' +
  Object.entries(apps).map(([id, d]) => id + ' ' + Math.round(Buffer.byteLength(d) / 1024) + 'KB').join(', ') + ' 암호화)');
