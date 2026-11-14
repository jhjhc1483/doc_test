// 배포 파일 만들기: src/의 HTML(평문 원본) → 루트의 HTML(화면·코드를 암호화한 파일)
//
//   node tools/seal.mjs            (처음 비밀번호 1234)
//   node tools/seal.mjs 비밀번호
//
// - src 파일의 '앱 화면' 표시 사이 HTML과 app-code 안의 코드를 묶어 AES-256-GCM으로 잠급니다
//   (키: 비밀번호 + 무작위 salt로 PBKDF2-SHA256 60만 번). 형식은 잠금 화면 스크립트와 같습니다:
//   ENC1.<salt>.<반복 횟수>.<iv>.<암호문+인증값> (Base64)
// - 잠금 화면(비밀번호 칸·비밀번호 바꾸기 창)과 화면 꾸밈(CSS), 탭 제목은 평문으로 남습니다.
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const FILES = ['daily_schedule.html', 'executive_schedule.html'];
const ITER = 600000;
const MARK_A = '<!-- ▼ 앱 화면', MARK_B = '<!-- ▲ 앱 화면 끝 -->';
const CODE_OPEN = '<script type="text/plain" id="app-code">', CLOSE = '</script>';
// 메모장으로 열었을 때 보이면 안 되는 글 (예시 부서·임원 이름, 코드)
const MUST_HIDE = ['기획조정과', '총무과', '김민수', '이정훈', 'function sampleState', 'function renderPreview', '주간 업무보고'];

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

for (const f of FILES) {
  const src = readFileSync(join(ROOT, 'src', f), 'utf8');
  const a = src.indexOf(MARK_A), b = src.indexOf(MARK_B);
  const cs = src.indexOf(CODE_OPEN), ce = cs < 0 ? -1 : src.indexOf(CLOSE, cs + CODE_OPEN.length);
  if (a < 0 || b < a || cs < b || ce < 0) throw new Error(f + ': 앱 화면 표시나 app-code를 찾지 못했습니다.');
  const markup = src.slice(src.indexOf('-->', a) + 3, b);
  const code = src.slice(cs + CODE_OPEN.length, ce);
  const env = await seal(JSON.stringify({ m: markup, c: code }));
  const out = src.slice(0, a) + '<!-- 앱 화면: 암호화되어 app-code 안에 있습니다 -->' +
    src.slice(b + MARK_B.length, cs + CODE_OPEN.length) + '\n' + env + '\n' + src.slice(ce);
  const leak = MUST_HIDE.filter((w) => out.includes(w));
  if (leak.length) throw new Error(f + ': 평문으로 남은 글 — ' + leak.join(', '));
  writeFileSync(join(ROOT, f), out);
  console.log(f + ': ' + Math.round(out.length / 1024) + 'KB (화면·코드 ' + Math.round((markup.length + code.length) / 1024) + 'KB 암호화)');
}
