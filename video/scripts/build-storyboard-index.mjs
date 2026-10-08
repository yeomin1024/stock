// VERSION: v2.4.0 — 2026-10-08 — 스토리보드 검수 페이지 생성 (가이드 v2 6-5). v2.4.0: 고지 카드 3.5초 반영
// 사용:
//   node scripts/build-storyboard-index.mjs                 → out/storyboard/index.html (같은 폴더의 PNG 를 바로 참조)
//   node scripts/build-storyboard-index.mjs --artifact DIR  → DIR/index.html (문서 껍데기 없는 조각) + DIR/img/*.jpg (1280px 미리보기)
// 입력: out/storyboard/manifest.json (render-storyboard.mjs 가 씀), src/data/timeline.ts, src/data/facts.ts
// 카드 하나 = still 하나: 이미지 / 장면 번호 / 자막 번호 / 자막 문장 / 연결 근거 한 줄. 장면별로 묶는다.
import fs from 'node:fs';
import path from 'node:path';
import {spawnSync} from 'node:child_process';
import {fileURLToPath} from 'node:url';
import {log} from './log.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SB = path.join(root, 'out', 'storyboard');
const argIdx = process.argv.indexOf('--artifact');
const artifactDir = argIdx > 0 ? path.resolve(process.argv[argIdx + 1] ?? '') : null;

const t0 = Date.now();
let manifest;
try {
	manifest = JSON.parse(fs.readFileSync(path.join(SB, 'manifest.json'), 'utf8'));
} catch (err) {
	log.error('LOAD', {file: 'out/storyboard/manifest.json', error: err.message, next: 'node scripts/render-storyboard.mjs 를 먼저 실행'});
	process.exit(1);
}
const tl = await import(path.join(root, 'src', 'data', 'timeline.ts'));
const {NOTICE} = await import(path.join(root, 'src', 'data', 'facts.ts'));
const missing = manifest.filter((m) => !fs.existsSync(path.join(SB, m.file)));
log.info('LOAD', {stills: manifest.length, scenes: tl.SCENES.length, missingPng: missing.length});
if (missing.length) {
	log.error('LOAD', {missing: missing.map((m) => m.file).join(','), next: 'node scripts/render-storyboard.mjs 로 다시 렌더'});
	process.exit(1);
}

// ---------------------------------------------------------------------------
// 검수 메모 (자체 검수 결과와 확인이 필요한 해석). 장면 번호는 바로가기 링크가 된다.
const FIXES = [
	['S16', '문장 31의 ✕가 "현금 0원"을 가려서 ✕를 "물타기" 제목 위로만 줄였습니다.'],
	['S17', '문장 36의 "필요"가 "+40%"에 붙어 있어 18px 띄웠습니다.'],
	['S19', '문장 47의 칸 3개가 JP모건 카드 아래 테두리와 겹쳐서 40px 내렸습니다.'],
	['S25', '문장 61–63의 "-2,000만 원"이 왼쪽 여백 96px를 넘어 60px에서 시작해서, 왼쪽 열 중심을 옮겼습니다.'],
	['S26', '카드 3장의 제목 높이를 맞추고, 문장 69 번개가 사라지지 않게 했습니다(가이드: 번개가 한 칩에 떨어짐). 번개 윗끝은 y 124입니다.'],
	['S27', '점 500개 그리드가 제목에 붙어 아래로 내리고, 도넛 안쪽을 넓혀 "90%"가 고리에 닿지 않게 했습니다.'],
	['S29', '작게 다시 나온 계좌 카드의 "MDB 100%"가 25px여서 카드 배율을 0.7에서 0.8로 올렸습니다(29px, 최소 28px 규칙).'],
];
const DECISIONS = [
	['', '자막 한 줄 기준: 두 줄을 이어 붙여 자막 띠 최대 폭(1432px, Noto Sans KR 700 46px로 실측) 안에 들어가면 한 줄로 합치고, 넘으면 원래 줄바꿈 위치에서 두 자막으로 나눴습니다(시간은 글자 수 비율). 합침 53, 나눔 26, 혼잣말 3개는 그대로.'],
	['', '읽기 속도는 1.1배(10% 빠르게)로 했습니다. 끝 여유 1초는 그대로입니다.'],
	['S11', '고지 카드는 가이드의 6초 대신 3.5초입니다(사용자 요청). 문구 3줄이 1초 안에 다 나오도록 등장을 앞당겼고, SRT에도 같은 3.5초 공백이 있어 SRT 시간 = 영상 시간입니다.'],
	['', '두 자막으로 나뉜 문장 26개는 뒷줄 내용에 해당하는 그림(예: S07 "최대 -26%", S28 화살표 2개)을 뒷줄 자막이 뜰 때 나오게 옮겼습니다.'],
	['', '출처 캡션은 가이드 예시의 24px 대신 28px입니다. 최소 글자 크기 28px 규칙을 우선했습니다.'],
	['S03', '이후 계좌 카드 머리글은 "내 계좌"만 씁니다. 잔고(예: 8,000만 원)는 데이터 시트에 없는 값이라 넣지 않았습니다.'],
	['', '번개 도형은 잉크색입니다. 노랑은 화면마다 강조 하나에만 씁니다.'],
	['S04', '색 빠짐은 CSS 필터 대신 색 보간으로 처리했습니다(무거운 효과 금지).'],
	['S13', '"장기 상승" 라벨은 문장 22에서 빠집니다. 한 화면 텍스트 덩어리를 3개 이하로 맞췄습니다.'],
	['S18', '문장 44의 커서는 노랑 4칸을 다 찾으면 사라지고, 찾은 칸에는 테두리가 남습니다.'],
	['S30', '문장 90–91 동안 그래픽이 천천히 사라지고, 마지막 1초는 잉크색으로 페이드됩니다.'],
	['', '가이드 목록 외에 S10·S22·S24·S30의 자막별 still과 S04 와이프 중간 still을 더 넣었습니다.'],
	['S24', 'S23(엔론)이 빠져 S22와 S24가 모두 찢어진 종이 와이프로 들어옵니다(와이프 연속 1곳). S24는 문제에서 해결책으로 넘어가는 파트 경계라 와이프를 유지했습니다.'],
	['', '입력에 내레이션·BGM 파일이 없어 영상은 무음입니다.'],
];

// ---------------------------------------------------------------------------
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const tc = (frame) => {
	const sec = frame / tl.FPS;
	return `${Math.floor(sec / 60)}:${(sec % 60).toFixed(1).padStart(4, '0')}`;
};
const BG = {S04: '크림 → 네이비', S11: '크림 + 노랑 띠', S24: '크림 → 밝은 크림'};
const TONE = {cream: '크림', navy: '네이비', bright: '밝은 크림'};
const ENTER = {cut: '컷', wipe: '찢어진 종이 와이프'};
const imgSrc = (file) => (artifactDir ? `img/${file.replace(/\.png$/, '.jpg')}` : file);
const sceneLink = (id) => (id ? `<a class="sid" href="#${id}">${id}</a> ` : '');

const labelText = (m) => (m.file === 'S04-wipe.png' ? '와이프 중간' : m.isSub ? `문장 ${m.subs[0].n}` : '장면 대표');
const cardLabel = (m) => `<span class="badge${!m.isSub && m.file !== 'S04-wipe.png' ? ' main' : ''}">${labelText(m)}</span>`;

const subsBlock = (m) => {
	if (!m.subs.length) {
		return `<p class="label">고지 문구 · 대본 [장면] 그대로</p><ol class="lines">${NOTICE.map((s) => `<li>${esc(s)}</li>`).join('')}</ol>`;
	}
	const srtOf = (s) => (s.srt.first === s.srt.last ? `${s.srt.first}` : `${s.srt.first}–${s.srt.last}`);
	const first = m.subs[0];
	const last = m.subs[m.subs.length - 1];
	const nums = `${m.subs.length > 1 ? `문장 ${first.n}–${last.n}` : `문장 ${first.n}`} · SRT ${srtOf({srt: {first: first.srt.first, last: last.srt.last}})}`;
	return `<p class="label">${nums}</p><ol class="lines">${m.subs
		.map((s) => `<li><span class="n">${s.n}</span><span>${esc(s.text.replace(/\n/g, ' '))}${s.srt.last > s.srt.first ? ` <span class="split">(화면 ${s.srt.last - s.srt.first + 1}줄)</span>` : ''}</span></li>`)
		.join('')}</ol>`;
};

let idx = 0;
const cards = [];
const sections = tl.SCENES.map((s) => {
	const shots = manifest.filter((m) => m.scene === s.id);
	const {start, end} = tl.sceneRange(s.id);
	const subsLabel = s.subs.length ? `문장 ${s.subs[0]}${s.subs.length > 1 ? `–${s.subs[s.subs.length - 1]}` : ''} · SRT ${tl.srtRange(s.subs[0]).first}–${tl.srtRange(s.subs[s.subs.length - 1]).last}` : `고지 카드 ${tl.DISCLAIMER_SEC}초`;
	const body = shots
		.map((m) => {
			const i = idx++;
			const why = m.why ?? s.why;
			cards.push({src: imgSrc(m.file), title: `${s.id} · ${labelText(m)}`, why});
			return `<figure class="card${m.isSub ? '' : ' main'}" data-main="${m.isSub ? 0 : 1}">
<button type="button" class="shot" data-i="${i}" aria-label="${esc(m.file)} 크게 보기"><img src="${imgSrc(m.file)}" alt="${esc(`${s.id} ${m.file} 스틸`)}" width="1920" height="1080" loading="lazy" decoding="async"></button>
<figcaption>
<div class="row">${cardLabel(m)}<span class="mono">${esc(m.file.replace('.png', ''))} · f${m.frame} · ${tc(m.frame)}</span></div>
${subsBlock(m)}
<p class="reason"><span class="k">연결 근거</span>${esc(why)}</p>
</figcaption>
</figure>`;
		})
		.join('\n');
	return `<section class="scene" id="${s.id}">
<header class="scene-head">
<h2>${s.id}</h2>
<span class="meta">${subsLabel}</span>
<span class="meta mono">${tc(start)}–${tc(end)}</span>
<span class="chip">${ENTER[s.enter]}</span>
<span class="chip">${BG[s.id] ?? TONE[s.tone]}</span>
<span class="meta">${shots.length}장</span>
${shots.length > 1 ? `<p class="scene-why"><span class="k">장면 연결 근거</span>${esc(s.why)}</p>` : ''}
</header>
<div class="grid">
${body}
</div>
</section>`;
});

const total = tl.TOTAL_FRAMES;
const mainCount = manifest.filter((m) => !m.isSub).length;
const listItems = (rows) => rows.map(([id, text]) => `<li>${sceneLink(id)}${esc(text)}</li>`).join('');

const STYLE = `
/* Layout: 리뷰 시트 — 요약과 검수 메모, 장면 번호 바로가기, 장면마다 머리줄 + still 카드 그리드 */
:root {
	--bg: #f2f4f7;
	--surface: #ffffff;
	--fg: #172033;
	--muted: #586377;
	--line: #d8dde6;
	--accent: #223d6c;
	--mark: #f2c21b;
	--mark-fg: #172033;
	--shade: rgba(9, 13, 22, 0.82);
	--font-display: 'Noto Serif KR', 'Nanum Myeongjo', 'AppleMyungjo', serif;
	--font-body: 'IBM Plex Sans KR', 'Apple SD Gothic Neo', 'Malgun Gothic', system-ui, sans-serif;
	--font-mono: 'IBM Plex Mono', ui-monospace, 'SFMono-Regular', Menlo, monospace;
}
@media (prefers-color-scheme: dark) {
	:root:not([data-theme='light']) {
		--bg: #0e131c;
		--surface: #161d2a;
		--fg: #e5e9f0;
		--muted: #98a2b3;
		--line: #273145;
		--accent: #9db8e8;
		--mark: #f2c21b;
		--mark-fg: #172033;
		--shade: rgba(3, 5, 9, 0.88);
		color-scheme: dark;
	}
}
:root[data-theme='dark'] {
	--bg: #0e131c;
	--surface: #161d2a;
	--fg: #e5e9f0;
	--muted: #98a2b3;
	--line: #273145;
	--accent: #9db8e8;
	--mark: #f2c21b;
	--mark-fg: #172033;
	--shade: rgba(3, 5, 9, 0.88);
	color-scheme: dark;
}
* { box-sizing: border-box; }
body { margin: 0; background: var(--bg); color: var(--fg); font-family: var(--font-body); font-size: 15px; line-height: 1.6; }
a { color: var(--accent); }
:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
.mono { font-family: var(--font-mono); font-variant-numeric: tabular-nums; font-size: 0.86em; color: var(--muted); }
.wrap { max-width: 1260px; margin: 0 auto; padding-inline: clamp(16px, 4vw, 40px); padding-block: 36px 96px; }
.top { display: grid; gap: 14px; padding-bottom: 28px; }
.eyebrow { margin: 0; font-size: 12.5px; letter-spacing: 0.06em; color: var(--muted); }
h1 { margin: 0; font-family: var(--font-display); font-weight: 900; font-size: clamp(30px, 4.4vw, 46px); line-height: 1.2; text-wrap: balance; }
.facts { display: flex; flex-wrap: wrap; gap: 10px 28px; margin: 0; }
.facts div { display: grid; gap: 0; }
.facts dt { font-size: 12.5px; color: var(--muted); }
.facts dd { margin: 0; font-family: var(--font-mono); font-variant-numeric: tabular-nums; font-size: 15px; }
.status { display: flex; flex-wrap: wrap; align-items: center; gap: 10px; margin: 4px 0 0; }
.notes { display: grid; grid-template-columns: repeat(auto-fit, minmax(min(100%, 420px), 1fr)); gap: 18px; padding-block: 8px 28px; }
.note { background: var(--surface); border: 1px solid var(--line); border-radius: 10px; padding: 18px 20px 6px; min-width: 0; }
.note h2 { margin: 0 0 8px; font-size: 16px; font-weight: 700; }
.note ul { margin: 0; padding-left: 18px; display: grid; gap: 6px; padding-bottom: 12px; }
.note li { font-size: 14px; }
.sid { font-family: var(--font-mono); font-weight: 600; text-decoration: none; }
.sid:hover { text-decoration: underline; }
.index { position: sticky; top: env(safe-area-inset-top, 0px); z-index: 5; background: var(--bg); border-bottom: 1px solid var(--line); padding-block: 10px; display: flex; flex-wrap: wrap; align-items: center; gap: 6px; }
.index a { font-family: var(--font-mono); font-size: 13px; text-decoration: none; color: var(--fg); border: 1px solid var(--line); background: var(--surface); border-radius: 6px; padding: 2px 7px; }
.index a:hover { border-color: var(--accent); }
.index label { margin-left: auto; display: inline-flex; align-items: center; gap: 6px; font-size: 13.5px; color: var(--muted); cursor: pointer; }
.scene { padding-block: 30px 10px; border-bottom: 1px solid var(--line); scroll-margin-top: 64px; }
.scene-head { display: flex; flex-wrap: wrap; align-items: baseline; gap: 6px 14px; margin-bottom: 16px; }
.scene-head h2 { margin: 0; font-family: var(--font-display); font-weight: 900; font-size: 30px; line-height: 1; }
.meta { font-size: 14px; color: var(--muted); }
.chip { font-size: 12.5px; border: 1px solid var(--line); border-radius: 999px; padding: 1px 9px; color: var(--muted); }
.scene-why { flex-basis: 100%; margin: 2px 0 0; font-size: 14.5px; }
.k { display: inline-block; margin-right: 8px; font-size: 12px; letter-spacing: 0.04em; color: var(--muted); }
.grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(min(100%, 360px), 1fr)); gap: 18px; }
.card { margin: 0; display: flex; flex-direction: column; background: var(--surface); border: 1px solid var(--line); border-radius: 10px; overflow: hidden; min-width: 0; }
.shot { display: block; padding: 0; border: 0; background: var(--line); cursor: zoom-in; width: 100%; }
.shot img { display: block; width: 100%; max-width: 100%; height: auto; aspect-ratio: 16 / 9; }
figcaption { display: grid; gap: 8px; padding: 12px 14px 14px; }
.row { display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 6px; }
.badge { font-size: 12.5px; font-weight: 700; border: 1px solid var(--line); border-radius: 6px; padding: 0 7px; }
.badge.main { background: var(--mark); color: var(--mark-fg); border-color: var(--mark); }
.label { margin: 0; font-size: 12.5px; color: var(--muted); }
.lines { list-style: none; margin: 0; padding: 0; display: grid; gap: 4px; }
.lines li { display: grid; grid-template-columns: 2.2em minmax(0, 1fr); font-size: 14.5px; line-height: 1.5; }
.lines .split { font-size: 12.5px; color: var(--muted); }
.lines .n { font-family: var(--font-mono); font-variant-numeric: tabular-nums; font-size: 12.5px; color: var(--muted); padding-top: 2px; }
.reason { margin: 0; font-size: 14px; padding-top: 8px; border-top: 1px dashed var(--line); }
body.only-main .card:not(.main) { display: none; }
dialog { width: min(96vw, 1440px); max-height: 96vh; padding: 0; border: 0; border-radius: 10px; background: var(--surface); color: var(--fg); overflow: auto; }
dialog::backdrop { background: var(--shade); }
dialog img { display: block; width: 100%; height: auto; aspect-ratio: 16 / 9; background: var(--line); }
.viewer-bar { display: flex; flex-wrap: wrap; align-items: center; gap: 8px 16px; padding: 12px 16px; }
.viewer-bar p { margin: 0; flex: 1 1 280px; min-width: 0; font-size: 14px; }
.viewer-bar button { font: inherit; font-size: 14px; color: var(--fg); background: var(--bg); border: 1px solid var(--line); border-radius: 6px; padding: 4px 12px; cursor: pointer; }
@media (max-width: 700px) {
	.index { position: static; }
	.index label { margin-left: 0; }
}
`;

const SCRIPT = `
(function () {
	var CARDS = ${JSON.stringify(cards).replace(/</g, "\\u003c")};
	var dlg = document.getElementById('viewer');
	var img = document.getElementById('viewer-img');
	var cap = document.getElementById('viewer-cap');
	var cur = 0;
	function visible() {
		return Array.prototype.filter.call(document.querySelectorAll('.shot'), function (b) { return b.offsetParent !== null; }).map(function (b) { return +b.dataset.i; });
	}
	function show(i) {
		cur = i;
		var c = CARDS[i];
		img.src = c.src;
		img.alt = c.title;
		cap.textContent = c.title + ' — ' + c.why;
	}
	function step(d) {
		var v = visible();
		var k = v.indexOf(cur);
		if (k < 0) return;
		show(v[(k + d + v.length) % v.length]);
	}
	document.addEventListener('click', function (e) {
		var b = e.target.closest('.shot');
		if (!b) return;
		show(+b.dataset.i);
		if (typeof dlg.showModal === 'function') dlg.showModal();
	});
	document.getElementById('viewer-prev').addEventListener('click', function () { step(-1); });
	document.getElementById('viewer-next').addEventListener('click', function () { step(1); });
	document.getElementById('viewer-close').addEventListener('click', function () { dlg.close(); });
	dlg.addEventListener('click', function (e) { if (e.target === dlg) dlg.close(); });
	dlg.addEventListener('keydown', function (e) {
		if (e.key === 'ArrowLeft') { e.preventDefault(); step(-1); }
		if (e.key === 'ArrowRight') { e.preventDefault(); step(1); }
	});
	var only = document.getElementById('only-main');
	only.addEventListener('change', function () { document.body.classList.toggle('only-main', only.checked); });
})();
`;

const CONTENT = `<main class="wrap">
<header class="top">
<p class="eyebrow">영상 가이드 v2 · 스토리보드 · 엔론 사례 제외 대본</p>
<h1>몰빵 영상 스토리보드</h1>
<dl class="facts">
<div><dt>장면</dt><dd>${tl.SCENES.length}</dd></div>
<div><dt>still</dt><dd>${manifest.length} (장면 대표 ${mainCount})</dd></div>
<div><dt>길이</dt><dd>${tc(total)} · ${total.toLocaleString('en-US')}f · ${tl.FPS}fps</dd></div>
<div><dt>화면</dt><dd>${tl.WIDTH}×${tl.HEIGHT}</dd></div>
<div><dt>자막</dt><dd>대본 문장 ${tl.EXPECTED_SENTENCES}개 = 한 줄 자막 ${tl.SUBTITLE_COUNT}개 + 고지 카드 ${tl.DISCLAIMER_SEC}초</dd></div>
</dl>
<p class="status"><span class="badge main">v2.4.0</span>고지 카드 3.5초, SRT 시간 = 영상 시간. 자막을 한 줄씩(혼잣말 3개만 두 줄) 나누고 읽기 속도를 1.1배로 했습니다. 카드의 번호는 대본 문장 번호(TXT 줄 순서)이고, SRT 자막 번호를 함께 적었습니다. 엔론 사례는 없습니다(장면 S23 없음).</p>
</header>
<div class="notes">
<section class="note"><h2>자체 검수에서 고친 것</h2><ul>${listItems(FIXES)}</ul></section>
<section class="note"><h2>확인이 필요한 해석</h2><ul>${listItems(DECISIONS)}</ul></section>
</div>
<nav class="index" aria-label="장면 바로가기">
${tl.SCENES.map((s) => `<a href="#${s.id}">${s.id}</a>`).join('')}
<label for="only-main"><input type="checkbox" id="only-main"> 장면 대표 still만 보기</label>
</nav>
${sections.join('\n')}
</main>
<dialog id="viewer" aria-label="스틸 크게 보기">
<img id="viewer-img" alt="" width="1920" height="1080">
<div class="viewer-bar"><p id="viewer-cap"></p><button type="button" id="viewer-prev">← 이전</button><button type="button" id="viewer-next">다음 →</button><button type="button" id="viewer-close">닫기</button></div>
</dialog>
<script>${SCRIPT}</script>`;

const TITLE = '몰빵 영상 스토리보드';
const FONTS =
	'<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>' +
	'<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans+KR:wght@400;700&family=Noto+Serif+KR:wght@900&display=swap">';

if (artifactDir) {
	// 아티팩트: 문서 껍데기 없이 <title> + <style> + 본문. 이미지는 1280px JPEG 로 줄여 함께 올린다.
	const imgDir = path.join(artifactDir, 'img');
	fs.mkdirSync(imgDir, {recursive: true});
	let bytes = 0;
	for (const m of manifest) {
		const out = path.join(imgDir, m.file.replace(/\.png$/, '.jpg'));
		const r = spawnSync('ffmpeg', ['-loglevel', 'error', '-y', '-i', path.join(SB, m.file), '-vf', 'scale=1280:-2', '-q:v', '3', out]);
		if (r.status !== 0) {
			log.error('JPEG', {file: m.file, error: String(r.stderr).trim(), next: 'ffmpeg 설치와 PNG 파일 확인'});
			process.exit(1);
		}
		bytes += fs.statSync(out).size;
	}
	log.info('JPEG', {files: manifest.length, totalMB: (bytes / 1e6).toFixed(1), dir: imgDir});
	fs.writeFileSync(path.join(artifactDir, 'index.html'), `<title>${TITLE}</title>\n${FONTS}\n<style>${STYLE}</style>\n${CONTENT}\n`);
	log.info('DONE', {out: path.join(artifactDir, 'index.html'), cards: cards.length, ms: Date.now() - t0});
} else {
	const html = `<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>${TITLE}</title>
${FONTS}
<style>${STYLE}</style>
</head>
<body>
${CONTENT}
</body>
</html>
`;
	fs.writeFileSync(path.join(SB, 'index.html'), html);
	log.info('DONE', {out: path.relative(root, path.join(SB, 'index.html')), cards: cards.length, ms: Date.now() - t0});
}
