# 몰빵 영상 — 모션그래픽 (Remotion)

> 연구·교육용 설명 영상입니다. 투자 권유가 아닙니다.

영상 제목: **한 종목에만 몰빵해서 얻은 천만원 수익이 -2천만원이 됐습니다. 한 종목 올인한 사람들의 최악의 결말**

`input/video_guide.md`(제작 지시서 v2.0)를 그대로 따라 만든 Vox 스타일 설명 영상입니다.
이미지·이모지·로고·사람 그림 없이, 모든 그래픽은 코드로 그린 벡터와 텍스트입니다.

| 항목 | 값 |
|---|---|
| 컴포지션 | `Molbbang` — 1920×1080, 30fps, 16451프레임 (9분 8.37초) |
| 길이 계산 | SRT 마지막 자막 끝 541.369초 + 고지 카드 6초(자막 18번 뒤) + 여유 1초 |
| 장면 | 30개 (S01–S30, 고지 카드 S11 포함) |
| 스토리보드 | `out/storyboard/index.html` — still 85장 (장면 대표 30 + 자막별 55) |
| 최종 출력 | `out/final_1080p.mp4` (h264, CRF 18) — **스토리보드 승인 후 렌더** |
| 오디오 | 내레이션/BGM 파일이 없어 무음 |

## 입력 파일 (`input/`)

- `대본_완성본.srt` — 자막 텍스트와 타이밍의 기준 (자막 96개)
- `대본_완성본.txt` — 파트 구조와 `[장면]` 고지 문구
- `video_guide.md` — 제작 지시서 v2.0

## 명령어

```console
npm i                       # 의존성 설치
npm run prepare-data        # SRT → src/data/subtitles.ts (96개 검증), 전체 길이/장면표 출력, 폰트 청크 목록 갱신
npm run dev                 # Remotion Studio 미리보기 (Scenes/ 폴더에 장면별 컴포지션)
npm run storyboard          # 장면별 still 85장 + manifest.json → out/storyboard/, 이어서 index.html 생성
npm run storyboard:index    # index.html 만 다시 생성
npm run render              # out/final_1080p.mp4 (가이드 렌더 명령, 승인 후)
npm run render:4k           # 4K 는 요청이 있을 때만 (--scale=2)
```

장면 몇 개만 다시 렌더하려면 `node --disable-warning=MODULE_TYPELESS_PACKAGE_JSON scripts/render-storyboard.mjs S16 S25`
처럼 장면 번호를 넘깁니다 (manifest 는 항상 전체 목록으로 다시 씁니다). 그다음 `npm run storyboard:index`.

## 스토리보드 still 규칙

- 장면 대표 `S01.png`–`S30.png`: 장면의 모든 요소가 나온 시점(장면 끝 10프레임 전, 다음 장면이 와이프로 들어오면 와이프 시작 전).
- 자막별 `S25-66.png`: 그 자막이 끝나기 10프레임 전. 가이드가 지정한 S14, S16–S19, S25–S29에 더해 S10, S22, S23, S24, S30.
- `S04-wipe.png`: S04는 끝 화면이 네이비 한 색이라 와이프 중간 1장을 더 둡니다.
- 카드 문구: 장면 연결 근거는 `src/data/timeline.ts`의 `SCENES[].why`, 자막별 근거는 `src/data/storyboard.ts`.

## 내레이션 / BGM을 넣을 때

가이드 2번: 내레이션 파일이 있으면 자막 18번 끝(01:33.29)에서 오디오를 나눠 뒷부분을 6초 뒤로 밀고,
BGM은 내레이션보다 -18dB 정도 낮게 깝니다. 현재 버전은 오디오 파일이 없어 오디오 트랙 코드를 넣지 않았습니다.

## 구조

```
input/                      입력 파일 (SRT, TXT, 지시서)
scripts/                    parse-srt / build-glyphs / render-storyboard / build-storyboard-index (구조화 로그)
src/data/subtitles.ts       SRT 파싱 결과 (자동 생성)
src/data/timeline.ts        모든 타이밍: 자막 번호 → 프레임, 장면표(30), 고지 카드 6초 반영, 스토리보드 still 시점
src/data/facts.ts           가이드 4번 데이터 시트 + 고지 문구 (화면 숫자는 여기 값만 사용)
src/data/storyboard.ts      스토리보드 카드의 자막별 연결 근거
src/design/                 색(빨강=수익, 파랑=손실, 노랑=강조 하나, 회색=보조), 폰트, 타이포, 모션 상수
src/components/             종이 질감, 자막, 계좌 카드·막대, 형광펜, 카운터, 선 그리기, 와플, 찢어진 종이 와이프 등
src/scenes/S01.tsx–S30.tsx  장면
src/Video.tsx               메인 타임라인
```

## 렌더 환경 메모

헤드리스 Chrome이 Google Fonts를 받을 때 이 클라우드 환경의 TLS 검사 프록시 CA를 신뢰해야 합니다.
CA는 `certutil`로 `~/.local/share/pki/nssdb`에 추가하고 `~/.pki/nssdb`를 그쪽으로 연결했습니다 (TLS 검증을 끄지 않음).
컨테이너가 재시작되면 이 연결이 지워지므로 `ln -sfn ~/.local/share/pki/nssdb ~/.pki/nssdb`로 다시 만듭니다. 일반 PC에서는 필요 없습니다.
