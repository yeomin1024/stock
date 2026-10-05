# 몰빵 영상 — 모션그래픽 (Remotion)

> 연구·교육용 설명 영상입니다. 투자 권유가 아닙니다.

영상 제목: **한 종목에만 몰빵해서 얻은 천만원 수익이 -3천만원이 됐습니다. 한 종목 올인한 사람들의 최악의 결말**

`input/video_guide.md`(제작 지시서)를 그대로 따라 만든 Vox 스타일 설명 영상입니다.
이미지·이모지·로고 없이, 모든 그래픽은 코드로 그린 벡터와 텍스트입니다.

| 항목 | 값 |
|---|---|
| 컴포지션 | `Molbbang` — 1920×1080, 30fps, 7547프레임 (4분 11.57초) |
| 길이 계산 | SRT 마지막 자막 끝 247.565초 + 고지 카드 3초 + 여유 1초 |
| 최종 출력 | `out/final_4k.mp4` (3840×2160, h264, CRF 16) |
| 미리보기 | `out/preview_1080p.mp4` |
| 오디오 | 내레이션/BGM 파일이 없어 무음 |

## 입력 파일 (`input/`)

- `대본_완성본.srt` — 자막 텍스트와 타이밍의 기준 (자막 43개)
- `대본_완성본.txt` — 파트 구조와 `[장면]` 지시
- `video_guide.md` — 제작 지시서

## 명령어

```console
npm i                     # 의존성 설치
npm run prepare-data      # SRT → src/data/subtitles.ts, 전체 길이/장면표 출력, 폰트 청크 목록 갱신
npm run dev               # Remotion Studio 미리보기 (Scenes/ 폴더에 장면별 컴포지션)
npm run stills            # 검수용 still: 장면마다 2장 → out/stills/
npm run render:preview    # out/preview_1080p.mp4
npm run render:4k         # out/final_4k.mp4 (--scale=2 --codec=h264 --crf=16)
npm run render:4k:resumable  # 같은 설정을 4구간으로 나눠 렌더 → 스트림 복사로 무손실 연결 (재시작 시 끝난 구간 건너뜀)
```

`npm run stills -- S05 S12 --extra=977,2695` 처럼 장면/프레임을 골라 렌더할 수도 있습니다.

## 채널명 바꾸기

채널명은 아직 정해지지 않아 대본 표기 그대로 `[채널명]`으로 나옵니다 (S09, S24).

1. `src/data/facts.ts`의 `CHANNEL_NAME_DEFAULT`를 바꾸거나, 렌더 때 `--props='{"channelName":"내 채널"}'`을 넘깁니다.
2. 새 이름에 처음 쓰는 글자가 있으면 `npm run prepare-data`로 폰트 청크 목록을 다시 만듭니다
   (`--props`로만 넘긴 이름은 스캔되지 않으므로 `facts.ts`에 적는 쪽을 권장).

자막 12번의 `[채널명]`은 SRT 원문이므로 자막 쪽은 SRT를 고쳐야 바뀝니다.

## 내레이션 / BGM을 넣을 때

가이드 2번: 내레이션 파일을 넣으면 자막 19번 끝(01:38.63)에서 오디오를 나눠 뒷부분을 3초 뒤로 밀어야 합니다.
현재 버전은 오디오 파일이 없어 오디오 트랙 코드를 넣지 않았습니다. 넣을 경우 `src/data/timeline.ts`의
`READ_PACE`를 1.0으로 바꾸면 그래픽 타이밍이 읽기 속도 대신 발화 속도 기준이 됩니다.

## 구조

```
input/                     입력 파일 (SRT, TXT, 지시서)
scripts/                   parse-srt / build-glyphs / render-stills (구조화 로그)
src/data/subtitles.ts      SRT 파싱 결과 (자동 생성)
src/data/timeline.ts       모든 타이밍: 자막 번호 → 프레임, 장면표, 고지 카드 3초 반영
src/data/facts.ts          가이드 4번 데이터 시트 (화면 숫자는 여기 값만 사용)
src/design/                색, 폰트, 타이포, 모션 상수
src/components/            종이 질감, 자막, 형광펜, 카운터, 선 그리기, 와플 차트, 찢어진 종이 와이프 등
src/scenes/S01.tsx–S24.tsx 장면
src/Video.tsx              메인 타임라인
```

## 렌더 환경 메모

헤드리스 Chrome이 Google Fonts를 받을 때 이 클라우드 환경의 TLS 검사 프록시 CA를 신뢰해야 합니다.
CA는 `certutil`로 `~/.local/share/pki/nssdb`에 추가하고 `~/.pki/nssdb`를 그쪽으로 연결했습니다 (TLS 검증을 끄지 않음).
컨테이너가 재시작되면 이 연결이 지워지므로 `ln -sfn ~/.local/share/pki/nssdb ~/.pki/nssdb`로 다시 만듭니다. 일반 PC에서는 필요 없습니다.
