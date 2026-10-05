#!/usr/bin/env bash
# VERSION: v1.0.1 — 2026-10-05 — 재시작에 강한 4K 렌더 (v1.0.1: ffprobe 프레임 수 파싱 수정 — csv 끝 쉼표 때문에 검증이 실패하던 문제)
# 가이드 명령(npx remotion render Molbbang out/final_4k.mp4 --scale=2 --codec=h264 --crf=16)과 같은 설정으로
# 프레임 구간을 나눠 렌더한 뒤 ffmpeg 스트림 복사(무손실, 재인코딩 없음)로 이어 붙인다.
# 끝난 구간은 .done 표시를 남겨, 컨테이너가 재시작돼도 남은 구간만 다시 렌더한다.
# 사용: bash scripts/render-4k.sh        (CHUNKS=4 기본)
set -euo pipefail
cd "$(dirname "$0")/.."

log() { echo "[$1] [$(date -u +%Y-%m-%dT%H:%M:%SZ)] ${*:2}"; }

TOTAL=$(node --disable-warning=MODULE_TYPELESS_PACKAGE_JSON -e 'import("./src/data/timeline.ts").then((t) => console.log(t.TOTAL_FRAMES))')
CHUNKS=${CHUNKS:-4}
DIR=out/4k_parts
OUT=out/final_4k.mp4
mkdir -p "$DIR"
SIZE=$(((TOTAL + CHUNKS - 1) / CHUNKS))
log PLAN "total_frames=$TOTAL chunks=$CHUNKS chunk_size=$SIZE out=$OUT"

: >"$DIR/list.txt"
for i in $(seq 0 $((CHUNKS - 1))); do
	a=$((i * SIZE))
	b=$(((i + 1) * SIZE - 1))
	[ "$b" -ge "$TOTAL" ] && b=$((TOTAL - 1))
	part="$DIR/part_$i.mp4"
	if [ -f "$part.done" ]; then
		log RENDER "chunk=$i frames=$a-$b status=skip_done"
	else
		log RENDER "chunk=$i frames=$a-$b status=start"
		t0=$(date +%s)
		if ! npx remotion render Molbbang "$part" --frames="$a-$b" --scale=2 --codec=h264 --crf=16 --concurrency=4 >"$DIR/part_$i.log" 2>&1; then
			log ERROR "chunk=$i frames=$a-$b log=$DIR/part_$i.log next=로그 확인 후 다시 실행(끝난 구간은 건너뜀)"
			exit 1
		fi
		touch "$part.done"
		log RENDER "chunk=$i frames=$a-$b status=done sec=$(($(date +%s) - t0))"
	fi
	echo "file 'part_$i.mp4'" >>"$DIR/list.txt"
done

log CONCAT "parts=$CHUNKS mode=stream_copy"
ffmpeg -loglevel error -y -f concat -safe 0 -i "$DIR/list.txt" -c copy -movflags +faststart "$OUT"
FRAMES=$(ffprobe -v error -count_frames -select_streams v:0 -show_entries stream=nb_read_frames -of default=nw=1:nk=1 "$OUT")
DIM=$(ffprobe -v error -select_streams v:0 -show_entries stream=width,height -of csv=p=0:s=x "$OUT" | tr -d "," | sed "s/x$//")
log VERIFY "out=$OUT frames=$FRAMES expected=$TOTAL size=$DIM"
if [ "$FRAMES" != "$TOTAL" ]; then
	log ERROR "frame_count_mismatch frames=$FRAMES expected=$TOTAL"
	exit 1
fi
log DONE "out=$OUT"
