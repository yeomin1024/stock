// 자동 생성 파일 — 직접 수정하지 말 것. 생성: node scripts/prepare-audio.mjs
// 원본: input/narration.*, input/bgm.* → public/ 사본을 staticFile() 로 읽는다. null 이면 그 트랙 없음.
export type AudioTrack = {readonly file: string; readonly durationSec: number};
export const NARRATION: AudioTrack | null = {"file":"narration.mp3","durationSec":486.456};
export const BGM: (AudioTrack & {readonly volumeDb: number}) | null = null;
