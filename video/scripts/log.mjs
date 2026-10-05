// VERSION: v1.0.0 — 2026-10-05 — 구조화 로그 헬퍼 (스크립트 공용)
// 형식: [STAGE] [TIMESTAMP] key=value key=value ...
// LOG_LEVEL 환경변수(DEBUG/INFO/WARNING)로 출력 수준을 조절한다.

const LEVELS = {DEBUG: 10, INFO: 20, WARNING: 30, ERROR: 40};
const threshold = LEVELS[(process.env.LOG_LEVEL ?? 'INFO').toUpperCase()] ?? 20;

const fmt = (v) => {
	if (typeof v === 'string') return /\s/.test(v) ? JSON.stringify(v) : v;
	return JSON.stringify(v);
};

const emit = (level, stage, fields) => {
	if (LEVELS[level] < threshold) return;
	const kv = Object.entries(fields)
		.map(([k, v]) => `${k}=${fmt(v)}`)
		.join(' ');
	const line = `[${stage}] [${new Date().toISOString()}] level=${level} ${kv}`;
	if (LEVELS[level] >= LEVELS.WARNING) console.error(line);
	else console.log(line);
};

export const log = {
	debug: (stage, fields = {}) => emit('DEBUG', stage, fields),
	info: (stage, fields = {}) => emit('INFO', stage, fields),
	warn: (stage, fields = {}) => emit('WARNING', stage, fields),
	error: (stage, fields = {}) => emit('ERROR', stage, fields),
};
