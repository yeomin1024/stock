// VERSION: v1.0.0 — 2026-11-02 — Initial implementation: structured stage logger
/**
 * Structured logger: `[STAGE] [ISO-TIMESTAMP] level=INFO event=name key=value ...`
 * Level is adjustable without code changes via the URL: ?log=debug | info | warn | error.
 */
export type LogLevel = 'DEBUG' | 'INFO' | 'WARN' | 'ERROR';
export type LogStage = 'BOOT' | 'CONFIG' | 'LAYOUT' | 'VALIDATE' | 'BUILD' | 'RENDER' | 'CAMERA' | 'UI' | 'DEMO' | 'PERF';
export type LogFields = Record<string, string | number | boolean | null | undefined | readonly (string | number)[]>;

const ORDER: Record<LogLevel, number> = { DEBUG: 10, INFO: 20, WARN: 30, ERROR: 40 };

export interface Logger {
  level: LogLevel;
  debug(stage: LogStage, event: string, fields?: LogFields): void;
  info(stage: LogStage, event: string, fields?: LogFields): void;
  warn(stage: LogStage, event: string, fields?: LogFields): void;
  error(stage: LogStage, event: string, fields?: LogFields): void;
}

function formatValue(v: LogFields[string]): string {
  if (v === null || v === undefined) return String(v);
  if (Array.isArray(v)) return v.map((x) => formatValue(x as string | number)).join(',');
  if (typeof v === 'number') return Number.isInteger(v) ? String(v) : String(Math.round(v * 1000) / 1000);
  const s = String(v);
  return /[\s"=]/.test(s) ? JSON.stringify(s) : s;
}

export function formatLine(level: LogLevel, stage: LogStage, event: string, fields: LogFields = {}, now = new Date()): string {
  const kv = Object.entries(fields)
    .map(([k, v]) => `${k}=${formatValue(v)}`)
    .join(' ');
  return `[${stage}] [${now.toISOString()}] level=${level} event=${event}${kv ? ' ' + kv : ''}`;
}

export function parseLogLevel(raw: string | null | undefined, fallback: LogLevel = 'INFO'): LogLevel {
  const up = (raw ?? '').toUpperCase();
  return up in ORDER ? (up as LogLevel) : fallback;
}

export function createLogger(level: LogLevel = 'INFO', sink: Pick<Console, 'debug' | 'info' | 'warn' | 'error'> = console): Logger {
  const logger: Logger = {
    level,
    debug: (s, e, f) => emit('DEBUG', s, e, f),
    info: (s, e, f) => emit('INFO', s, e, f),
    warn: (s, e, f) => emit('WARN', s, e, f),
    error: (s, e, f) => emit('ERROR', s, e, f),
  };
  function emit(lv: LogLevel, stage: LogStage, event: string, fields?: LogFields): void {
    if (ORDER[lv] < ORDER[logger.level]) return;
    const line = formatLine(lv, stage, event, fields);
    if (lv === 'DEBUG') sink.debug(line);
    else if (lv === 'INFO') sink.info(line);
    else if (lv === 'WARN') sink.warn(line);
    else sink.error(line);
  }
  return logger;
}
