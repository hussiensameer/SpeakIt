// The agent: executes tasks and fires them every day at their scheduled time.
import { EventEmitter } from 'node:events';

const OUT_MAX = 4000;
const clip = (s) => (s.length > OUT_MAX ? s.slice(0, OUT_MAX) + '…' : s);

// Each runner returns the text output, or throws to mark the run as failed.
export const runners = {
  async reminder(cfg) {
    return cfg.message || 'Reminder delivered.';
  },
  async http(cfg) {
    const res = await fetch(cfg.url, {
      method: cfg.method || 'GET',
      headers: cfg.headers,
      body: cfg.body,
      signal: AbortSignal.timeout(15000),
    });
    const text = await res.text();
    const line = `${res.status} ${res.statusText}\n${text}`;
    if (!res.ok) throw new Error(line);
    return line;
  },
  async claude(cfg) {
    const key = process.env.ANTHROPIC_API_KEY;
    if (!key) throw new Error('ANTHROPIC_API_KEY is not set on the server');
    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'content-type': 'application/json', 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({
        model: process.env.CLAUDE_MODEL || 'claude-sonnet-5-5',
        max_tokens: 1024,
        messages: [{ role: 'user', content: cfg.prompt }],
      }),
      signal: AbortSignal.timeout(60000),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(`Anthropic API ${res.status}: ${data?.error?.message ?? 'request failed'}`);
    return (data.content ?? []).filter((b) => b.type === 'text').map((b) => b.text).join('\n');
  },
};

export const KINDS = Object.keys(runners);

const pad = (n) => String(n).padStart(2, '0');
export const dayKey = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

/** Is a task's daily slot due at `now` (local time) and not yet run today? */
export function isDue(task, now = new Date()) {
  if (!task.enabled || !task.days.includes(now.getDay())) return false;
  if (task.last_slot === dayKey(now)) return false;
  return `${pad(now.getHours())}:${pad(now.getMinutes())}` >= task.time;
}

export class Agent extends EventEmitter {
  constructor(store, { tickMs = 15000 } = {}) {
    super();
    this.store = store;
    this.tickMs = tickMs;
    this.running = new Set();
  }

  start() {
    this.timer = setInterval(() => this.tick(), this.tickMs);
    this.timer.unref();
    this.tick();
  }

  stop() { clearInterval(this.timer); }

  tick(now = new Date()) {
    for (const task of this.store.listTasks()) {
      if (isDue(task, now) && !this.running.has(task.id)) {
        this.store.markSlot(task.id, dayKey(now));
        this.run(task.id, 'schedule');
      }
    }
  }

  /** Runs a task now. Resolves with the finished run; rejects if the task is missing or busy. */
  async run(id, trigger = 'manual') {
    const task = this.store.getTask(id);
    if (!task) throw Object.assign(new Error('Task not found'), { code: 'NOT_FOUND' });
    if (this.running.has(id)) throw Object.assign(new Error('Task is already running'), { code: 'BUSY' });
    this.running.add(id);
    const runId = this.store.startRun(id, trigger);
    this.store.setStatus(id, 'working');
    this.emit('change', { type: 'start', task_id: id, run_id: runId });
    let status = 'done', output;
    try {
      output = clip(String(await runners[task.kind](task.config)));
    } catch (e) {
      status = 'error';
      output = clip(e?.message ?? String(e));
    }
    this.store.finishRun(runId, status, output);
    this.store.finishTask(id, status);
    this.running.delete(id);
    this.emit('change', { type: 'finish', task_id: id, run_id: runId, status });
    return this.store.getRun(runId);
  }
}
