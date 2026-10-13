import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join, normalize, extname, sep } from 'node:path';
import { openDb, makeStore } from './db.js';
import { Agent, KINDS, dayKey } from './agent.js';
import { buildSpec } from './openapi.js';

const PUBLIC = join(dirname(fileURLToPath(import.meta.url)), 'public');
const NAMES = ['Ava', 'Ben', 'Chloe', 'Dan', 'Eva', 'Finn', 'Gia', 'Hugo', 'Iris', 'Jack', 'Kira', 'Leo', 'Mia', 'Noah', 'Opal', 'Pete'];
const COLORS = ['#e76f51', '#2a9d8f', '#e9c46a', '#6d6fd8', '#f4a261', '#43aa8b', '#d65db1', '#4ea8de'];
const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.css': 'text/css', '.svg': 'image/svg+xml' };

class HttpError extends Error {
  constructor(status, message) { super(message); this.status = status; }
}

const str = (v, f, max, req = false) => {
  if (v === undefined || v === null) { if (req) throw new HttpError(400, `${f} is required`); return undefined; }
  if (typeof v !== 'string') throw new HttpError(400, `${f} must be a string`);
  if (req && !v.trim()) throw new HttpError(400, `${f} is required`);
  if (v.length > max) throw new HttpError(400, `${f} is too long`);
  return v;
};

/** Validates a task payload merged over `base` (the existing task, for updates). */
function buildTask(body, base, store) {
  if (!body || typeof body !== 'object' || Array.isArray(body)) throw new HttpError(400, 'Body must be a JSON object');
  const b = { ...base, ...body };
  const t = {
    name: str(b.name, 'name', 100, true).trim(),
    description: str(b.description ?? '', 'description', 500),
    kind: b.kind,
    config: b.config ?? {},
    time: b.time,
    days: b.days ?? [0, 1, 2, 3, 4, 5, 6],
    enabled: b.enabled ?? true,
  };
  if (!KINDS.includes(t.kind)) throw new HttpError(400, `kind must be one of: ${KINDS.join(', ')}`);
  if (typeof t.time !== 'string' || !/^([01]\d|2[0-3]):[0-5]\d$/.test(t.time)) throw new HttpError(400, 'time must be HH:MM (24h)');
  if (!Array.isArray(t.days) || !t.days.length || !t.days.every((d) => Number.isInteger(d) && d >= 0 && d <= 6))
    throw new HttpError(400, 'days must be a non-empty array of integers 0-6');
  t.days = [...new Set(t.days)].sort();
  if (typeof t.enabled !== 'boolean') throw new HttpError(400, 'enabled must be a boolean');
  if (!t.config || typeof t.config !== 'object' || Array.isArray(t.config)) throw new HttpError(400, 'config must be an object');
  if (JSON.stringify(t.config).length > 10000) throw new HttpError(400, 'config is too large');
  if (t.kind === 'http') {
    let u;
    try { u = new URL(t.config.url); } catch { throw new HttpError(400, 'config.url must be a valid URL'); }
    if (!/^https?:$/.test(u.protocol)) throw new HttpError(400, 'config.url must be http(s)');
  }
  if (t.kind === 'claude') str(t.config.prompt, 'config.prompt', 5000, true);
  if (t.kind === 'reminder') str(t.config.message ?? '', 'config.message', 1000);

  const n = store.listTasks().length;
  t.person_name = str(b.person_name, 'person_name', 30)?.trim() || NAMES[n % NAMES.length];
  t.person_color = b.person_color ?? COLORS[n % COLORS.length];
  if (typeof t.person_color !== 'string' || !/^#[0-9a-fA-F]{6}$/.test(t.person_color)) throw new HttpError(400, 'person_color must be #rrggbb');
  return t;
}

export function createApp({ dbPath = process.env.DB_PATH || './data/office.db', tickMs } = {}) {
  const db = openDb(dbPath);
  const store = makeStore(db);
  const agent = new Agent(store, { tickMs });
  const spec = buildSpec(KINDS);
  const clients = new Set();

  agent.on('change', (e) => {
    for (const res of clients) res.write(`event: change\ndata: ${JSON.stringify(e)}\n\n`);
  });
  const notify = (e) => agent.emit('change', e);

  const taskOr404 = (id) => store.getTask(id) ?? (() => { throw new HttpError(404, 'Task not found'); })();

  async function api(req, res, url) {
    const p = url.pathname;
    const m = req.method;
    let match;

    if (p === '/api/events' && m === 'GET') {
      res.writeHead(200, { 'content-type': 'text/event-stream', 'cache-control': 'no-cache', connection: 'keep-alive' });
      res.write('retry: 3000\n\n');
      clients.add(res);
      req.on('close', () => clients.delete(res));
      return;
    }
    if (p === '/api/tasks') {
      if (m === 'GET') return send(res, 200, store.listTasks());
      if (m === 'POST') {
        const t = buildTask(await readBody(req), {}, store);
        const now = new Date();
        const hhmm = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
        // A task created after today's slot waits for tomorrow instead of firing instantly.
        t.last_slot = hhmm >= t.time ? dayKey(now) : null;
        const created = store.insertTask(t);
        notify({ type: 'task', task_id: created.id });
        return send(res, 201, created);
      }
    }
    if ((match = p.match(/^\/api\/tasks\/(\d+)$/))) {
      const id = Number(match[1]);
      if (m === 'GET') return send(res, 200, taskOr404(id));
      if (m === 'PUT' || m === 'PATCH') {
        const old = taskOr404(id);
        const t = buildTask(await readBody(req), old, store);
        const updated = store.updateTask(id, t);
        if (updated.time !== old.time) store.markSlot(id, null);
        notify({ type: 'task', task_id: id });
        return send(res, 200, store.getTask(id));
      }
      if (m === 'DELETE') {
        taskOr404(id);
        store.deleteTask(id);
        notify({ type: 'task', task_id: id });
        res.writeHead(204).end();
        return;
      }
    }
    if ((match = p.match(/^\/api\/tasks\/(\d+)\/run$/)) && m === 'POST') {
      const id = Number(match[1]);
      taskOr404(id);
      if (agent.running.has(id)) throw new HttpError(409, 'Task is already running');
      agent.run(id, 'manual').catch(() => {});
      return send(res, 202, { started: true });
    }
    if ((match = p.match(/^\/api\/tasks\/(\d+)\/runs$/)) && m === 'GET') {
      const id = Number(match[1]);
      taskOr404(id);
      return send(res, 200, store.listRuns(id));
    }
    if (p === '/api/runs' && m === 'GET') {
      const limit = Math.min(Math.max(parseInt(url.searchParams.get('limit')) || 50, 1), 200);
      return send(res, 200, store.listRuns(null, limit));
    }
    throw new HttpError(404, 'Not found');
  }

  const server = createServer(async (req, res) => {
    try {
      const url = new URL(req.url, 'http://x');
      if (url.pathname === '/openapi.json') return send(res, 200, spec);
      if (url.pathname.startsWith('/api/')) return await api(req, res, url);
      if (req.method !== 'GET') throw new HttpError(405, 'Method not allowed');
      const file = url.pathname === '/' ? 'index.html' : url.pathname === '/docs' ? 'docs.html' : url.pathname.slice(1);
      const full = normalize(join(PUBLIC, file));
      if (!full.startsWith(PUBLIC + sep)) throw new HttpError(404, 'Not found');
      const body = await readFile(full).catch(() => { throw new HttpError(404, 'Not found'); });
      res.writeHead(200, { 'content-type': MIME[extname(full)] || 'application/octet-stream' }).end(body);
    } catch (e) {
      if (!(e instanceof HttpError)) console.error(e);
      if (res.headersSent) return res.end();
      send(res, e instanceof HttpError ? e.status : 500, { error: e instanceof HttpError ? e.message : 'Internal error' });
    }
  });

  return {
    server, agent, store,
    close() { agent.stop(); for (const c of clients) c.end(); server.close(); server.closeAllConnections?.(); db.close(); },
  };
}

function send(res, status, data) {
  res.writeHead(status, { 'content-type': 'application/json' }).end(JSON.stringify(data));
}

async function readBody(req) {
  let size = 0;
  const chunks = [];
  for await (const c of req) {
    size += c.length;
    if (size > 100_000) throw new HttpError(413, 'Body too large');
    chunks.push(c);
  }
  try { return JSON.parse(Buffer.concat(chunks).toString() || 'null'); } catch { throw new HttpError(400, 'Invalid JSON'); }
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const port = Number(process.env.PORT) || 3000;
  const host = process.env.HOST || '127.0.0.1';
  const app = createApp();
  app.agent.start();
  app.server.listen(port, host, () => console.log(`Task Office running at http://${host}:${port}  (API docs: /docs)`));
}
