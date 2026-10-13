// SQLite store ("apilite"): saves routine tasks and their run history.
import { DatabaseSync } from 'node:sqlite';
import { mkdirSync } from 'node:fs';
import { dirname } from 'node:path';

export function openDb(path) {
  if (path !== ':memory:') mkdirSync(dirname(path), { recursive: true });
  const db = new DatabaseSync(path);
  db.exec(`
    PRAGMA journal_mode = WAL;
    CREATE TABLE IF NOT EXISTS tasks (
      id           INTEGER PRIMARY KEY AUTOINCREMENT,
      name         TEXT NOT NULL,
      description  TEXT NOT NULL DEFAULT '',
      kind         TEXT NOT NULL,
      config       TEXT NOT NULL DEFAULT '{}',
      time         TEXT NOT NULL,
      days         TEXT NOT NULL DEFAULT '0,1,2,3,4,5,6',
      enabled      INTEGER NOT NULL DEFAULT 1,
      person_name  TEXT NOT NULL,
      person_color TEXT NOT NULL,
      status       TEXT NOT NULL DEFAULT 'idle',
      last_slot    TEXT,
      last_run_at  TEXT,
      created_at   TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS runs (
      id          INTEGER PRIMARY KEY AUTOINCREMENT,
      task_id     INTEGER NOT NULL REFERENCES tasks(id) ON DELETE CASCADE,
      started_at  TEXT NOT NULL,
      finished_at TEXT,
      status      TEXT NOT NULL,
      output      TEXT NOT NULL DEFAULT '',
      trigger     TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS runs_task ON runs(task_id, id DESC);
  `);
  db.exec('PRAGMA foreign_keys = ON');
  // A crash mid-run must not leave people "working" forever.
  db.exec(`UPDATE runs SET status='error', finished_at=datetime('now'), output='interrupted by restart' WHERE status='running'`);
  db.exec(`UPDATE tasks SET status='idle' WHERE status='working'`);
  return db;
}

const parseTask = (r) => r && { ...r, config: JSON.parse(r.config), days: r.days.split(',').filter(Boolean).map(Number), enabled: !!r.enabled };

export function makeStore(db) {
  const q = (sql) => db.prepare(sql);
  return {
    listTasks: () => q('SELECT * FROM tasks ORDER BY id').all().map(parseTask),
    getTask: (id) => parseTask(q('SELECT * FROM tasks WHERE id=?').get(id)),
    insertTask(t) {
      const r = q(`INSERT INTO tasks (name,description,kind,config,time,days,enabled,person_name,person_color,last_slot,created_at)
        VALUES (?,?,?,?,?,?,?,?,?,?,?)`).run(t.name, t.description, t.kind, JSON.stringify(t.config), t.time, t.days.join(','),
        t.enabled ? 1 : 0, t.person_name, t.person_color, t.last_slot ?? null, new Date().toISOString());
      return this.getTask(Number(r.lastInsertRowid));
    },
    updateTask(id, t) {
      q(`UPDATE tasks SET name=?,description=?,kind=?,config=?,time=?,days=?,enabled=?,person_name=?,person_color=? WHERE id=?`)
        .run(t.name, t.description, t.kind, JSON.stringify(t.config), t.time, t.days.join(','), t.enabled ? 1 : 0, t.person_name, t.person_color, id);
      return this.getTask(id);
    },
    deleteTask: (id) => q('DELETE FROM tasks WHERE id=?').run(id).changes > 0,
    setStatus: (id, status) => q('UPDATE tasks SET status=? WHERE id=?').run(status, id),
    markSlot: (id, slot) => q('UPDATE tasks SET last_slot=? WHERE id=?').run(slot, id),
    finishTask: (id, status) => q('UPDATE tasks SET status=?, last_run_at=? WHERE id=?').run(status, new Date().toISOString(), id),
    startRun: (taskId, trigger) =>
      Number(q('INSERT INTO runs (task_id,started_at,status,trigger) VALUES (?,?,?,?)').run(taskId, new Date().toISOString(), 'running', trigger).lastInsertRowid),
    finishRun: (id, status, output) =>
      q('UPDATE runs SET status=?, output=?, finished_at=? WHERE id=?').run(status, output, new Date().toISOString(), id),
    getRun: (id) => q('SELECT * FROM runs WHERE id=?').get(id),
    listRuns: (taskId, limit = 50) =>
      (taskId
        ? q('SELECT * FROM runs WHERE task_id=? ORDER BY id DESC LIMIT ?').all(taskId, limit)
        : q('SELECT r.*, t.name AS task_name, t.person_name FROM runs r JOIN tasks t ON t.id=r.task_id ORDER BY r.id DESC LIMIT ?').all(limit)),
  };
}
