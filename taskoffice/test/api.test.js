import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createApp } from '../server.js';
import { isDue } from '../agent.js';

async function boot() {
  const app = createApp({ dbPath: ':memory:' });
  await new Promise((r) => app.server.listen(0, '127.0.0.1', r));
  const base = `http://127.0.0.1:${app.server.address().port}`;
  const call = async (method, path, body) => {
    const res = await fetch(base + path, { method, headers: { 'content-type': 'application/json' }, body: body && JSON.stringify(body) });
    return { status: res.status, body: res.status === 204 ? null : await res.json() };
  };
  return { app, call };
}

test('create, read, update, run, delete a task', async () => {
  const { app, call } = await boot();
  try {
    let r = await call('POST', '/api/tasks', { name: 'Stretch', kind: 'reminder', time: '09:00', config: { message: 'stretch' } });
    assert.equal(r.status, 201);
    const id = r.body.id;
    assert.ok(r.body.person_name && r.body.person_color);

    r = await call('PATCH', `/api/tasks/${id}`, { time: '10:15', enabled: false });
    assert.equal(r.body.time, '10:15');
    assert.equal(r.body.enabled, false);

    assert.equal((await call('POST', `/api/tasks/${id}/run`)).status, 202);
    await new Promise((r) => setTimeout(r, 50));
    r = await call('GET', `/api/tasks/${id}/runs`);
    assert.equal(r.body[0].status, 'done');
    assert.equal(r.body[0].output, 'stretch');
    assert.equal((await call('GET', '/api/runs')).body[0].task_name, 'Stretch');

    assert.equal((await call('DELETE', `/api/tasks/${id}`)).status, 204);
    assert.equal((await call('GET', `/api/tasks/${id}`)).status, 404);
  } finally { app.close(); }
});

test('validation rejects bad input', async () => {
  const { app, call } = await boot();
  try {
    for (const bad of [
      { name: 'x', kind: 'nope', time: '09:00' },
      { name: 'x', kind: 'reminder', time: '9am' },
      { name: '', kind: 'reminder', time: '09:00' },
      { name: 'x', kind: 'reminder', time: '09:00', days: [9] },
      { name: 'x', kind: 'http', time: '09:00', config: { url: 'file:///etc/passwd' } },
      { name: 'x', kind: 'claude', time: '09:00', config: {} },
    ]) assert.equal((await call('POST', '/api/tasks', bad)).status, 400, JSON.stringify(bad));
  } finally { app.close(); }
});

test('failed runs are recorded as errors', async () => {
  const { app, call } = await boot();
  try {
    const { body } = await call('POST', '/api/tasks', { name: 'dead', kind: 'http', time: '09:00', config: { url: 'http://127.0.0.1:1/' } });
    await call('POST', `/api/tasks/${body.id}/run`);
    await new Promise((r) => setTimeout(r, 300));
    assert.equal((await call('GET', `/api/tasks/${body.id}`)).body.status, 'error');
  } finally { app.close(); }
});

test('scheduler fires once per day at the task time', () => {
  const t = { enabled: true, days: [0, 1, 2, 3, 4, 5, 6], time: '09:00', last_slot: null };
  const at = (h, m) => new Date(2026, 9, 13, h, m);
  assert.equal(isDue(t, at(8, 59)), false);
  assert.equal(isDue(t, at(9, 0)), true);
  assert.equal(isDue({ ...t, last_slot: '2026-10-13' }, at(9, 30)), false);
  assert.equal(isDue({ ...t, enabled: false }, at(9, 30)), false);
  assert.equal(isDue({ ...t, days: [(at(9, 0).getDay() + 1) % 7] }, at(9, 30)), false);
});

test('agent tick runs due tasks', async () => {
  const { app, call } = await boot();
  try {
    const { body } = await call('POST', '/api/tasks', { name: 'a', kind: 'reminder', time: '00:00' });
    // created after 00:00 today, so it waits for tomorrow
    assert.equal(body.last_slot !== null, true);
    app.agent.tick(new Date(Date.now() + 86400_000));
    await new Promise((r) => setTimeout(r, 50));
    assert.equal((await call('GET', `/api/tasks/${body.id}/runs`)).body.length, 1);
  } finally { app.close(); }
});
