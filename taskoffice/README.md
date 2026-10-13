# Task Office

A daily task agent with a "god view" mini-office UI. Every task is a person sitting at a desk:
they type when the task runs, show ✓ when it succeeds, ! when it fails, and go home when disabled.

- **Agent** (`agent.js`): checks every 15 s and runs each enabled task once a day at its time (server local time).
- **Store** (`db.js`): SQLite via Node's built-in `node:sqlite` — saves routine tasks and run history in `data/office.db`.
- **API** (`server.js`, `openapi.js`): REST + OpenAPI 3 spec at `/openapi.json`, docs at `/docs`.
- **UI** (`public/index.html`): drag to pan, wheel to zoom, click a person to select, add/edit tasks in the side panel.

## Run

Requires Node ≥ 22.13 (no npm install needed).

```
cd taskoffice
npm start          # http://127.0.0.1:3000
npm test
```

Env: `PORT`, `HOST` (default `127.0.0.1`), `DB_PATH`, `ANTHROPIC_API_KEY` / `CLAUDE_MODEL` (for `claude` tasks).

## Task kinds

| kind | config | does |
|---|---|---|
| `reminder` | `{message}` | logs the message |
| `http` | `{url, method?, headers?, body?}` | calls the URL; non-2xx = error |
| `claude` | `{prompt}` | sends the prompt to Claude and stores the reply |

## API

```
curl -X POST localhost:3000/api/tasks -H 'content-type: application/json' \
  -d '{"name":"Stand-up","kind":"reminder","time":"09:00","config":{"message":"Stand up!"}}'
```

`GET/POST /api/tasks`, `GET/PUT/PATCH/DELETE /api/tasks/{id}`, `POST /api/tasks/{id}/run`,
`GET /api/tasks/{id}/runs`, `GET /api/runs`, `GET /api/events` (SSE).

Notes: there is no authentication, so keep it on localhost (or behind your own proxy). A task created after
today's time waits until tomorrow; if the server was down at the scheduled time, the task runs when it comes back up that day.
