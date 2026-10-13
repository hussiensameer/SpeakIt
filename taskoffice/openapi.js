// OpenAPI 3.0 description of the REST API, served at /openapi.json.
const ref = (n) => ({ $ref: `#/components/schemas/${n}` });
const json = (schema) => ({ 'application/json': { schema } });
const idParam = { name: 'id', in: 'path', required: true, schema: { type: 'integer' } };
const err = { description: 'Error', content: json(ref('Error')) };

export function buildSpec(kinds) {
  return {
    openapi: '3.0.3',
    info: {
      title: 'Task Office API',
      version: '1.0.0',
      description: 'Save routine tasks and create new ones. Each task is a person in the office; the agent runs it daily at its time.',
    },
    paths: {
      '/api/tasks': {
        get: { summary: 'List tasks', operationId: 'listTasks', responses: { 200: { description: 'All tasks', content: json({ type: 'array', items: ref('Task') }) } } },
        post: {
          summary: 'Create a task', operationId: 'createTask',
          requestBody: { required: true, content: json(ref('TaskInput')) },
          responses: { 201: { description: 'Created', content: json(ref('Task')) }, 400: err },
        },
      },
      '/api/tasks/{id}': {
        parameters: [idParam],
        get: { summary: 'Get a task', operationId: 'getTask', responses: { 200: { description: 'The task', content: json(ref('Task')) }, 404: err } },
        put: {
          summary: 'Update a task (send only the fields to change)', operationId: 'updateTask',
          requestBody: { required: true, content: json(ref('TaskInput')) },
          responses: { 200: { description: 'Updated', content: json(ref('Task')) }, 400: err, 404: err },
        },
        delete: { summary: 'Delete a task and its history', operationId: 'deleteTask', responses: { 204: { description: 'Deleted' }, 404: err } },
      },
      '/api/tasks/{id}/run': {
        parameters: [idParam],
        post: { summary: 'Run a task now', operationId: 'runTask', responses: { 202: { description: 'Started' }, 404: err, 409: err } },
      },
      '/api/tasks/{id}/runs': {
        parameters: [idParam],
        get: { summary: 'Run history of a task', operationId: 'listTaskRuns', responses: { 200: { description: 'Newest first', content: json({ type: 'array', items: ref('Run') }) }, 404: err } },
      },
      '/api/runs': {
        get: {
          summary: 'Recent runs across all tasks', operationId: 'listRuns',
          parameters: [{ name: 'limit', in: 'query', schema: { type: 'integer', default: 50, maximum: 200 } }],
          responses: { 200: { description: 'Newest first', content: json({ type: 'array', items: ref('Run') }) } },
        },
      },
      '/api/events': {
        get: { summary: 'Server-sent events: a "change" event whenever a task starts or finishes', operationId: 'events', responses: { 200: { description: 'text/event-stream' } } },
      },
    },
    components: {
      schemas: {
        Error: { type: 'object', properties: { error: { type: 'string' } } },
        TaskInput: {
          type: 'object',
          required: ['name', 'kind', 'time'],
          properties: {
            name: { type: 'string', example: 'Morning standup reminder' },
            description: { type: 'string' },
            kind: { type: 'string', enum: kinds, description: 'reminder: logs `message`. http: calls `url`. claude: sends `prompt` to Claude (needs ANTHROPIC_API_KEY).' },
            config: {
              type: 'object', description: 'Depends on kind',
              example: { message: 'Stand up!' },
              properties: {
                message: { type: 'string' }, url: { type: 'string' }, method: { type: 'string' },
                headers: { type: 'object', additionalProperties: { type: 'string' } }, body: { type: 'string' }, prompt: { type: 'string' },
              },
            },
            time: { type: 'string', pattern: '^([01]\\d|2[0-3]):[0-5]\\d$', example: '09:00', description: 'Daily run time, server local time (HH:MM)' },
            days: { type: 'array', items: { type: 'integer', minimum: 0, maximum: 6 }, default: [0, 1, 2, 3, 4, 5, 6], description: '0 = Sunday' },
            enabled: { type: 'boolean', default: true },
            person_name: { type: 'string', description: 'Who sits at this desk (auto-assigned if omitted)' },
            person_color: { type: 'string', pattern: '^#[0-9a-fA-F]{6}$' },
          },
        },
        Task: {
          allOf: [ref('TaskInput'), {
            type: 'object',
            properties: {
              id: { type: 'integer' },
              status: { type: 'string', enum: ['idle', 'working', 'done', 'error'] },
              last_run_at: { type: 'string', format: 'date-time', nullable: true },
              created_at: { type: 'string', format: 'date-time' },
            },
          }],
        },
        Run: {
          type: 'object',
          properties: {
            id: { type: 'integer' }, task_id: { type: 'integer' },
            status: { type: 'string', enum: ['running', 'done', 'error'] },
            trigger: { type: 'string', enum: ['schedule', 'manual'] },
            started_at: { type: 'string', format: 'date-time' },
            finished_at: { type: 'string', format: 'date-time', nullable: true },
            output: { type: 'string' },
            task_name: { type: 'string', description: 'Only on /api/runs' },
            person_name: { type: 'string', description: 'Only on /api/runs' },
          },
        },
      },
    },
  };
}
