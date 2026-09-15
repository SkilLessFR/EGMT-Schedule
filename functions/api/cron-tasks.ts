// functions/api/cron-tasks.ts
// Alias endpoint for automated Cron services (e.g. cron-job.org, Cloudflare Cron, GitHub Actions)
export { onRequestGet, onRequestPost, onRequestOptions } from './send-task-push';
