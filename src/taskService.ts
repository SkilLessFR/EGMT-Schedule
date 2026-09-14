// src/taskService.ts
/**
 * Real-time synchronization service for recurring tasks.
 * Stores tasks persistently in Cloudflare KV (site backend)
 * with optimistic local storage caching for offline support.
 */

export type ScheduleType = 'once' | 'daily' | 'weekly';

export interface Task {
  id: string;
  title: string;
  times: string[];
  scheduleType: ScheduleType;
  daysOfWeek?: number[];
  dateCreated?: string;
}

export interface TimeSelection {
  hour: string;
  minute: string;
}

export const TASKS_STORAGE_KEY = 'work-schedule-site-tasks';
export const TASKS_UPDATED_AT_KEY = 'work-schedule-tasks-updated-at';

export function loadInitialTasks(): Task[] | null {
  try {
    const raw = localStorage.getItem(TASKS_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed) && parsed.length > 0) {
      return parsed;
    }
  } catch {
    // Ignore invalid JSON
  }
  return null;
}

export function saveTasksLocal(tasks: Task[]): void {
  try {
    localStorage.setItem(TASKS_STORAGE_KEY, JSON.stringify(tasks));
  } catch {
    // Ignore storage quota errors
  }
}

export async function fetchTasksFromCloud(): Promise<{ tasks: Task[]; updatedAt: string; source: string }> {
  try {
    const res = await fetch('/api/tasks', {
      headers: { 'Accept': 'application/json' },
      cache: 'no-store',
    });

    if (res.ok) {
      const data = await res.json();
      if (Array.isArray(data.tasks)) {
        saveTasksLocal(data.tasks);
        if (data.updatedAt) {
          localStorage.setItem(TASKS_UPDATED_AT_KEY, data.updatedAt);
        }
        return {
          tasks: data.tasks,
          updatedAt: data.updatedAt || '',
          source: data.source || 'kv',
        };
      }
    }
  } catch (err) {
    console.warn('Could not fetch tasks from Cloudflare KV, falling back to static/local:', err);
  }

  // Fallback 1: Local storage
  const local = loadInitialTasks();
  if (local && local.length > 0) {
    return {
      tasks: local,
      updatedAt: localStorage.getItem(TASKS_UPDATED_AT_KEY) || '',
      source: 'local',
    };
  }

  // Fallback 2: Static tasks.json bundle file
  try {
    const baseRes = await fetch('/tasks.json', { cache: 'no-store' });
    if (baseRes.ok) {
      const baseTasks = (await baseRes.json()) as Task[];
      saveTasksLocal(baseTasks);
      return {
        tasks: baseTasks,
        updatedAt: '',
        source: 'static-file',
      };
    }
  } catch {
    // Ignore
  }

  return {
    tasks: [],
    updatedAt: '',
    source: 'empty',
  };
}

export async function saveTasksToCloud(
  tasks: Task[],
  updatedBy?: string
): Promise<{ success: boolean; updatedAt: string }> {
  // 1. Optimistically persist locally
  saveTasksLocal(tasks);

  // 2. Dispatch to Cloudflare KV site storage
  try {
    const res = await fetch('/api/tasks', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        tasks,
        updatedBy: updatedBy || 'user',
      }),
    });

    if (res.ok) {
      const data = await res.json();
      const updatedAt = data.updatedAt || new Date().toISOString();
      localStorage.setItem(TASKS_UPDATED_AT_KEY, updatedAt);
      return { success: true, updatedAt };
    }
  } catch (err) {
    console.warn('Network task sync failed, stored locally:', err);
  }

  return {
    success: false,
    updatedAt: new Date().toISOString(),
  };
}

/**
 * Downloads tasks as a JSON file upon explicit user request.
 */
export function exportTasksAsJsonFile(tasks: Task[]): void {
  try {
    const jsonString = JSON.stringify(tasks, null, 2);
    const blob = new Blob([jsonString], { type: 'application/json' });
    const url = URL.createObjectURL(blob);

    const tempLink = document.createElement('a');
    tempLink.href = url;
    tempLink.download = 'tasks.json';
    document.body.appendChild(tempLink);
    tempLink.click();

    document.body.removeChild(tempLink);
    URL.revokeObjectURL(url);
  } catch (error) {
    console.error('Failed to export tasks backup file:', error);
  }
}
