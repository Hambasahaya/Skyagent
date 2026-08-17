const DEFAULT_API_BASE_URL = import.meta.env.VITE_API_BASE_URL ?? "http://localhost:8080";
const API_BASE_URL_KEY = "skyagen_api_base_url";
const TOKEN_KEY = "skyagen_access_token";
const API_DISCONNECTED_AT_KEY = "skyagen_api_disconnected_at";

export type ApiList<T> = { data: T[]; page?: number; limit?: number; total?: number };

export function getApiBaseUrl() {
  return localStorage.getItem(API_BASE_URL_KEY) || DEFAULT_API_BASE_URL;
}

export function setApiBaseUrl(url: string) {
  const normalized = url.trim().replace(/\/+$/, "");
  localStorage.setItem(API_BASE_URL_KEY, normalized || DEFAULT_API_BASE_URL);
  localStorage.removeItem(API_DISCONNECTED_AT_KEY);
}

export function notifyApiDisconnected() {
  localStorage.setItem(API_DISCONNECTED_AT_KEY, new Date().toISOString());
  window.dispatchEvent(new CustomEvent("skyagen:api-disconnected", { detail: { baseUrl: getApiBaseUrl() } }));
}

export async function checkApiConnection() {
  const controller = new AbortController();
  const timeout = window.setTimeout(() => controller.abort(), 3500);
  try {
    const url = new URL("/api/dashboard/summary", getApiBaseUrl());
    const response = await fetch(url, { method: "GET", cache: "no-store", headers: authHeaders(false), signal: controller.signal });
    if (!response.ok && response.status !== 401 && response.status !== 403) throw new Error(`API health check failed with status ${response.status}`);
    return true;
  } catch {
    notifyApiDisconnected();
    return false;
  } finally {
    window.clearTimeout(timeout);
  }
}
export function getToken() {
  return localStorage.getItem(TOKEN_KEY) ?? "";
}

export function setToken(token: string) {
  localStorage.setItem(TOKEN_KEY, token);
}

export function clearToken() {
  localStorage.removeItem(TOKEN_KEY);
}

function authHeaders(json = true) {
  const token = getToken();
  const headers: Record<string, string> = {};
  if (json) headers["Content-Type"] = "application/json";
  if (token) headers.Authorization = `Bearer ${token}`;
  return headers;
}

function apiUrl(path: string) {
  return new URL(path, getApiBaseUrl());
}

async function safeFetch(input: URL, init: RequestInit) {
  try {
    return await fetch(input, init);
  } catch (err) {
    notifyApiDisconnected();
    throw new Error(`Cannot connect to API at ${getApiBaseUrl()}`);
  }
}

async function parseResponse<T>(response: Response): Promise<T> {
  const rawText = await response.text();
  const text = rawText.trim();
  let data: any = null;
  if (text) {
    try {
      data = JSON.parse(text);
    } catch {
      // Safely extract the primary valid JSON payload if extra characters exist
      const match = text.match(/^({[\s\S]*\}|\[[\s\S]*\]|".*?"|true|false|null|\d+)/);
      if (match) {
        try {
          data = JSON.parse(match[0]);
        } catch {
          data = { message: text };
        }
      } else {
        data = { message: text };
      }
    }
  }
  if (!response.ok) {
    throw new Error(data?.message ?? (typeof data === "string" ? data : `Request failed with status ${response.status}`));
  }
  return data as T;
}

export async function apiGet<T>(path: string, params?: Record<string, string | number | undefined>) {
  const url = apiUrl(path);
  Object.entries(params ?? {}).forEach(([key, value]) => {
    if (value !== undefined && value !== "") url.searchParams.set(key, String(value));
  });
  const response = await safeFetch(url, { headers: authHeaders(false) });
  return parseResponse<T>(response);
}

export async function apiJson<T>(path: string, method: string, body?: unknown) {
  const response = await safeFetch(apiUrl(path), {
    method,
    headers: authHeaders(true),
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  return parseResponse<T>(response);
}

export async function login(email: string, password: string) {
  const data = await apiJson<{ access_token: string }>("/api/auth/login", "POST", { email, password });
  setToken(data.access_token);
  return data;
}

export async function registerUser(body: { name: string; email: string; password: string; role: string }) {
  return apiJson("/api/auth/register", "POST", body);
}

export async function importExcel(path: string, file: File, sheet: string, batchSize: string) {
  const form = new FormData();
  form.append("file", file);
  if (sheet) form.append("sheet", sheet);
  if (batchSize) form.append("batch_size", batchSize);
  const response = await safeFetch(apiUrl(path), {
    method: "POST",
    headers: authHeaders(false),
    body: form,
  });
  return parseResponse<{ message: string; result: { inserted: number; skipped: number; errors: string[] } }>(response);
}


export async function uploadDocument(file: File) {
  const form = new FormData();
  form.append("file", file);
  const response = await safeFetch(apiUrl("/api/upload/document"), {
    method: "POST",
    headers: authHeaders(false),
    body: form,
  });
  return parseResponse<{
    message: string;
    file_name: string;
    original_name: string;
    file_url: string;
    file_path: string;
    file_size: number;
  }>(response);
}
export async function exportExcel(path: string, fields?: string, filename = "export.xlsx") {
  const url = apiUrl(path);
  if (fields) url.searchParams.set("fields", fields);
  const response = await safeFetch(url, { headers: authHeaders(false) });
  if (!response.ok) {
    const text = await response.text();
    let message = `Export failed with status ${response.status}`;
    try { message = JSON.parse(text)?.message ?? message; } catch {}
    throw new Error(message);
  }
  const blob = await response.blob();
  const objectUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = objectUrl;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(objectUrl);
}

export function toDateInput(value?: string | null) {
  if (!value) return "";
  return value.slice(0, 10);
}



