import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(__dirname, "..");
const map = JSON.parse(fs.readFileSync(path.join(projectRoot, "live-test-map.json"), "utf8"));

const baseUrl = (process.env.API_BASE_URL || map.baseUrlDefault).replace(/\/$/, "");
const email = process.env.ADMIN_EMAIL || "";
const password = process.env.ADMIN_PASSWORD || "";
const timeoutMs = Number(process.env.LIVE_TEST_TIMEOUT_MS || map.timeoutMs || 6000);

if (!email || !password) {
  console.error("Missing ADMIN_EMAIL or ADMIN_PASSWORD.");
  console.error("Example: $env:ADMIN_EMAIL='admin@example.com'; $env:ADMIN_PASSWORD='password'; npm.cmd run live:smoke");
  process.exit(1);
}

let token = "";
const results = [];

function now() {
  return Number(process.hrtime.bigint() / 1000000n);
}

async function requestJson(step) {
  const started = now();
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  const headers = { Accept: "application/json" };

  if (step.method !== "GET") headers["Content-Type"] = "application/json";
  if (token) headers.Authorization = `Bearer ${token}`;

  const options = {
    method: step.method,
    headers,
    signal: controller.signal
  };

  if (step.id === "auth.login") {
    options.body = JSON.stringify({ email, password });
  }

  try {
    const response = await fetch(`${baseUrl}${step.path}`, options);
    const text = await response.text();
    let payload = null;
    if (text) {
      try {
        payload = JSON.parse(text);
      } catch {
        payload = { raw: text.slice(0, 160) };
      }
    }

    if (step.id === "auth.login") {
      token = payload?.token || payload?.access_token || payload?.data?.token || payload?.data?.access_token || "";
      if (!token) throw new Error("Login succeeded but token was not found in response");
    }

    const ok = response.ok;
    results.push({ id: step.id, method: step.method, path: step.path, ok, status: response.status, ms: now() - started, critical: step.critical });
    return ok;
  } catch (error) {
    results.push({ id: step.id, method: step.method, path: step.path, ok: false, status: "ERR", ms: now() - started, critical: step.critical, error: error.message });
    return false;
  } finally {
    clearTimeout(timeout);
  }
}

console.log(`Live smoke test: ${baseUrl}`);
console.log(`Timeout per request: ${timeoutMs}ms`);

for (const step of map.fastPath) {
  const ok = await requestJson(step);
  if (!ok && step.critical) {
    break;
  }
}

const printable = results.map((item) => ({
  id: item.id,
  method: item.method,
  status: item.status,
  ms: item.ms,
  critical: item.critical ? "yes" : "no",
  ok: item.ok ? "PASS" : "FAIL",
  path: item.path
}));

console.table(printable);

const failed = results.filter((item) => !item.ok);
const failedCritical = failed.filter((item) => item.critical);

if (failed.length) {
  console.log("Failed steps:");
  for (const item of failed) {
    console.log(`- ${item.id} ${item.status} ${item.error || ""}`.trim());
  }
}

if (failedCritical.length) {
  console.error(`Live smoke failed: ${failedCritical.length} critical endpoint(s) failed.`);
  process.exit(1);
}

console.log("Live smoke passed for all critical endpoints.");
