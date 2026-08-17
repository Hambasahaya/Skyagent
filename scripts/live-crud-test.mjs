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
const stamp = `LIVE_TEST_${Date.now()}`;

if (!email || !password) {
  console.error("Missing ADMIN_EMAIL or ADMIN_PASSWORD.");
  process.exit(1);
}

let token = "";
const created = [];
const results = [];

function startedAt() {
  return Number(process.hrtime.bigint() / 1000000n);
}

function extractId(payload) {
  return payload?.id || payload?.data?.id || payload?.item?.id || payload?.document_type?.id || payload?.document_name?.id || payload?.principal?.id || payload?.vessel?.id || payload?.status?.id || payload?.requirement?.id || payload?.custom_field?.id;
}

async function call(label, method, urlPath, body, critical = true) {
  const started = startedAt();
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  const headers = { Accept: "application/json" };
  if (token) headers.Authorization = `Bearer ${token}`;
  if (body !== undefined) headers["Content-Type"] = "application/json";

  try {
    const response = await fetch(`${baseUrl}${urlPath}`, {
      method,
      headers,
      body: body === undefined ? undefined : JSON.stringify(body),
      signal: controller.signal
    });
    const text = await response.text();
    let payload = null;
    if (text) {
      try { payload = JSON.parse(text); } catch { payload = { raw: text.slice(0, 160) }; }
    }
    if (!response.ok) {
      throw new Error(payload?.message || payload?.error || `HTTP ${response.status}`);
    }
    results.push({ label, method, status: response.status, ms: startedAt() - started, ok: "PASS" });
    return payload;
  } catch (error) {
    results.push({ label, method, status: "ERR", ms: startedAt() - started, ok: "FAIL", error: error.message, critical });
    if (critical) throw error;
    return null;
  } finally {
    clearTimeout(timer);
  }
}

async function cleanup() {
  for (const item of created.reverse()) {
    try {
      await call(`cleanup ${item.label}`, "DELETE", item.path, undefined, false);
    } catch {}
  }
}

try {
  const login = await call("auth.login", "POST", "/api/auth/login", { email, password });
  token = login?.access_token || login?.token || login?.data?.access_token || login?.data?.token || "";
  if (!token) throw new Error("Token not found after login");

  const docType = await call("document-type.create", "POST", "/api/admin/document-types", {
    name: `${stamp}_TYPE`, description: "Live CRUD test document type", is_active: true
  });
  const docTypeId = extractId(docType);
  created.push({ label: "document-type", path: `/api/admin/document-types/${docTypeId}` });
  await call("document-type.read", "GET", `/api/admin/document-types/${docTypeId}`);
  await call("document-type.update", "PUT", `/api/admin/document-types/${docTypeId}`, {
    name: `${stamp}_TYPE_UPDATED`, description: "Updated live CRUD test document type", is_active: true
  });

  const docName = await call("document-name.create", "POST", "/api/admin/document-names", {
    document_type_id: docTypeId, name: `${stamp}_BST`, code: `${stamp}_BST`, description: "Live CRUD test document", requires_expiry: true, default_validity_months: 60, metadata_schema: "{}", is_active: true
  });
  const docNameId = extractId(docName);
  created.push({ label: "document-name", path: `/api/admin/document-names/${docNameId}` });
  await call("document-name.read", "GET", `/api/admin/document-names/${docNameId}`);
  await call("document-name.update", "PUT", `/api/admin/document-names/${docNameId}`, {
    document_type_id: docTypeId, name: `${stamp}_BST_UPDATED`, code: `${stamp}_BST_UPD`, description: "Updated live CRUD test document", requires_expiry: false, default_validity_months: 0, metadata_schema: "{}", is_active: true
  });

  const principal = await call("principal.create", "POST", "/api/admin/principals", {
    name: `${stamp} Principal`, code: stamp.slice(-12), contact_name: "Live Tester", contact_phone: "0800000000", email: "live.test@example.com", address: "Test Address", custom_fields: "{}", is_active: true
  });
  const principalId = extractId(principal);
  created.push({ label: "principal", path: `/api/admin/principals/${principalId}` });
  await call("principal.read", "GET", `/api/admin/principals/${principalId}`);
  await call("principal.update", "PUT", `/api/admin/principals/${principalId}`, {
    name: `${stamp} Principal Updated`, code: stamp.slice(-12), contact_name: "Live Tester 2", contact_phone: "0811111111", email: "live.test@example.com", address: "Updated Test Address", custom_fields: "{\"payment_term\":\"30 days\"}", is_active: true
  });

  const vessel = await call("vessel.create", "POST", "/api/admin/vessels", {
    principal_id: principalId, name: `${stamp} Vessel`, code: `${stamp.slice(-8)}V`, vessel_type: "General Cargo", flag: "Indonesia", imo: `IMO${String(Date.now()).slice(-7)}`, mmsi: String(Date.now()).slice(-9), custom_fields: "{}", is_active: true
  });
  const vesselId = extractId(vessel);
  created.push({ label: "vessel", path: `/api/admin/vessels/${vesselId}` });
  await call("vessel.read", "GET", `/api/admin/vessels/${vesselId}`);
  await call("vessel.update", "PUT", `/api/admin/vessels/${vesselId}`, {
    principal_id: principalId, name: `${stamp} Vessel Updated`, code: `${stamp.slice(-8)}VU`, vessel_type: "Bulk Carrier", flag: "Indonesia", imo: `IMO${String(Date.now()).slice(-7)}`, mmsi: String(Date.now()).slice(-9), custom_fields: "{\"engine_type\":\"Diesel\"}", is_active: true
  });

  const customField = await call("custom-field.create", "POST", "/api/admin/custom-fields", {
    entity_type: "principal", field_key: `${stamp.toLowerCase()}_payment_term`, label: "Live Payment Term", field_type: "text", options_json: "{}", is_required: false, is_active: true
  });
  const customFieldId = extractId(customField);
  created.push({ label: "custom-field", path: `/api/admin/custom-fields/${customFieldId}` });
  await call("custom-field.update", "PUT", `/api/admin/custom-fields/${customFieldId}`, {
    entity_type: "principal", field_key: `${stamp.toLowerCase()}_payment_term`, label: "Live Payment Term Updated", field_type: "text", options_json: "{}", is_required: false, is_active: true
  });
  await call("principal.detail", "GET", `/api/admin/principals/${principalId}/detail`);
  await call("principal.vessels", "GET", `/api/admin/principals/${principalId}/vessels?active=true`);
  await call("principal.custom-fields", "GET", `/api/admin/principals/${principalId}/custom-fields`);

  const status = await call("joining-status.create", "POST", "/api/admin/joining-statuses", {
    name: `${stamp}_STATUS`, description: "Live CRUD status", is_active: true
  });
  const statusId = extractId(status);
  created.push({ label: "joining-status", path: `/api/admin/joining-statuses/${statusId}` });
  await call("joining-status.read", "GET", `/api/admin/joining-statuses/${statusId}`);
  await call("joining-status.update", "PUT", `/api/admin/joining-statuses/${statusId}`, {
    name: `${stamp}_STATUS_UPDATED`, description: "Updated live CRUD status", is_active: true
  });

  const requirement = await call("requirement.create", "POST", "/api/admin/principal-requirements", {
    principal_id: principalId, vessel_id: vesselId, requirement_type: "custom", title: `${stamp} Requirement`, description: "Live CRUD requirement", is_mandatory: true, custom_fields: "{\"minimum_contracts\":1}", is_active: true
  });
  const requirementId = extractId(requirement);
  created.push({ label: "requirement", path: `/api/admin/principal-requirements/${requirementId}` });
  await call("requirement.read", "GET", `/api/admin/principal-requirements/${requirementId}`);
  await call("requirement.update", "PUT", `/api/admin/principal-requirements/${requirementId}`, {
    principal_id: principalId, vessel_id: vesselId, requirement_type: "custom", title: `${stamp} Requirement Updated`, description: "Updated live CRUD requirement", is_mandatory: false, custom_fields: "{\"minimum_contracts\":2}", is_active: true
  });

  await cleanup();
  console.table(results.map(({ label, method, status, ms, ok }) => ({ label, method, status, ms, ok })));
  const failed = results.filter((result) => result.ok !== "PASS" && result.critical !== false);
  if (failed.length) process.exit(1);
  console.log("Live CRUD passed and temporary data was cleaned up.");
} catch (error) {
  console.error(`Live CRUD failed: ${error.message}`);
  await cleanup();
  console.table(results.map(({ label, method, status, ms, ok, error }) => ({ label, method, status, ms, ok, error })));
  process.exit(1);
}