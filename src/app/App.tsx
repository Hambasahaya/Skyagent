import React, { useCallback, useEffect, useMemo, useState } from "react";
import {
  LayoutDashboard, Users, Ship, FileText, Settings, ChevronLeft, ChevronRight,
  ChevronDown, ChevronUp, Search, Bell, X, Plus, Download, Upload, Pencil,
  Trash2, Eye, AlertTriangle, Clock, Ban, UserCheck, UserX, Building2, Anchor,
  BarChart3, Shield, Globe, Phone, Mail, Calendar, Briefcase, DollarSign,
  MapPin, CheckCircle, XCircle, AlertCircle, RefreshCw, LogOut, User, Lock,
  ArrowRight, MoreVertical, Info, TrendingUp, Activity, Database, ClipboardList,
  Layers, FileCheck, UserPlus, Filter, Hash, Key, FileWarning, Menu, Check,
  ArrowLeft, Send
} from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, Legend
} from "recharts";
import { Toaster, toast } from "sonner";
import { apiGet, apiJson, checkApiConnection, clearToken, exportExcel, getApiBaseUrl, getToken, importExcel, login, registerUser, setApiBaseUrl, toDateInput, uploadDocument } from "./api";
import skyagenLogoUrl from "../assets/skyagen-logo.png";

// ─── Types ───────────────────────────────────────────────────────────────────

type AuthPage = "login" | "register" | "verify-email" | "verify-success" | "verify-failed" | "update-password" | "error-401" | "error-403" | "error-404";
type Page = string;

interface Crew {
  id: number; code: string; name: string; phone: string; nik: string;
  passport: string; seamanBook: string; status: string; rank: string;
  nationality: string; dob: string; email: string;
}
interface Vessel {
  id: number; code: string; name: string; principal: string; type: string;
  flag: string; imo: string; mmsi: string; active: boolean; customFields: Record<string, unknown>;
}
interface Principal {
  id: number; code: string; name: string; contact: string; phone: string;
  email: string; active: boolean; customFields: Record<string, unknown>;
}
interface JoiningRecord {
  id: number; crewId: number; crewName: string; rank: string; principal: string;
  vessel: string; status: string; signOn: string | null; signOff: string | null;
  port: string; salary: number;
}

type ApiListResponse<T> = { data: T[]; total?: number };

type ApiSeafarer = {
  id: number; seafarer_code?: string; name: string; phone?: string; nik?: string;
  passport_no?: string; seaman_book_no?: string; status?: string; certificate_no?: string;
  birth_date?: string | null; photo_url?: string;
};

type ApiPrincipal = {
  id: number; code?: string; name: string; contact_name?: string; contact_phone?: string;
  email?: string; is_active?: boolean; custom_fields?: Record<string, unknown> | null;
};

type ApiVessel = {
  id: number; code?: string; name: string; principal_id?: number; principal?: ApiPrincipal | null;
  vessel_type?: string; flag?: string; imo?: string; mmsi?: string; is_active?: boolean; custom_fields?: Record<string, unknown> | null;
};

type ApiDashboardSummary = { total_crew: number; onboard: number; available: number; waiting: number; blacklisted: number; total_principals: number; total_vessels: number; total_joining: number };
type ApiDashboardCharts = { months: { month: string; sign_on_count: number; sign_off_count: number }[]; total_sign_on: number; total_sign_off: number; period: string };
type ApiExpiringDoc = { id: number; seafarer_id: number; seafarer_name: string; document_name?: { id: number; name: string } | null; document_no: string; expired_at?: string | null; is_lifetime?: boolean };
type ApiDocumentReport = { total_documents: number; expired_count: number; expiring_count: number; valid_count: number; type_breakdown?: { document_type_name: string; total: number; expired: number; expiring: number; valid?: number }[]; top_expired?: { seafarer_id: number; seafarer_name: string; expired_count: number }[] };
type ApiCrewReport = { total_crew: number; by_status: { status: string; count: number }[]; by_marital: { marital_status: string; count: number }[]; age_distribution: { range: string; count: number }[] };
type ApiJoiningReport = { total: number; by_status: { status: string; count: number }[]; by_principal: { principal_name: string; count: number }[]; by_vessel_type: { vessel_type: string; count: number }[]; monthly_trend: { month: string; count: number }[] };
type ApiUser = { id: number; name: string; email: string; role: string; is_email_verified: boolean; email_verified_at?: string | null; created_at?: string };
type ApiProfile = { id: number; name: string; email: string; role: string };
type ApiStorageSettings = { storage_path: string; absolute_path: string; exists: boolean; allowed_extensions: string[]; max_file_size_mb: number };
type ApiUploadedDocument = { message: string; file_name: string; original_name: string; file_url: string; file_path: string; file_size: number };
type ApiPrincipalRequirement = {
  id: number; principal_id: number; vessel_id?: number | null; requirement_type?: string;
  document_name_id?: number | null; is_mandatory?: boolean; requires_valid_document?: boolean; is_active?: boolean;
  principal?: ApiPrincipal | null; vessel?: ApiVessel | null; document_name?: ApiDocumentName | null;
};

type ApiRequirementCheck = {
  passed: boolean; total_requirements?: number; passed_count?: number; failed_count?: number;
  requirements?: { name?: string; document_name?: string; status?: string; passed?: boolean; reason?: string; document_no?: string; expired_at?: string | null }[];
};
type ApiCustomField = {
  id: number; entity_type: "principal" | "vessel"; field_key: string; label: string;
  field_type?: string; options_json?: string; is_required?: boolean; is_active?: boolean;
};
type ApiDocumentType = {
  id: number; name: string; description?: string; is_active?: boolean;
};

type ApiPayslipTemplate = {
  id: number;
  name: string;
  file_url: string;
  placeholders_json?: string;
  is_default?: boolean;
  is_active?: boolean;
  created_at?: string;
};

type ApiPayslip = {
  id: number;
  seafarer_id: number;
  seafarer?: { id: number; name: string; rank?: string } | null;
  month: number;
  year: number;
  basic_salary: number;
  allowances: number;
  deductions: number;
  net_salary: number;
  status: "generated" | "sent" | string;
  sent_at?: string | null;
  details_json?: string;
};

async function uploadPayslipTemplate(file: File, name?: string, isDefault: boolean = false) {
  const formData = new FormData();
  formData.append("file", file);
  if (name) formData.append("name", name);
  if (isDefault) formData.append("is_default", "true");

  const response = await fetch(`${getApiBaseUrl()}/api/admin/payslips/templates/upload`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${getToken()}`,
    },
    body: formData,
  });

  const resJson = await response.json();
  if (!response.ok) {
    throw new Error(resJson.message || resJson.error || "Upload template failed");
  }
  return resJson;
}

async function generatePayslips(payload: { month: number; year: number; template_id?: number; items: Array<{ seafarer_id: number; basic_salary: number; allowances: number; deductions: number; net_salary: number; details_json?: string }> }) {
  return apiJson<{ message: string; created_count: number; updated_count: number }>("/api/admin/payslips/generate", "POST", payload);
}

async function updatePayslip(id: number, payload: { basic_salary: number; allowances: number; deductions: number; net_salary: number; details_json?: string }) {
  return apiJson(`/api/admin/payslips/${id}`, "PUT", payload);
}

async function sendSinglePayslip(id: number) {
  return apiJson<{ message: string; data: ApiPayslip }>(`/api/admin/payslips/${id}/send`, "POST", {});
}

async function sendBulkPayslips(payload: { payslip_ids?: number[]; month?: number; year?: number }) {
  return apiJson<{ message: string; sent_count: number }>("/api/admin/payslips/send", "POST", payload);
}

type ApiDocumentName = {
  id: number; document_type_id: number; document_type?: ApiDocumentType | null;
  name: string; description?: string; is_required?: boolean; has_expiry?: boolean; is_active?: boolean;
};

type ApiBlacklist = {
  id: number; seafarer_id: number; seafarer?: ApiSeafarer | null; reason: string; notes?: string;
  is_active?: boolean; released_reason?: string;
};
type ApiJoining = {
  id: number; seafarer_id?: number; seafarer?: ApiSeafarer | null; crew_code?: string;
  name: string; rank?: string; principal?: ApiPrincipal | null; vessel?: ApiVessel | null;
  vessel_name?: string; vessel_type?: string; joining_status?: { name?: string } | null;
  status?: string; sign_on?: string | null; sign_off?: string | null; port_join?: string;
  salary_crew?: number; total_salary?: number; shipowner?: string;
};

function mapCrew(item: ApiSeafarer): Crew {
  return {
    id: item.id,
    code: item.seafarer_code || `SEA${String(item.id || 0).padStart(3, "0")}`,
    name: item.name || "-",
    phone: item.phone || "-",
    nik: item.nik || "-",
    passport: item.passport_no || "-",
    seamanBook: item.seaman_book_no || "-",
    status: item.status || "available",
    rank: item.certificate_no || "-",
    nationality: "-",
    dob: toDateInput(item.birth_date),
    email: "",
  };
}

function parseJsonField(val: any): Record<string, unknown> {
  if (!val) return {};
  if (typeof val === "object") return val;
  try {
    return JSON.parse(val);
  } catch {
    return {};
  }
}

function mapPrincipal(item: ApiPrincipal): Principal {
  return {
    id: item.id,
    code: item.code || `PC${String(item.id).padStart(3, "0")}`,
    name: item.name || "-",
    contact: item.contact_name || "-",
    phone: item.contact_phone || "-",
    email: item.email || "-",
    active: item.is_active ?? true,
    customFields: parseJsonField(item.custom_fields),
  };
}

function mapVessel(item: ApiVessel): Vessel {
  return {
    id: item.id,
    code: item.code || `VS${String(item.id).padStart(3, "0")}`,
    name: item.name || "-",
    principal: item.principal?.name || (item.principal_id ? `Principal #${item.principal_id}` : "-"),
    type: item.vessel_type || "-",
    flag: item.flag || "-",
    imo: item.imo || "-",
    mmsi: item.mmsi || "-",
    active: item.is_active ?? true,
    customFields: parseJsonField(item.custom_fields),
  };
}

function mapDocType(item: ApiDocumentType) {
  return { id: item.id, code: `DOC${String(item.id).padStart(3, "0")}`, name: item.name, description: item.description || "", active: item.is_active ?? true, count: 0 };
}

function mapPrincipalRequirement(item: ApiPrincipalRequirement) {
  return {
    id: item.id,
    principalId: item.principal_id,
    principalName: item.principal?.name || `Principal #${item.principal_id}`,
    vesselId: item.vessel_id ?? null,
    vesselName: item.vessel?.name || (item.vessel_id ? `Vessel #${item.vessel_id}` : "All vessels"),
    type: item.requirement_type || "document",
    documentNameId: item.document_name_id ?? null,
    documentName: item.document_name?.name || (item.requirement_type === "custom" ? "Custom requirement" : `Document #${item.document_name_id ?? "-"}`),
    mandatory: item.is_mandatory ?? true,
    requiresValid: item.requires_valid_document ?? true,
    active: item.is_active ?? true,
  };
}

type PrincipalRequirement = ReturnType<typeof mapPrincipalRequirement>;
function mapCustomField(item: ApiCustomField) {
  return {
    id: item.id,
    entityType: item.entity_type,
    key: item.field_key,
    label: item.label,
    type: item.field_type || "text",
    optionsJson: item.options_json || "{}",
    required: item.is_required ?? false,
    active: item.is_active ?? true,
  };
}

type CustomField = ReturnType<typeof mapCustomField>;

function parseCustomFieldOptions(optionsJson: string) {
  try {
    const parsed = JSON.parse(optionsJson || "{}");
    if (Array.isArray(parsed)) return parsed.map(String);
    if (Array.isArray(parsed.options)) return parsed.options.map(String);
    return Object.keys(parsed).length ? Object.values(parsed).map(String) : [];
  } catch {
    return [];
  }
}

function customValueToString(value: unknown) {
  if (value === null || value === undefined) return "";
  if (typeof value === "boolean") return value ? "true" : "false";
  return String(value);
}

function mapDocName(item: ApiDocumentName) {
  return { id: item.id, type: item.document_type?.name || `Type #${item.document_type_id}`, documentTypeId: item.document_type_id, name: item.name, required: item.is_required ?? false, hasExpiry: item.has_expiry ?? false, active: item.is_active ?? true };
}

function mapBlacklist(item: ApiBlacklist) {
  return { id: item.id, seafarerId: item.seafarer_id, crew: item.seafarer?.name || `Seafarer #${item.seafarer_id}`, code: item.seafarer?.seafarer_code || String(item.seafarer_id), reason: item.reason, notes: item.notes || item.released_reason || "", status: item.is_active === false ? "released" : "active" };
}
function mapJoining(item: ApiJoining): JoiningRecord {
  return {
    id: item.id,
    crewId: item.seafarer_id ?? item.seafarer?.id ?? 0,
    crewName: item.name || item.seafarer?.name || "-",
    rank: item.rank || "-",
    principal: item.principal?.name || item.shipowner || "-",
    vessel: item.vessel_name || item.vessel?.name || "-",
    status: item.status || item.joining_status?.name || "-",
    signOn: toDateInput(item.sign_on) || null,
    signOff: toDateInput(item.sign_off) || null,
    port: item.port_join || "-",
    salary: item.total_salary || item.salary_crew || 0,
  };
}

function useApiList<TApi, TUi>(path: string, mapper: (item: TApi) => TUi, fallback: TUi[] = []) {
  const [data, setData] = useState<TUi[]>([]);
  const [loading, setLoading] = useState(false);
  const [reloadKey, setReloadKey] = useState(0);

  useEffect(() => {
    if (!getToken()) { setData([]); return; }
    let active = true;
    setLoading(true);
    apiGet<ApiListResponse<TApi>>(path)
      .then(res => { if (active) setData((res.data ?? []).map(mapper)); })
      .catch(err => {
        if (active) {
          setData([]);
          if (!err.message?.includes("404")) {
            toast.error(err.message);
          }
        }
      })
      .finally(() => { if (active) setLoading(false); });
    return () => { active = false; };
  }, [path, reloadKey]);

  return { data, loading, refresh: () => setReloadKey(v => v + 1) };
}

// ─── Mock Data ────────────────────────────────────────────────────────────────

const MOCK_CREW: Crew[] = [
  { id: 1, code: "SEA001", name: "Budi Santoso", phone: "+62 812-3456-7890", nik: "3201011234567890", passport: "A1234567", seamanBook: "D123456", status: "onboard", rank: "Chief Officer", nationality: "Indonesian", dob: "1988-03-15", email: "budi.santoso@email.com" },
  { id: 2, code: "SEA002", name: "Juan Dela Cruz", phone: "+63 917-234-5678", nik: "PHL9876543210", passport: "P9876543", seamanBook: "PH234567", status: "available", rank: "Master", nationality: "Filipino", dob: "1982-07-22", email: "juan.delacruz@email.com" },
  { id: 3, code: "SEA003", name: "Ahmad Fauzi", phone: "+62 813-4567-8901", nik: "3271011234567891", passport: "B2345678", seamanBook: "D234567", status: "waiting", rank: "Second Officer", nationality: "Indonesian", dob: "1991-11-08", email: "ahmad.fauzi@email.com" },
  { id: 4, code: "SEA004", name: "Roberto Santos", phone: "+63 918-345-6789", nik: "PHL8765432109", passport: "P8765432", seamanBook: "PH345678", status: "available", rank: "Chief Engineer", nationality: "Filipino", dob: "1985-04-30", email: "roberto.santos@email.com" },
  { id: 5, code: "SEA005", name: "Dedi Kurniawan", phone: "+62 814-5678-9012", nik: "3301021234567892", passport: "C3456789", seamanBook: "D345678", status: "onboard", rank: "Third Officer", nationality: "Indonesian", dob: "1993-09-17", email: "dedi.kurniawan@email.com" },
  { id: 6, code: "SEA006", name: "Mario Reyes", phone: "+63 919-456-7890", nik: "PHL7654321098", passport: "P7654321", seamanBook: "PH456789", status: "ex-crew", rank: "Second Engineer", nationality: "Filipino", dob: "1980-12-03", email: "mario.reyes@email.com" },
  { id: 7, code: "SEA007", name: "Slamet Wahyudi", phone: "+62 815-6789-0123", nik: "3401031234567893", passport: "D4567890", seamanBook: "D456789", status: "blacklisted", rank: "Bosun", nationality: "Indonesian", dob: "1987-06-25", email: "slamet.w@email.com" },
  { id: 8, code: "SEA008", name: "Jose Garcia", phone: "+63 920-567-8901", nik: "PHL6543210987", passport: "P6543210", seamanBook: "PH567890", status: "available", rank: "Able Seaman", nationality: "Filipino", dob: "1995-02-14", email: "jose.garcia@email.com" },
  { id: 9, code: "SEA009", name: "Rizky Pratama", phone: "+62 816-7890-1234", nik: "3501041234567894", passport: "E5678901", seamanBook: "D567890", status: "waiting", rank: "Ordinary Seaman", nationality: "Indonesian", dob: "1997-08-11", email: "rizky.p@email.com" },
  { id: 10, code: "SEA010", name: "Carlos Mendoza", phone: "+63 921-678-9012", nik: "PHL5432109876", passport: "P5432109", seamanBook: "PH678901", status: "onboard", rank: "Cook", nationality: "Filipino", dob: "1989-01-28", email: "carlos.mendoza@email.com" },
  { id: 11, code: "SEA011", name: "Eko Prasetyo", phone: "+62 817-8901-2345", nik: "3601051234567895", passport: "F6789012", seamanBook: "D678901", status: "available", rank: "Engine Cadet", nationality: "Indonesian", dob: "2000-05-19", email: "eko.prasetyo@email.com" },
  { id: 12, code: "SEA012", name: "Benjamin Cruz", phone: "+63 922-789-0123", nik: "PHL4321098765", passport: "P4321098", seamanBook: "PH789012", status: "available", rank: "Deck Cadet", nationality: "Filipino", dob: "2001-10-07", email: "ben.cruz@email.com" },
  { id: 13, code: "SEA013", name: "Wahyu Hidayat", phone: "+62 818-9012-3456", nik: "3701061234567896", passport: "G7890123", seamanBook: "D789012", status: "waiting", rank: "Third Engineer", nationality: "Indonesian", dob: "1994-03-22", email: "wahyu.h@email.com" },
  { id: 14, code: "SEA014", name: "Mark Villanueva", phone: "+63 923-890-1234", nik: "PHL3210987654", passport: "P3210987", seamanBook: "PH890123", status: "available", rank: "Electrician", nationality: "Filipino", dob: "1990-07-14", email: "mark.v@email.com" },
];

const MOCK_VESSELS: Vessel[] = [
  { id: 1, code: "VS001", name: "MV Ocean Pioneer", principal: "Pacific Shipping Co", type: "Bulk Carrier", flag: "Panama", imo: "9234567", mmsi: "371234567", active: true },
  { id: 2, code: "VS002", name: "MT Sea Dragon", principal: "Atlas Maritime Ltd", type: "Oil Tanker", flag: "Liberia", imo: "9345678", mmsi: "636345678", active: true },
  { id: 3, code: "VS003", name: "MV Harbor Star", principal: "Pacific Shipping Co", type: "Container Ship", flag: "Marshall Islands", imo: "9456789", mmsi: "538456789", active: true },
  { id: 4, code: "VS004", name: "MV Southern Cross", principal: "Nordic Ocean AS", type: "General Cargo", flag: "Bahamas", imo: "9567890", mmsi: "311567890", active: false },
  { id: 5, code: "VS005", name: "MT Blue Horizon", principal: "Atlas Maritime Ltd", type: "Chemical Tanker", flag: "Singapore", imo: "9678901", mmsi: "566678901", active: true },
  { id: 6, code: "VS006", name: "MV Pacific Arrow", principal: "Eastern Carriers Inc", type: "Bulk Carrier", flag: "Cyprus", imo: "9789012", mmsi: "212789012", active: true },
  { id: 7, code: "VS007", name: "MV Gulf Fortune", principal: "Gulf Marine Services", type: "Platform Supply", flag: "UAE", imo: "9890123", mmsi: "470890123", active: true },
];

const MOCK_PRINCIPALS: Principal[] = [
  { id: 1, code: "PC001", name: "Pacific Shipping Co", contact: "Robert Wilson", phone: "+1 555-0100", email: "ops@pacificshipping.com", active: true },
  { id: 2, code: "PC002", name: "Atlas Maritime Ltd", contact: "Sarah Johnson", phone: "+44 20 7000-1234", email: "crew@atlasmaritme.com", active: true },
  { id: 3, code: "PC003", name: "Nordic Ocean AS", contact: "Erik Andersen", phone: "+47 22 000-111", email: "manning@nordicocean.no", active: true },
  { id: 4, code: "PC004", name: "Eastern Carriers Inc", contact: "Michael Chen", phone: "+852 2000-3456", email: "operations@easterncarriers.hk", active: false },
  { id: 5, code: "PC005", name: "Gulf Marine Services", contact: "Ahmed Al-Rashid", phone: "+971 4 000-5678", email: "crew@gulfmarine.ae", active: true },
];

const MOCK_JOINING: JoiningRecord[] = [
  { id: 1, crewId: 1, crewName: "Budi Santoso", rank: "Chief Officer", principal: "Pacific Shipping Co", vessel: "MV Ocean Pioneer", status: "onboard", signOn: "2024-01-15", signOff: null, port: "Singapore", salary: 3200 },
  { id: 2, crewId: 5, crewName: "Dedi Kurniawan", rank: "Third Officer", principal: "Atlas Maritime Ltd", vessel: "MT Sea Dragon", status: "onboard", signOn: "2024-02-01", signOff: null, port: "Rotterdam", salary: 1800 },
  { id: 3, crewId: 10, crewName: "Carlos Mendoza", rank: "Cook", principal: "Nordic Ocean AS", vessel: "MV Harbor Star", status: "onboard", signOn: "2024-01-28", signOff: null, port: "Shanghai", salary: 1200 },
  { id: 4, crewId: 2, crewName: "Juan Dela Cruz", rank: "Master", principal: "Pacific Shipping Co", vessel: "MV Harbor Star", status: "completed", signOn: "2023-08-01", signOff: "2024-01-05", port: "Houston", salary: 7500 },
  { id: 5, crewId: 4, crewName: "Roberto Santos", rank: "Chief Engineer", principal: "Eastern Carriers Inc", vessel: "MV Pacific Arrow", status: "pending", signOn: null, signOff: null, port: "Manila", salary: 5500 },
  { id: 6, crewId: 6, crewName: "Mario Reyes", rank: "Second Engineer", principal: "Atlas Maritime Ltd", vessel: "MT Blue Horizon", status: "completed", signOn: "2023-06-10", signOff: "2023-12-20", port: "Santos", salary: 3800 },
  { id: 7, crewId: 8, crewName: "Jose Garcia", rank: "Able Seaman", principal: "Gulf Marine Services", vessel: "MV Gulf Fortune", status: "approved", signOn: null, signOff: null, port: "Dubai", salary: 950 },
];

const MOCK_EXPIRING_DOCS = [
  { id: 1, crew: "Budi Santoso", document: "Medical Certificate", number: "MED-2024-001", expiry: "2024-03-15", daysLeft: 12, status: "expiring" },
  { id: 2, crew: "Juan Dela Cruz", document: "GMDSS Certificate", number: "GMDSS-PH-4567", expiry: "2024-02-28", daysLeft: -5, status: "expired" },
  { id: 3, crew: "Ahmad Fauzi", document: "Passport", number: "B2345678", expiry: "2024-04-30", daysLeft: 58, status: "expiring" },
  { id: 4, crew: "Dedi Kurniawan", document: "Basic Safety Training", number: "BST-2024-789", expiry: "2024-03-01", daysLeft: 2, status: "expiring" },
  { id: 5, crew: "Roberto Santos", document: "Seaman Book", number: "PH234567", expiry: "2029-12-31", daysLeft: 2150, status: "valid" },
  { id: 6, crew: "Carlos Mendoza", document: "Yellow Fever Certificate", number: "YF-2023-456", expiry: "2024-02-10", daysLeft: -23, status: "expired" },
  { id: 7, crew: "Jose Garcia", document: "STCW Certificate", number: "STCW-2024-321", expiry: "2024-05-15", daysLeft: 73, status: "valid" },
  { id: 8, crew: "Eko Prasetyo", document: "Seafarer ID", number: "ID-2024-654", expiry: "2024-03-20", daysLeft: 17, status: "expiring" },
  { id: 9, crew: "Wahyu Hidayat", document: "Medical Certificate", number: "MED-2024-009", expiry: "2024-02-25", daysLeft: -8, status: "expired" },
  { id: 10, crew: "Mark Villanueva", document: "Officer Watch Cert", number: "OWK-2024-010", expiry: "2026-08-15", daysLeft: 365, status: "valid" },
];

const MOCK_DOC_TYPES = [
  { id: 1, code: "CERT", name: "Certificate", active: true, count: 8 },
  { id: 2, code: "ID", name: "Identification", active: true, count: 5 },
  { id: 3, code: "MED", name: "Medical", active: true, count: 3 },
  { id: 4, code: "TRVL", name: "Travel", active: true, count: 4 },
  { id: 5, code: "TRAIN", name: "Training", active: false, count: 2 },
];

const MOCK_DOC_NAMES = [
  { id: 1, type: "Certificate", name: "STCW Certificate", required: true, hasExpiry: true },
  { id: 2, type: "Certificate", name: "Basic Safety Training (BST)", required: true, hasExpiry: true },
  { id: 3, type: "Certificate", name: "GMDSS Certificate", required: false, hasExpiry: true },
  { id: 4, type: "Medical", name: "Medical Certificate", required: true, hasExpiry: true },
  { id: 5, type: "Identification", name: "National ID / KTP", required: true, hasExpiry: true },
  { id: 6, type: "Identification", name: "Seafarer ID", required: true, hasExpiry: true },
  { id: 7, type: "Medical", name: "Yellow Fever Certificate", required: false, hasExpiry: true },
  { id: 8, type: "Travel", name: "Passport", required: true, hasExpiry: true },
  { id: 9, type: "Travel", name: "Seaman Book", required: true, hasExpiry: true },
  { id: 10, type: "Certificate", name: "Officer Watch Certificate", required: true, hasExpiry: false },
  { id: 11, type: "Certificate", name: "Ship Security Officer Cert", required: false, hasExpiry: true },
  { id: 12, type: "Training", name: "HUET Training Certificate", required: false, hasExpiry: true },
];

const CHART_DATA = [
  { month: "Aug", signOn: 8, signOff: 5 },
  { month: "Sep", signOn: 12, signOff: 9 },
  { month: "Oct", signOn: 6, signOff: 7 },
  { month: "Nov", signOn: 15, signOff: 11 },
  { month: "Dec", signOn: 9, signOff: 13 },
  { month: "Jan", signOn: 11, signOff: 8 },
  { month: "Feb", signOn: 14, signOff: 10 },
];

const STATUS_DIST = [
  { name: "Available", value: 5, color: "#22c55e" },
  { name: "Onboard", value: 3, color: "#3b82f6" },
  { name: "Waiting", value: 3, color: "#f59e0b" },
  { name: "Ex-Crew", value: 1, color: "#94a3b8" },
  { name: "Blacklisted", value: 1, color: "#ef4444" },
  { name: "Unknown", value: 1, color: "#cbd5e1" },
];

const MOCK_BLACKLIST = [
  { id: 1, crew: "Slamet Wahyudi", code: "SEA007", reason: "Contract breach — abandoned vessel in Rotterdam", notes: "Left vessel without notice during cargo operations. Deemed unfit for re-deployment.", status: "active" },
  { id: 2, crew: "Andika Prasetya", code: "SEA019", reason: "Drug test positive — cannabis detected", notes: "Failed pre-joining drug screening. Released pending medical clearance.", status: "active" },
  { id: 3, crew: "Renato Castillo", code: "SEA024", reason: "Misconduct — insubordination to Master", notes: "Multiple documented incidents. Reviewed by company disciplinary board.", status: "released" },
];

const MOCK_JOINING_STATUSES = [
  { id: 1, code: "PENDING", label: "Pending", color: "#f59e0b", isFinal: false },
  { id: 2, code: "APPROVED", label: "Approved", color: "#0ea5e9", isFinal: false },
  { id: 3, code: "ONBOARD", label: "Onboard", color: "#3b82f6", isFinal: false },
  { id: 4, code: "COMPLETED", label: "Completed", color: "#22c55e", isFinal: true },
  { id: 5, code: "CANCELLED", label: "Cancelled", color: "#ef4444", isFinal: true },
];

const NAV_SECTIONS = [
  {
    section: null,
    items: [{ id: "dashboard", label: "Dashboard", icon: LayoutDashboard }],
  },
  {
    section: "CREW MANAGEMENT",
    items: [
      { id: "crew-database", label: "Crew Database", icon: Database },
      { id: "search-crew", label: "Search Crew", icon: Search },
      { id: "available-crew", label: "Available Crew", icon: UserCheck },
      { id: "onboard-crew", label: "Onboard Crew", icon: Ship },
      { id: "joining-principal", label: "Joining Principal", icon: Briefcase },
      { id: "waiting-crew", label: "Waiting Crew", icon: Clock },
      { id: "blacklist", label: "Blacklist", icon: Ban },
    ],
  },
  {
    section: "DOCUMENT CONTROL",
    items: [
      { id: "documents", label: "Documents", icon: FileText },
      { id: "document-types", label: "Document Types", icon: Layers },
      { id: "document-names", label: "Document Names", icon: FileCheck },
      { id: "expiring-documents", label: "Expiring Documents", icon: FileWarning },
    ],
  },
  {
    section: "COMPANY & VESSEL",
    items: [
      { id: "vessels", label: "Vessel", icon: Anchor },
      { id: "principals", label: "Company / Principal", icon: Building2 },
    ],
  },
  {
    section: "REPORTS",
    items: [
      { id: "crew-reports", label: "Crew Reports", icon: BarChart3 },
      { id: "document-reports", label: "Document Reports", icon: ClipboardList },
    ],
  },
  {
    section: "SYSTEM",
    items: [
      { id: "users", label: "Users", icon: Users },
      { id: "settings", label: "Settings", icon: Settings },
    ],
  },
];

const BREADCRUMBS: Record<string, string[]> = {
  dashboard: ["Dashboard"],
  "crew-database": ["Crew Management", "Crew Database"],
  "search-crew": ["Crew Management", "Search Crew"],
  "available-crew": ["Crew Management", "Available Crew"],
  "onboard-crew": ["Crew Management", "Onboard Crew"],
  "joining-principal": ["Crew Management", "Joining Principal"],
  "waiting-crew": ["Crew Management", "Waiting Crew"],
  blacklist: ["Crew Management", "Blacklist"],
  "ex-crew": ["Crew Management", "Ex Crew"],
  "sign-on": ["Operations", "Sign On"],
  "sign-off": ["Operations", "Sign Off"],
  documents: ["Document Control", "Documents"],
  "expiring-documents": ["Document Control", "Expiring Documents"],
  "document-types": ["Document Control", "Document Types"],
  "document-names": ["Document Control", "Document Names"],
  vessels: ["Company & Vessel", "Vessel"],
  principals: ["Company & Vessel", "Company / Principal"],
  "principal-detail": ["Company & Vessel", "Principal Detail"],
  "crew-reports": ["Reports", "Crew Reports"],
  "document-reports": ["Reports", "Document Reports"],
  users: ["System", "Users"],
  settings: ["System", "Settings"],
  "crew-detail": ["Crew Management", "Crew Detail"],
  "crew-form": ["Crew Management", "Add / Edit Crew"],
};

// ─── Utilities ────────────────────────────────────────────────────────────────

function cn(...classes: (string | boolean | undefined | null)[]): string {
  return classes.filter(Boolean).join(" ");
}

function fmtDate(d: string | null): string {
  if (!d) return "—";
  return new Date(d).toLocaleDateString("en-US", { year: "numeric", month: "short", day: "2-digit" });
}

function statusBadgeClass(status: string): string {
  const m: Record<string, string> = {
    available: "bg-emerald-50 text-emerald-700 border-emerald-200",
    onboard: "bg-blue-50 text-blue-700 border-blue-200",
    waiting: "bg-amber-50 text-amber-700 border-amber-200",
    blacklisted: "bg-red-50 text-red-700 border-red-200",
    "ex-crew": "bg-slate-100 text-slate-600 border-slate-200",
    pending: "bg-orange-50 text-orange-700 border-orange-200",
    approved: "bg-cyan-50 text-cyan-700 border-cyan-200",
    completed: "bg-slate-100 text-slate-600 border-slate-200",
    cancelled: "bg-red-50 text-red-600 border-red-200",
    valid: "bg-emerald-50 text-emerald-700 border-emerald-200",
    expiring: "bg-amber-50 text-amber-700 border-amber-200",
    expired: "bg-red-50 text-red-700 border-red-200",
    active: "bg-emerald-50 text-emerald-700 border-emerald-200",
    released: "bg-slate-100 text-slate-600 border-slate-200",
    inactive: "bg-slate-100 text-slate-500 border-slate-200",
  };
  return m[status] || "bg-slate-100 text-slate-600 border-slate-200";
}

// ─── UI Primitives ────────────────────────────────────────────────────────────

function Badge({ status, label }: { status: string; label?: string }) {
  return (
    <span className={cn("inline-flex items-center px-2 py-0.5 rounded text-xs font-medium border capitalize", statusBadgeClass(status))}>
      {label ?? status.replace("-", " ")}
    </span>
  );
}

function Btn({
  children, variant = "primary", size = "md", onClick, type = "button", disabled = false, className = "",
}: {
  children: React.ReactNode; variant?: "primary" | "secondary" | "ghost" | "danger" | "outline";
  size?: "sm" | "md" | "lg"; onClick?: () => void; type?: "button" | "submit" | "reset";
  disabled?: boolean; className?: string;
}) {
  const base = "inline-flex items-center gap-1.5 font-medium rounded-lg transition-all focus:outline-none focus:ring-2 focus:ring-offset-1 cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed";
  const variants = {
    primary: "bg-blue-600 text-white hover:bg-blue-700 focus:ring-blue-500",
    secondary: "bg-white text-slate-700 border border-slate-200 hover:bg-slate-50 focus:ring-slate-300",
    ghost: "text-slate-600 hover:bg-slate-100 focus:ring-slate-300",
    danger: "bg-red-600 text-white hover:bg-red-700 focus:ring-red-500",
    outline: "border border-blue-600 text-blue-600 hover:bg-blue-50 focus:ring-blue-400",
  };
  const sizes = { sm: "px-2.5 py-1.5 text-xs", md: "px-3.5 py-2 text-sm", lg: "px-5 py-2.5 text-sm" };
  return (
    <button type={type} disabled={disabled} onClick={onClick} className={cn(base, variants[variant], sizes[size], className)}>
      {children}
    </button>
  );
}

function Input({
  label, value, onChange, placeholder = "", type = "text", error, required = false, className = "", readOnly = false,
}: {
  label?: string; value: string; onChange: (v: string) => void; placeholder?: string;
  type?: string; error?: string; required?: boolean; className?: string; readOnly?: boolean;
}) {
  return (
    <div className={cn("flex flex-col gap-1", className)}>
      {label && <label className="text-xs font-medium text-slate-600">{label}{required && <span className="text-red-500 ml-0.5">*</span>}</label>}
      <input
        type={type} value={value} readOnly={readOnly}
        onChange={e => onChange(e.target.value)}
        placeholder={placeholder}
        className={cn(
          "w-full px-3 py-2 text-sm bg-white border rounded-lg outline-none transition-colors",
          "placeholder:text-slate-400 focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500",
          error ? "border-red-400" : "border-slate-200",
          readOnly && "bg-slate-50 cursor-default"
        )}
      />
      {error && <span className="text-xs text-red-500">{error}</span>}
    </div>
  );
}

function Select({
  label, value, onChange, options, placeholder = "Select…", error, required = false, className = "",
}: {
  label?: string; value: string; onChange: (v: string) => void;
  options: { value: string; label: string }[]; placeholder?: string;
  error?: string; required?: boolean; className?: string;
}) {
  return (
    <div className={cn("flex flex-col gap-1", className)}>
      {label && <label className="text-xs font-medium text-slate-600">{label}{required && <span className="text-red-500 ml-0.5">*</span>}</label>}
      <select
        value={value} onChange={e => onChange(e.target.value)}
        className={cn(
          "w-full px-3 py-2 text-sm bg-white border rounded-lg outline-none transition-colors",
          "focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 cursor-pointer",
          error ? "border-red-400" : "border-slate-200"
        )}
      >
        <option value="">{placeholder}</option>
        {options.map(o => <option key={o.value} value={o.value}>{o.label}</option>)}
      </select>
      {error && <span className="text-xs text-red-500">{error}</span>}
    </div>
  );
}

function SelectWithOther({
  label, value, onChange, options, placeholder = "Select…", required = false, error, className = "",
}: {
  label?: string; value: string; onChange: (v: string) => void;
  options: { value: string; label: string }[]; placeholder?: string;
  required?: boolean; error?: string; className?: string;
}) {
  const isCustom = Boolean(value && !options.some(o => o.value === value));
  const [mode, setMode] = useState<string>(isCustom ? "__OTHER__" : value);

  useEffect(() => {
    if (!value) {
      setMode("");
    } else if (options.some(o => o.value === value)) {
      setMode(value);
    } else {
      setMode("__OTHER__");
    }
  }, [value, options]);

  return (
    <div className={cn("flex flex-col gap-1.5", className)}>
      <Select
        label={label}
        value={mode}
        onChange={v => {
          setMode(v);
          if (v !== "__OTHER__") {
            onChange(v);
          }
        }}
        options={[
          ...options,
          { value: "__OTHER__", label: "✨ Lainnya / Ketik Manual…" },
        ]}
        placeholder={placeholder}
        required={required}
        error={error}
      />
      {mode === "__OTHER__" && (
        <Input
          value={isCustom ? value : ""}
          onChange={v => onChange(v)}
          placeholder={`Ketik ${label ? label.replace("*", "").trim() : "opsi"} baru di sini…`}
          required={required}
        />
      )}
    </div>
  );
}

function Textarea({
  label, value, onChange, placeholder = "", rows = 3, className = "",
}: {
  label?: string; value: string; onChange: (v: string) => void;
  placeholder?: string; rows?: number; className?: string;
}) {
  return (
    <div className={cn("flex flex-col gap-1", className)}>
      {label && <label className="text-xs font-medium text-slate-600">{label}</label>}
      <textarea
        value={value} rows={rows} onChange={e => onChange(e.target.value)} placeholder={placeholder}
        className="w-full px-3 py-2 text-sm bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 resize-none placeholder:text-slate-400"
      />
    </div>
  );
}

function DynamicCustomFields({ fields, values, onChange }: { fields: CustomField[]; values: Record<string, unknown>; onChange: (values: Record<string, unknown>) => void }) {
  const activeFields = fields.filter(field => field.active);
  if (activeFields.length === 0) return null;

  function setValue(key: string, value: unknown) {
    onChange({ ...values, [key]: value });
  }

  return (
    <div className="col-span-2 grid grid-cols-2 gap-4 rounded-xl border border-slate-100 bg-slate-50 p-4">
      <div className="col-span-2 flex items-center gap-2">
        <Database size={13} className="text-slate-400" />
        <p className="text-xs font-semibold text-slate-700">Custom Fields</p>
      </div>
      {activeFields.map(field => {
        const value = customValueToString(values[field.key]);
        const options = parseCustomFieldOptions(field.optionsJson);
        if (["select", "dropdown"].includes(field.type)) {
          return <Select key={field.id} label={field.label} value={value} onChange={v => setValue(field.key, v)} required={field.required} options={options.map(option => ({ value: option, label: option }))} />;
        }
        if (["textarea", "multiline"].includes(field.type)) {
          return <Textarea key={field.id} label={field.label} value={value} onChange={v => setValue(field.key, v)} className="col-span-2" />;
        }
        if (["boolean", "checkbox"].includes(field.type)) {
          return <Select key={field.id} label={field.label} value={value} onChange={v => setValue(field.key, v === "true")} required={field.required} options={[{ value: "true", label: "Yes" }, { value: "false", label: "No" }]} />;
        }
        return <Input key={field.id} label={field.label} value={value} onChange={v => setValue(field.key, v)} required={field.required} type={field.type === "number" ? "number" : field.type === "date" ? "date" : "text"} />;
      })}
    </div>
  );
}
function SearchBar({ value, onChange, placeholder = "Search…" }: { value: string; onChange: (v: string) => void; placeholder?: string }) {
  return (
    <div className="relative">
      <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={15} />
      <input
        value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder}
        className="pl-9 pr-3 py-2 text-sm bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 w-64 placeholder:text-slate-400"
      />
    </div>
  );
}

function Skeleton({ className = "" }: { className?: string }) {
  return <div className={cn("animate-pulse bg-slate-200 rounded", className)} />;
}

function EmptyState({ title, description, action }: { title: string; description?: string; action?: React.ReactNode }) {
  return (
    <div className="flex flex-col items-center justify-center py-16 text-center">
      <div className="w-12 h-12 bg-slate-100 rounded-xl flex items-center justify-center mb-3">
        <Database size={22} className="text-slate-400" />
      </div>
      <p className="text-sm font-medium text-slate-700">{title}</p>
      {description && <p className="text-xs text-slate-400 mt-1">{description}</p>}
      {action && <div className="mt-4">{action}</div>}
    </div>
  );
}

function Pagination({ page, total, pageSize, onPage }: { page: number; total: number; pageSize: number; onPage: (p: number) => void }) {
  const pages = Math.ceil(total / pageSize);
  if (pages <= 1) return null;
  return (
    <div className="flex items-center justify-between px-4 py-3 border-t border-slate-100">
      <span className="text-xs text-slate-500">
        Showing {(page - 1) * pageSize + 1}–{Math.min(page * pageSize, total)} of {total}
      </span>
      <div className="flex items-center gap-1">
        <button onClick={() => onPage(page - 1)} disabled={page === 1}
          className="p-1.5 rounded hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed transition-colors">
          <ArrowLeft size={14} className="text-slate-600" />
        </button>
        {Array.from({ length: Math.min(pages, 5) }, (_, i) => {
          const p = i + 1;
          return (
            <button key={p} onClick={() => onPage(p)}
              className={cn("px-2.5 py-1 text-xs rounded font-medium transition-colors", page === p ? "bg-blue-600 text-white" : "hover:bg-slate-100 text-slate-600")}>
              {p}
            </button>
          );
        })}
        <button onClick={() => onPage(page + 1)} disabled={page === pages}
          className="p-1.5 rounded hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed transition-colors">
          <ArrowRight size={14} className="text-slate-600" />
        </button>
      </div>
    </div>
  );
}

function Modal({
  open, onClose, title, children, size = "md", footer,
}: {
  open: boolean; onClose: () => void; title: string; children: React.ReactNode;
  size?: "sm" | "md" | "lg" | "xl"; footer?: React.ReactNode;
}) {
  if (!open) return null;
  const widths = { sm: "max-w-sm", md: "max-w-lg", lg: "max-w-2xl", xl: "max-w-4xl" };
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
      <div className="absolute inset-0 bg-black/40 backdrop-blur-sm" onClick={onClose} />
      <div className={cn("relative bg-white rounded-xl shadow-2xl w-full flex flex-col max-h-[90vh]", widths[size])}>
        <div className="flex items-center justify-between px-5 py-4 border-b border-slate-100">
          <h3 className="font-semibold text-slate-800 text-sm">{title}</h3>
          <button onClick={onClose} className="p-1 rounded-md hover:bg-slate-100 transition-colors"><X size={16} className="text-slate-500" /></button>
        </div>
        <div className="overflow-y-auto flex-1 px-5 py-4">{children}</div>
        {footer && <div className="px-5 py-3 border-t border-slate-100 flex items-center justify-end gap-2 bg-slate-50/50 rounded-b-xl">{footer}</div>}
      </div>
    </div>
  );
}

function ConfirmModal({ open, onClose, onConfirm, title, message, danger = false }: {
  open: boolean; onClose: () => void; onConfirm: () => void;
  title: string; message: string; danger?: boolean;
}) {
  return (
    <Modal open={open} onClose={onClose} title={title} size="sm"
      footer={<><Btn variant="secondary" size="sm" onClick={onClose}>Cancel</Btn><Btn variant={danger ? "danger" : "primary"} size="sm" onClick={() => { onConfirm(); onClose(); }}>Confirm</Btn></>}>
      <p className="text-sm text-slate-600">{message}</p>
    </Modal>
  );
}

function KPICard({ label, value, icon: Icon, color, trend, sub }: {
  label: string; value: string | number; icon: React.ElementType;
  color: string; trend?: string; sub?: string;
}) {
  return (
    <div className="bg-white rounded-xl border border-slate-100 p-4 flex flex-col gap-3 hover:shadow-md transition-shadow">
      <div className="flex items-center justify-between">
        <span className="text-xs font-medium text-slate-500 uppercase tracking-wide">{label}</span>
        <div className={cn("w-8 h-8 rounded-lg flex items-center justify-center", color)}>
          <Icon size={16} className="text-white" />
        </div>
      </div>
      <div>
        <span className="text-2xl font-bold text-slate-800">{value}</span>
        {sub && <p className="text-xs text-slate-400 mt-0.5">{sub}</p>}
      </div>
      {trend && (
        <div className="flex items-center gap-1 text-xs text-emerald-600">
          <TrendingUp size={11} />
          <span>{trend}</span>
        </div>
      )}
    </div>
  );
}

function PageHeader({
  title, children,
}: { title: string; children?: React.ReactNode }) {
  return (
    <div className="flex items-center justify-between mb-5">
      <h1 className="text-lg font-semibold text-slate-800">{title}</h1>
      {children && <div className="flex items-center gap-2">{children}</div>}
    </div>
  );
}

function Card({ children, className = "" }: { children: React.ReactNode; className?: string }) {
  return <div className={cn("bg-white rounded-xl border border-slate-100", className)}>{children}</div>;
}

class AppErrorBoundary extends React.Component<{ children: React.ReactNode }, { error: Error | null }> {
  state = { error: null as Error | null };

  static getDerivedStateFromError(error: Error) {
    return { error };
  }

  componentDidCatch(error: Error) {
    console.error("SKYagen render error", error);
  }

  render() {
    if (!this.state.error) return this.props.children;
    return (
      <div className="flex min-h-[420px] items-center justify-center p-6">
        <div className="max-w-md rounded-xl border border-red-100 bg-white p-5 shadow-sm">
          <div className="mb-3 flex items-center gap-2 text-red-600">
            <AlertTriangle size={18} />
            <h2 className="text-sm font-semibold">Page failed to render</h2>
          </div>
          <p className="text-xs leading-relaxed text-slate-500">A data format issue stopped this page from rendering. The app stayed open so the issue can be recovered without a blank screen.</p>
          <pre className="mt-3 max-h-32 overflow-auto rounded-lg bg-slate-50 p-3 text-[11px] text-slate-500">{this.state.error.message}</pre>
          <div className="mt-4 flex justify-end gap-2">
            <Btn variant="secondary" size="sm" onClick={() => this.setState({ error: null })}>Try Again</Btn>
            <Btn variant="primary" size="sm" onClick={() => window.location.reload()}>Reload App</Btn>
          </div>
        </div>
      </div>
    );
  }
}
function ApiConnectionModal() {
  const [open, setOpen] = useState(false);
  const [baseUrl, setBaseUrlState] = useState(() => getApiBaseUrl());

  useEffect(() => {
    const handler = () => {
      setBaseUrlState(getApiBaseUrl());
      setOpen(true);
    };
    window.addEventListener("skyagen:api-disconnected", handler);
    return () => window.removeEventListener("skyagen:api-disconnected", handler);
  }, []);

  function save() {
    setApiBaseUrl(baseUrl);
    toast.success("API Base URL saved");
    setOpen(false);
    window.location.reload();
  }

  return (
    <Modal open={open} onClose={() => {}} title="API Connection Lost" size="sm"
      footer={<><Btn variant="secondary" size="sm" onClick={() => setBaseUrlState("http://localhost:8080")}>Use Local</Btn><Btn variant="primary" size="sm" onClick={save}>Save & Retry</Btn></>}>
      <div className="space-y-3">
        <div className="flex items-start gap-3 p-3 bg-red-50 border border-red-100 rounded-xl">
          <AlertTriangle size={16} className="text-red-500 mt-0.5 shrink-0" />
          <p className="text-xs text-red-700">Frontend cannot connect to the API. Enter the backend Base URL below. This URL will be used until another disconnection happens.</p>
        </div>
        <Input label="API Base URL" value={baseUrl} onChange={setBaseUrlState} placeholder="http://localhost:8080" required />
      </div>
    </Modal>
  );
}
function ApiPendingBanner({ endpoint }: { endpoint?: string }) {
  return (
    <div className="flex items-start gap-3 p-4 bg-amber-50 border border-amber-200 rounded-xl mb-4">
      <AlertTriangle size={16} className="text-amber-500 mt-0.5 shrink-0" />
      <div>
        <p className="text-sm font-medium text-amber-800">API Integration Pending</p>
        {endpoint && <p className="text-xs text-amber-600 mt-0.5 font-mono">{endpoint}</p>}
        <p className="text-xs text-amber-600 mt-1">This section will be connected to the backend once the endpoint is available.</p>
      </div>
    </div>
  );
}

function SectionLabel({ children }: { children: React.ReactNode }) {
  return <p className="text-[10px] font-semibold text-slate-400 uppercase tracking-widest mb-1 mt-4 px-1 first:mt-0">{children}</p>;
}

// ─── Sidebar ──────────────────────────────────────────────────────────────────

function Sidebar({ currentPage, setPage, collapsed, setCollapsed }: {
  currentPage: string; setPage: (p: string) => void;
  collapsed: boolean; setCollapsed: (v: boolean) => void;
}) {
  const [expanded, setExpanded] = useState<Set<string>>(new Set(["CREW MANAGEMENT", "OPERATIONS", "DOCUMENT CONTROL", "COMPANY & VESSEL", "REPORTS", "SYSTEM"]));

  function toggle(section: string) {
    setExpanded(prev => {
      const next = new Set(prev);
      if (next.has(section)) next.delete(section);
      else next.add(section);
      return next;
    });
  }

  return (
    <aside
      className="flex flex-col border-r border-slate-200 bg-white shrink-0 transition-all duration-200 z-30"
      style={{ width: collapsed ? 56 : 220 }}
    >
      {/* Logo */}
      <div className="flex items-center h-14 px-3 border-b border-slate-100 shrink-0">
        <div className="flex items-center gap-2.5 overflow-hidden">
          <img src={skyagenLogoUrl} alt="SKYagen" className="w-8 h-8 rounded-lg object-contain bg-slate-50 p-0.5 border border-slate-200 shrink-0" />
          {!collapsed && (
            <div className="overflow-hidden">
              <span className="text-slate-900 font-bold text-sm tracking-wide block">SKYagen</span>
              <span className="text-slate-400 text-[10px] block -mt-0.5 font-medium">Crew Management</span>
            </div>
          )}
        </div>
      </div>

      {/* Nav */}
      <nav className="flex-1 overflow-y-auto py-3 px-2 scrollbar-none">
        {NAV_SECTIONS.map(({ section, items }) => (
          <div key={section ?? "main"}>
            {section && !collapsed && (
              <button
                onClick={() => toggle(section)}
                className="w-full flex items-center justify-between px-2 py-1 mb-1 mt-3"
              >
                <SectionLabel>{section}</SectionLabel>
                {expanded.has(section)
                  ? <ChevronUp size={11} className="text-slate-400 shrink-0" />
                  : <ChevronDown size={11} className="text-slate-400 shrink-0" />
                }
              </button>
            )}
            {section && collapsed && <div className="my-2 border-t border-slate-100" />}
            {(!section || collapsed || expanded.has(section)) && items.map(({ id, label, icon: Icon }) => {
              const active = currentPage === id;
              return (
                <button
                  key={id}
                  onClick={() => setPage(id)}
                  title={collapsed ? label : undefined}
                  className={cn(
                    "w-full flex items-center gap-2.5 px-2.5 py-2 rounded-lg text-sm transition-all mb-0.5 font-medium",
                    active
                      ? "bg-blue-600 text-white shadow-xs font-semibold"
                      : "text-slate-600 hover:bg-slate-50 hover:text-slate-900"
                  )}
                >
                  <Icon size={16} className={cn("shrink-0", active ? "text-white" : "text-slate-500")} />
                  {!collapsed && <span className="text-[13px] truncate">{label}</span>}
                </button>
              );
            })}
          </div>
        ))}
      </nav>

      {/* Collapse toggle */}
      <div className="border-t border-slate-100 p-2">
        <button
          onClick={() => setCollapsed(!collapsed)}
          className="w-full flex items-center justify-center gap-2 px-2 py-2 rounded-lg text-slate-500 hover:text-slate-800 hover:bg-slate-100 transition-colors"
        >
          {collapsed ? <ChevronRight size={15} /> : <><ChevronLeft size={15} /><span className="text-xs font-medium">Collapse</span></>}
        </button>
      </div>
    </aside>
  );
}

// ─── Topbar ───────────────────────────────────────────────────────────────────

function Topbar({ currentPage, onLogout, onSearch }: { currentPage: string; onLogout: () => void; onSearch?: (v: string) => void }) {
  const [profileOpen, setProfileOpen] = useState(false);
  const [notifOpen, setNotifOpen] = useState(false);
  const [search, setSearch] = useState("");
  const crumbs = BREADCRUMBS[currentPage] ?? [currentPage];

  return (
    <header className="h-14 bg-white border-b border-slate-100 flex items-center justify-between px-5 shrink-0 z-20">
      {/* Breadcrumb */}
      <div className="flex items-center gap-1.5 text-sm">
        {crumbs.map((c, i) => (
          <React.Fragment key={i}>
            {i > 0 && <ChevronRight size={13} className="text-slate-300" />}
            <span className={i === crumbs.length - 1 ? "text-slate-800 font-medium" : "text-slate-400"}>
              {c}
            </span>
          </React.Fragment>
        ))}
      </div>

      {/* Right controls */}
      <div className="flex items-center gap-2">
        {/* Search */}
        <div className="relative">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={13} />
          <input
            value={search}
            onChange={e => { setSearch(e.target.value); onSearch?.(e.target.value); }}
            placeholder="Quick search…"
            className="pl-8 pr-3 py-1.5 text-xs bg-slate-50 border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-400 w-48 placeholder:text-slate-400"
          />
        </div>

        {/* Notifications */}
        <div className="relative">
          <button
            onClick={() => setNotifOpen(!notifOpen)}
            className="relative p-2 rounded-lg hover:bg-slate-100 transition-colors"
          >
            <Bell size={17} className="text-slate-600" />
            <span className="absolute top-1 right-1 w-4 h-4 bg-red-500 rounded-full text-[9px] text-white flex items-center justify-center font-bold">3</span>
          </button>
          {notifOpen && (
            <>
              <div className="fixed inset-0 z-30" onClick={() => setNotifOpen(false)} />
              <div className="absolute right-0 top-10 w-80 bg-white rounded-xl shadow-xl border border-slate-100 z-40 overflow-hidden">
                <div className="flex items-center justify-between px-4 py-3 border-b border-slate-100">
                  <span className="text-sm font-semibold text-slate-800">Notifications</span>
                  <span className="text-[10px] bg-red-50 text-red-600 px-1.5 py-0.5 rounded font-medium">3 new</span>
                </div>
                {[
                  { icon: FileWarning, color: "text-amber-500", text: "Budi Santoso — Medical Cert expires in 12 days", time: "10m ago" },
                  { icon: AlertCircle, color: "text-red-500", text: "Juan Dela Cruz — GMDSS Certificate expired", time: "1h ago" },
                  { icon: UserCheck, color: "text-blue-500", text: "Jose Garcia joining application approved", time: "3h ago" },
                ].map(({ icon: Icon, color, text, time }, i) => (
                  <div key={i} className="flex gap-3 px-4 py-3 hover:bg-slate-50 border-b border-slate-50 cursor-pointer">
                    <Icon size={16} className={cn("mt-0.5 shrink-0", color)} />
                    <div className="flex-1 min-w-0">
                      <p className="text-xs text-slate-700 leading-snug">{text}</p>
                      <p className="text-[10px] text-slate-400 mt-0.5">{time}</p>
                    </div>
                  </div>
                ))}
              </div>
            </>
          )}
        </div>

        {/* Avatar */}
        <div className="relative">
          <button
            onClick={() => setProfileOpen(!profileOpen)}
            className="flex items-center gap-2 pl-1 pr-2 py-1 rounded-lg hover:bg-slate-100 transition-colors"
          >
            <div className="w-7 h-7 bg-blue-600 rounded-full flex items-center justify-center">
              <span className="text-white text-xs font-bold">A</span>
            </div>
            <div className="text-left">
              <p className="text-xs font-medium text-slate-700 leading-none">Admin User</p>
              <p className="text-[10px] text-slate-400 mt-0.5">Super Admin</p>
            </div>
            <ChevronDown size={12} className="text-slate-400" />
          </button>
          {profileOpen && (
            <>
              <div className="fixed inset-0 z-30" onClick={() => setProfileOpen(false)} />
              <div className="absolute right-0 top-10 w-44 bg-white rounded-xl shadow-xl border border-slate-100 z-40 overflow-hidden py-1">
                {[
                  { icon: User, label: "Profile" },
                  { icon: Lock, label: "Change Password" },
                  { icon: Settings, label: "Settings" },
                ].map(({ icon: Icon, label }) => (
                  <button key={label} className="w-full flex items-center gap-2.5 px-3 py-2 text-xs text-slate-600 hover:bg-slate-50 transition-colors">
                    <Icon size={13} className="text-slate-400" />
                    {label}
                  </button>
                ))}
                <div className="border-t border-slate-100 mt-1 pt-1">
                  <button onClick={onLogout} className="w-full flex items-center gap-2.5 px-3 py-2 text-xs text-red-600 hover:bg-red-50 transition-colors">
                    <LogOut size={13} />
                    Sign Out
                  </button>
                </div>
              </div>
            </>
          )}
        </div>
      </div>
    </header>
  );
}

// ─── Layout ───────────────────────────────────────────────────────────────────

function AdminLayout({ currentPage, setPage, onLogout, children }: {
  currentPage: string; setPage: (p: string) => void;
  onLogout: () => void; children: React.ReactNode;
}) {
  const [collapsed, setCollapsed] = useState(false);
  return (
    <div className="flex h-screen overflow-hidden bg-slate-100">
      <Sidebar currentPage={currentPage} setPage={setPage} collapsed={collapsed} setCollapsed={setCollapsed} />
      <div className="flex flex-col flex-1 overflow-hidden">
        <Topbar currentPage={currentPage} onLogout={onLogout} />
        <main className="flex-1 overflow-y-auto p-5">
          {children}
        </main>
      </div>
    </div>
  );
}

// ─── Auth Pages ───────────────────────────────────────────────────────────────

function AuthShell({ children }: { children: React.ReactNode }) {
  return (
    <div className="min-h-screen bg-slate-900 flex">
      {/* Left panel */}
      <div className="hidden lg:flex w-96 flex-col justify-between p-10" style={{ background: "#0D1B2E" }}>
        <div className="flex items-center gap-3">
          <img src={skyagenLogoUrl} alt="SKYagen" className="w-12 h-12 rounded-xl object-contain bg-white p-1 shadow-sm" />
          <div>
            <p className="text-white font-bold text-base">SKYagen</p>
            <p className="text-slate-400 text-xs">Crew Management System</p>
          </div>
        </div>
        <div>
          <blockquote className="text-slate-300 text-sm leading-relaxed italic border-l-2 border-blue-500 pl-4 mb-6">
            "Streamlining maritime crew operations — from sign-on to sign-off, every seafarer accounted for."
          </blockquote>
          <div className="grid grid-cols-2 gap-3">
            {[
              { label: "Active Crew", value: "14" },
              { label: "Vessels", value: "7" },
              { label: "Principals", value: "5" },
              { label: "Expiring Docs", value: "6" },
            ].map(({ label, value }) => (
              <div key={label} className="bg-white/5 rounded-lg p-3 border border-white/10">
                <p className="text-xl font-bold text-white">{value}</p>
                <p className="text-xs text-slate-400 mt-0.5">{label}</p>
              </div>
            ))}
          </div>
        </div>
        <p className="text-slate-600 text-xs">© 2024 SKYagen. All rights reserved.</p>
      </div>
      {/* Right panel */}
      <div className="flex-1 flex items-center justify-center p-8 bg-slate-50">
        <div className="w-full max-w-sm">{children}</div>
      </div>
    </div>
  );
}

function LoginPage({ setAuth, setAuthPage }: { setAuth: () => void; setAuthPage: (p: AuthPage) => void }) {
  const [email, setEmail] = useState("admin@skyagen.com");
  const [password, setPassword] = useState("password123");
  const [loading, setLoading] = useState(false);
  const [errors, setErrors] = useState<Record<string, string>>({});

  async function submit(e: React.FormEvent) {
    e.preventDefault();
    const errs: Record<string, string> = {};
    if (!email) errs.email = "Email is required";
    if (!password) errs.password = "Password is required";
    if (Object.keys(errs).length) { setErrors(errs); return; }
    setLoading(true);
    try {
      await login(email, password);
      toast.success("Welcome back");
      setAuth();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Login failed");
    } finally {
      setLoading(false);
    }
  }

  return (
    <AuthShell>
      <div className="mb-7">
        <div className="flex lg:hidden items-center gap-2 mb-6">
          <img src={skyagenLogoUrl} alt="SKYagen" className="w-9 h-9 rounded-xl object-contain bg-white p-1 shadow-sm" />
          <span className="font-bold text-slate-800">SKYagen</span>
        </div>
        <h2 className="text-2xl font-bold text-slate-800">Sign in</h2>
        <p className="text-slate-500 text-sm mt-1">Enter your credentials to access the system</p>
      </div>
      <form onSubmit={submit} className="flex flex-col gap-4">
        <Input label="Email address" type="email" value={email} onChange={setEmail} placeholder="admin@skyagen.com" required error={errors.email} />
        <div className="flex flex-col gap-1">
          <div className="flex items-center justify-between">
            <label className="text-xs font-medium text-slate-600">Password<span className="text-red-500 ml-0.5">*</span></label>
            <button type="button" onClick={() => setAuthPage("update-password")} className="text-xs text-blue-600 hover:underline">Forgot password?</button>
          </div>
          <input
            type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="••••••••"
            className={cn("w-full px-3 py-2 text-sm bg-white border rounded-lg outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 placeholder:text-slate-400", errors.password ? "border-red-400" : "border-slate-200")}
          />
          {errors.password && <span className="text-xs text-red-500">{errors.password}</span>}
        </div>
        <Btn type="submit" variant="primary" size="lg" disabled={loading} className="w-full justify-center mt-1">
          {loading ? <><RefreshCw size={14} className="animate-spin" />Signing in…</> : <>Sign In <ArrowRight size={14} /></>}
        </Btn>
      </form>
      <p className="text-center text-xs text-slate-400 mt-5">
        Don&apos;t have an account?{" "}
        <button onClick={() => setAuthPage("register")} className="text-blue-600 hover:underline font-medium">Register</button>
      </p>
    </AuthShell>
  );
}

function RegisterPage({ setAuthPage }: { setAuthPage: (p: AuthPage) => void }) {
  const [form, setForm] = useState({ name: "", email: "", password: "", confirm: "", role: "" });
  const [loading, setLoading] = useState(false);

  async function submit(e: React.FormEvent) {
    e.preventDefault();
    if (form.password !== form.confirm) {
      toast.error("Password confirmation does not match");
      return;
    }
    setLoading(true);
    try {
      await registerUser({ name: form.name, email: form.email, password: form.password, role: form.role || "admin" });
      toast.success("Account registered. Please verify email before login.");
      setAuthPage("verify-email");
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Register failed");
    } finally {
      setLoading(false);
    }
  }

  function f(k: string) { return (v: string) => setForm(p => ({ ...p, [k]: v })); }

  return (
    <AuthShell>
      <div className="mb-6">
        <h2 className="text-2xl font-bold text-slate-800">Create account</h2>
        <p className="text-slate-500 text-sm mt-1">Register a new user account</p>
      </div>
      <form onSubmit={submit} className="flex flex-col gap-3.5">
        <Input label="Full Name" value={form.name} onChange={f("name")} placeholder="John Smith" required />
        <Input label="Email address" type="email" value={form.email} onChange={f("email")} placeholder="john@skyagen.com" required />
        <Select label="Role" value={form.role} onChange={f("role")} required
          options={[{ value: "admin", label: "Admin" }, { value: "operator", label: "Operator" }, { value: "viewer", label: "Viewer" }]} />
        <Input label="Password" type="password" value={form.password} onChange={f("password")} placeholder="Min. 8 characters" required />
        <Input label="Confirm Password" type="password" value={form.confirm} onChange={f("confirm")} placeholder="Repeat password" required />
        <Btn type="submit" variant="primary" size="lg" disabled={loading} className="w-full justify-center mt-1">
          {loading ? <><RefreshCw size={14} className="animate-spin" />Registering…</> : "Create Account"}
        </Btn>
      </form>
      <p className="text-center text-xs text-slate-400 mt-4">
        Already have an account?{" "}
        <button onClick={() => setAuthPage("login")} className="text-blue-600 hover:underline font-medium">Sign in</button>
      </p>
    </AuthShell>
  );
}

function VerifyEmailPage({ setAuthPage }: { setAuthPage: (p: AuthPage) => void }) {
  return (
    <AuthShell>
      <div className="text-center">
        <div className="w-16 h-16 bg-blue-50 rounded-2xl flex items-center justify-center mx-auto mb-5">
          <Mail size={28} className="text-blue-600" />
        </div>
        <h2 className="text-xl font-bold text-slate-800 mb-2">Check your email</h2>
        <p className="text-slate-500 text-sm mb-5">We sent a verification link to your email address. Click the link to activate your account.</p>
        <Btn variant="primary" size="md" className="w-full justify-center" onClick={() => setAuthPage("verify-success")}>
          Simulate Verification Success
        </Btn>
        <button onClick={() => setAuthPage("verify-failed")} className="block text-xs text-slate-400 hover:text-slate-600 mt-3 mx-auto">
          Simulate Verification Failed
        </button>
        <button onClick={() => setAuthPage("login")} className="block text-xs text-blue-600 hover:underline mt-3 mx-auto">
          Back to sign in
        </button>
      </div>
    </AuthShell>
  );
}

function VerifySuccessPage({ setAuthPage }: { setAuthPage: (p: AuthPage) => void }) {
  return (
    <AuthShell>
      <div className="text-center">
        <div className="w-16 h-16 bg-emerald-50 rounded-2xl flex items-center justify-center mx-auto mb-5">
          <CheckCircle size={28} className="text-emerald-500" />
        </div>
        <h2 className="text-xl font-bold text-slate-800 mb-2">Email verified!</h2>
        <p className="text-slate-500 text-sm mb-6">Your account has been successfully activated. You can now sign in.</p>
        <Btn variant="primary" size="lg" className="w-full justify-center" onClick={() => setAuthPage("login")}>
          Go to Sign In <ArrowRight size={14} />
        </Btn>
      </div>
    </AuthShell>
  );
}

function VerifyFailedPage({ setAuthPage }: { setAuthPage: (p: AuthPage) => void }) {
  return (
    <AuthShell>
      <div className="text-center">
        <div className="w-16 h-16 bg-red-50 rounded-2xl flex items-center justify-center mx-auto mb-5">
          <XCircle size={28} className="text-red-500" />
        </div>
        <h2 className="text-xl font-bold text-slate-800 mb-2">Verification failed</h2>
        <p className="text-slate-500 text-sm mb-6">The verification link is invalid or has expired. Please request a new one.</p>
        <Btn variant="primary" size="lg" className="w-full justify-center mb-2" onClick={() => setAuthPage("verify-email")}>
          Resend Verification Email
        </Btn>
        <Btn variant="ghost" size="md" className="w-full justify-center" onClick={() => setAuthPage("login")}>
          Back to Sign In
        </Btn>
      </div>
    </AuthShell>
  );
}

function UpdatePasswordPage({ setAuthPage }: { setAuthPage: (p: AuthPage) => void }) {
  const [form, setForm] = useState({ current: "", next: "", confirm: "" });
  const [loading, setLoading] = useState(false);

  function submit(e: React.FormEvent) {
    e.preventDefault();
    setLoading(true);
    setTimeout(() => { setLoading(false); toast.success("Password updated successfully"); setAuthPage("login"); }, 1000);
  }
  function f(k: string) { return (v: string) => setForm(p => ({ ...p, [k]: v })); }

  return (
    <AuthShell>
      <div className="mb-6">
        <h2 className="text-xl font-bold text-slate-800">Update password</h2>
        <p className="text-slate-500 text-sm mt-1">Choose a new secure password</p>
      </div>
      <form onSubmit={submit} className="flex flex-col gap-4">
        <Input label="Current Password" type="password" value={form.current} onChange={f("current")} placeholder="••••••••" required />
        <Input label="New Password" type="password" value={form.next} onChange={f("next")} placeholder="Min. 8 characters" required />
        <Input label="Confirm New Password" type="password" value={form.confirm} onChange={f("confirm")} placeholder="Repeat new password" required />
        <Btn type="submit" variant="primary" size="lg" disabled={loading} className="w-full justify-center">
          {loading ? <><RefreshCw size={14} className="animate-spin" />Updating…</> : "Update Password"}
        </Btn>
        <button type="button" onClick={() => setAuthPage("login")} className="text-xs text-slate-400 hover:text-slate-600 text-center">
          Back to Sign In
        </button>
      </form>
    </AuthShell>
  );
}

function ErrorPage({ code, setAuthPage }: { code: 401 | 403 | 404; setAuthPage: (p: AuthPage) => void }) {
  const info = {
    401: { icon: Lock, title: "Unauthorized", desc: "You need to sign in to access this resource.", color: "text-amber-500", bg: "bg-amber-50" },
    403: { icon: Shield, title: "Forbidden", desc: "You don't have permission to view this page.", color: "text-red-500", bg: "bg-red-50" },
    404: { icon: AlertCircle, title: "Page Not Found", desc: "The page you're looking for doesn't exist or has been moved.", color: "text-slate-400", bg: "bg-slate-100" },
  }[code];

  return (
    <AuthShell>
      <div className="text-center">
        <div className={cn("w-16 h-16 rounded-2xl flex items-center justify-center mx-auto mb-4", info.bg)}>
          <info.icon size={28} className={info.color} />
        </div>
        <span className="text-5xl font-black text-slate-200 block mb-2">{code}</span>
        <h2 className="text-xl font-bold text-slate-800 mb-2">{info.title}</h2>
        <p className="text-slate-500 text-sm mb-6">{info.desc}</p>
        <Btn variant="primary" size="lg" className="w-full justify-center" onClick={() => setAuthPage("login")}>
          {code === 401 ? "Sign In" : "Go Home"}
        </Btn>
      </div>
    </AuthShell>
  );
}

// ─── Dashboard ────────────────────────────────────────────────────────────────

function DashboardPage({ setPage }: { setPage: (p: string) => void }) {
  const [summary, setSummary] = useState<ApiDashboardSummary | null>(null);
  const [charts, setCharts] = useState<ApiDashboardCharts | null>(null);
  const [docs, setDocs] = useState<ApiExpiringDoc[]>([]);
  const [joinings, setJoinings] = useState<ApiJoining[]>([]);
  const [crews, setCrews] = useState<Crew[]>([]);
  const [lastSync, setLastSync] = useState<string>("");

  const fetchData = useCallback(() => {
    Promise.all([
      apiGet<ApiDashboardSummary>("/api/dashboard/summary").catch(() => null),
      apiGet<ApiDashboardCharts>("/api/dashboard/charts").catch(() => null),
      apiGet<{ data: ApiExpiringDoc[] }>("/api/dashboard/recent-documents").catch(() => ({ data: [] })),
      apiGet<{ data: ApiJoining[] }>("/api/dashboard/recent-joinings").catch(() => ({ data: [] })),
      apiGet<ApiListResponse<ApiSeafarer>>("/api/seafarers").catch(() => ({ data: [] })),
      apiGet<ApiListResponse<ApiJoining>>("/api/joining-principals").catch(() => ({ data: [] })),
    ]).then(([summaryRes, chartRes, docRes, joiningRes, seafarersRes, allJoiningsRes]) => {
      if (summaryRes) setSummary(summaryRes);
      if (chartRes) setCharts(chartRes);
      setDocs(docRes.data ?? []);
      setJoinings(joiningRes.data ?? (allJoiningsRes.data ?? []));
      setCrews((seafarersRes.data ?? []).map(mapCrew));
      setLastSync(new Date().toLocaleTimeString());
    }).catch(() => {});
  }, []);

  useEffect(() => {
    fetchData();
    const timer = setInterval(fetchData, 5000);
    return () => clearInterval(timer);
  }, [fetchData]);

  // Compute REAL-TIME counts from actual seafarers list
  const totalCrewCount = crews.length || summary?.total_crew || 0;
  const availableCount = crews.filter(c => (c.status || "available").toLowerCase() === "available").length;
  const onboardCount = crews.filter(c => String(c.status).toLowerCase() === "onboard").length;
  const waitingCount = crews.filter(c => String(c.status).toLowerCase() === "waiting").length;
  const blacklistedCount = crews.filter(c => String(c.status).toLowerCase() === "blacklisted").length;
  const joiningsCount = joinings.length || summary?.total_joining || 0;

  const chartData = (charts?.months ?? []).map(m => ({ month: m.month, signOn: m.sign_on_count, signOff: m.sign_off_count }));
  const statusDist = [
    { name: "Available", value: availableCount, color: "#22c55e" },
    { name: "Onboard", value: onboardCount, color: "#3b82f6" },
    { name: "Waiting", value: waitingCount, color: "#f59e0b" },
    { name: "Blacklisted", value: blacklistedCount, color: "#ef4444" },
  ];

  const tooltip = ({ active, payload, label }: any) => active && payload?.length ? (
    <div className="bg-white border border-slate-100 rounded-xl shadow-lg p-3 text-xs">
      <p className="font-medium text-slate-700 mb-1">{label}</p>
      {payload.map((e: any) => (
        <div key={e.dataKey} className="flex items-center gap-2">
          <span className="w-2 h-2 rounded-full" style={{ background: e.color }} />
          <span className="text-slate-500 capitalize">{e.dataKey}:</span>
          <span className="font-semibold text-slate-800">{e.value}</span>
        </div>
      ))}
    </div>
  ) : null;

  return (
    <div>
      <div className="flex items-center justify-between mb-4">
        <div>
          <h1 className="text-lg font-semibold text-slate-800">Dashboard</h1>
        </div>
        <div className="flex items-center gap-2 bg-emerald-50 border border-emerald-200 text-emerald-700 px-3 py-1 rounded-full text-xs font-semibold shadow-2xs">
          <span className="h-2 w-2 rounded-full bg-emerald-500 animate-pulse" />
          <span>Realtime Live Sync</span>
          {lastSync && <span className="text-[10px] opacity-75 font-mono ml-1">({lastSync})</span>}
        </div>
      </div>

      <div className="grid grid-cols-6 gap-3 mb-5">
        <KPICard label="Total Crew" value={totalCrewCount} icon={Users} color="bg-slate-700" sub="All registered seafarers" />
        <KPICard label="Available" value={availableCount} icon={UserCheck} color="bg-emerald-600" />
        <KPICard label="Onboard" value={onboardCount} icon={Ship} color="bg-blue-600" sub="Currently deployed" />
        <KPICard label="Waiting" value={waitingCount} icon={Clock} color="bg-amber-500" />
        <KPICard label="Blacklisted" value={blacklistedCount} icon={Ban} color="bg-red-500" />
        <KPICard label="Joinings" value={joiningsCount} icon={Briefcase} color="bg-orange-500" />
      </div>

      <div className="grid grid-cols-3 gap-4 mb-4">
        <Card className="col-span-2 p-5">
          <div className="flex items-center justify-between mb-4">
            <div>
              <h3 className="font-semibold text-slate-800 text-sm">Crew Activity (Monthly)</h3>
              <p className="text-xs text-slate-400 mt-0.5">{charts?.period ?? "Last 12 months"}</p>
            </div>
            <Btn variant="secondary" size="sm" onClick={() => exportExcel("/api/reports/export?type=joining", undefined, "joining-report.xlsx").catch(err => toast.error(err.message))}>
              <Download size={12} />Export
            </Btn>
          </div>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={chartData} barSize={14} barGap={4}>
              <CartesianGrid strokeDasharray="3 3" stroke="#F1F5F9" vertical={false} />
              <XAxis dataKey="month" tick={{ fontSize: 11, fill: "#94A3B8" }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: "#94A3B8" }} axisLine={false} tickLine={false} width={24} />
              <Tooltip content={tooltip} />
              <Bar dataKey="signOn" name="Sign On" fill="#3B82F6" radius={[3, 3, 0, 0]} />
              <Bar dataKey="signOff" name="Sign Off" fill="#94A3B8" radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </Card>

        <Card className="p-5">
          <h3 className="font-semibold text-slate-800 text-sm mb-1">Crew Status Distribution</h3>
          <ResponsiveContainer width="100%" height={150}>
            <PieChart>
              <Pie data={statusDist} cx="50%" cy="50%" innerRadius={45} outerRadius={65} paddingAngle={3} dataKey="value">
                {statusDist.map((entry, i) => <Cell key={i} fill={entry.color} />)}
              </Pie>
              <Tooltip />
            </PieChart>
          </ResponsiveContainer>
          {statusDist.map(({ name, value, color }) => (
            <div key={name} className="flex items-center justify-between">
              <span className="text-xs text-slate-600">
                <span className="inline-block w-2 h-2 rounded-full mr-1" style={{ background: color }} />
                {name}
              </span>
              <span className="text-xs font-semibold text-slate-700">{value}</span>
            </div>
          ))}
        </Card>
      </div>

      <div className="grid grid-cols-2 gap-4">
        <Card>
          <div className="flex items-center justify-between px-4 py-3 border-b border-slate-100">
            <h3 className="font-semibold text-slate-800 text-sm">Expiring Documents</h3>
            <Btn variant="ghost" size="sm" onClick={() => setPage("expiring-documents")}>
              View all <ArrowRight size={12} />
            </Btn>
          </div>
          <div className="divide-y divide-slate-50">
            {docs.slice(0, 5).map(doc => (
              <div key={doc.id} className="flex items-center gap-3 px-4 py-2.5 hover:bg-slate-50">
                <FileWarning size={13} className="text-amber-500" />
                <div className="flex-1 min-w-0">
                  <p className="text-xs font-medium text-slate-700 truncate">{doc.seafarer_name}</p>
                  <p className="text-[11px] text-slate-400 truncate">{doc.document_name?.name || "-"}</p>
                </div>
                <p className="text-[10px] text-slate-400">{fmtDate(toDateInput(doc.expired_at))}</p>
              </div>
            ))}
            {docs.length === 0 && <p className="text-xs text-slate-400 text-center py-6">Tidak ada dokumen yang akan expired</p>}
          </div>
        </Card>

        <Card>
          <div className="flex items-center justify-between px-4 py-3 border-b border-slate-100">
            <h3 className="font-semibold text-slate-800 text-sm">Recent Joining Activity</h3>
            <Btn variant="ghost" size="sm" onClick={() => setPage("joining-principal")}>
              View all <ArrowRight size={12} />
            </Btn>
          </div>
          <div className="divide-y divide-slate-50">
            {joinings.slice(0, 5).map(j => (
              <div key={j.id} className="flex items-center gap-3 px-4 py-2.5 hover:bg-slate-50">
                <Briefcase size={13} className="text-blue-500" />
                <div className="flex-1 min-w-0">
                  <p className="text-xs font-medium text-slate-700 truncate">{j.name}</p>
                  <p className="text-[11px] text-slate-400 truncate">{j.rank || "-"} - {j.vessel?.name || j.vessel_name || "-"}</p>
                </div>
                <Badge status={j.status || "-"} />
              </div>
            ))}
            {joinings.length === 0 && <p className="text-xs text-slate-400 text-center py-6">Belum ada aktivitas joining</p>}
          </div>
        </Card>
      </div>

      <Card className="mt-4 p-4">
        <h3 className="font-semibold text-slate-800 text-sm mb-3">Quick Actions</h3>
        <div className="grid grid-cols-6 gap-2">
          {[
            { label: "Add Crew", icon: UserPlus, page: "crew-form", color: "bg-blue-50 text-blue-700 hover:bg-blue-100" },
            { label: "New Joining", icon: Briefcase, page: "joining-principal", color: "bg-cyan-50 text-cyan-700 hover:bg-cyan-100" },
            { label: "Sign On", icon: CheckCircle, page: "sign-on", color: "bg-emerald-50 text-emerald-700 hover:bg-emerald-100" },
            { label: "Sign Off", icon: XCircle, page: "sign-off", color: "bg-amber-50 text-amber-700 hover:bg-amber-100" },
            { label: "Documents", icon: FileText, page: "documents", color: "bg-purple-50 text-purple-700 hover:bg-purple-100" },
            { label: "Reports", icon: BarChart3, page: "crew-reports", color: "bg-slate-100 text-slate-700 hover:bg-slate-200" },
          ].map(({ label, icon: Icon, page, color }) => (
            <button key={label} onClick={() => setPage(page)} className={cn("flex flex-col items-center gap-1.5 p-3 rounded-xl transition-colors", color)}>
              <Icon size={18} />
              <span className="text-xs font-medium">{label}</span>
            </button>
          ))}
        </div>
      </Card>
    </div>
  );
}
function CrewTable({
  data, onView, onEdit, onDelete, loading = false,
}: {
  data: Crew[]; onView?: (c: Crew) => void; onEdit?: (c: Crew) => void;
  onDelete?: (c: Crew) => void; loading?: boolean;
}) {
  const cols = ["Photo", "Code", "Name", "Phone", "NIK", "Passport", "Seaman Book", "Status", "Actions"];
  if (loading) {
    return (
      <div className="divide-y divide-slate-50">
        {Array.from({ length: 5 }).map((_, i) => (
          <div key={i} className="flex items-center gap-3 px-4 py-3">
            <Skeleton className="w-8 h-8 rounded-full" />
            <Skeleton className="h-3 w-16 rounded" />
            <Skeleton className="h-3 w-32 rounded" />
            <Skeleton className="h-3 w-28 rounded" />
          </div>
        ))}
      </div>
    );
  }
  if (!data.length) return <EmptyState title="No crew members found" description="Try adjusting your filters or add a new crew member" />;

  return (
    <div className="overflow-x-auto">
      <table className="w-full text-sm">
        <thead>
          <tr className="border-b border-slate-100">
            {cols.map(c => (
              <th key={c} className="px-4 py-3 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-wide whitespace-nowrap">
                {c}
              </th>
            ))}
          </tr>
        </thead>
        <tbody className="divide-y divide-slate-50">
          {data.map(crew => (
            <tr key={crew.id} className="hover:bg-slate-50 transition-colors group">
              <td className="px-4 py-3">
                <div className="w-8 h-8 bg-blue-100 rounded-full flex items-center justify-center">
                  <span className="text-blue-700 text-[10px] font-bold">
                    {String(crew.name || "?").split(" ").map(n => n[0]).join("").slice(0, 2)}
                  </span>
                </div>
              </td>
              <td className="px-4 py-3 font-mono text-xs text-slate-500">{crew.code}</td>
              <td className="px-4 py-3">
                <p className="font-medium text-slate-800 text-sm">{crew.name}</p>
                <p className="text-[11px] text-slate-400">{crew.rank} · {crew.nationality}</p>
              </td>
              <td className="px-4 py-3 text-xs text-slate-600 whitespace-nowrap">{crew.phone}</td>
              <td className="px-4 py-3 font-mono text-xs text-slate-500">{crew.nik}</td>
              <td className="px-4 py-3 font-mono text-xs text-slate-500">{crew.passport}</td>
              <td className="px-4 py-3 font-mono text-xs text-slate-500">{crew.seamanBook}</td>
              <td className="px-4 py-3"><Badge status={crew.status} /></td>
              <td className="px-4 py-3">
                <div className="flex items-center gap-1">
                  {onView && <button onClick={() => onView(crew)} className="p-1.5 rounded hover:bg-blue-50 text-slate-400 hover:text-blue-600 transition-colors" title="View / Preview All Data"><Eye size={13} /></button>}
                  {onEdit && <button onClick={() => onEdit(crew)} className="p-1.5 rounded hover:bg-amber-50 text-slate-400 hover:text-amber-600 transition-colors" title="Edit Crew"><Pencil size={13} /></button>}
                  {onDelete && <button onClick={() => onDelete(crew)} className="p-1.5 rounded hover:bg-red-50 text-slate-400 hover:text-red-600 transition-colors" title="Delete Crew"><Trash2 size={13} /></button>}
                </div>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function CrewPreviewModal({ crew, open, onClose, onEdit }: { crew: Crew | null; open: boolean; onClose: () => void; onEdit: () => void }) {
  if (!crew) return null;

  return (
    <Modal open={open} onClose={onClose} title={`Preview Detail Seafarer — ${crew.name}`} size="lg"
      footer={
        <div className="flex items-center justify-between w-full">
          <Btn variant="secondary" size="sm" onClick={onClose}>Tutup</Btn>
          <Btn variant="primary" size="sm" onClick={onEdit}><Pencil size={12} />Edit Data Seafarer</Btn>
        </div>
      }>
      <div className="space-y-4">
        {/* Profile Header */}
        <div className="flex items-center gap-4 bg-slate-50 p-4 rounded-xl border border-slate-100">
          <div className="w-14 h-14 bg-blue-600 text-white rounded-xl flex items-center justify-center font-bold text-lg shrink-0 shadow-xs">
            {String(crew.name || "?").split(" ").map(n => n[0]).join("").slice(0, 2)}
          </div>
          <div className="flex-1 min-w-0">
            <div className="flex items-center justify-between">
              <h3 className="font-bold text-slate-900 text-base">{crew.name}</h3>
              <Badge status={crew.status} />
            </div>
            <p className="text-xs text-slate-500 mt-0.5">{crew.rank || "Seafarer"} · {crew.nationality || "Indonesian"}</p>
            <p className="text-[11px] font-mono text-blue-600 font-semibold mt-1">Kode: {crew.code}</p>
          </div>
        </div>

        {/* Data Grid */}
        <div className="grid grid-cols-2 gap-4">
          <div className="bg-white p-4 rounded-xl border border-slate-200 space-y-3">
            <h4 className="font-semibold text-xs text-slate-800 uppercase tracking-wider border-b border-slate-100 pb-2">Informasi Pribadi &amp; Kontak</h4>
            <div className="space-y-2 text-xs">
              <div className="flex justify-between"><span className="text-slate-400">Nama Lengkap:</span><span className="font-medium text-slate-800">{crew.name}</span></div>
              <div className="flex justify-between"><span className="text-slate-400">Tanggal Lahir:</span><span className="font-medium text-slate-800">{fmtDate(crew.dob) || "-"}</span></div>
              <div className="flex justify-between"><span className="text-slate-400">Kewarganegaraan:</span><span className="font-medium text-slate-800">{crew.nationality || "-"}</span></div>
              <div className="flex justify-between"><span className="text-slate-400">Email:</span><span className="font-medium text-slate-800">{crew.email || "-"}</span></div>
              <div className="flex justify-between"><span className="text-slate-400">Telepon / WhatsApp:</span><span className="font-medium text-slate-800">{crew.phone || "-"}</span></div>
            </div>
          </div>

          <div className="bg-white p-4 rounded-xl border border-slate-200 space-y-3">
            <h4 className="font-semibold text-xs text-slate-800 uppercase tracking-wider border-b border-slate-100 pb-2">Dokumen Identitas (IDs)</h4>
            <div className="space-y-2 text-xs font-mono">
              <div className="flex justify-between"><span className="text-slate-400 font-sans">Kode Pelaut:</span><span className="font-bold text-blue-600">{crew.code}</span></div>
              <div className="flex justify-between"><span className="text-slate-400 font-sans">NIK (KTP):</span><span className="font-semibold text-slate-800">{crew.nik || "-"}</span></div>
              <div className="flex justify-between"><span className="text-slate-400 font-sans">No. Passport:</span><span className="font-semibold text-slate-800">{crew.passport || "-"}</span></div>
              <div className="flex justify-between"><span className="text-slate-400 font-sans">No. Seaman Book:</span><span className="font-semibold text-slate-800">{crew.seamanBook || "-"}</span></div>
              <div className="flex justify-between"><span className="text-slate-400 font-sans">Status Pelaut:</span><Badge status={crew.status} /></div>
            </div>
          </div>
        </div>
      </div>
    </Modal>
  );
}

function CrewDatabasePage({ setPage, setSelectedCrew }: { setPage: (p: string) => void; setSelectedCrew: (c: Crew) => void }) {
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [nationalityFilter, setNationalityFilter] = useState("");
  const [page, setPageNum] = useState(1);
  const [showImport, setShowImport] = useState(false);
  const [previewCrew, setPreviewCrew] = useState<Crew | null>(null);
  const [confirmDelete, setConfirmDelete] = useState<Crew | null>(null);
  const PAGE_SIZE = 8;

  const { data: crews, loading, refresh } = useApiList<ApiSeafarer, Crew>("/api/seafarers", mapCrew, []);

  const filtered = useMemo(() => {
    const q = search.toLowerCase();
    return crews.filter(c => {
      const name = String(c.name ?? "").toLowerCase();
      const code = String(c.code ?? "").toLowerCase();
      const passport = String(c.passport ?? "").toLowerCase();
      const status = String(c.status ?? "");
      const nationality = String(c.nationality ?? "");
      return (q === "" || name.includes(q) || code.includes(q) || passport.includes(q)) &&
        (statusFilter === "" || status === statusFilter) &&
        (nationalityFilter === "" || nationality === nationalityFilter);
    });
  }, [crews, search, statusFilter, nationalityFilter]);

  const paginated = filtered.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);

  return (
    <div>
      <PageHeader title="Crew Database">
        <Btn variant="secondary" size="sm" onClick={() => setShowImport(true)}><Upload size={13} />Import</Btn>
        <Btn variant="secondary" size="sm" onClick={() => exportExcel("/api/admin/export/seafarers", "id,name,photo_url,phone,nik,passport_no,seaman_book_no,status", "seafarers.xlsx").catch(err => toast.error(err.message))}><Download size={13} />Export</Btn>
        <Btn variant="primary" size="sm" onClick={() => setPage("crew-form")}><Plus size={13} />Add Crew</Btn>
      </PageHeader>

      <Card>
        {/* Filters */}
        <div className="flex items-center gap-3 px-4 py-3 border-b border-slate-100">
          <SearchBar value={search} onChange={v => { setSearch(v); setPageNum(1); }} placeholder="Search crew, code, passport…" />
          <Select value={statusFilter} onChange={v => { setStatusFilter(v); setPageNum(1); }}
            options={["available", "onboard", "waiting", "blacklisted", "ex-crew"].map(s => ({ value: s, label: s.replace("-", " ").replace(/\b\w/g, l => l.toUpperCase()) }))}
            placeholder="All Status" />
          <Select value={nationalityFilter} onChange={v => { setNationalityFilter(v); setPageNum(1); }}
            options={[{ value: "Indonesian", label: "Indonesian" }, { value: "Filipino", label: "Filipino" }]}
            placeholder="All Nationalities" />
          {(statusFilter || nationalityFilter || search) && (
            <button onClick={() => { setSearch(""); setStatusFilter(""); setNationalityFilter(""); setPageNum(1); }}
              className="text-xs text-blue-600 hover:underline whitespace-nowrap">Clear filters</button>
          )}
          <div className="ml-auto text-xs text-slate-400">{loading ? "Loading..." : `${filtered.length} result${filtered.length !== 1 ? "s" : ""}`}</div>
        </div>

        <CrewTable
          data={paginated}
          onView={c => { setSelectedCrew(c); setPage("crew-detail"); }}
          onEdit={c => { setSelectedCrew(c); setPage("crew-form"); }}
          onDelete={c => setConfirmDelete(c)}
        />
        <Pagination page={page} total={filtered.length} pageSize={PAGE_SIZE} onPage={setPageNum} />
      </Card>

      {/* Preview Modal */}
      <CrewPreviewModal
        crew={previewCrew}
        open={!!previewCrew}
        onClose={() => setPreviewCrew(null)}
        onEdit={() => { if (previewCrew) { setSelectedCrew(previewCrew); setPage("crew-form"); setPreviewCrew(null); } }}
      />

      {/* Import Modal */}
      <ImportModal open={showImport} onClose={() => { setShowImport(false); refresh(); }} entity="Seafarers" endpoint="/api/admin/import/seafarers" />

      <ConfirmModal
        open={!!confirmDelete}
        onClose={() => setConfirmDelete(null)}
        onConfirm={async () => { if (!confirmDelete) return; try { await apiJson(`/api/seafarers/${confirmDelete.id}`, "DELETE"); toast.success(`${confirmDelete.name} removed from database`); refresh(); setConfirmDelete(null); } catch (err) { toast.error(err instanceof Error ? err.message : "Delete failed"); } }}
        title="Remove Crew Member"
        message={`Are you sure you want to remove ${confirmDelete?.name}? This action cannot be undone.`}
        danger
      />
    </div>
  );
}

// ─── Crew Detail ──────────────────────────────────────────────────────────────

function CrewDetailPage({ crew, setPage }: { crew: Crew; setPage: (p: string) => void }) {
  const [tab, setTab] = useState("overview");
  const tabs = ["overview", "documents", "contacts", "joining-history", "blacklist"];
  const [showAddDoc, setShowAddDoc] = useState(false);
  const [showAddContact, setShowAddContact] = useState(false);

  const [docForm, setDocForm] = useState({ name: "", number: "", expiry: "", isLifetime: false });
  const [contactForm, setContactForm] = useState({ name: "", relation: "Spouse", phone: "", email: "" });

  const [localDocs, setLocalDocs] = useState<{ id: number; name: string; number: string; expiry: string; status: string }[]>([]);
  const [contacts, setContacts] = useState<{ id: number; name: string; relation: string; phone: string; email: string }[]>([]);

  // Fetch real documents from API
  const { data: apiDocs } = useApiList<ApiExpiringDoc, ApiExpiringDoc>(
    `/api/documents/expiring?days=36500&limit=200`,
    item => item,
    [crew.id]
  );

  // Filter API docs for current seafarer
  const filteredApiDocs = useMemo(() => {
    return apiDocs.filter(d => d.seafarer_id === crew.id || d.seafarer_name.toLowerCase() === crew.name.toLowerCase());
  }, [apiDocs, crew.id, crew.name]);

  // Combine real native docs (Passport, Seaman Book if populated) with API documents & local docs
  const realCrewDocs = useMemo(() => {
    const list: { id: number; name: string; number: string; expiry: string; status: string }[] = [];

    if (crew.passport && crew.passport !== "-") {
      list.push({ id: -1, name: "Passport", number: crew.passport, expiry: "-", status: "valid" });
    }
    if (crew.seamanBook && crew.seamanBook !== "-") {
      list.push({ id: -2, name: "Seaman Book", number: crew.seamanBook, expiry: "-", status: "valid" });
    }

    filteredApiDocs.forEach(d => {
      const expiry = toDateInput(d.expired_at);
      const isExpired = d.expired_at && new Date(d.expired_at).getTime() < Date.now();
      list.push({
        id: d.id,
        name: d.document_name?.name || "Dokumen",
        number: d.document_no || "-",
        expiry: expiry || (d.is_lifetime ? "Seumur Hidup" : "-"),
        status: isExpired ? "expired" : "valid"
      });
    });

    localDocs.forEach(ld => list.push(ld));

    return list;
  }, [crew.passport, crew.seamanBook, filteredApiDocs, localDocs]);

  // Fetch real joining history
  const { data: allJoinings } = useApiList<ApiJoining, JoiningRecord>("/api/joining-principals", mapJoining, []);
  const realHistory = useMemo(() => {
    return allJoinings.filter(j => j.crewName.toLowerCase() === crew.name.toLowerCase() || (j as any).seafarer_id === crew.id);
  }, [allJoinings, crew.id, crew.name]);

  // Fetch real Blacklist
  const { data: allBlacklist } = useApiList<any, any>("/api/admin/blacklists", item => item, []);
  const realBlacklist = useMemo(() => {
    return (allBlacklist || []).filter((b: any) => b.seafarer_id === crew.id || String(b.crew || "").toLowerCase() === crew.name.toLowerCase());
  }, [allBlacklist, crew.id, crew.name]);

  function handleAddDocument() {
    if (!docForm.name.trim()) {
      toast.error("Nama dokumen wajib diisi");
      return;
    }
    const isExpired = docForm.expiry && new Date(docForm.expiry).getTime() < Date.now();
    setLocalDocs(prev => [
      ...prev,
      {
        id: Date.now(),
        name: docForm.name.trim(),
        number: docForm.number.trim() || "-",
        expiry: docForm.isLifetime ? "Seumur Hidup" : docForm.expiry || "-",
        status: isExpired ? "expired" : "valid"
      }
    ]);
    toast.success(`Dokumen ${docForm.name} berhasil ditambahkan`);
    setShowAddDoc(false);
    setDocForm({ name: "", number: "", expiry: "", isLifetime: false });
  }

  function handleAddContact() {
    if (!contactForm.name.trim()) {
      toast.error("Nama kontak wajib diisi");
      return;
    }
    setContacts(prev => [
      ...prev,
      {
        id: Date.now(),
        name: contactForm.name.trim(),
        relation: contactForm.relation,
        phone: contactForm.phone.trim() || "-",
        email: contactForm.email.trim() || "-"
      }
    ]);
    toast.success(`Kontak ${contactForm.name} berhasil ditambahkan`);
    setShowAddContact(false);
    setContactForm({ name: "", relation: "Spouse", phone: "", email: "" });
  }

  return (
    <div>
      <div className="flex items-center gap-3 mb-5">
        <button onClick={() => setPage("crew-database")} className="p-1.5 rounded-lg hover:bg-white border border-slate-200 transition-colors">
          <ArrowLeft size={14} className="text-slate-600" />
        </button>
        <h1 className="text-lg font-semibold text-slate-800">Crew Detail</h1>
      </div>

      {/* Profile Header */}
      <Card className="p-5 mb-4">
        <div className="flex items-start gap-5">
          <div className="w-16 h-16 bg-blue-100 rounded-xl flex items-center justify-center shrink-0">
            <span className="text-blue-700 text-xl font-bold">{String(crew.name || "?").split(" ").map(n => n[0]).join("").slice(0, 2)}</span>
          </div>
          <div className="flex-1">
            <div className="flex items-start justify-between">
              <div>
                <h2 className="text-base font-bold text-slate-800">{crew.name}</h2>
                <p className="text-sm text-slate-500 mt-0.5">{crew.rank} · {crew.nationality}</p>
                <div className="flex items-center gap-3 mt-2">
                  <Badge status={crew.status} />
                  <span className="text-xs text-slate-400 font-mono">{crew.code}</span>
                </div>
              </div>
              <div className="flex gap-2">
                <Btn variant="secondary" size="sm" onClick={() => setPage("crew-form")}><Pencil size={12} />Edit</Btn>
                <Btn variant="primary" size="sm" onClick={() => setPage("joining-principal")}><Briefcase size={12} />New Joining</Btn>
              </div>
            </div>
            <div className="grid grid-cols-4 gap-4 mt-4 pt-4 border-t border-slate-100">
              {[
                { icon: Mail, label: crew.email || "-" },
                { icon: Phone, label: crew.phone || "-" },
                { icon: Calendar, label: fmtDate(crew.dob) || "-" },
                { icon: Hash, label: crew.nik || "-" },
              ].map(({ icon: Icon, label }, i) => (
                <div key={i} className="flex items-center gap-1.5 text-xs text-slate-500">
                  <Icon size={12} className="text-slate-400 shrink-0" />
                  <span className="truncate">{label}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      </Card>

      {/* Tabs */}
      <div className="flex gap-0.5 mb-4 bg-white border border-slate-100 rounded-xl p-1 w-fit">
        {tabs.map(t => (
          <button key={t} onClick={() => setTab(t)}
            className={cn("px-4 py-1.5 rounded-lg text-xs font-medium transition-colors capitalize",
              tab === t ? "bg-blue-600 text-white" : "text-slate-500 hover:text-slate-700 hover:bg-slate-50")}>
            {t.replace("-", " ")}
          </button>
        ))}
      </div>

      {/* Tab Content */}
      {tab === "overview" && (
        <div className="grid grid-cols-2 gap-4">
          <Card className="p-4">
            <h3 className="font-semibold text-slate-700 text-sm mb-3">Personal Information</h3>
            <div className="space-y-2.5">
              {[
                ["Full Name", crew.name], ["Date of Birth", fmtDate(crew.dob) || "-"],
                ["Nationality", crew.nationality || "-"], ["Email", crew.email || "-"], ["Phone", crew.phone || "-"],
              ].map(([k, v]) => (
                <div key={k} className="flex items-start gap-2">
                  <span className="text-xs text-slate-400 w-28 shrink-0">{k}</span>
                  <span className="text-xs text-slate-700 font-medium">{v}</span>
                </div>
              ))}
            </div>
          </Card>
          <Card className="p-4">
            <h3 className="font-semibold text-slate-700 text-sm mb-3">Identification</h3>
            <div className="space-y-2.5">
              {[
                ["Crew Code", crew.code], ["NIK", crew.nik || "-"],
                ["Passport No.", crew.passport || "-"], ["Seaman Book", crew.seamanBook || "-"],
                ["Current Status", crew.status],
              ].map(([k, v]) => (
                <div key={k} className="flex items-start gap-2">
                  <span className="text-xs text-slate-400 w-28 shrink-0">{k}</span>
                  {k === "Current Status"
                    ? <Badge status={v} />
                    : <span className="text-xs text-slate-700 font-mono">{v}</span>
                  }
                </div>
              ))}
            </div>
          </Card>
        </div>
      )}

      {tab === "documents" && (
        <Card>
          <div className="flex items-center justify-between px-4 py-3 border-b border-slate-100">
            <h3 className="font-semibold text-slate-700 text-sm">Dokumen Pelaut</h3>
            <Btn variant="primary" size="sm" onClick={() => setShowAddDoc(true)}><Plus size={12} />Add Document</Btn>
          </div>
          {realCrewDocs.length > 0 ? (
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-slate-100">
                  {["Document", "Number", "Expiry Date", "Status", "Actions"].map(h => (
                    <th key={h} className="px-4 py-3 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-wide">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {realCrewDocs.map((d, i) => (
                  <tr key={`${d.id}-${i}`} className="hover:bg-slate-50">
                    <td className="px-4 py-3 font-medium text-slate-700 text-xs">{d.name}</td>
                    <td className="px-4 py-3 font-mono text-xs text-slate-500">{d.number}</td>
                    <td className="px-4 py-3 text-xs text-slate-600">{d.expiry === "-" ? "-" : fmtDate(d.expiry)}</td>
                    <td className="px-4 py-3"><Badge status={d.status} /></td>
                    <td className="px-4 py-3">
                      <div className="flex gap-1">
                        <button className="p-1.5 rounded hover:bg-slate-100 text-slate-400 transition-colors" title="Lihat Dokumen"><Eye size={12} /></button>
                        <button className="p-1.5 rounded hover:bg-amber-50 text-slate-400 hover:text-amber-600 transition-colors" title="Edit Dokumen"><Pencil size={12} /></button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          ) : (
            <div className="py-8">
              <EmptyState title="Belum Ada Dokumen Tercatat" description="Pelaut ini belum memiliki dokumen tercatat di sistem. Klik tombol di bawah untuk menambahkan dokumen baru." action={<Btn variant="primary" size="sm" onClick={() => setShowAddDoc(true)}><Plus size={12} />Add Document</Btn>} />
            </div>
          )}
        </Card>
      )}

      {tab === "contacts" && (
        <Card className="p-4">
          <div className="flex items-center justify-between mb-3">
            <h3 className="font-semibold text-slate-700 text-sm">Emergency Contacts</h3>
            {contacts.length < 2 && <Btn variant="primary" size="sm" onClick={() => setShowAddContact(true)}><Plus size={12} />Add Contact</Btn>}
          </div>
          <div className="space-y-3">
            {contacts.map((c, i) => (
              <div key={i} className="flex items-start gap-3 p-3 bg-slate-50 rounded-xl">
                <div className="w-8 h-8 bg-blue-100 rounded-full flex items-center justify-center shrink-0">
                  <span className="text-blue-700 text-xs font-bold">{c.name.split(" ").map(n => n[0]).join("").slice(0, 2)}</span>
                </div>
                <div className="flex-1">
                  <div className="flex items-center gap-2">
                    <p className="text-xs font-semibold text-slate-700">{c.name}</p>
                    <span className="text-[10px] bg-blue-50 text-blue-600 px-1.5 py-0.5 rounded font-medium">{c.relation}</span>
                  </div>
                  <p className="text-[11px] text-slate-400 mt-0.5">{c.phone} · {c.email}</p>
                </div>
                <button className="p-1.5 rounded hover:bg-white text-slate-400 transition-colors"><Pencil size={12} /></button>
              </div>
            ))}
            {contacts.length === 0 && <EmptyState title="Belum Ada Kontak Darurat" description="Belum ada kontak darurat yang didaftarkan untuk pelaut ini." action={<Btn variant="primary" size="sm" onClick={() => setShowAddContact(true)}><Plus size={12} />Add Contact</Btn>} />}
            {contacts.length > 0 && <p className="text-[10px] text-slate-300 text-center mt-2">Maksimal 2 kontak darurat</p>}
          </div>
        </Card>
      )}

      {tab === "joining-history" && (
        <Card>
          <div className="px-4 py-3 border-b border-slate-100">
            <h3 className="font-semibold text-slate-700 text-sm">Joining History</h3>
          </div>
          {realHistory.length === 0
            ? <EmptyState title="Belum Ada Riwayat Joining" description="Pelaut ini belum memiliki riwayat penugasan / joining kapal." />
            : (
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b border-slate-100">
                    {["Vessel", "Principal", "Rank", "Sign On", "Sign Off", "Port", "Status"].map(h => (
                      <th key={h} className="px-4 py-3 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-wide">{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  {realHistory.map(j => (
                    <tr key={j.id} className="hover:bg-slate-50">
                      <td className="px-4 py-3 font-medium text-slate-700 text-xs">{j.vessel}</td>
                      <td className="px-4 py-3 text-xs text-slate-500">{j.principal}</td>
                      <td className="px-4 py-3 text-xs text-slate-600">{j.rank}</td>
                      <td className="px-4 py-3 text-xs text-slate-600">{fmtDate(j.signOn)}</td>
                      <td className="px-4 py-3 text-xs text-slate-600">{fmtDate(j.signOff)}</td>
                      <td className="px-4 py-3 text-xs text-slate-600">{j.port}</td>
                      <td className="px-4 py-3"><Badge status={j.status} /></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )
          }
        </Card>
      )}

      {tab === "blacklist" && (
        <Card className="p-4">
          <div className="flex items-center gap-3 p-4 bg-slate-50 border border-slate-200 rounded-xl">
            <AlertTriangle size={16} className={cn("shrink-0", crew.status === "blacklisted" || realBlacklist.length > 0 ? "text-red-500" : "text-slate-400")} />
            <p className="text-xs text-slate-700 font-medium">
              {crew.status === "blacklisted" || realBlacklist.length > 0
                ? "Pelaut ini terdaftar pada catatan blacklist aktif."
                : "Pelaut ini tidak memiliki catatan blacklist."}
            </p>
          </div>
          {realBlacklist.length > 0 && (
            <div className="mt-4 space-y-3">
              {realBlacklist.map((b: any) => (
                <div key={b.id} className="p-3 bg-red-50 rounded-xl border border-red-100">
                  <p className="text-xs font-semibold text-red-800 mb-1">{b.reason || "Contract Breach"}</p>
                  <p className="text-xs text-red-600">{b.notes || "-"}</p>
                  <div className="flex gap-2 mt-3">
                    <Badge status={b.status || "active"} />
                  </div>
                </div>
              ))}
            </div>
          )}
        </Card>
      )}

      {/* Add Document Modal */}
      <Modal open={showAddDoc} onClose={() => setShowAddDoc(false)} title={`Tambah Dokumen — ${crew.name}`} size="md"
        footer={<><Btn variant="secondary" size="sm" onClick={() => setShowAddDoc(false)}>Batal</Btn><Btn variant="primary" size="sm" onClick={handleAddDocument}>Simpan Dokumen</Btn></>}>
        <div className="space-y-3">
          <Input label="Nama Dokumen" value={docForm.name} onChange={v => setDocForm(p => ({ ...p, name: v }))} placeholder="misal: STCW Certificate, Medical, BST" required />
          <Input label="Nomor Dokumen" value={docForm.number} onChange={v => setDocForm(p => ({ ...p, number: v }))} placeholder="misal: DOC-2026-991" />
          {!docForm.isLifetime && (
            <Input label="Tanggal Expiry (Masa Berlaku)" type="date" value={docForm.expiry} onChange={v => setDocForm(p => ({ ...p, expiry: v }))} />
          )}
          <label className="flex items-center gap-2 cursor-pointer pt-1">
            <input type="checkbox" checked={docForm.isLifetime} onChange={e => setDocForm(p => ({ ...p, isLifetime: e.target.checked }))} className="rounded border-slate-300 text-blue-600 focus:ring-blue-500" />
            <span className="text-xs text-slate-700 font-medium">Dokumen Seumur Hidup (Lifetime)</span>
          </label>
        </div>
      </Modal>

      {/* Add Contact Modal */}
      <Modal open={showAddContact} onClose={() => setShowAddContact(false)} title={`Tambah Kontak Darurat — ${crew.name}`} size="md"
        footer={<><Btn variant="secondary" size="sm" onClick={() => setShowAddContact(false)}>Batal</Btn><Btn variant="primary" size="sm" onClick={handleAddContact}>Simpan Kontak</Btn></>}>
        <div className="space-y-3">
          <Input label="Nama Kontak" value={contactForm.name} onChange={v => setContactForm(p => ({ ...p, name: v }))} placeholder="Nama lengkap wali / keluarga" required />
          <Select label="Hubungan / Relasi" value={contactForm.relation} onChange={v => setContactForm(p => ({ ...p, relation: v }))} options={[{ value: "Spouse", label: "Suami / Istri (Spouse)" }, { value: "Parent", label: "Orang Tua (Parent)" }, { value: "Child", label: "Anak (Child)" }, { value: "Kin", label: "Kerabat / Sdr (Kin)" }]} />
          <Input label="Nomor Telepon / HP" value={contactForm.phone} onChange={v => setContactForm(p => ({ ...p, phone: v }))} placeholder="+62 812-xxxx-xxxx" />
          <Input label="Email" type="email" value={contactForm.email} onChange={v => setContactForm(p => ({ ...p, email: v }))} placeholder="email@domain.com" />
        </div>
      </Modal>
    </div>
  );
}

// ─── Crew Form (Add/Edit) ─────────────────────────────────────────────────────

function CrewFormPage({ crew, setPage }: { crew: Crew | null; setPage: (p: string) => void }) {
  const isEdit = !!crew;
  const [tab, setTab] = useState("personal");
  const [loading, setLoading] = useState(false);
  const [form, setForm] = useState({
    name: crew?.name ?? "", code: crew?.code ?? "", phone: crew?.phone ?? "",
    email: crew?.email ?? "", dob: crew?.dob ?? "", nationality: crew?.nationality ?? "Indonesian",
    rank: crew?.rank ?? "", nik: crew?.nik ?? "", passport: crew?.passport ?? "",
    seamanBook: crew?.seamanBook ?? "", status: crew?.status ?? "available",
  });

  function f(k: string) { return (v: string) => setForm(p => ({ ...p, [k]: v })); }

  const [docState, setDocState] = useState<Record<string, { file_name?: string; original_name?: string; file_url?: string; uploading?: boolean }>>({});
  const { data: crewDocumentNames, loading: crewDocumentsLoading } = useApiList<ApiDocumentName, ReturnType<typeof mapDocName>>("/api/admin/document-names", mapDocName, []);
  const activeCrewDocuments = useMemo(() => crewDocumentNames.filter(doc => doc.active), [crewDocumentNames]);

  async function handleFileUpload(docType: string, e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    setDocState(p => ({ ...p, [docType]: { ...p[docType], uploading: true } }));
    try {
      const res = await uploadDocument(file);
      setDocState(p => ({
        ...p,
        [docType]: {
          file_name: res.file_name,
          original_name: res.original_name,
          file_url: res.file_url,
          uploading: false,
        },
      }));
      toast.success(`${docType} uploaded successfully`);
    } catch (err) {
      setDocState(p => ({ ...p, [docType]: { ...p[docType], uploading: false } }));
      toast.error(err instanceof Error ? err.message : "Upload failed");
    }
  }

  async function submit(e: React.FormEvent) {
    e.preventDefault();
    setLoading(true);
    const payload = {
      name: form.name,
      seafarer_code: form.code,
      phone: form.phone,
      nik: form.nik,
      birth_date: form.dob ? `${form.dob}T00:00:00Z` : null,
      passport_no: form.passport,
      seaman_book_no: form.seamanBook,
      certificate_no: form.rank,
      status: form.status,
    };
    try {
      await apiJson(isEdit ? `/api/seafarers/${crew.id}` : "/api/seafarers", isEdit ? "PUT" : "POST", payload);
      toast.success(isEdit ? "Crew member updated successfully" : "Crew member added successfully");
      setPage("crew-database");
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Save failed");
    } finally {
      setLoading(false);
    }
  }

  const tabs = ["personal", "identification", "family", "contract", "emergency", "documents"];

  return (
    <div>
      <div className="flex items-center gap-3 mb-5">
        <button onClick={() => setPage("crew-database")} className="p-1.5 rounded-lg hover:bg-white border border-slate-200 transition-colors">
          <ArrowLeft size={14} className="text-slate-600" />
        </button>
        <h1 className="text-lg font-semibold text-slate-800">{isEdit ? `Edit: ${crew.name}` : "Add New Crew"}</h1>
      </div>

      <form onSubmit={submit}>
        <div className="flex gap-1 mb-4 bg-white border border-slate-100 rounded-xl p-1 w-fit">
          {tabs.map(t => (
            <button type="button" key={t} onClick={() => setTab(t)}
              className={cn("px-4 py-1.5 rounded-lg text-xs font-medium transition-colors capitalize",
                tab === t ? "bg-blue-600 text-white" : "text-slate-500 hover:text-slate-700 hover:bg-slate-50")}>
              {t === "emergency" ? "Emergency" : t.charAt(0).toUpperCase() + t.slice(1)} {t === "emergency" ? "Contacts" : ""}
            </button>
          ))}
        </div>

        <Card className="p-5">
          {tab === "personal" && (
            <div className="grid grid-cols-2 gap-4">
              <Input label="Full Name" value={form.name} onChange={f("name")} required placeholder="Enter full name" />
              <Input label="Crew Code" value={form.code} onChange={f("code")} placeholder="Auto-generated if empty" />
              <Input label="Email" type="email" value={form.email} onChange={f("email")} placeholder="crew@email.com" />
              <Input label="Phone" value={form.phone} onChange={f("phone")} placeholder="+62 812-3456-7890" />
              <Input label="Date of Birth" type="date" value={form.dob} onChange={f("dob")} required />
              <Select label="Nationality" value={form.nationality} onChange={f("nationality")} required
                options={[{ value: "Indonesian", label: "Indonesian" }, { value: "Filipino", label: "Filipino" }, { value: "Indian", label: "Indian" }, { value: "Myanmar", label: "Myanmar" }]} />
              <SelectWithOther label="Rank" value={form.rank} onChange={f("rank")} required
                options={["Master", "Chief Officer", "Second Officer", "Third Officer", "Chief Engineer", "Second Engineer", "Third Engineer", "Fourth Engineer", "Bosun", "Able Seaman", "Ordinary Seaman", "Chief Cook", "Cook", "Messman", "Oiler", "Wiper", "Deck Cadet", "Engine Cadet", "Electrician", "Fitter", "Pumpman"].map(r => ({ value: r, label: r }))} />
              <SelectWithOther label="Status" value={form.status} onChange={f("status")} required
                options={["Available", "Waiting", "Onboard", "Blacklisted"].map(s => ({ value: s.toLowerCase(), label: s }))} />
            </div>
          )}

          {tab === "identification" && (
            <div className="grid grid-cols-2 gap-4">
              <Input label="NIK / National ID" value={form.nik} onChange={f("nik")} placeholder="16-digit ID number" />
              <Input label="Passport Number" value={form.passport} onChange={f("passport")} placeholder="e.g. A1234567" />
              <Input label="Seaman Book Number" value={form.seamanBook} onChange={f("seamanBook")} placeholder="e.g. D123456" />
              <Input label="Passport Expiry" type="date" value="" onChange={() => {}} />
              <Input label="Seaman Book Expiry" type="date" value="" onChange={() => {}} />
              <Input label="Place of Birth" value="" onChange={() => {}} placeholder="City of birth" />
            </div>
          )}

          {tab === "family" && (
            <div className="grid grid-cols-2 gap-4">
              <Input label="Father's Name" value="" onChange={() => {}} placeholder="Full name" />
              <Input label="Mother's Name" value="" onChange={() => {}} placeholder="Full name" />
              <Input label="Spouse's Name" value="" onChange={() => {}} placeholder="Full name (if married)" />
              <Input label="Number of Children" value="" onChange={() => {}} type="number" placeholder="0" />
              <Input label="Home Address" value="" onChange={() => {}} placeholder="Full address" className="col-span-2" />
              <Select label="Marital Status" value="" onChange={() => {}}
                options={["Single", "Married", "Divorced", "Widowed"].map(s => ({ value: s, label: s }))} />
            </div>
          )}

          {tab === "contract" && (
            <div className="grid grid-cols-2 gap-4">
              <Select label="Contract Type" value="" onChange={() => {}}
                options={["9+3", "6+3", "4+2", "12+3"].map(s => ({ value: s, label: s + " months" }))} />
              <Input label="Base Salary (USD)" value="" onChange={() => {}} type="number" placeholder="0.00" />
              <Input label="Allotment" value="" onChange={() => {}} type="number" placeholder="Allotment amount" />
              <Input label="Contract Start" type="date" value="" onChange={() => {}} />
              <Input label="Contract End" type="date" value="" onChange={() => {}} />
              <Select label="Preferred Principal" value="" onChange={() => {}}
                options={MOCK_PRINCIPALS.map(p => ({ value: p.code, label: p.name }))} />
            </div>
          )}

          {tab === "emergency" && (
            <div className="space-y-5">
              {[1, 2].map(n => (
                <div key={n} className="p-4 bg-slate-50 rounded-xl border border-slate-100">
                  <p className="text-xs font-semibold text-slate-600 mb-3">Emergency Contact #{n}</p>
                  <div className="grid grid-cols-2 gap-3">
                    <Input label="Full Name" value="" onChange={() => {}} placeholder="Contact name" />
                    <Select label="Relationship" value="" onChange={() => {}}
                      options={["Spouse", "Parent", "Sibling", "Child", "Other"].map(r => ({ value: r, label: r }))} />
                    <Input label="Phone Number" value="" onChange={() => {}} placeholder="+62 812-3456-7890" />
                    <Input label="Email" type="email" value="" onChange={() => {}} placeholder="email@example.com" />
                    <Input label="Address" value="" onChange={() => {}} placeholder="Home address" className="col-span-2" />
                  </div>
                </div>
              ))}
              <p className="text-[10px] text-slate-400 text-center">Maximum 2 emergency contacts allowed</p>
            </div>
          )}

          {tab === "documents" && (
            <div>
              <p className="text-xs text-slate-500 mb-4">Upload and manage crew documents. Supported: PDF, JPG, PNG, DOC.</p>
              <div className="space-y-3">
                {crewDocumentsLoading && <p className="text-xs text-slate-400">Loading document master...</p>}
                {activeCrewDocuments.map(doc => {
                  const docKey = String(doc.id);
                  const item = docState[docKey];
                  return (
                    <div key={doc.id} className="flex items-center justify-between p-3 bg-slate-50 rounded-xl border border-slate-100">
                      <div className="flex items-center gap-2.5">
                        <FileCheck size={14} className={item?.file_url ? "text-emerald-500" : "text-slate-400"} />
                        <div>
                          <div className="flex items-center gap-2">
                            <p className="text-xs font-medium text-slate-700">{doc.name}</p>
                            {doc.required && <span className="text-[10px] rounded bg-amber-50 px-1.5 py-0.5 font-medium text-amber-700">Required</span>}
                            {!doc.hasExpiry && <span className="text-[10px] rounded bg-slate-100 px-1.5 py-0.5 font-medium text-slate-500">Lifetime</span>}
                          </div>
                          <p className="text-[10px] text-slate-400">{doc.type}</p>
                          {item?.uploading ? (
                            <p className="text-[10px] text-blue-500 font-medium animate-pulse">Uploading file...</p>
                          ) : item?.original_name ? (
                            <a href={`${getApiBaseUrl()}${item.file_url}`} target="_blank" rel="noreferrer" className="text-[10px] text-blue-600 hover:underline font-mono">{item.original_name}</a>
                          ) : (
                            <p className="text-[10px] text-slate-400">No file uploaded</p>
                          )}
                        </div>
                      </div>
                      <label className={cn("inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium border cursor-pointer transition-colors", item?.uploading ? "bg-slate-100 text-slate-400 border-slate-200" : "bg-white text-slate-700 border-slate-200 hover:bg-slate-50")}>
                        <Upload size={11} />
                        {item?.uploading ? "Uploading..." : item?.original_name ? "Change" : "Upload"}
                        <input type="file" className="hidden" disabled={item?.uploading} accept=".pdf,.jpg,.jpeg,.png,.doc,.docx,.xls,.xlsx" onChange={e => handleFileUpload(docKey, e)} />
                      </label>
                    </div>
                  );
                })}
                {!crewDocumentsLoading && activeCrewDocuments.length === 0 && <EmptyState title="No document master" description="Tambahkan Document Names terlebih dahulu agar daftar upload muncul di form crew." />}
              </div>
            </div>
          )}
        </Card>

        {/* Footer actions */}
        <div className="flex items-center justify-end gap-2 mt-4">
          <Btn variant="secondary" onClick={() => setPage("crew-database")}>Cancel</Btn>
          <Btn type="submit" variant="primary" disabled={loading}>
            {loading ? <><RefreshCw size={13} className="animate-spin" />Saving…</> : isEdit ? "Update Crew" : "Add Crew"}
          </Btn>
        </div>
      </form>
    </div>
  );
}

// ─── Status Pages (Available / Onboard / Waiting / Ex-Crew) ──────────────────

function CrewStatusPage({
  status, title, setPage, setSelectedCrew,
}: { status: string; title: string; setPage: (p: string) => void; setSelectedCrew: (c: Crew) => void }) {
  const [search, setSearch] = useState("");
  const [page, setPageNum] = useState(1);
  const PAGE_SIZE = 10;

  const { data: crews, loading } = useApiList<ApiSeafarer, Crew>("/api/seafarers", mapCrew, MOCK_CREW);

  const filtered = useMemo(() =>
    crews.filter(c => c.status === status && (search === "" || c.name.toLowerCase().includes(search.toLowerCase()))),
    [crews, status, search]
  );

  return (
    <div>
      <PageHeader title={title}>
        <Btn variant="secondary" size="sm" onClick={() => exportExcel("/api/admin/export/seafarers", "id,name,photo_url,phone,nik,passport_no,seaman_book_no,status", "seafarers.xlsx").catch(err => toast.error(err.message))}><Download size={13} />Export</Btn>
      </PageHeader>
      <Card>
        <div className="flex items-center gap-3 px-4 py-3 border-b border-slate-100">
          <SearchBar value={search} onChange={v => { setSearch(v); setPageNum(1); }} placeholder={`Search ${title.toLowerCase()}…`} />
          <div className="ml-auto text-xs text-slate-400">{loading ? "Loading..." : `${filtered.length} crew`}</div>
        </div>
        <CrewTable
          data={filtered.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE)}
          onView={c => { setSelectedCrew(c); setPage("crew-detail"); }}
        />
        <Pagination page={page} total={filtered.length} pageSize={PAGE_SIZE} onPage={setPageNum} />
      </Card>
    </div>
  );
}

function SearchCrewPage({ setPage, setSelectedCrew }: { setPage: (p: string) => void; setSelectedCrew: (c: Crew) => void }) {
  const [query, setQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [rankFilter, setRankFilter] = useState("");
  const [pageNum, setPageNum] = useState(1);
  const PAGE_SIZE = 8;

  const { data: crews, loading } = useApiList<ApiSeafarer, Crew>("/api/seafarers", mapCrew, []);

  const results = useMemo(() => {
    const q = query.toLowerCase();
    return crews.filter(c => {
      const matchQuery = q === "" ||
        c.name.toLowerCase().includes(q) ||
        c.code.toLowerCase().includes(q) ||
        c.passport.toLowerCase().includes(q) ||
        c.phone.includes(q) ||
        c.nik.includes(q);
      const matchStatus = statusFilter === "" || c.status === statusFilter;
      const matchRank = rankFilter === "" || c.rank === rankFilter;
      return matchQuery && matchStatus && matchRank;
    });
  }, [crews, query, statusFilter, rankFilter]);

  const paginated = results.slice((pageNum - 1) * PAGE_SIZE, pageNum * PAGE_SIZE);

  return (
    <div>
      <PageHeader title="Search Crew">
        <Btn variant="primary" size="sm" onClick={() => setPage("crew-form")}><Plus size={13} />Add Crew</Btn>
      </PageHeader>

      <Card>
        <div className="flex items-center gap-3 px-4 py-3 border-b border-slate-100">
          <SearchBar value={query} onChange={v => { setQuery(v); setPageNum(1); }} placeholder="Search crew, code, passport, phone…" />
          <Select value={statusFilter} onChange={v => { setStatusFilter(v); setPageNum(1); }}
            options={["available", "onboard", "waiting", "blacklisted"].map(s => ({ value: s, label: s.charAt(0).toUpperCase() + s.slice(1) }))}
            placeholder="All Status" />
          <Select value={rankFilter} onChange={v => { setRankFilter(v); setPageNum(1); }}
            options={["Master", "Chief Officer", "Second Officer", "Third Officer", "Chief Engineer", "Second Engineer", "Bosun", "Able Seaman", "Cook"].map(r => ({ value: r, label: r }))}
            placeholder="All Ranks" />
          {(query || statusFilter || rankFilter) && (
            <button onClick={() => { setQuery(""); setStatusFilter(""); setRankFilter(""); setPageNum(1); }}
              className="text-xs text-blue-600 hover:underline whitespace-nowrap">Clear filters</button>
          )}
          <div className="ml-auto text-xs text-slate-400">{loading ? "Loading..." : `${results.length} result${results.length !== 1 ? "s" : ""}`}</div>
        </div>

        <CrewTable
          data={paginated}
          onView={c => { setSelectedCrew(c); setPage("crew-detail"); }}
          onEdit={c => { setSelectedCrew(c); setPage("crew-form"); }}
          onDelete={async c => {
            if (window.confirm(`Are you sure you want to delete ${c.name}?`)) {
              try {
                await apiJson(`/api/seafarers/${c.id}`, "DELETE");
                toast.success("Crew deleted successfully");
                window.location.reload();
              } catch (err) {
                toast.error(err instanceof Error ? err.message : "Delete failed");
              }
            }
          }}
        />
        <Pagination page={pageNum} total={results.length} pageSize={PAGE_SIZE} onPage={setPageNum} />
      </Card>
    </div>
  );
}

// ─── Blacklist ────────────────────────────────────────────────────────────────

function BlacklistPage() {
  const { data, loading, refresh } = useApiList<ApiBlacklist, ReturnType<typeof mapBlacklist>>("/api/admin/blacklists", mapBlacklist, []);
  const { data: crews } = useApiList<ApiSeafarer, Crew>("/api/seafarers", mapCrew, []);
  const [showAdd, setShowAdd] = useState(false);
  const [form, setForm] = useState({ crew: "", reason: "", notes: "" });
  const [confirmRelease, setConfirmRelease] = useState<number | null>(null);

  async function addBlacklist() {
    if (!form.crew || !form.reason) { toast.error("Crew and reason are required"); return; }
    try {
      await apiJson("/api/admin/blacklists", "POST", { seafarer_id: Number(form.crew), reason: form.reason, notes: form.notes });
      toast.success("Crew added to blacklist");
      setShowAdd(false);
      setForm({ crew: "", reason: "", notes: "" });
      refresh();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Save failed");
    }
  }

  async function releaseBlacklist(id: number) {
    try {
      await apiJson(`/api/admin/blacklists/${id}/release`, "PUT", { released_reason: "Released from FE" });
      toast.success("Crew released from blacklist");
      setConfirmRelease(null);
      refresh();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Release failed");
    }
  }

  async function deleteBlacklist(id: number) {
    try {
      await apiJson(`/api/admin/blacklists/${id}`, "DELETE");
      toast.success("Blacklist record deleted");
      refresh();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Delete failed");
    }
  }

  return (
    <div>
      <PageHeader title="Blacklist">
        <Btn variant="danger" size="sm" onClick={() => setShowAdd(true)}><Ban size={13} />Add to Blacklist</Btn>
      </PageHeader>

      <Card>
        <div className="px-4 py-3 border-b border-slate-100">
          <span className="text-xs text-slate-500">{loading ? "Loading..." : `${data.filter(d => d.status === "active").length} active blacklist records`}</span>
        </div>
        <table className="w-full text-sm">
          <thead>
            <tr className="border-b border-slate-100">
              {["Crew", "Reason", "Notes", "Status", "Actions"].map(h => (
                <th key={h} className="px-4 py-3 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-wide">{h}</th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-50">
            {data.map(b => (
              <tr key={b.id} className="hover:bg-slate-50">
                <td className="px-4 py-3"><p className="font-medium text-slate-800 text-xs">{b.crew}</p><p className="text-[10px] text-slate-400 font-mono">{b.code}</p></td>
                <td className="px-4 py-3 text-xs text-slate-600 max-w-xs truncate">{b.reason}</td>
                <td className="px-4 py-3 text-xs text-slate-500 max-w-xs truncate">{b.notes || "-"}</td>
                <td className="px-4 py-3"><Badge status={b.status} /></td>
                <td className="px-4 py-3">
                  <div className="flex gap-1">
                    {b.status === "active" && <button onClick={() => setConfirmRelease(b.id)} className="px-2 py-1 text-[10px] font-medium rounded bg-emerald-50 text-emerald-700 hover:bg-emerald-100 transition-colors">Release</button>}
                    <button onClick={() => deleteBlacklist(b.id)} className="p-1.5 rounded hover:bg-red-50 text-slate-400 hover:text-red-500 transition-colors"><Trash2 size={12} /></button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {!loading && data.length === 0 && <EmptyState title="No blacklist records" description="No crew members are currently blacklisted." />}
      </Card>

      <Modal open={showAdd} onClose={() => setShowAdd(false)} title="Add to Blacklist"
        footer={<><Btn variant="secondary" size="sm" onClick={() => setShowAdd(false)}>Cancel</Btn><Btn variant="danger" size="sm" onClick={addBlacklist}>Add to Blacklist</Btn></>}>
        <div className="space-y-4">
          <Select label="Crew Member" value={form.crew} onChange={v => setForm(p => ({ ...p, crew: v }))} required options={crews.filter(c => c.status !== "blacklisted").map(c => ({ value: String(c.id), label: `${c.name} (${c.code})` }))} />
          <Input label="Reason" value={form.reason} onChange={v => setForm(p => ({ ...p, reason: v }))} required placeholder="Brief reason for blacklisting" />
          <Textarea label="Notes" value={form.notes} onChange={v => setForm(p => ({ ...p, notes: v }))} placeholder="Detailed notes..." />
        </div>
      </Modal>

      <ConfirmModal open={!!confirmRelease} onClose={() => setConfirmRelease(null)} onConfirm={() => confirmRelease && releaseBlacklist(confirmRelease)} title="Release from Blacklist" message="Are you sure you want to release this crew member from the blacklist?" />
    </div>
  );
}
// ─── Payslip (Slip Gaji) Management ──────────────────────────────────────────

function createDocxBlob(documentXml: string): Blob {
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

  const files = [
    { name: "[Content_Types].xml", data: new TextEncoder().encode(contentTypes) },
    { name: "_rels/.rels", data: new TextEncoder().encode(rels) },
    { name: "word/document.xml", data: new TextEncoder().encode(documentXml) },
  ];

  const crcTable = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    crcTable[i] = c;
  }
  function crc32(buf: Uint8Array): number {
    let crc = 0xffffffff;
    for (let i = 0; i < buf.length; i++) crc = crcTable[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
    return (crc ^ 0xffffffff) >>> 0;
  }

  const parts: Uint8Array[] = [];
  const cdEntries: Uint8Array[] = [];
  let offset = 0;

  for (const f of files) {
    const nameBytes = new TextEncoder().encode(f.name);
    const crc = crc32(f.data);
    const size = f.data.length;

    const lh = new Uint8Array(30 + nameBytes.length);
    const view = new DataView(lh.buffer);
    view.setUint32(0, 0x04034b50, true);
    view.setUint16(4, 20, true);
    view.setUint16(6, 0, true);
    view.setUint16(8, 0, true);
    view.setUint16(10, 0x4800, true);
    view.setUint16(12, 0x58a1, true);
    view.setUint32(14, crc, true);
    view.setUint32(18, size, true);
    view.setUint32(22, size, true);
    view.setUint16(26, nameBytes.length, true);
    view.setUint16(28, 0, true);
    lh.set(nameBytes, 30);

    parts.push(lh, f.data);

    const cd = new Uint8Array(46 + nameBytes.length);
    const cdView = new DataView(cd.buffer);
    cdView.setUint32(0, 0x02014b50, true);
    cdView.setUint16(4, 20, true);
    cdView.setUint16(6, 20, true);
    cdView.setUint16(8, 0, true);
    cdView.setUint16(10, 0, true);
    cdView.setUint16(12, 0x4800, true);
    cdView.setUint16(14, 0x58a1, true);
    cdView.setUint32(16, crc, true);
    cdView.setUint32(20, size, true);
    cdView.setUint32(24, size, true);
    cdView.setUint16(28, nameBytes.length, true);
    cdView.setUint16(30, 0, true);
    cdView.setUint16(32, 0, true);
    cdView.setUint16(34, 0, true);
    cdView.setUint16(36, 0, true);
    cdView.setUint32(38, 0, true);
    cdView.setUint32(42, offset, true);
    cd.set(nameBytes, 46);

    cdEntries.push(cd);
    offset += lh.length + size;
  }

  const cdOffset = offset;
  let cdSize = 0;
  for (const cd of cdEntries) {
    parts.push(cd);
    cdSize += cd.length;
  }

  const eocd = new Uint8Array(22);
  const eocdView = new DataView(eocd.buffer);
  eocdView.setUint32(0, 0x06054b50, true);
  eocdView.setUint16(4, 0, true);
  eocdView.setUint16(6, 0, true);
  eocdView.setUint16(8, files.length, true);
  eocdView.setUint16(10, files.length, true);
  eocdView.setUint32(12, cdSize, true);
  eocdView.setUint32(16, cdOffset, true);
  eocdView.setUint16(20, 0, true);
  parts.push(eocd);

  return new Blob(parts, { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

function downloadSamplePayslipDoc() {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="1E3A8A"/></w:rPr><w:t>SKYAGEN MARITIME LOGISTICS</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r><w:rPr><w:sz w:val="20"/><w:color w:val="64748B"/></w:rPr><w:t>Official Seafarer Payslip Statement</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>----------------------------------------------------------------------------------------------------</w:t></w:r></w:p>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="26"/><w:color w:val="0F172A"/></w:rPr><w:t>SLIP GAJI PELAUT / SEAFARER PAYSLIP TEMPLATE</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>----------------------------------------------------------------------------------------------------</w:t></w:r></w:p>
    <w:p/>
    <w:p>
      <w:r><w:rPr><w:b/></w:rPr><w:t>PERIODE GAJI: </w:t></w:r>
      <w:r><w:rPr><w:b/><w:color w:val="2563EB"/></w:rPr><w:t>{{month}} / {{year}}</w:t></w:r>
    </w:p>
    <w:p/>
    <w:p><w:r><w:rPr><w:b/><w:color w:val="0F172A"/></w:rPr><w:t>INFORMASI PELAUT:</w:t></w:r></w:p>
    <w:p><w:r><w:t>- Nama Pelaut     : {{seafarer_name}}</w:t></w:r></w:p>
    <w:p><w:r><w:t>- Jabatan / Rank  : {{rank}}</w:t></w:r></w:p>
    <w:p><w:r><w:t>- Kapal / Vessel  : {{vessel_name}}</w:t></w:r></w:p>
    <w:p/>
    <w:p><w:r><w:rPr><w:b/><w:color w:val="0F172A"/></w:rPr><w:t>RINCIAN GAJI &amp; POTONGAN:</w:t></w:r></w:p>
    <w:p><w:r><w:t>1. Gaji Pokok (Basic Salary)            : $ {{basic_salary}}</w:t></w:r></w:p>
    <w:p><w:r><w:t>2. Tunjangan Layar &amp; Transpor           : $ {{allowances}}</w:t></w:r></w:p>
    <w:p><w:r><w:t>3. Potongan BPJS / Asuransi / Pajak     : $ -{{deductions}}</w:t></w:r></w:p>
    <w:p><w:r><w:t>----------------------------------------------------------------------------------------------------</w:t></w:r></w:p>
    <w:p>
      <w:r><w:rPr><w:b/><w:sz w:val="24"/><w:color w:val="166534"/></w:rPr><w:t>TOTAL TAKE HOME PAY (NET SALARY)       : $ {{net_salary}}</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>====================================================================================================</w:t></w:r></w:p>
    <w:p/>
    <w:p/>
    <w:p>
      <w:r><w:t>Disetujui Oleh,                                                   Diterima Oleh,</w:t></w:r>
    </w:p>
    <w:p/>
    <w:p/>
    <w:p>
      <w:r><w:t>( Finance / Admin Manager )                             ( {{seafarer_name}} )</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

  const blob = createDocxBlob(documentXml);
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "Template_Slip_Gaji_SKYagen_2026.docx";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
  toast.success("Contoh Template Word (.docx) asli berhasil di-download!");
}

function PayslipTemplateGuideModal({ open, onClose }: { open: boolean; onClose: () => void }) {
  return (
    <Modal open={open} onClose={onClose} title="📘 Tutorial & Panduan Pengisian Template Word (.docx)" size="lg"
      footer={<Btn variant="primary" size="sm" onClick={onClose}>Mengerti &amp; Tutup</Btn>}>
      <div className="space-y-4 text-xs text-slate-700">
        <div className="p-3 bg-blue-50 border border-blue-200 rounded-xl">
          <p className="font-bold text-blue-900 text-sm mb-1">💡 Konsep Utama Template Slip Gaji Word</p>
          <p className="text-blue-800 leading-relaxed">
            Admin mendesain tampilan slip gaji di Microsoft Word (.docx) satu kali saja. Sistem akan secara otomatis mengganti kode placeholder seperti <code className="bg-white px-1.5 py-0.5 rounded font-mono text-blue-700 font-bold border border-blue-300">{"{{seafarer_name}}"}</code> dengan data asli pelaut saat slip gaji dicetak atau dikirim.
          </p>
        </div>

        <div className="grid grid-cols-2 gap-3">
          <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-xl space-y-2">
            <h4 className="font-bold text-emerald-900 text-xs uppercase tracking-wider flex items-center gap-1">
              <CheckCircle size={14} className="text-emerald-600" />
              1. Yang HARUS Diisi Kode Placeholder
            </h4>
            <ul className="list-disc pl-4 space-y-1.5 text-[11px] text-emerald-800">
              <li><code className="font-mono font-bold bg-white px-1 rounded">{"{{seafarer_name}}"}</code> — Nama pelaut penerima gaji.</li>
              <li><code className="font-mono font-bold bg-white px-1 rounded">{"{{rank}}"}</code> — Jabatan / Rank pelaut.</li>
              <li><code className="font-mono font-bold bg-white px-1 rounded">{"{{month}}"}</code> &amp; <code className="font-mono font-bold bg-white px-1 rounded">{"{{year}}"}</code> — Bulan &amp; tahun slip.</li>
              <li><code className="font-mono font-bold bg-white px-1 rounded">{"{{net_salary}}"}</code> — Gaji bersih / Take Home Pay.</li>
              <li><code className="font-mono font-bold bg-white px-1 rounded">{"{{basic_salary}}"}</code> — Gaji pokok (jika ada rincian).</li>
              <li><code className="font-mono font-bold bg-white px-1 rounded">{"{{allowances}}"}</code> — Tunjangan layar / transpor.</li>
              <li><code className="font-mono font-bold bg-white px-1 rounded">{"{{deductions}}"}</code> — Potongan BPJS / Pajak.</li>
              <li><code className="font-mono font-bold bg-white px-1 rounded">{"{{vessel_name}}"}</code> — Nama kapal pelaut.</li>
            </ul>
          </div>

          <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl space-y-2">
            <h4 className="font-bold text-slate-900 text-xs uppercase tracking-wider flex items-center gap-1">
              <Info size={14} className="text-slate-600" />
              2. Yang TIDAK HARUS Diisi Tanda {"{{ }}"}
            </h4>
            <ul className="list-disc pl-4 space-y-1.5 text-[11px] text-slate-700">
              <li><strong>Kop / Logo Perusahaan</strong>: Ketik atau sisipkan gambar logo biasa di Word.</li>
              <li><strong>Judul Dokumen</strong>: Ketik biasa (misal: <em>"SLIP GAJI PELAUT"</em>).</li>
              <li><strong>Desain Tabel &amp; Warna</strong>: Buat tabel Word, garis, dan warna sel sesuai selera Admin.</li>
              <li><strong>Label Teks Statis</strong>: Kata-kata seperti <em>"Disetujui Oleh"</em>, <em>"Diterima Oleh"</em>, <em>"PT SKYAGEN MARITIME"</em>.</li>
            </ul>
          </div>
        </div>

        <div className="p-3 bg-red-50 border border-red-200 rounded-xl space-y-1.5">
          <h4 className="font-bold text-red-900 text-xs uppercase tracking-wider flex items-center gap-1">
            <XCircle size={14} className="text-red-600" />
            3. Larangan &amp; Hal yang TIDAK BOLEH Dilakukan
          </h4>
          <ul className="list-disc pl-4 space-y-1 text-[11px] text-red-800">
            <li><strong>❌ JANGAN mengetik nama pelaut asli</strong> (seperti <em>"Budi Santoso"</em>) secara manual di template Word jika template ingin dipakai berulang kali untuk semua pelaut. Gunakan <code className="font-mono font-bold bg-white px-1 rounded">{"{{seafarer_name}}"}</code>.</li>
            <li><strong>❌ JANGAN mengubah penulisan kurung kurawal ganda</strong>. Harus persis <code className="font-mono font-bold bg-white px-1 rounded">{"{{seafarer_name}}"}</code> (bukan <code className="font-mono">{"{seafarer_name}"}</code> atau <code className="font-mono">{"{{seafarer name}}"}</code> dengan spasi).</li>
          </ul>
        </div>
      </div>
    </Modal>
  );
}

function PayslipTemplatesManager() {
  const [showUpload, setShowUpload] = useState(false);
  const [showPreview, setShowPreview] = useState(true);
  const [showGuideModal, setShowGuideModal] = useState(false);
  const [uploadFile, setUploadFile] = useState<File | null>(null);
  const [templateName, setTemplateName] = useState("");
  const [isDefault, setIsDefault] = useState(false);
  const [uploading, setUploading] = useState(false);

  const { data: templates, loading, refresh } = useApiList<ApiPayslipTemplate, ApiPayslipTemplate>(
    "/api/admin/payslips/templates?active=true",
    item => item,
    []
  );

  async function handleUpload() {
    if (!uploadFile) {
      toast.error("Pilih file .docx terlebih dahulu");
      return;
    }
    setUploading(true);
    try {
      await uploadPayslipTemplate(uploadFile, templateName, isDefault);
      toast.success("Template slip gaji berhasil di-upload");
      setShowUpload(false);
      setUploadFile(null);
      setTemplateName("");
      setIsDefault(false);
      refresh();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Upload template gagal");
    } finally {
      setUploading(false);
    }
  }

  const standardPlaceholders = [
    "{{seafarer_name}}",
    "{{rank}}",
    "{{basic_salary}}",
    "{{allowances}}",
    "{{deductions}}",
    "{{net_salary}}",
    "{{month}}",
    "{{year}}",
  ];

  return (
    <Card className="p-5">
      <div className="flex items-center justify-between mb-4">
        <div>
          <h3 className="font-semibold text-slate-800 text-sm">Payslip Word Templates (.docx)</h3>
          <p className="text-xs text-slate-400 mt-1">Upload dan atur template Microsoft Word untuk cetak / kirim slip gaji</p>
        </div>
        <div className="flex items-center gap-2">
          <Btn variant="outline" size="sm" onClick={() => setShowGuideModal(true)}>
            <Info size={13} /> Tutorial &amp; Panduan Word
          </Btn>
          <Btn variant="secondary" size="sm" onClick={() => setShowPreview(!showPreview)}>
            <Eye size={13} /> {showPreview ? "Sembunyikan Preview" : "Lihat Contoh Layout"}
          </Btn>
          <Btn variant="secondary" size="sm" onClick={downloadSamplePayslipDoc}>
            <Download size={13} /> Download Contoh Template (.docx)
          </Btn>
          <Btn variant="primary" size="sm" onClick={() => setShowUpload(true)}>
            <Upload size={13} /> Upload Template (.docx)
          </Btn>
        </div>
      </div>

      <PayslipTemplateGuideModal open={showGuideModal} onClose={() => setShowGuideModal(false)} />

      <div className="mb-4 rounded-xl border border-blue-100 bg-blue-50/60 p-3.5">
        <div className="flex items-center gap-2 mb-1.5">
          <FileText size={13} className="text-blue-600 shrink-0" />
          <span className="text-xs font-semibold text-blue-900">Panduan Placeholder Template Word</span>
        </div>
        <p className="text-[11px] text-blue-700 mb-2">Gunakan placeholder berikut di dalam berkas Word (.docx) Anda untuk pengisian otomatis:</p>
        <div className="flex flex-wrap gap-1.5">
          {standardPlaceholders.map(p => (
            <span key={p} className="font-mono text-[10px] bg-white border border-blue-200 text-blue-800 px-2 py-0.5 rounded font-medium shadow-2xs">
              {p}
            </span>
          ))}
        </div>
      </div>

      {showPreview && (
        <div className="mb-5 rounded-xl border border-slate-200 bg-slate-50/50 p-4 space-y-3">
          <div className="flex items-center justify-between border-b border-slate-200 pb-2">
            <div className="flex items-center gap-2">
              <span className="h-2 w-2 rounded-full bg-emerald-500 animate-pulse" />
              <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wider">Contoh Visual Layout Template Word (.docx)</h4>
            </div>
            <span className="text-[10px] font-semibold text-blue-600 bg-blue-50 border border-blue-200 px-2 py-0.5 rounded">Visual Preview</span>
          </div>

          <div className="mx-auto max-w-2xl bg-white border border-slate-200 rounded-lg p-6 text-slate-800 space-y-4 shadow-sm text-xs">
            <div className="text-center border-b border-slate-300 pb-3">
              <img src={skyagenLogoUrl} alt="SKYagen Logo" className="h-10 mx-auto object-contain mb-1" />
              <p className="text-[11px] font-sans text-slate-500 font-medium">Official Seafarer Payslip Statement</p>
            </div>

            <div className="grid grid-cols-2 gap-2 text-[11px] font-sans bg-slate-50 p-3 rounded border border-slate-200">
              <div>
                <span className="text-slate-400">Nama Pelaut: </span>
                <span className="font-mono font-bold text-blue-700 bg-blue-50 px-1 py-0.5 rounded border border-blue-200">{"{{seafarer_name}}"}</span>
              </div>
              <div>
                <span className="text-slate-400">Jabatan / Rank: </span>
                <span className="font-mono font-bold text-blue-700 bg-blue-50 px-1 py-0.5 rounded border border-blue-200">{"{{rank}}"}</span>
              </div>
              <div>
                <span className="text-slate-400">Bulan: </span>
                <span className="font-mono font-bold text-blue-700 bg-blue-50 px-1 py-0.5 rounded border border-blue-200">{"{{month}}"}</span>
              </div>
              <div>
                <span className="text-slate-400">Tahun: </span>
                <span className="font-mono font-bold text-blue-700 bg-blue-50 px-1 py-0.5 rounded border border-blue-200">{"{{year}}"}</span>
              </div>
            </div>

            <table className="w-full text-[11px] font-sans border-collapse border border-slate-200">
              <thead>
                <tr className="bg-slate-100 border-b border-slate-200">
                  <th className="text-left px-3 py-2 border-r border-slate-200">Deskripsi Komponen Gaji</th>
                  <th className="text-right px-3 py-2">Nominal ($)</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 font-mono">
                <tr>
                  <td className="px-3 py-1.5 font-sans border-r border-slate-200">Gaji Pokok (Basic Salary)</td>
                  <td className="px-3 py-1.5 text-right font-bold text-slate-800">{"{{basic_salary}}"}</td>
                </tr>
                <tr>
                  <td className="px-3 py-1.5 font-sans border-r border-slate-200">Tunjangan Layar &amp; Transpor (Allowances)</td>
                  <td className="px-3 py-1.5 text-right font-bold text-emerald-700">{"{{allowances}}"}</td>
                </tr>
                <tr>
                  <td className="px-3 py-1.5 font-sans border-r border-slate-200">Potongan BPJS / Asuransi (Deductions)</td>
                  <td className="px-3 py-1.5 text-right font-bold text-red-600">-{"{{deductions}}"}</td>
                </tr>
                <tr className="bg-blue-50/80 font-sans font-bold">
                  <td className="px-3 py-2 text-slate-900 border-r border-slate-200">TOTAL TAKE HOME PAY (NET SALARY)</td>
                  <td className="px-3 py-2 text-right text-blue-700 text-sm font-mono">{"{{net_salary}}"}</td>
                </tr>
              </tbody>
            </table>

            <div className="pt-4 flex justify-between text-[11px] font-sans text-slate-600 text-center">
              <div>
                <p className="mb-8">Disetujui Oleh,</p>
                <p className="font-bold border-t border-slate-300 pt-1 px-4">( Finance &amp; HR Manager )</p>
              </div>
              <div>
                <p className="mb-8">Diterima Oleh,</p>
                <p className="font-bold border-t border-slate-300 pt-1 px-4">( {"{{seafarer_name}}"} )</p>
              </div>
            </div>
          </div>
        </div>
      )}

      <div className="overflow-x-auto rounded-xl border border-slate-100">
        <table className="w-full text-sm">
          <thead>
            <tr className="border-b border-slate-100 bg-slate-50">
              {["ID", "Nama Template", "File Word (.docx)", "Default", "Status", "Actions"].map(h => (
                <th key={h} className="px-4 py-3 text-left text-[11px] font-semibold uppercase tracking-wide text-slate-400">{h}</th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-50">
            {templates.map(t => (
              <tr key={t.id} className="hover:bg-slate-50">
                <td className="px-4 py-3 font-mono text-xs text-slate-500">#{t.id}</td>
                <td className="px-4 py-3 text-xs font-medium text-slate-800">{t.name || "Template Standard"}</td>
                <td className="px-4 py-3 text-xs">
                  <a href={`${getApiBaseUrl()}${t.file_url}`} target="_blank" rel="noreferrer" className="text-blue-600 font-mono hover:underline flex items-center gap-1">
                    <Download size={11} /> {t.file_url.split("/").pop()}
                  </a>
                </td>
                <td className="px-4 py-3">
                  {t.is_default ? (
                    <span className="text-[10px] bg-emerald-50 text-emerald-700 px-2 py-0.5 rounded font-semibold border border-emerald-200">Default</span>
                  ) : (
                    <span className="text-[10px] text-slate-400">-</span>
                  )}
                </td>
                <td className="px-4 py-3"><Badge status={t.is_active !== false ? "active" : "inactive"} /></td>
                <td className="px-4 py-3">
                  <a href={`${getApiBaseUrl()}${t.file_url}`} target="_blank" rel="noreferrer" className="px-2.5 py-1 text-[10px] rounded bg-slate-100 text-slate-600 hover:bg-slate-200 transition-colors">
                    Download
                  </a>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {!loading && templates.length === 0 && (
          <EmptyState title="Belum ada template Word" description="Upload template Word (.docx) pertama Anda untuk mengaktifkan pembuat slip gaji." action={<Btn variant="primary" size="sm" onClick={() => setShowUpload(true)}><Upload size={13} />Upload Template</Btn>} />
        )}
      </div>

      <Modal open={showUpload} onClose={() => setShowUpload(false)} title="Upload Payslip Word Template (.docx)" size="md"
        footer={<><Btn variant="secondary" size="sm" onClick={() => setShowUpload(false)}>Cancel</Btn><Btn variant="primary" size="sm" onClick={handleUpload} disabled={uploading}>{uploading ? "Uploading..." : "Upload Template"}</Btn></>}>
        <div className="space-y-4">
          <Input label="Nama Template" value={templateName} onChange={setTemplateName} placeholder="Contoh: Template Standard 2026" />
          <div className="space-y-1.5">
            <label className="text-xs font-medium text-slate-600">File Template Word (.docx) *</label>
            <input type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" onChange={e => setUploadFile(e.target.files?.[0] || null)} className="w-full px-3 py-2 text-xs bg-white border border-slate-200 rounded-lg outline-none cursor-pointer" />
          </div>
          <div className="flex items-center gap-2 pt-2">
            <input type="checkbox" id="is_default_cb" checked={isDefault} onChange={e => setIsDefault(e.target.checked)} className="rounded border-slate-300 text-blue-600 focus:ring-blue-500 cursor-pointer" />
            <label htmlFor="is_default_cb" className="text-xs text-slate-700 font-medium cursor-pointer">Jadikan sebagai template bawaan (Default)</label>
          </div>
        </div>
      </Modal>
    </Card>
  );
}

function PayslipModal({
  open,
  onClose,
  joiningRecords,
  targetRecord,
}: {
  open: boolean;
  onClose: () => void;
  joiningRecords: JoiningRecord[];
  targetRecord?: JoiningRecord | null;
}) {
  const [month, setMonth] = useState<number>(8);
  const [year, setYear] = useState<number>(2026);
  const [selectedTemplateId, setSelectedTemplateId] = useState<number | "">("");
  const [salaries, setSalaries] = useState<Record<number, { basic: number; allowances: number; deductions: number }>>({});
  const [generating, setGenerating] = useState(false);
  const [sendingBulk, setSendingBulk] = useState(false);
  const [sendingSingleId, setSendingSingleId] = useState<number | null>(null);

  const { data: templates } = useApiList<ApiPayslipTemplate, ApiPayslipTemplate>(
    "/api/admin/payslips/templates?active=true",
    item => item,
    []
  );

  const { data: payslips, refresh: refreshPayslips } = useApiList<ApiPayslip, ApiPayslip>(
    `/api/admin/payslips?month=${month}&year=${year}&limit=100`,
    item => item,
    []
  );

  useEffect(() => {
    if (templates.length > 0 && !selectedTemplateId) {
      const def = templates.find(t => t.is_default) || templates[0];
      if (def) setSelectedTemplateId(def.id);
    }
  }, [templates, selectedTemplateId]);

  // Initialize salary forms for joining records
  useEffect(() => {
    const next: Record<number, { basic: number; allowances: number; deductions: number }> = {};
    joiningRecords.forEach(j => {
      const existing = payslips.find(p => p.seafarer_id === j.crewId);
      next[j.crewId] = {
        basic: existing ? existing.basic_salary : (j.salary || 5000000),
        allowances: existing ? existing.allowances : 1500000,
        deductions: existing ? existing.deductions : 500000,
      };
    });
    setSalaries(next);
  }, [joiningRecords, payslips]);

  async function handleGenerateBulk() {
    setGenerating(true);
    try {
      const items = joiningRecords.map(j => {
        const sal = salaries[j.crewId] || { basic: j.salary || 5000000, allowances: 1500000, deductions: 500000 };
        const net = Math.max(0, sal.basic + sal.allowances - sal.deductions);
        return {
          seafarer_id: j.crewId,
          basic_salary: sal.basic,
          allowances: sal.allowances,
          deductions: sal.deductions,
          net_salary: net,
          details_json: JSON.stringify({ gaji_pokok: sal.basic, tunjangan: sal.allowances, potongan: sal.deductions }),
        };
      });

      const res = await generatePayslips({
        month: Number(month),
        year: Number(year),
        template_id: selectedTemplateId ? Number(selectedTemplateId) : undefined,
        items,
      });

      toast.success(res.message || "Slip gaji berhasil di-generate");
      refreshPayslips();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Generate slip gaji gagal");
    } finally {
      setGenerating(false);
    }
  }

  async function handleSingleSend(payslipId: number, crewName: string) {
    setSendingSingleId(payslipId);
    try {
      const res = await sendSinglePayslip(payslipId);
      toast.success(res.message || `Slip gaji untuk ${crewName} berhasil dikirim`);
      refreshPayslips();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Gagal mengirim slip gaji");
    } finally {
      setSendingSingleId(null);
    }
  }

  async function handleBulkSend() {
    setSendingBulk(true);
    try {
      const res = await sendBulkPayslips({ month: Number(month), year: Number(year) });
      toast.success(res.message || `${res.sent_count} slip gaji berhasil dikirim masal`);
      refreshPayslips();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Gagal mengirim slip gaji masal");
    } finally {
      setSendingBulk(false);
    }
  }

  const filteredRecords = targetRecord ? joiningRecords.filter(j => j.id === targetRecord.id) : joiningRecords;

  return (
    <Modal open={open} onClose={onClose} title={`Input & Kirim Slip Gaji (${filteredRecords.length} Pelaut)`} size="xl"
      footer={
        <div className="flex items-center justify-between w-full">
          <Btn variant="secondary" size="sm" onClick={onClose}>Tutup</Btn>
          <div className="flex items-center gap-2">
            <Btn variant="secondary" size="sm" onClick={handleBulkSend} disabled={sendingBulk}>
              {sendingBulk ? <RefreshCw size={13} className="animate-spin" /> : <Send size={13} />}
              Kirim Semua Slip Gaji (Bulan {month}/{year})
            </Btn>
            <Btn variant="primary" size="sm" onClick={handleGenerateBulk} disabled={generating}>
              {generating ? <RefreshCw size={13} className="animate-spin" /> : <FileText size={13} />}
              Generate / Update Slip Gaji
            </Btn>
          </div>
        </div>
      }>
      <div className="space-y-4">
        <div className="grid grid-cols-3 gap-4 bg-slate-50 p-3.5 rounded-xl border border-slate-100">
          <Select label="Bulan Slip Gaji" value={String(month)} onChange={v => setMonth(Number(v))}
            options={Array.from({ length: 12 }, (_, i) => ({ value: String(i + 1), label: `Bulan ${i + 1}` }))} />
          <Select label="Tahun" value={String(year)} onChange={v => setYear(Number(v))}
            options={[2025, 2026, 2027, 2028].map(y => ({ value: String(y), label: String(y) }))} />
          <Select label="Template Word (.docx)" value={String(selectedTemplateId)} onChange={v => setSelectedTemplateId(Number(v))}
            options={templates.map(t => ({ value: String(t.id), label: `${t.name} ${t.is_default ? "(Default)" : ""}` }))} />
        </div>

        <div className="overflow-x-auto rounded-xl border border-slate-100">
          <table className="w-full text-sm">
            <thead>
              <tr className="border-b border-slate-100 bg-slate-50 text-[11px] font-semibold text-slate-400 uppercase tracking-wide">
                <th className="px-3 py-2.5 text-left">Pelaut &amp; Kapal</th>
                <th className="px-3 py-2.5 text-left">Gaji Pokok ($)</th>
                <th className="px-3 py-2.5 text-left">Tunjangan ($)</th>
                <th className="px-3 py-2.5 text-left">Potongan ($)</th>
                <th className="px-3 py-2.5 text-left">Take Home Pay</th>
                <th className="px-3 py-2.5 text-left">Status Slip</th>
                <th className="px-3 py-2.5 text-left">Aksi Single Send</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-50">
              {filteredRecords.map(j => {
                const sal = salaries[j.crewId] || { basic: j.salary || 5000000, allowances: 1500000, deductions: 500000 };
                const net = Math.max(0, sal.basic + sal.allowances - sal.deductions);
                const existingPayslip = payslips.find(p => p.seafarer_id === j.crewId);
                const isSent = existingPayslip?.status === "sent";

                return (
                  <tr key={j.id} className="hover:bg-slate-50/80">
                    <td className="px-3 py-2.5">
                      <p className="font-semibold text-slate-800 text-xs">{j.crewName}</p>
                      <p className="text-[10px] text-slate-400">{j.rank} · {j.vessel}</p>
                    </td>
                    <td className="px-3 py-2.5">
                      <input type="number" value={sal.basic} onChange={e => setSalaries(prev => ({ ...prev, [j.crewId]: { ...sal, basic: Number(e.target.value) } }))} className="w-28 px-2 py-1 text-xs border border-slate-200 rounded outline-none focus:ring-1 focus:ring-blue-500 font-mono" />
                    </td>
                    <td className="px-3 py-2.5">
                      <input type="number" value={sal.allowances} onChange={e => setSalaries(prev => ({ ...prev, [j.crewId]: { ...sal, allowances: Number(e.target.value) } }))} className="w-24 px-2 py-1 text-xs border border-slate-200 rounded outline-none focus:ring-1 focus:ring-blue-500 font-mono text-emerald-700" />
                    </td>
                    <td className="px-3 py-2.5">
                      <input type="number" value={sal.deductions} onChange={e => setSalaries(prev => ({ ...prev, [j.crewId]: { ...sal, deductions: Number(e.target.value) } }))} className="w-24 px-2 py-1 text-xs border border-slate-200 rounded outline-none focus:ring-1 focus:ring-blue-500 font-mono text-red-600" />
                    </td>
                    <td className="px-3 py-2.5 font-bold text-xs text-slate-800 font-mono">
                      ${net.toLocaleString()}
                    </td>
                    <td className="px-3 py-2.5">
                      {isSent ? (
                        <span className="inline-flex items-center gap-1 text-[10px] bg-emerald-50 text-emerald-700 px-2 py-0.5 rounded font-semibold border border-emerald-200">
                          <CheckCircle size={10} /> Terkirim
                        </span>
                      ) : existingPayslip ? (
                        <span className="inline-flex items-center gap-1 text-[10px] bg-blue-50 text-blue-700 px-2 py-0.5 rounded font-semibold border border-blue-200">
                          Generated
                        </span>
                      ) : (
                        <span className="text-[10px] text-slate-400">Belum Buat</span>
                      )}
                    </td>
                    <td className="px-3 py-2.5">
                      {existingPayslip ? (
                        <Btn variant={isSent ? "secondary" : "primary"} size="sm" onClick={() => handleSingleSend(existingPayslip.id, j.crewName)} disabled={sendingSingleId === existingPayslip.id}>
                          {sendingSingleId === existingPayslip.id ? <RefreshCw size={11} className="animate-spin" /> : <Send size={11} />}
                          {isSent ? "Kirim Ulang" : "Kirim Slip"}
                        </Btn>
                      ) : (
                        <span className="text-[11px] text-slate-400 italic">Generate dulu</span>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </Modal>
  );
}

// ─── Joining Principal ────────────────────────────────────────────────────────

function JoiningPrincipalPage() {
  const { data, loading } = useApiList<ApiJoining, JoiningRecord>("/api/joining-principals", mapJoining, []);
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [showForm, setShowForm] = useState(false);
  const [showRequirements, setShowRequirements] = useState(false);
  const [showPayslipModal, setShowPayslipModal] = useState(false);
  const [targetPayslipJoining, setTargetPayslipJoining] = useState<JoiningRecord | null>(null);
  const [page, setPageNum] = useState(1);
  const PAGE_SIZE = 8;

  const filtered = useMemo(() =>
    data.filter(j =>
      (search === "" || j.crewName.toLowerCase().includes(search.toLowerCase()) || j.vessel.toLowerCase().includes(search.toLowerCase())) &&
      (statusFilter === "" || j.status === statusFilter)
    ), [data, search, statusFilter]);

  return (
    <div>
      <PageHeader title="Joining Principal">
        <Btn variant="secondary" size="sm" onClick={() => { setTargetPayslipJoining(null); setShowPayslipModal(true); }}>
          <FileText size={13} /> Slip Gaji / Payslips
        </Btn>
        <Btn variant="primary" size="sm" onClick={() => setShowForm(true)}><Plus size={13} />New Joining</Btn>
      </PageHeader>

      <Card>
        <div className="flex items-center gap-3 px-4 py-3 border-b border-slate-100">
          <SearchBar value={search} onChange={v => { setSearch(v); setPageNum(1); }} placeholder="Search crew, vessel…" />
          <Select value={statusFilter} onChange={v => { setStatusFilter(v); setPageNum(1); }} placeholder="All Status"
            options={["pending", "approved", "onboard", "completed", "cancelled"].map(s => ({ value: s, label: s.charAt(0).toUpperCase() + s.slice(1) }))} />
          <div className="ml-auto text-xs text-slate-400">{loading ? "Loading..." : `${filtered.length} records`}</div>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="border-b border-slate-100">
                {["Crew", "Rank", "Principal", "Vessel", "Status", "Sign On", "Sign Off", "Port", "Salary", "Actions"].map(h => (
                  <th key={h} className="px-4 py-3 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-wide whitespace-nowrap">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-50">
              {filtered.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE).map(j => (
                <tr key={j.id} className="hover:bg-slate-50 group">
                  <td className="px-4 py-3">
                    <div className="flex items-center gap-2">
                      <div className="w-7 h-7 bg-blue-100 rounded-full flex items-center justify-center shrink-0">
                        <span className="text-blue-700 text-[10px] font-bold">{j.crewName.split(" ").map(n => n[0]).join("").slice(0, 2)}</span>
                      </div>
                      <span className="text-xs font-medium text-slate-800">{j.crewName}</span>
                    </div>
                  </td>
                  <td className="px-4 py-3 text-xs text-slate-600">{j.rank}</td>
                  <td className="px-4 py-3 text-xs text-slate-600 max-w-[120px] truncate">{j.principal}</td>
                  <td className="px-4 py-3 text-xs text-slate-600">{j.vessel}</td>
                  <td className="px-4 py-3"><Badge status={j.status} /></td>
                  <td className="px-4 py-3 text-xs text-slate-600 whitespace-nowrap">{fmtDate(j.signOn)}</td>
                  <td className="px-4 py-3 text-xs text-slate-600 whitespace-nowrap">{fmtDate(j.signOff)}</td>
                  <td className="px-4 py-3 text-xs text-slate-600">{j.port}</td>
                  <td className="px-4 py-3 text-xs font-semibold text-slate-700">${j.salary.toLocaleString()}</td>
                  <td className="px-4 py-3">
                    <div className="flex gap-1">
                      <button onClick={() => { setTargetPayslipJoining(j); setShowPayslipModal(true); }} className="p-1.5 rounded hover:bg-emerald-50 text-slate-400 hover:text-emerald-600 transition-colors" title="Slip Gaji (Payslip)"><FileText size={12} /></button>
                      <button onClick={() => setShowRequirements(true)} className="p-1.5 rounded hover:bg-blue-50 text-slate-400 hover:text-blue-600 transition-colors" title="Check requirements"><Shield size={12} /></button>
                      <button className="p-1.5 rounded hover:bg-amber-50 text-slate-400 hover:text-amber-600 transition-colors" title="Edit Joining"><Pencil size={12} /></button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          {filtered.length === 0 && <EmptyState title="No joining records found" description="Adjust your filters or create a new joining record." />}
        </div>
        <Pagination page={page} total={filtered.length} pageSize={PAGE_SIZE} onPage={setPageNum} />
      </Card>

      {/* New Joining Form */}
      <JoiningFormModal open={showForm} onClose={() => setShowForm(false)} onRequirementsCheck={() => { setShowForm(false); setShowRequirements(true); }} />

      {/* Requirements Modal */}
      <RequirementsModal open={showRequirements} onClose={() => setShowRequirements(false)} />

      {/* Payslips Modal */}
      <PayslipModal open={showPayslipModal} onClose={() => setShowPayslipModal(false)} joiningRecords={data} targetRecord={targetPayslipJoining} />
    </div>
  );
}

function JoiningFormModal({ open, onClose, onRequirementsCheck, onCreated }: {
  open: boolean;
  onClose: () => void;
  onRequirementsCheck: () => void;
  onCreated?: () => void;
}) {
  const [form, setForm] = useState({
    crew: "",
    principal: "",
    vessel: "",
    rank: "",
    port: "",
    signOn: "",
    signOff: "",
    salary: "",
    agentFee: "",
    status: "pending"
  });
  const [loading, setLoading] = useState(false);
  const { data: crews } = useApiList<ApiSeafarer, Crew>("/api/seafarers", mapCrew, []);
  const { data: principals } = useApiList<ApiPrincipal, Principal>("/api/admin/principals", mapPrincipal, []);
  const { data: vessels } = useApiList<ApiVessel, Vessel>("/api/admin/vessels", mapVessel, []);

  function f(k: string) { return (v: string) => setForm(p => ({ ...p, [k]: v })); }
  function check() { onRequirementsCheck(); }

  async function submit(e?: React.FormEvent) {
    if (e) e.preventDefault();
    if (!form.crew) {
      toast.error("Crew Member wajib dipilih");
      return;
    }
    if (!form.principal) {
      toast.error("Principal wajib dipilih");
      return;
    }
    if (!form.vessel) {
      toast.error("Vessel (Kapal) wajib dipilih");
      return;
    }
    if (!form.rank) {
      toast.error("Rank / Jabatan wajib dipilih");
      return;
    }
    if (!form.port.trim()) {
      toast.error("Port of Departure (Pelabuhan Keberangkatan) wajib diisi");
      return;
    }
    if (!form.signOn) {
      toast.error("Sign On Date wajib diisi");
      return;
    }

    setLoading(true);
    try {
      const selectedCrew = crews.find(c => c.code === form.crew || String(c.id) === form.crew);
      const selectedPrincipal = principals.find(p => p.code === form.principal || String(p.id) === form.principal);
      const selectedVessel = vessels.find(v => v.code === form.vessel || String(v.id) === form.vessel);

      await apiJson("/api/joining-principals", "POST", {
        seafarer_id: selectedCrew?.id,
        principal_id: selectedPrincipal?.id,
        vessel_id: selectedVessel?.id,
        vessel_name: selectedVessel?.name || form.vessel,
        rank: form.rank,
        port_join: form.port,
        sign_on: form.signOn || null,
        sign_off: form.signOff || null,
        salary_crew: parseFloat(form.salary) || 0,
        agency_fee: parseFloat(form.agentFee) || 0,
        status: form.status,
      });

      toast.success("Joining record berhasil dibuat!");
      setForm({ crew: "", principal: "", vessel: "", rank: "", port: "", signOn: "", signOff: "", salary: "", agentFee: "", status: "pending" });
      onClose();
      if (onCreated) onCreated();
      else window.location.reload();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Gagal membuat record Joining");
    } finally {
      setLoading(false);
    }
  }

  return (
    <Modal open={open} onClose={onClose} title="New Joining Principal" size="xl"
      footer={
        <>
          <Btn variant="secondary" size="sm" onClick={onClose}>Cancel</Btn>
          <Btn variant="outline" size="sm" onClick={check}><Shield size={12} />Check Requirements</Btn>
          <Btn variant="primary" size="sm" onClick={() => submit()} disabled={loading}>
            {loading ? <><RefreshCw size={12} className="animate-spin" />Saving…</> : "Create Joining"}
          </Btn>
        </>
      }>
      <div className="grid grid-cols-2 gap-4 text-xs">
        <Select label="Crew Member" value={form.crew} onChange={f("crew")} required
          options={crews.filter(c => ["available", "waiting"].includes(c.status)).map(c => ({ value: c.code, label: `${c.name} (${c.rank})` }))} />
        <Select label="Principal" value={form.principal} onChange={f("principal")} required
          options={principals.filter(p => p.active).map(p => ({ value: p.code, label: p.name }))} />
        <Select label="Vessel" value={form.vessel} onChange={f("vessel")} required
          options={vessels.filter(v => v.active).map(v => ({ value: v.code, label: v.name }))} />
        <SelectWithOther label="Rank" value={form.rank} onChange={f("rank")} required
          options={["Master", "Chief Officer", "Second Officer", "Third Officer", "Chief Engineer", "Second Engineer", "Third Engineer", "Fourth Engineer", "Bosun", "Able Seaman", "Ordinary Seaman", "Chief Cook", "Cook", "Messman", "Oiler", "Wiper", "Deck Cadet", "Engine Cadet", "Electrician", "Fitter", "Pumpman"].map(r => ({ value: r, label: r }))} />
        <SelectWithOther label="Status" value={form.status} onChange={f("status")}
          options={["Pending", "Approved", "Onboard", "Completed", "Cancelled"].map(s => ({ value: s.toLowerCase(), label: s }))} />
        <Input label="Port of Departure" value={form.port} onChange={f("port")} required placeholder="e.g. Singapore" />
        <Input label="Sign On Date" type="date" value={form.signOn} onChange={f("signOn")} required />
        <Input label="Expected Sign Off" type="date" value={form.signOff} onChange={f("signOff")} />
        <Input label="Basic Salary (USD)" type="number" value={form.salary} onChange={f("salary")} placeholder="0.00" />
        <Input label="Agent Fee (USD)" type="number" value={form.agentFee} onChange={f("agentFee")} placeholder="0.00" />
        <div className="col-span-2 p-3 bg-slate-50 rounded-lg border border-slate-100 flex items-center justify-between">
          <span className="text-xs text-slate-500">Total Salary (Salary + Agent Fee)</span>
          <span className="text-sm font-bold text-slate-800">
            ${((parseFloat(form.salary) || 0) + (parseFloat(form.agentFee) || 0)).toFixed(2)}
          </span>
        </div>
        <div className="col-span-2">
          <label className="text-xs font-medium text-slate-600 block mb-1">Upload CV / Documents</label>
          <div className="border-2 border-dashed border-slate-200 rounded-xl p-6 text-center hover:border-blue-300 cursor-pointer transition-colors">
            <Upload size={20} className="text-slate-300 mx-auto mb-1" />
            <p className="text-xs text-slate-400">Drag &amp; drop or click to upload PDF, DOC</p>
          </div>
        </div>
      </div>
    </Modal>
  );
}

function RequirementsModal({ open, onClose }: { open: boolean; onClose: () => void }) {
  const requirements = [
    { name: "Passport", status: "passed", doc: "A1234567", expiry: "2028-03-15" },
    { name: "Seaman Book", status: "passed", doc: "D123456", expiry: "2027-11-20" },
    { name: "STCW Certificate", status: "passed", doc: "STCW-SEA001", expiry: "2025-08-01" },
    { name: "Medical Certificate", status: "manual-review", doc: "MED-2024-001", expiry: "2024-03-15" },
    { name: "Basic Safety Training", status: "passed", doc: "BST-SEA001", expiry: "2026-06-30" },
    { name: "GMDSS Certificate", status: "failed", doc: null, expiry: null },
  ];

  const icons = {
    passed: <CheckCircle size={14} className="text-emerald-500" />,
    failed: <XCircle size={14} className="text-red-500" />,
    "manual-review": <AlertTriangle size={14} className="text-amber-500" />,
  };

  const colors = {
    passed: "bg-emerald-50 border-emerald-100",
    failed: "bg-red-50 border-red-100",
    "manual-review": "bg-amber-50 border-amber-100",
  };

  const summary = {
    passed: requirements.filter(r => r.status === "passed").length,
    failed: requirements.filter(r => r.status === "failed").length,
    review: requirements.filter(r => r.status === "manual-review").length,
  };

  return (
    <Modal open={open} onClose={onClose} title="Principal Requirements Check" size="lg"
      footer={<Btn variant="primary" size="sm" onClick={onClose}>Close</Btn>}>
      <div className="grid grid-cols-3 gap-3 mb-5">
        {[
          { label: "Passed", count: summary.passed, color: "bg-emerald-50 text-emerald-700 border-emerald-200" },
          { label: "Failed", count: summary.failed, color: "bg-red-50 text-red-700 border-red-200" },
          { label: "Manual Review", count: summary.review, color: "bg-amber-50 text-amber-700 border-amber-200" },
        ].map(({ label, count, color }) => (
          <div key={label} className={cn("p-3 rounded-xl border text-center", color)}>
            <p className="text-xl font-bold">{count}</p>
            <p className="text-xs mt-0.5">{label}</p>
          </div>
        ))}
      </div>
      <div className="space-y-2">
        {requirements.map((r, i) => (
          <div key={i} className={cn("flex items-center gap-3 p-3 rounded-xl border", colors[r.status as keyof typeof colors])}>
            {icons[r.status as keyof typeof icons]}
            <div className="flex-1">
              <p className="text-xs font-medium text-slate-800">{r.name}</p>
              {r.doc && <p className="text-[10px] text-slate-500 font-mono mt-0.5">{r.doc} · Exp: {fmtDate(r.expiry)}</p>}
              {!r.doc && <p className="text-[10px] text-red-500 mt-0.5">Document not found</p>}
            </div>
            <span className={cn("text-[10px] font-semibold capitalize px-2 py-0.5 rounded",
              r.status === "passed" ? "bg-emerald-100 text-emerald-700" :
              r.status === "failed" ? "bg-red-100 text-red-700" : "bg-amber-100 text-amber-700")}>
              {r.status.replace("-", " ")}
            </span>
          </div>
        ))}
      </div>
    </Modal>
  );
}

// ─── Operations ───────────────────────────────────────────────────────────────

function SignOnPage() {
  const [form, setForm] = useState({ joining: "", date: "", port: "", note: "" });
  const [loading, setLoading] = useState(false);
  const { data: joinings, refresh } = useApiList<ApiJoining, JoiningRecord>("/api/joining-principals", mapJoining, []);
  function f(k: string) { return (v: string) => setForm(p => ({ ...p, [k]: v })); }
  async function submit(e: React.FormEvent) {
    e.preventDefault();
    if (!form.joining || !form.date || !form.port) { toast.error("Joining record, date, and port are required"); return; }
    setLoading(true);
    try {
      await apiJson("/api/operations/sign-on", "POST", { joining_principal_id: Number(form.joining), sign_on_date: form.date, port: form.port, note: form.note });
      toast.success("Sign On recorded successfully");
      setForm({ joining: "", date: "", port: "", note: "" });
      refresh();
    } catch (err) { toast.error(err instanceof Error ? err.message : "Sign on failed"); } finally { setLoading(false); }
  }
  const candidates = joinings.filter(j => !j.signOn || ["pending", "approved", "waiting", "-"].includes(j.status.toLowerCase()));
  const recent = joinings.filter(j => j.signOn && !j.signOff);
  return <div><PageHeader title="Sign On" /><div className="grid grid-cols-3 gap-4"><div className="col-span-2"><Card className="p-5"><h3 className="font-semibold text-slate-700 text-sm mb-4">Sign On Details</h3><form onSubmit={submit} className="grid grid-cols-2 gap-4"><Select label="Joining Record" value={form.joining} onChange={f("joining")} required placeholder="Select joining" options={candidates.map(j => ({ value: String(j.id), label: `${j.crewName} - ${j.vessel}` }))} className="col-span-2" /><Input label="Sign On Date" type="date" value={form.date} onChange={f("date")} required /><Input label="Port of Embarkation" value={form.port} onChange={f("port")} required placeholder="e.g. Singapore" /><Textarea label="Note" value={form.note} onChange={f("note")} placeholder="Optional remarks..." className="col-span-2" /><div className="col-span-2 flex justify-end gap-2 pt-2 border-t border-slate-100"><Btn variant="secondary" type="button" onClick={() => setForm({ joining: "", date: "", port: "", note: "" })}>Clear</Btn><Btn variant="primary" type="submit" disabled={loading}>{loading ? <><RefreshCw size={13} className="animate-spin" />Processing...</> : <><CheckCircle size={13} />Record Sign On</>}</Btn></div></form></Card></div><Card className="p-4 self-start"><h3 className="font-semibold text-slate-700 text-sm mb-3">Recent Sign Ons</h3><div className="space-y-2">{recent.map(j => <div key={j.id} className="p-2.5 bg-slate-50 rounded-lg"><p className="text-xs font-medium text-slate-700">{j.crewName}</p><p className="text-[10px] text-slate-400 mt-0.5">{j.vessel} - {fmtDate(j.signOn)}</p></div>)}</div></Card></div></div>;
}

function SignOffPage() {
  const [form, setForm] = useState({ joining: "", date: "", location: "", note: "" });
  const [loading, setLoading] = useState(false);
  const { data: joinings, refresh } = useApiList<ApiJoining, JoiningRecord>("/api/joining-principals", mapJoining, []);
  function f(k: string) { return (v: string) => setForm(p => ({ ...p, [k]: v })); }
  async function submit(e: React.FormEvent) {
    e.preventDefault();
    if (!form.joining || !form.date) { toast.error("Joining record and sign off date are required"); return; }
    setLoading(true);
    try {
      await apiJson("/api/operations/sign-off", "POST", { joining_principal_id: Number(form.joining), sign_off_date: form.date, port: form.location, note: form.note });
      toast.success("Sign Off recorded successfully");
      setForm({ joining: "", date: "", location: "", note: "" });
      refresh();
    } catch (err) { toast.error(err instanceof Error ? err.message : "Sign off failed"); } finally { setLoading(false); }
  }
  const onboardCrew = joinings.filter(j => j.status.toLowerCase().includes("board") || (j.signOn && !j.signOff));
  return <div><PageHeader title="Sign Off" /><div className="grid grid-cols-3 gap-4"><div className="col-span-2"><Card className="p-5"><h3 className="font-semibold text-slate-700 text-sm mb-4">Sign Off Details</h3><form onSubmit={submit} className="grid grid-cols-2 gap-4"><Select label="Crew Member (Onboard)" value={form.joining} onChange={f("joining")} required placeholder="Select crew" options={onboardCrew.map(j => ({ value: String(j.id), label: `${j.crewName} - ${j.vessel}` }))} className="col-span-2" /><Input label="Sign Off Date" type="date" value={form.date} onChange={f("date")} required /><Input label="Port of Disembarkation" value={form.location} onChange={f("location")} placeholder="e.g. Singapore" /><Textarea label="Note" value={form.note} onChange={f("note")} placeholder="Optional remarks..." className="col-span-2" /><div className="col-span-2 flex justify-end gap-2 pt-2 border-t border-slate-100"><Btn variant="secondary" type="button" onClick={() => setForm({ joining: "", date: "", location: "", note: "" })}>Clear</Btn><Btn variant="primary" type="submit" disabled={loading}>{loading ? <><RefreshCw size={13} className="animate-spin" />Processing...</> : <><XCircle size={13} />Record Sign Off</>}</Btn></div></form></Card></div><Card className="p-4 self-start"><h3 className="font-semibold text-slate-700 text-sm mb-3">Currently Onboard</h3><div className="space-y-2">{onboardCrew.map(j => <div key={j.id} className="p-2.5 bg-blue-50 rounded-lg border border-blue-100"><p className="text-xs font-medium text-blue-800">{j.crewName}</p><p className="text-[10px] text-blue-500 mt-0.5">{j.vessel}</p><p className="text-[10px] text-slate-400 mt-0.5">Since {fmtDate(j.signOn)}</p></div>)}</div></Card></div></div>;
}
// ─── Documents ────────────────────────────────────────────────────────────────

function DocumentsPage({ setPage }: { setPage: (p: string) => void }) {
  const [summary, setSummary] = useState<ApiDocumentReport | null>(null);
  const { data: expiring } = useApiList<ApiExpiringDoc, ApiExpiringDoc>("/api/documents/expiring?days=365&limit=20", item => item, []);
  const { data: docTypes } = useApiList<ApiDocumentType, ReturnType<typeof mapDocType>>("/api/admin/document-types", mapDocType, []);
  const { data: docNames } = useApiList<ApiDocumentName, ReturnType<typeof mapDocName>>("/api/admin/document-names", mapDocName, []);
  useEffect(() => { apiGet<ApiDocumentReport>("/api/documents/report").then(setSummary).catch(err => toast.error(err.message)); }, []);
  const rows = expiring.map(d => ({ id: d.id, crew: d.seafarer_name, document: d.document_name?.name || "-", number: d.document_no || "-", expiry: toDateInput(d.expired_at), status: d.expired_at && new Date(d.expired_at).getTime() < Date.now() ? "expired" : "expiring" }));
  return <div><PageHeader title="Documents" /><div className="grid grid-cols-2 gap-4 mb-4">{[
    { label: "Expiring Documents", count: summary?.expiring_count ?? 0, color: "bg-amber-500", page: "expiring-documents", icon: FileWarning },
    { label: "Expired Documents", count: summary?.expired_count ?? 0, color: "bg-red-500", page: "expiring-documents", icon: AlertCircle },
    { label: "Document Types", count: docTypes.length, color: "bg-blue-500", page: "document-types", icon: Layers },
    { label: "Document Names", count: docNames.length, color: "bg-slate-600", page: "document-names", icon: FileText },
  ].map(({ label, count, color, page: p, icon: Icon }) => <button key={label} onClick={() => setPage(p)} className="bg-white border border-slate-100 rounded-xl p-4 flex items-center gap-4 hover:shadow-md transition-shadow text-left"><div className={cn("w-10 h-10 rounded-xl flex items-center justify-center", color)}><Icon size={18} className="text-white" /></div><div><p className="text-xl font-bold text-slate-800">{count}</p><p className="text-xs text-slate-500">{label}</p></div><ArrowRight size={14} className="text-slate-300 ml-auto" /></button>)}</div><Card><div className="px-4 py-3 border-b border-slate-100 flex items-center justify-between"><h3 className="font-semibold text-slate-700 text-sm">Tracked Crew Documents</h3><Btn variant="secondary" size="sm" onClick={() => exportExcel("/api/reports/export?type=documents", undefined, "document-report.xlsx").catch(err => toast.error(err.message))}><Download size={12} />Export</Btn></div><table className="w-full text-sm"><thead><tr className="border-b border-slate-100">{["Crew", "Document", "Number", "Expiry", "Status"].map(h => <th key={h} className="px-4 py-3 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-wide">{h}</th>)}</tr></thead><tbody className="divide-y divide-slate-50">{rows.map(d => <tr key={d.id} className="hover:bg-slate-50"><td className="px-4 py-3 font-medium text-slate-800 text-xs">{d.crew}</td><td className="px-4 py-3 text-xs text-slate-600">{d.document}</td><td className="px-4 py-3 font-mono text-xs text-slate-500">{d.number}</td><td className="px-4 py-3 text-xs text-slate-600">{fmtDate(d.expiry)}</td><td className="px-4 py-3"><Badge status={d.status} /></td></tr>)}</tbody></table>{rows.length === 0 && <EmptyState title="No tracked documents" />}</Card></div>;
}
function ExpiringDocumentsPage() {
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("expiring");
  const [page, setPageNum] = useState(1);
  const PAGE_SIZE = 8;
  const path = statusFilter === "expired" ? "/api/documents/expired" : "/api/documents/expiring";
  const { data, loading } = useApiList<ApiExpiringDoc, ApiExpiringDoc>(`${path}?days=30&limit=100`, item => item, []);
  const [summary, setSummary] = useState<ApiDocumentReport | null>(null);

  useEffect(() => {
    apiGet<ApiDocumentReport>("/api/documents/report").then(setSummary).catch(err => toast.error(err.message));
  }, []);

  const rows = useMemo(() => data.map(d => {
    const expiry = toDateInput(d.expired_at);
    const daysLeft = d.expired_at ? Math.ceil((new Date(d.expired_at).getTime() - Date.now()) / 86400000) : 0;
    return { id: d.id, crew: d.seafarer_name, document: d.document_name?.name || "-", number: d.document_no || "-", expiry, daysLeft, status: daysLeft < 0 ? "expired" : "expiring" };
  }).filter(d => search === "" || d.crew.toLowerCase().includes(search.toLowerCase()) || d.document.toLowerCase().includes(search.toLowerCase())), [data, search]);

  return (
    <div>
      <PageHeader title="Expiring Documents">
        <Btn variant="secondary" size="sm" onClick={() => exportExcel("/api/reports/export?type=documents", undefined, "document-report.xlsx").catch(err => toast.error(err.message))}><Download size={13} />Export</Btn>
      </PageHeader>
      <div className="grid grid-cols-4 gap-3 mb-4">
        {[
          { label: "Expired", value: summary?.expired_count ?? 0, color: "text-red-600 bg-red-50" },
          { label: "Expiring (<30d)", value: summary?.expiring_count ?? 0, color: "text-amber-600 bg-amber-50" },
          { label: "Valid", value: summary?.valid_count ?? 0, color: "text-emerald-600 bg-emerald-50" },
          { label: "Total Tracked", value: summary?.total_documents ?? 0, color: "text-blue-600 bg-blue-50" },
        ].map(({ label, value, color }) => <div key={label} className={cn("rounded-xl p-3 text-center", color)}><p className="text-xl font-bold">{value}</p><p className="text-xs mt-0.5 opacity-80">{label}</p></div>)}
      </div>
      <Card>
        <div className="flex items-center gap-3 px-4 py-3 border-b border-slate-100">
          <SearchBar value={search} onChange={v => { setSearch(v); setPageNum(1); }} placeholder="Search crew, document..." />
          <Select value={statusFilter} onChange={v => { setStatusFilter(v || "expiring"); setPageNum(1); }} placeholder="Status" options={[{ value: "expiring", label: "Expiring" }, { value: "expired", label: "Expired" }]} />
          <div className="ml-auto text-xs text-slate-400">{loading ? "Loading..." : `${rows.length} records`}</div>
        </div>
        <table className="w-full text-sm"><thead><tr className="border-b border-slate-100">{["Crew", "Document", "Number", "Expiry Date", "Days Remaining", "Status"].map(h => <th key={h} className="px-4 py-3 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-wide">{h}</th>)}</tr></thead>
          <tbody className="divide-y divide-slate-50">{rows.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE).map(d => <tr key={d.id} className="hover:bg-slate-50"><td className="px-4 py-3 font-medium text-slate-800 text-xs">{d.crew}</td><td className="px-4 py-3 text-xs text-slate-600">{d.document}</td><td className="px-4 py-3 font-mono text-xs text-slate-500">{d.number}</td><td className="px-4 py-3 text-xs text-slate-600">{fmtDate(d.expiry)}</td><td className="px-4 py-3"><span className={cn("text-xs font-semibold font-mono", d.daysLeft < 0 ? "text-red-600" : "text-amber-600")}>{d.daysLeft < 0 ? `${Math.abs(d.daysLeft)}d ago` : `${d.daysLeft}d`}</span></td><td className="px-4 py-3"><Badge status={d.status} /></td></tr>)}</tbody>
        </table>
        {rows.length === 0 && <EmptyState title="No documents found" description="No documents match your search." />}
        <Pagination page={page} total={rows.length} pageSize={PAGE_SIZE} onPage={setPageNum} />
      </Card>
    </div>
  );
}
function DocumentTypesPage() {
  const { data, loading, refresh } = useApiList<ApiDocumentType, ReturnType<typeof mapDocType>>("/api/admin/document-types", mapDocType, []);
  const [search, setSearch] = useState("");
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState<ReturnType<typeof mapDocType> | null>(null);
  const [form, setForm] = useState({ name: "", description: "", active: "true" });

  const filtered = useMemo(() => {
    const q = search.toLowerCase();
    return data.filter(item => q === "" || item.name.toLowerCase().includes(q) || item.description.toLowerCase().includes(q));
  }, [data, search]);

  function openCreate() {
    setEditing(null);
    setForm({ name: "", description: "", active: "true" });
    setShowForm(true);
  }

  function openEdit(item: ReturnType<typeof mapDocType>) {
    setEditing(item);
    setForm({ name: item.name, description: item.description, active: item.active ? "true" : "false" });
    setShowForm(true);
  }

  function closeForm() {
    setShowForm(false);
    setEditing(null);
    setForm({ name: "", description: "", active: "true" });
  }

  async function saveType() {
    if (!form.name.trim()) { toast.error("Type name is required"); return; }
    try {
      const body = { name: form.name.trim(), description: form.description.trim(), is_active: form.active === "true" };
      if (editing) {
        await apiJson(`/api/admin/document-types/${editing.id}`, "PUT", body);
        toast.success("Document type updated");
      } else {
        await apiJson("/api/admin/document-types", "POST", body);
        toast.success("Document type added");
      }
      closeForm();
      refresh();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Save failed");
    }
  }

  async function toggleType(item: ReturnType<typeof mapDocType>) {
    try {
      await apiJson(`/api/admin/document-types/${item.id}`, "PUT", { name: item.name, description: item.description, is_active: !item.active });
      refresh();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Update failed");
    }
  }

  async function deleteType(id: number) {
    if (!confirm("Delete this document type?")) return;
    try {
      await apiJson(`/api/admin/document-types/${id}`, "DELETE");
      toast.success("Document type deleted");
      refresh();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Delete failed");
    }
  }

  return (
    <div>
      <PageHeader title="Document Types">
        <Btn variant="primary" size="sm" onClick={openCreate}><Plus size={13} />Add Type</Btn>
      </PageHeader>
      <Card>
        <div className="flex items-center gap-3 px-4 py-3 border-b border-slate-100">
          <SearchBar value={search} onChange={setSearch} placeholder="Search document categories..." />
          <div className="ml-auto text-xs text-slate-400">{loading ? "Loading..." : `${filtered.length} document types`}</div>
        </div>
        <table className="w-full text-sm">
          <thead>
            <tr className="border-b border-slate-100">
              {["Code", "Category / Type", "Description", "Status", "Actions"].map(h => (
                <th key={h} className="px-4 py-3 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-wide">{h}</th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-50">
            {filtered.map(t => (
              <tr key={t.id} className="hover:bg-slate-50 group">
                <td className="px-4 py-3 font-mono text-xs font-semibold text-slate-600">{t.code}</td>
                <td className="px-4 py-3 font-medium text-slate-800 text-xs">{t.name}</td>
                <td className="px-4 py-3 text-xs text-slate-500 max-w-md truncate">{t.description || "-"}</td>
                <td className="px-4 py-3"><Badge status={t.active ? "active" : "inactive"} /></td>
                <td className="px-4 py-3">
                  <div className="flex gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                    <button onClick={() => openEdit(t)} className="p-1.5 rounded hover:bg-amber-50 text-slate-400 hover:text-amber-600 transition-colors"><Pencil size={12} /></button>
                    <button onClick={() => toggleType(t)} className="px-2 py-1 text-[10px] rounded bg-slate-50 text-slate-600">{t.active ? "Disable" : "Enable"}</button>
                    <button onClick={() => deleteType(t.id)} className="p-1.5 rounded hover:bg-red-50 text-slate-400 hover:text-red-500 transition-colors"><Trash2 size={12} /></button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {!loading && filtered.length === 0 && <EmptyState title="No document types" description="Belum ada kategori dokumen." action={<Btn variant="primary" size="sm" onClick={openCreate}><Plus size={13} />Add Type</Btn>} />}
      </Card>
      <Modal open={showForm} onClose={closeForm} title={editing ? "Edit Document Type" : "Add Document Type"}
        footer={<><Btn variant="secondary" size="sm" onClick={closeForm}>Cancel</Btn><Btn variant="primary" size="sm" onClick={saveType}>{editing ? "Save Changes" : "Add Type"}</Btn></>}>
        <div className="space-y-4">
          <Input label="Category / Type Name" value={form.name} onChange={v => setForm(p => ({ ...p, name: v }))} required placeholder="Travel & Identity" />
          <Textarea label="Description" value={form.description} onChange={v => setForm(p => ({ ...p, description: v }))} placeholder="Passport, seaman book, visa, or related identity documents" />
          <Select label="Status" value={form.active} onChange={v => setForm(p => ({ ...p, active: v }))} options={[{ value: "true", label: "Active" }, { value: "false", label: "Inactive" }]} />
        </div>
      </Modal>
    </div>
  );
}
function DocumentNamesPage() {
  const [search, setSearch] = useState("");
  const [showAdd, setShowAdd] = useState(false);
  const [form, setForm] = useState({ type: "", name: "", required: "true", hasExpiry: "true" });
  const { data: docTypes } = useApiList<ApiDocumentType, ReturnType<typeof mapDocType>>("/api/admin/document-types", mapDocType, []);
  const { data, loading, refresh } = useApiList<ApiDocumentName, ReturnType<typeof mapDocName>>("/api/admin/document-names", mapDocName, []);

  const filtered = useMemo(() =>
    data.filter(d => search === "" || d.name.toLowerCase().includes(search.toLowerCase())),
    [data, search]
  );

  async function addName() {
    if (!form.type || !form.name) { toast.error("Type and name are required"); return; }
    try {
      await apiJson("/api/admin/document-names", "POST", { document_type_id: Number(form.type), name: form.name, is_required: form.required === "true", has_expiry: form.hasExpiry === "true", is_active: true });
      toast.success("Document name added");
      setShowAdd(false);
      setForm({ type: "", name: "", required: "true", hasExpiry: "true" });
      refresh();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Save failed");
    }
  }

  async function deleteName(id: number) {
    try {
      await apiJson(`/api/admin/document-names/${id}`, "DELETE");
      toast.success("Document name deleted");
      refresh();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Delete failed");
    }
  }

  return (
    <div>
      <PageHeader title="Document Names">
        <Btn variant="primary" size="sm" onClick={() => setShowAdd(true)}><Plus size={13} />Add Name</Btn>
      </PageHeader>
      <Card>
        <div className="flex items-center gap-3 px-4 py-3 border-b border-slate-100">
          <SearchBar value={search} onChange={setSearch} placeholder="Search document names..." />
          <div className="ml-auto text-xs text-slate-400">{loading ? "Loading..." : `${filtered.length} documents`}</div>
        </div>
        {filtered.length > 0 ? (
          <table className="w-full text-sm">
            <thead>
              <tr className="border-b border-slate-100">
                {["Type", "Document Name", "Required", "Has Expiry", "Actions"].map(h => (
                  <th key={h} className="px-4 py-3 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-wide">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-50">
              {filtered.map(d => (
                <tr key={d.id} className="hover:bg-slate-50 group">
                  <td className="px-4 py-3 text-xs text-slate-500">{d.type}</td>
                  <td className="px-4 py-3 font-medium text-slate-800 text-xs">{d.name}</td>
                  <td className="px-4 py-3">{d.required ? <span className="text-emerald-600 text-xs font-medium flex items-center gap-1"><Check size={11} />Yes</span> : <span className="text-slate-400 text-xs">No</span>}</td>
                  <td className="px-4 py-3">{d.hasExpiry ? <span className="text-blue-600 text-xs font-medium flex items-center gap-1"><Calendar size={11} />Yes</span> : <span className="text-slate-400 text-xs">Lifetime</span>}</td>
                  <td className="px-4 py-3"><button className="p-1.5 rounded hover:bg-red-50 text-slate-400 hover:text-red-500 transition-colors opacity-0 group-hover:opacity-100" onClick={() => deleteName(d.id)}><Trash2 size={12} /></button></td>
                </tr>
              ))}
            </tbody>
          </table>
        ) : !loading ? (
          <div className="py-12">
            <EmptyState title="Belum Ada Nama Dokumen" description="Admin belum mengatur nama dokumen. Silakan atur atau tambahkan nama dokumen baru terlebih dahulu." action={<Btn variant="primary" size="sm" onClick={() => setShowAdd(true)}><Plus size={13} />Add Name</Btn>} />
          </div>
        ) : null}
      </Card>
      <Modal open={showAdd} onClose={() => setShowAdd(false)} title="Add Document Name"
        footer={<><Btn variant="secondary" size="sm" onClick={() => setShowAdd(false)}>Cancel</Btn><Btn variant="primary" size="sm" onClick={addName}>Add</Btn></>}>
        <div className="space-y-4">
          <Select label="Document Type" value={form.type} onChange={v => setForm(p => ({ ...p, type: v }))} required options={docTypes.map(t => ({ value: String(t.id), label: t.name }))} />
          <Input label="Document Name" value={form.name} onChange={v => setForm(p => ({ ...p, name: v }))} required placeholder="e.g. STCW Certificate" />
          <Select label="Required" value={form.required} onChange={v => setForm(p => ({ ...p, required: v }))} options={[{ value: "true", label: "Yes - Required" }, { value: "false", label: "No - Optional" }]} />
          <Select label="Has Expiry Date" value={form.hasExpiry} onChange={v => setForm(p => ({ ...p, hasExpiry: v }))} options={[{ value: "true", label: "Yes - Has Expiry" }, { value: "false", label: "No - Lifetime" }]} />
        </div>
      </Modal>
    </div>
  );
}
// ─── Company & Vessel ─────────────────────────────────────────────────────────

function VesselsPage() {
  const [search, setSearch] = useState("");
  const [showAdd, setShowAdd] = useState(false);
  const [showImport, setShowImport] = useState(false);
  const [vesselForm, setVesselForm] = useState({ code: "", name: "", principalId: "", type: "", flag: "", imo: "", mmsi: "", active: "true", customFields: {} as Record<string, unknown> });
  const [page, setPageNum] = useState(1);
  const PAGE_SIZE = 8;

  const { data: principalOptions } = useApiList<ApiPrincipal, Principal>("/api/admin/principals", mapPrincipal, []);
  const { data: vesselCustomFields } = useApiList<ApiCustomField, CustomField>("/api/admin/custom-fields?entity_type=vessel", mapCustomField, []);
  const { data: vessels, loading, refresh } = useApiList<ApiVessel, Vessel>("/api/admin/vessels", mapVessel, []);

  const activeVesselCustomFields = useMemo(() => vesselCustomFields.filter(field => field.active), [vesselCustomFields]);

  const vesselTypeOptions = useMemo(() => {
    const defaults = ["Bulk Carrier", "Oil Tanker", "Container Ship", "General Cargo", "Chemical Tanker", "Platform Supply"];
    const set = new Set(defaults);
    vessels.forEach(v => { if (v.type) set.add(v.type); });
    return Array.from(set).map(t => ({ value: t, label: t }));
  }, [vessels]);

  const filtered = useMemo(() =>
    vessels.filter(v => search === "" || v.name.toLowerCase().includes(search.toLowerCase()) || v.code.toLowerCase().includes(search.toLowerCase())),
    [vessels, search]
  );

  return (
    <div>
      <PageHeader title="Vessels">
        <Btn variant="secondary" size="sm" onClick={() => setShowImport(true)}><Upload size={13} />Import</Btn>
        <Btn variant="secondary" size="sm" onClick={() => exportExcel("/api/admin/export/seafarers", "id,name,photo_url,phone,nik,passport_no,seaman_book_no,status", "seafarers.xlsx").catch(err => toast.error(err.message))}><Download size={13} />Export</Btn>
        <Btn variant="primary" size="sm" onClick={() => setShowAdd(true)}><Plus size={13} />Add Vessel</Btn>
      </PageHeader>
      <Card>
        <div className="flex items-center gap-3 px-4 py-3 border-b border-slate-100">
          <SearchBar value={search} onChange={v => { setSearch(v); setPageNum(1); }} placeholder="Search vessels…" />
          <div className="ml-auto text-xs text-slate-400">{loading ? "Loading..." : `${filtered.length} vessels`}</div>
        </div>
        {filtered.length > 0 ? (
          <>
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-slate-100">
                  {["Code", "Vessel", "Principal", "Type", "Flag", "IMO", "MMSI", ...activeVesselCustomFields.map(field => field.label), "Status", "Actions"].map(h => (
                    <th key={h} className="px-4 py-3 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-wide whitespace-nowrap">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {filtered.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE).map(v => (
                  <tr key={v.id} className="hover:bg-slate-50 group">
                    <td className="px-4 py-3 font-mono text-xs text-slate-500">{v.code}</td>
                    <td className="px-4 py-3">
                      <div className="flex items-center gap-2">
                        <div className="w-7 h-7 bg-blue-50 rounded-lg flex items-center justify-center shrink-0">
                          <Anchor size={12} className="text-blue-500" />
                        </div>
                        <span className="font-medium text-slate-800 text-xs">{v.name}</span>
                      </div>
                    </td>
                    <td className="px-4 py-3 text-xs text-slate-600">{v.principal}</td>
                    <td className="px-4 py-3 text-xs text-slate-600">{v.type}</td>
                    <td className="px-4 py-3 text-xs text-slate-600">{v.flag}</td>
                    <td className="px-4 py-3 font-mono text-xs text-slate-500">{v.imo}</td>
                    <td className="px-4 py-3 font-mono text-xs text-slate-500">{v.mmsi}</td>
                    {activeVesselCustomFields.map(field => <td key={field.key} className="px-4 py-3 text-xs text-slate-600">{customValueToString(v.customFields[field.key]) || "-"}</td>)}
                    <td className="px-4 py-3"><Badge status={v.active ? "active" : "inactive"} /></td>
                    <td className="px-4 py-3">
                      <div className="flex gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                        <button className="p-1.5 rounded hover:bg-amber-50 text-slate-400 hover:text-amber-600 transition-colors"><Pencil size={12} /></button>
                        <button className="p-1.5 rounded hover:bg-red-50 text-slate-400 hover:text-red-500 transition-colors"><Trash2 size={12} /></button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <Pagination page={page} total={filtered.length} pageSize={PAGE_SIZE} onPage={setPageNum} />
          </>
        ) : !loading ? (
          <div className="py-12">
            <EmptyState title="Belum Ada Data Vessel" description="Admin belum mengatur data vessel. Silakan atur atau tambahkan data vessel baru terlebih dahulu." action={<Btn variant="primary" size="sm" onClick={() => setShowAdd(true)}><Plus size={13} />Add Vessel</Btn>} />
          </div>
        ) : null}
      </Card>

      <Modal open={showAdd} onClose={() => setShowAdd(false)} title="Add Vessel" size="lg"
        footer={<><Btn variant="secondary" size="sm" onClick={() => setShowAdd(false)}>Cancel</Btn><Btn variant="primary" size="sm" onClick={async () => { try { await apiJson("/api/admin/vessels", "POST", { code: vesselForm.code, name: vesselForm.name, principal_id: Number(vesselForm.principalId), vessel_type: vesselForm.type, flag: vesselForm.flag, imo: vesselForm.imo, mmsi: vesselForm.mmsi, is_active: vesselForm.active === "true", custom_fields: JSON.stringify(vesselForm.customFields || {}) }); toast.success("Vessel added"); setShowAdd(false); refresh(); } catch (err) { toast.error(err instanceof Error ? err.message : "Save failed"); } }}>Add Vessel</Btn></>}>
        <div className="grid grid-cols-2 gap-4">
          <Input label="Vessel Code" value={vesselForm.code} onChange={v => setVesselForm(p => ({ ...p, code: v }))} required placeholder="" />
          <Input label="Vessel Name" value={vesselForm.name} onChange={v => setVesselForm(p => ({ ...p, name: v }))} required placeholder="" />
          <Select label="Principal" value={vesselForm.principalId} onChange={v => setVesselForm(p => ({ ...p, principalId: v }))} required options={principalOptions.map(p => ({ value: String(p.id), label: p.name }))} />
          <SelectWithOther label="Vessel Type" value={vesselForm.type} onChange={v => setVesselForm(p => ({ ...p, type: v }))} required options={vesselTypeOptions} />
          <Select label="Status" value={vesselForm.active} onChange={v => setVesselForm(p => ({ ...p, active: v }))} options={[{ value: "true", label: "Active" }, { value: "false", label: "Inactive" }]} />
          <DynamicCustomFields fields={vesselCustomFields} values={vesselForm.customFields} onChange={values => setVesselForm(p => ({ ...p, customFields: values }))} />
        </div>
      </Modal>

      <ImportModal open={showImport} onClose={() => { setShowImport(false); refresh(); }} entity="Vessels" endpoint="/api/admin/import/vessels" />
    </div>
  );
}

function PrincipalsPage({ setPage, setSelectedPrincipal }: { setPage: (p: string) => void; setSelectedPrincipal: (p: Principal) => void }) {
  const [search, setSearch] = useState("");
  const [showAdd, setShowAdd] = useState(false);
  const [showImport, setShowImport] = useState(false);
  const [page, setPageNum] = useState(1);
  const [principalForm, setPrincipalForm] = useState({
    code: "",
    name: "",
    contact: "",
    phone: "",
    email: "",
    address: "",
    website: "",
    customFields: {} as Record<string, unknown>,
  });
  const PAGE_SIZE = 8;

  const { data: principals, loading, refresh } = useApiList<ApiPrincipal, Principal>("/api/admin/principals", mapPrincipal, []);
  const { data: principalCustomFields } = useApiList<ApiCustomField, CustomField>("/api/admin/custom-fields?entity_type=principal", mapCustomField, []);

  const activePrincipalCustomFields = useMemo(() => principalCustomFields.filter(field => field.active), [principalCustomFields]);

  const filtered = useMemo(() =>
    principals.filter(p => search === "" || p.name.toLowerCase().includes(search.toLowerCase()) || p.code.toLowerCase().includes(search.toLowerCase())),
    [principals, search]
  );

  return (
    <div>
      <PageHeader title="Company / Principal">
        <Btn variant="secondary" size="sm" onClick={() => setShowImport(true)}><Upload size={13} />Import</Btn>
        <Btn variant="secondary" size="sm" onClick={() => exportExcel("/api/admin/export/seafarers", "id,name,photo_url,phone,nik,passport_no,seaman_book_no,status", "seafarers.xlsx").catch(err => toast.error(err.message))}><Download size={13} />Export</Btn>
        <Btn variant="primary" size="sm" onClick={() => setShowAdd(true)}><Plus size={13} />Add Principal</Btn>
      </PageHeader>
      <Card>
        <div className="flex items-center gap-3 px-4 py-3 border-b border-slate-100">
          <SearchBar value={search} onChange={v => { setSearch(v); setPageNum(1); }} placeholder="Search principals…" />
          <div className="ml-auto text-xs text-slate-400">{loading ? "Loading..." : `${filtered.length} principals`}</div>
        </div>
        {filtered.length > 0 ? (
          <>
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-slate-100">
                  {["Code", "Name", "Contact", "Phone", "Email", ...activePrincipalCustomFields.map(field => field.label), "Status", "Actions"].map(h => (
                    <th key={h} className="px-4 py-3 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-wide">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {filtered.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE).map(p => (
                  <tr key={p.id} className="hover:bg-slate-50 group">
                    <td className="px-4 py-3 font-mono text-xs text-slate-500">{p.code}</td>
                    <td className="px-4 py-3">
                      <button onClick={() => { setSelectedPrincipal(p); setPage("principal-detail"); }}
                        className="flex items-center gap-2 hover:underline text-left">
                        <div className="w-7 h-7 bg-slate-100 rounded-lg flex items-center justify-center shrink-0">
                          <Building2 size={12} className="text-slate-500" />
                        </div>
                        <span className="font-medium text-blue-600 text-xs">{p.name}</span>
                      </button>
                    </td>
                    <td className="px-4 py-3 text-xs text-slate-600">{p.contact}</td>
                    <td className="px-4 py-3 text-xs text-slate-600">{p.phone}</td>
                    <td className="px-4 py-3 text-xs text-slate-600">{p.email}</td>
                    {activePrincipalCustomFields.map(field => <td key={field.key} className="px-4 py-3 text-xs text-slate-600">{customValueToString(p.customFields[field.key]) || "-"}</td>)}
                    <td className="px-4 py-3"><Badge status={p.active ? "active" : "inactive"} /></td>
                    <td className="px-4 py-3">
                      <div className="flex gap-1">
                        <button onClick={() => { setSelectedPrincipal(p); setPage("principal-detail"); }} className="p-1.5 rounded hover:bg-blue-50 text-slate-400 hover:text-blue-600 transition-colors" title="View / Preview Principal"><Eye size={12} /></button>
                        <button className="p-1.5 rounded hover:bg-amber-50 text-slate-400 hover:text-amber-600 transition-colors" title="Edit Principal"><Pencil size={12} /></button>
                        <button className="p-1.5 rounded hover:bg-red-50 text-slate-400 hover:text-red-500 transition-colors" title="Delete Principal"><Trash2 size={12} /></button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <Pagination page={page} total={filtered.length} pageSize={PAGE_SIZE} onPage={setPageNum} />
          </>
        ) : !loading ? (
          <div className="py-12">
            <EmptyState title="Belum Ada Data Principal" description="Admin belum mengatur data principal. Silakan atur atau tambahkan data principal baru terlebih dahulu." action={<Btn variant="primary" size="sm" onClick={() => setShowAdd(true)}><Plus size={13} />Add Principal</Btn>} />
          </div>
        ) : null}
      </Card>

      <Modal open={showAdd} onClose={() => { setShowAdd(false); setPrincipalForm({ code: "", name: "", contact: "", phone: "", email: "", address: "", website: "", customFields: {} }); }} title="Add Principal" size="lg"
        footer={<><Btn variant="secondary" size="sm" onClick={() => { setShowAdd(false); setPrincipalForm({ code: "", name: "", contact: "", phone: "", email: "", address: "", website: "", customFields: {} }); }}>Cancel</Btn><Btn variant="primary" size="sm" onClick={async () => { try { await apiJson("/api/admin/principals", "POST", { code: principalForm.code, name: principalForm.name, contact_name: principalForm.contact, contact_phone: principalForm.phone, email: principalForm.email, address: principalForm.address, is_active: true, custom_fields: JSON.stringify(principalForm.customFields || {}) }); toast.success("Principal added"); setShowAdd(false); setPrincipalForm({ code: "", name: "", contact: "", phone: "", email: "", address: "", website: "", customFields: {} }); refresh(); } catch (err) { toast.error(err instanceof Error ? err.message : "Save failed"); } }}>Add Principal</Btn></>}>
        <div className="grid grid-cols-2 gap-4">
          <Input label="Principal Code" value={principalForm.code} onChange={v => setPrincipalForm(p => ({ ...p, code: v }))} required placeholder="" />
          <Input label="Company Name" value={principalForm.name} onChange={v => setPrincipalForm(p => ({ ...p, name: v }))} required placeholder="" />
          <DynamicCustomFields fields={principalCustomFields} values={principalForm.customFields} onChange={values => setPrincipalForm(p => ({ ...p, customFields: values }))} />
        </div>
      </Modal>

      <ImportModal open={showImport} onClose={() => { setShowImport(false); refresh(); }} entity="Principals" endpoint="/api/admin/import/principals" />
    </div>
  );
}

function PrincipalDetailPage({ principal, setPage }: { principal: Principal; setPage: (p: string) => void }) {
  const [tab, setTab] = useState("overview");
  const { data: allVessels } = useApiList<ApiVessel, Vessel>("/api/admin/vessels", mapVessel, []);
  const { data: allDeployments } = useApiList<ApiJoining, JoiningRecord>("/api/joining-principals", mapJoining, []);
  const { data: allRequirements, loading: reqLoading } = useApiList<ApiPrincipalRequirement, PrincipalRequirement>("/api/admin/principal-requirements", mapPrincipalRequirement, []);
  const { data: principalCustomFields } = useApiList<ApiCustomField, CustomField>("/api/admin/custom-fields?entity_type=principal", mapCustomField, []);

  const vessels = useMemo(() => allVessels.filter(v => v.principal === principal.name || (v as any).principal_id === principal.id), [allVessels, principal]);
  const deployments = useMemo(() => allDeployments.filter(j => j.principal === principal.name), [allDeployments, principal]);
  const requirements = useMemo(() => allRequirements.filter(r => r.principalId === principal.id), [allRequirements, principal]);

  return (
    <div>
      <div className="flex items-center gap-3 mb-5">
        <button onClick={() => setPage("principals")} className="p-1.5 rounded-lg hover:bg-white border border-slate-200 transition-colors">
          <ArrowLeft size={14} className="text-slate-600" />
        </button>
        <h1 className="text-lg font-semibold text-slate-800">Principal Detail</h1>
      </div>

      <Card className="p-5 mb-4">
        <div className="flex items-start gap-5">
          <div className="w-14 h-14 bg-slate-100 rounded-xl flex items-center justify-center shrink-0">
            <Building2 size={24} className="text-slate-500" />
          </div>
          <div className="flex-1">
            <div className="flex items-start justify-between">
              <div>
                <h2 className="text-base font-bold text-slate-800">{principal.name}</h2>
                <p className="text-sm text-slate-500 mt-0.5">{principal.code} · {principal.contact}</p>
                <Badge status={principal.active ? "active" : "inactive"} />
              </div>
              <Btn variant="secondary" size="sm"><Pencil size={12} />Edit</Btn>
            </div>
            <div className="grid grid-cols-3 gap-4 mt-4 pt-4 border-t border-slate-100">
              {[
                { icon: Phone, label: principal.phone },
                { icon: Mail, label: principal.email },
                { icon: Ship, label: `${vessels.length} vessel${vessels.length !== 1 ? "s" : ""}` },
              ].map(({ icon: Icon, label }, i) => (
                <div key={i} className="flex items-center gap-1.5 text-xs text-slate-500">
                  <Icon size={12} className="text-slate-400 shrink-0" />
                  <span>{label}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      </Card>

      <div className="flex gap-0.5 mb-4 bg-white border border-slate-100 rounded-xl p-1 w-fit">
        {["overview", "vessels", "requirements", "custom-fields"].map(t => (
          <button key={t} onClick={() => setTab(t)}
            className={cn("px-4 py-1.5 rounded-lg text-xs font-medium transition-colors capitalize",
              tab === t ? "bg-blue-600 text-white" : "text-slate-500 hover:text-slate-700 hover:bg-slate-50")}>
            {t.replace("-", " ")}
          </button>
        ))}
      </div>

      {tab === "overview" && (
        <div className="grid grid-cols-2 gap-4">
          <Card className="p-4">
            <h3 className="font-semibold text-slate-700 text-sm mb-3">Contact Information</h3>
            <div className="space-y-2.5">
              {[
                ["Company Name", principal.name], ["Contact Person", principal.contact],
                ["Phone", principal.phone], ["Email", principal.email], ["Code", principal.code],
              ].map(([k, v]) => (
                <div key={k} className="flex gap-2">
                  <span className="text-xs text-slate-400 w-32 shrink-0">{k}</span>
                  <span className="text-xs text-slate-700 font-medium">{v}</span>
                </div>
              ))}
            </div>
          </Card>
          <Card className="p-4">
            <h3 className="font-semibold text-slate-700 text-sm mb-3">Deployment Summary</h3>
            <div className="space-y-2">
              {[
                ["Active Deployments", deployments.filter(d => d.status === "onboard").length],
                ["Total Joinings", deployments.length],
                ["Vessels Assigned", vessels.length],
                ["Completed", deployments.filter(d => d.status === "completed").length],
              ].map(([k, v]) => (
                <div key={String(k)} className="flex items-center justify-between">
                  <span className="text-xs text-slate-500">{k}</span>
                  <span className="text-xs font-bold text-slate-800">{v}</span>
                </div>
              ))}
            </div>
          </Card>
        </div>
      )}

      {tab === "vessels" && (
        <Card>
          <div className="px-4 py-3 border-b border-slate-100 flex items-center justify-between">
            <h3 className="font-semibold text-slate-700 text-sm">Assigned Vessels</h3>
            <Btn variant="primary" size="sm"><Plus size={12} />Assign Vessel</Btn>
          </div>
          {vessels.length === 0
            ? <EmptyState title="No vessels assigned" description="No vessels are assigned to this principal." />
            : (
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b border-slate-100">
                    {["Code", "Name", "Type", "Flag", "IMO", "Status"].map(h => (
                      <th key={h} className="px-4 py-3 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-wide">{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  {vessels.map(v => (
                    <tr key={v.id} className="hover:bg-slate-50">
                      <td className="px-4 py-3 font-mono text-xs text-slate-500">{v.code}</td>
                      <td className="px-4 py-3 font-medium text-slate-800 text-xs">{v.name}</td>
                      <td className="px-4 py-3 text-xs text-slate-600">{v.type}</td>
                      <td className="px-4 py-3 text-xs text-slate-600">{v.flag}</td>
                      <td className="px-4 py-3 font-mono text-xs text-slate-500">{v.imo}</td>
                      <td className="px-4 py-3"><Badge status={v.active ? "active" : "inactive"} /></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )
          }
        </Card>
      )}

      {tab === "requirements" && (
        <Card className="p-4">
          <div className="flex items-center justify-between mb-4">
            <h3 className="font-semibold text-slate-700 text-sm">Document Requirements</h3>
            <Btn variant="primary" size="sm"><Plus size={12} />Add Requirement</Btn>
          </div>
          {reqLoading ? (
            <p className="text-xs text-slate-400 py-4 text-center">Loading requirements...</p>
          ) : requirements.length === 0 ? (
            <EmptyState title="No requirements set" description="Belum ada persyaratan dokumen yang diatur untuk principal ini." />
          ) : (
            <div className="space-y-2">
              {requirements.map((req) => (
                <div key={req.id} className="flex items-center justify-between p-3 bg-slate-50 rounded-xl border border-slate-100">
                  <div className="flex items-center gap-2">
                    <CheckCircle size={13} className="text-emerald-500" />
                    <span className="text-xs font-medium text-slate-700">{req.documentName}</span>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className={cn("text-[10px] px-1.5 py-0.5 rounded font-medium", req.mandatory ? "bg-emerald-50 text-emerald-600" : "bg-slate-100 text-slate-600")}>
                      {req.mandatory ? "Required" : "Optional"}
                    </span>
                  </div>
                </div>
              ))}
            </div>
          )}
        </Card>
      )}

      {tab === "custom-fields" && (
        <Card className="p-4">
          <h3 className="font-semibold text-slate-700 text-sm mb-3">Custom Fields</h3>
          {principalCustomFields.filter(f => f.active).length === 0 ? (
            <EmptyState title="No custom fields configured" description="Belum ada field tambahan yang diatur di Settings -> Masters -> Custom Fields." />
          ) : (
            <div className="space-y-2.5">
              {principalCustomFields.filter(f => f.active).map(f => (
                <div key={f.id} className="flex gap-2 py-1 border-b border-slate-50">
                  <span className="text-xs text-slate-400 w-36 shrink-0">{f.label}</span>
                  <span className="text-xs text-slate-700 font-medium">{customValueToString(principal.customFields?.[f.key]) || "-"}</span>
                </div>
              ))}
            </div>
          )}
        </Card>
      )}
    </div>
  );
}

// ─── Reports ──────────────────────────────────────────────────────────────────

function CrewReportsPage() {
  const [report, setReport] = useState<ApiCrewReport | null>(null);
  const [joining, setJoining] = useState<ApiJoiningReport | null>(null);
  const [loading, setLoading] = useState(false);
  const [ran, setRan] = useState(false);

  async function generate() {
    setLoading(true);
    try {
      const [crewReport, joiningReport] = await Promise.all([apiGet<ApiCrewReport>("/api/reports/crew"), apiGet<ApiJoiningReport>("/api/reports/joining")]);
      setReport(crewReport);
      setJoining(joiningReport);
      setRan(true);
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Failed to generate report");
    } finally {
      setLoading(false);
    }
  }

  const crewByStatus = Array.isArray(report?.by_status) ? report.by_status : [];
  const ageDistribution = Array.isArray(report?.age_distribution) ? report.age_distribution : [];
  const maritalBreakdown = Array.isArray(report?.by_marital) ? report.by_marital : [];
  const joiningByStatus = Array.isArray(joining?.by_status) ? joining.by_status : [];

  return (
    <div>
      <PageHeader title="Crew Reports">
        {ran && <Btn variant="secondary" size="sm" onClick={() => exportExcel("/api/reports/export?type=crew", undefined, "crew-report.xlsx").catch(err => toast.error(err.message))}><Download size={13} />Export Excel</Btn>}
      </PageHeader>
      <Card className="p-5 mb-4">
        <h3 className="font-semibold text-slate-700 text-sm mb-4">Report Generator</h3>
        <Btn variant="primary" onClick={generate} disabled={loading}>{loading ? <><RefreshCw size={13} className="animate-spin" />Generating...</> : <><BarChart3 size={13} />Generate Report</>}</Btn>
      </Card>
      {ran && report && (
        <div className="grid grid-cols-2 gap-4">
          <Card className="p-4"><h3 className="font-semibold text-slate-800 text-sm mb-3">Crew by Status</h3>{crewByStatus.map(r => <div key={r.status} className="flex items-center justify-between py-2 border-b border-slate-50"><Badge status={r.status} /><span className="text-sm font-bold text-slate-800">{r.count}</span></div>)}</Card>
          <Card className="p-4"><h3 className="font-semibold text-slate-800 text-sm mb-3">Age Distribution</h3>{ageDistribution.map(r => <div key={r.range} className="flex items-center justify-between py-2 border-b border-slate-50"><span className="text-xs text-slate-500">{r.range}</span><span className="text-sm font-bold text-slate-800">{r.count}</span></div>)}</Card>
          <Card className="p-4"><h3 className="font-semibold text-slate-800 text-sm mb-3">Marital Status</h3>{maritalBreakdown.map(r => <div key={r.marital_status} className="flex items-center justify-between py-2 border-b border-slate-50"><span className="text-xs text-slate-500">{r.marital_status || "-"}</span><span className="text-sm font-bold text-slate-800">{r.count}</span></div>)}</Card>
          <Card className="p-4"><h3 className="font-semibold text-slate-800 text-sm mb-3">Joining Summary</h3><p className="text-3xl font-bold text-slate-800 mb-3">{joining?.total ?? 0}</p>{joiningByStatus.map(r => <div key={r.status} className="flex items-center justify-between py-2 border-b border-slate-50"><Badge status={r.status} /><span className="text-sm font-bold text-slate-800">{r.count}</span></div>)}</Card>
        </div>
      )}
    </div>
  );
}

function DocumentReportsPage() {
  const [report, setReport] = useState<ApiDocumentReport | null>(null);
  const [loading, setLoading] = useState(false);
  async function generate() {
    setLoading(true);
    try {
      setReport(await apiGet<ApiDocumentReport>("/api/reports/documents"));
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Failed to generate report");
    } finally {
      setLoading(false);
    }
  }
  return (
    <div>
      <PageHeader title="Document Reports">
        {report && <Btn variant="secondary" size="sm" onClick={() => exportExcel("/api/reports/export?type=documents", undefined, "document-report.xlsx").catch(err => toast.error(err.message))}><Download size={13} />Export Excel</Btn>}
      </PageHeader>
      <Card className="p-5 mb-4"><h3 className="font-semibold text-slate-700 text-sm mb-4">Report Generator</h3><Btn variant="primary" onClick={generate} disabled={loading}>{loading ? <><RefreshCw size={13} className="animate-spin" />Generating...</> : <><BarChart3 size={13} />Generate</>}</Btn></Card>
      {report && <div className="grid grid-cols-4 gap-3 mb-4">{[{ label: "Total", value: report.total_documents, color: "text-blue-600 bg-blue-50" }, { label: "Expired", value: report.expired_count, color: "text-red-600 bg-red-50" }, { label: "Expiring", value: report.expiring_count, color: "text-amber-600 bg-amber-50" }, { label: "Valid", value: report.valid_count, color: "text-emerald-600 bg-emerald-50" }].map(x => <div key={x.label} className={cn("rounded-xl p-3 text-center", x.color)}><p className="text-xl font-bold">{x.value}</p><p className="text-xs mt-0.5 opacity-80">{x.label}</p></div>)}</div>}
      {report && <Card><table className="w-full text-sm"><thead><tr className="border-b border-slate-100">{["Type", "Total", "Expired", "Expiring", "Valid"].map(h => <th key={h} className="px-4 py-3 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-wide">{h}</th>)}</tr></thead><tbody className="divide-y divide-slate-50">{(report.type_breakdown ?? []).map(r => <tr key={r.document_type_name}><td className="px-4 py-3 text-xs font-medium text-slate-800">{r.document_type_name}</td><td className="px-4 py-3 text-xs">{r.total}</td><td className="px-4 py-3 text-xs text-red-600">{r.expired}</td><td className="px-4 py-3 text-xs text-amber-600">{r.expiring}</td><td className="px-4 py-3 text-xs text-emerald-600">{r.valid ?? Math.max(0, r.total - r.expired - r.expiring)}</td></tr>)}</tbody></table></Card>}
    </div>
  );
}
// ─── Settings ─────────────────────────────────────────────────────────────────

function PrincipalRequirementsManager() {
  const [showAdd, setShowAdd] = useState(false);
  const [checkForm, setCheckForm] = useState({ seafarerId: "", principalId: "", vesselId: "" });
  const [checkResult, setCheckResult] = useState<ApiRequirementCheck | null>(null);
  const [checking, setChecking] = useState(false);
  const [form, setForm] = useState({ principalId: "", vesselId: "", type: "document", documentNameId: "", mandatory: "true", requiresValid: "true", active: "true" });
  const { data, loading, refresh } = useApiList<ApiPrincipalRequirement, PrincipalRequirement>("/api/admin/principal-requirements", mapPrincipalRequirement, []);
  const { data: principals } = useApiList<ApiPrincipal, Principal>("/api/admin/principals", mapPrincipal, []);
  const { data: vessels } = useApiList<ApiVessel, Vessel>("/api/admin/vessels", mapVessel, []);
  const { data: documentNames } = useApiList<ApiDocumentName, ReturnType<typeof mapDocName>>("/api/admin/document-names", mapDocName, []);

  function resetForm() {
    setForm({ principalId: "", vesselId: "", type: "document", documentNameId: "", mandatory: "true", requiresValid: "true", active: "true" });
  }

  async function addRequirement() {
    try {
      if (!form.principalId) {
        toast.error("Principal wajib dipilih");
        return;
      }
      if (form.type === "document" && !form.documentNameId) {
        toast.error("Document name wajib dipilih");
        return;
      }
      await apiJson("/api/admin/principal-requirements", "POST", {
        principal_id: Number(form.principalId),
        vessel_id: form.vesselId ? Number(form.vesselId) : null,
        requirement_type: form.type,
        document_name_id: form.type === "document" ? Number(form.documentNameId) : null,
        is_mandatory: form.mandatory === "true",
        requires_valid_document: form.requiresValid === "true",
        is_active: form.active === "true",
      });
      toast.success("Requirement added");
      setShowAdd(false);
      resetForm();
      refresh();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Save failed");
    }
  }

  async function toggleRequirement(rule: PrincipalRequirement) {
    try {
      await apiJson(`/api/admin/principal-requirements/${rule.id}`, "PUT", {
        principal_id: rule.principalId,
        vessel_id: rule.vesselId,
        requirement_type: rule.type,
        document_name_id: rule.documentNameId,
        is_mandatory: rule.mandatory,
        requires_valid_document: rule.requiresValid,
        is_active: !rule.active,
      });
      refresh();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Update failed");
    }
  }

  async function deleteRequirement(id: number) {
    if (!confirm("Delete this requirement?")) return;
    try {
      await apiJson(`/api/admin/principal-requirements/${id}`, "DELETE");
      toast.success("Requirement deleted");
      refresh();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Delete failed");
    }
  }

  async function runComplianceCheck() {
    try {
      if (!checkForm.seafarerId || !checkForm.principalId) {
        toast.error("Seafarer ID dan Principal wajib diisi");
        return;
      }
      setChecking(true);
      const params: Record<string, string | number | undefined> = {
        seafarer_id: checkForm.seafarerId,
        principal_id: checkForm.principalId,
        vessel_id: checkForm.vesselId || undefined,
      };
      setCheckResult(await apiGet<ApiRequirementCheck>("/api/principal-requirements/check", params));
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Check failed");
    } finally {
      setChecking(false);
    }
  }

  const checkRows = Array.isArray(checkResult?.requirements) ? checkResult.requirements : [];

  return (
    <div className="grid grid-cols-[1fr_360px] gap-4">
      <Card className="p-5">
        <div className="flex items-center justify-between mb-4">
          <div><h3 className="font-semibold text-slate-800 text-sm">Principal Requirements</h3><p className="text-xs text-slate-400 mt-1">Aturan dokumen dan kualifikasi wajib per Principal atau Vessel</p></div>
          <Btn variant="primary" size="sm" onClick={() => setShowAdd(true)}><Plus size={13} />Add Requirement</Btn>
        </div>
        <div className="overflow-x-auto rounded-xl border border-slate-100">
          <table className="w-full text-sm"><thead><tr className="border-b border-slate-100 bg-slate-50">{["Principal", "Vessel", "Type", "Document", "Mandatory", "Valid Doc", "Status", "Actions"].map(h => <th key={h} className="px-4 py-3 text-left text-[11px] font-semibold uppercase tracking-wide text-slate-400">{h}</th>)}</tr></thead><tbody className="divide-y divide-slate-50">{data.map(rule => <tr key={rule.id} className="hover:bg-slate-50"><td className="px-4 py-3 text-xs font-medium text-slate-800">{rule.principalName}</td><td className="px-4 py-3 text-xs text-slate-600">{rule.vesselName}</td><td className="px-4 py-3 text-xs text-slate-600">{rule.type}</td><td className="px-4 py-3 text-xs text-slate-600">{rule.documentName}</td><td className="px-4 py-3 text-xs text-slate-600">{rule.mandatory ? "Yes" : "No"}</td><td className="px-4 py-3 text-xs text-slate-600">{rule.requiresValid ? "Yes" : "No"}</td><td className="px-4 py-3"><Badge status={rule.active ? "active" : "inactive"} /></td><td className="px-4 py-3"><div className="flex gap-1"><button onClick={() => toggleRequirement(rule)} className="px-2 py-1 text-[10px] rounded bg-slate-50 text-slate-600">{rule.active ? "Disable" : "Enable"}</button><button onClick={() => deleteRequirement(rule.id)} className="p-1.5 rounded hover:bg-red-50 text-slate-400 hover:text-red-500"><Trash2 size={12} /></button></div></td></tr>)}</tbody></table>
          {!loading && data.length === 0 && <EmptyState title="No requirements" description="Belum ada aturan Principal/Vessel." />}
        </div>
      </Card>
      <Card className="p-5">
        <h3 className="font-semibold text-slate-800 text-sm mb-4">Compliance Check</h3>
        <div className="space-y-3">
          <Input label="Seafarer ID" value={checkForm.seafarerId} onChange={v => setCheckForm(p => ({ ...p, seafarerId: v }))} type="number" placeholder="1" />
          <Select label="Principal" value={checkForm.principalId} onChange={v => setCheckForm(p => ({ ...p, principalId: v }))} options={principals.map(p => ({ value: String(p.id), label: p.name }))} />
          <Select label="Vessel" value={checkForm.vesselId} onChange={v => setCheckForm(p => ({ ...p, vesselId: v }))} options={vessels.map(v => ({ value: String(v.id), label: v.name }))} placeholder="All vessels" />
          <Btn variant="primary" className="w-full justify-center" onClick={runComplianceCheck} disabled={checking}>{checking ? <><RefreshCw size={13} className="animate-spin" />Checking...</> : <><Shield size={13} />Check</>}</Btn>
        </div>
        {checkResult && <div className="mt-4 rounded-xl border border-slate-100 bg-slate-50 p-3"><div className="flex items-center justify-between"><span className="text-xs font-semibold text-slate-700">Result</span><Badge status={checkResult.passed ? "active" : "failed"} label={checkResult.passed ? "Passed" : "Failed"} /></div><div className="mt-3 space-y-2">{checkRows.map((row, index) => <div key={index} className="rounded-lg bg-white p-2 text-xs"><p className="font-medium text-slate-700">{row.name || row.document_name || `Requirement ${index + 1}`}</p><p className="mt-0.5 text-slate-400">{row.reason || row.status || (row.passed ? "Passed" : "Failed")}</p></div>)}</div>{checkRows.length === 0 && <p className="mt-2 text-xs text-slate-400">No detailed rows returned.</p>}</div>}
      </Card>
      <Modal open={showAdd} onClose={() => { setShowAdd(false); resetForm(); }} title="Add Principal Requirement" size="lg" footer={<><Btn variant="secondary" size="sm" onClick={() => { setShowAdd(false); resetForm(); }}>Cancel</Btn><Btn variant="primary" size="sm" onClick={addRequirement}>Save Requirement</Btn></>}>
        <div className="grid grid-cols-2 gap-4">
          <Select label="Principal" value={form.principalId} onChange={v => setForm(p => ({ ...p, principalId: v, vesselId: "" }))} required options={principals.map(p => ({ value: String(p.id), label: p.name }))} />
          <Select label="Vessel" value={form.vesselId} onChange={v => setForm(p => ({ ...p, vesselId: v }))} placeholder="All vessels" options={vessels.map(v => ({ value: String(v.id), label: v.name }))} />
          <Select label="Requirement Type" value={form.type} onChange={v => setForm(p => ({ ...p, type: v }))} options={[{ value: "document", label: "Document" }, { value: "custom", label: "Custom" }]} />
          <Select label="Document Name" value={form.documentNameId} onChange={v => setForm(p => ({ ...p, documentNameId: v }))} required={form.type === "document"} options={documentNames.map(d => ({ value: String(d.id), label: d.name }))} />
          <Select label="Mandatory" value={form.mandatory} onChange={v => setForm(p => ({ ...p, mandatory: v }))} options={[{ value: "true", label: "Mandatory" }, { value: "false", label: "Optional" }]} />
          <Select label="Requires Valid Document" value={form.requiresValid} onChange={v => setForm(p => ({ ...p, requiresValid: v }))} options={[{ value: "true", label: "Must be valid" }, { value: "false", label: "Allow expired/not checked" }]} />
          <Select label="Status" value={form.active} onChange={v => setForm(p => ({ ...p, active: v }))} options={[{ value: "true", label: "Active" }, { value: "false", label: "Inactive" }]} />
        </div>
      </Modal>
    </div>
  );
}
function CustomFieldsManager() {
  const [entityType, setEntityType] = useState<"principal" | "vessel">("principal");
  const [showAdd, setShowAdd] = useState(false);
  const [form, setForm] = useState({ fieldKey: "", label: "", type: "text", optionsJson: "{}", required: "false", active: "true" });
  const { data, loading, refresh } = useApiList<ApiCustomField, CustomField>(`/api/admin/custom-fields?entity_type=${entityType}`, mapCustomField, []);

  function resetForm() {
    setForm({ fieldKey: "", label: "", type: "text", optionsJson: "{}", required: "false", active: "true" });
  }

  async function addField() {
    try {
      if (!form.fieldKey.trim() || !form.label.trim()) {
        toast.error("Field key dan label wajib diisi");
        return;
      }
      JSON.parse(form.optionsJson || "{}");
      await apiJson("/api/admin/custom-fields", "POST", {
        entity_type: entityType,
        field_key: form.fieldKey.trim(),
        label: form.label.trim(),
        field_type: form.type,
        options_json: form.optionsJson || "{}",
        is_required: form.required === "true",
        is_active: form.active === "true",
      });
      toast.success("Custom field added");
      setShowAdd(false);
      resetForm();
      refresh();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Save failed");
    }
  }

  async function toggleField(field: CustomField) {
    try {
      await apiJson(`/api/admin/custom-fields/${field.id}`, "PUT", {
        entity_type: field.entityType,
        field_key: field.key,
        label: field.label,
        field_type: field.type,
        options_json: field.optionsJson,
        is_required: field.required,
        is_active: !field.active,
      });
      refresh();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Update failed");
    }
  }

  async function deleteField(id: number) {
    if (!confirm("Delete this custom field definition? Existing JSON values on Principal/Vessel will not be removed automatically.")) return;
    try {
      await apiJson(`/api/admin/custom-fields/${id}`, "DELETE");
      toast.success("Custom field deleted");
      refresh();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Delete failed");
    }
  }

  return (
    <Card className="p-5">
      <div className="flex items-center justify-between mb-4">
        <div>
          <h3 className="font-semibold text-slate-800 text-sm">Custom Fields</h3>
          <p className="text-xs text-slate-400 mt-1">Definisi field tambahan untuk Principal dan Vessel</p>
        </div>
        <Btn variant="primary" size="sm" onClick={() => setShowAdd(true)}><Plus size={13} />Add Field</Btn>
      </div>
      <div className="mb-4 flex w-fit rounded-lg border border-slate-200 bg-slate-50 p-1">
        {(["principal", "vessel"] as const).map(type => (
          <button key={type} onClick={() => setEntityType(type)} className={cn("rounded-md px-3 py-1.5 text-xs font-semibold capitalize", entityType === type ? "bg-blue-600 text-white" : "text-slate-500 hover:text-slate-700")}>{type}</button>
        ))}
      </div>

      <div className="mb-4 rounded-xl border border-slate-200 bg-slate-50/80 p-3.5">
        <div className="flex items-center justify-between mb-2">
          <div className="flex items-center gap-2">
            <CheckCircle size={13} className="text-emerald-600" />
            <span className="text-xs font-semibold text-slate-700">Standard Built-in Fields ({entityType === "principal" ? "Company / Principal" : "Vessel"})</span>
          </div>
          <span className="text-[10px] font-medium text-slate-500 bg-white px-2 py-0.5 rounded border border-slate-200">Bawaan Sistem (Selalu Ada)</span>
        </div>
        <p className="text-[11px] text-slate-500 mb-2">Field bawaan standar ini sudah secara otomatis tersedia di modal &amp; tabel {entityType}:</p>
        <div className="flex flex-wrap gap-2">
          {(entityType === "principal"
            ? [
                { key: "code", label: "Principal Code *" },
                { key: "name", label: "Company Name *" },
              ]
            : [
                { key: "code", label: "Vessel Code *" },
                { key: "name", label: "Vessel Name *" },
                { key: "principal_id", label: "Principal *" },
                { key: "vessel_type", label: "Vessel Type *" },
                { key: "is_active", label: "Status" },
              ]
          ).map(f => (
            <div key={f.key} className="flex items-center gap-1.5 bg-white border border-slate-200 px-2.5 py-1 rounded-lg text-xs font-medium text-slate-700 shadow-xs">
              <span className="font-mono text-[10px] text-blue-600 font-semibold">{f.key}</span>
              <span className="text-slate-300">·</span>
              <span>{f.label}</span>
            </div>
          ))}
        </div>
      </div>

      <div className="overflow-x-auto rounded-xl border border-slate-100">
        <table className="w-full text-sm">
          <thead><tr className="border-b border-slate-100 bg-slate-50">{["Field Key", "Label", "Type", "Required", "Status", "Actions"].map(h => <th key={h} className="px-4 py-3 text-left text-[11px] font-semibold uppercase tracking-wide text-slate-400">{h}</th>)}</tr></thead>
          <tbody className="divide-y divide-slate-50">
            {data.map(field => <tr key={field.id} className="hover:bg-slate-50"><td className="px-4 py-3 font-mono text-xs text-slate-600">{field.key}</td><td className="px-4 py-3 text-xs font-medium text-slate-800">{field.label}</td><td className="px-4 py-3 text-xs text-slate-600">{field.type}</td><td className="px-4 py-3 text-xs text-slate-600">{field.required ? "Yes" : "No"}</td><td className="px-4 py-3"><Badge status={field.active ? "active" : "inactive"} /></td><td className="px-4 py-3"><div className="flex gap-1"><button onClick={() => toggleField(field)} className="px-2 py-1 text-[10px] rounded bg-slate-50 text-slate-600">{field.active ? "Disable" : "Enable"}</button><button onClick={() => deleteField(field.id)} className="p-1.5 rounded hover:bg-red-50 text-slate-400 hover:text-red-500"><Trash2 size={12} /></button></div></td></tr>)}
          </tbody>
        </table>
        {!loading && data.length === 0 && <EmptyState title="No custom fields" description={`Belum ada field tambahan untuk ${entityType}.`} />}
      </div>
      <Modal open={showAdd} onClose={() => { setShowAdd(false); resetForm(); }} title="Add Custom Field" size="lg" footer={<><Btn variant="secondary" size="sm" onClick={() => { setShowAdd(false); resetForm(); }}>Cancel</Btn><Btn variant="primary" size="sm" onClick={addField}>Save Field</Btn></>}>
        <div className="grid grid-cols-2 gap-4">
          <Select label="Entity" value={entityType} onChange={v => setEntityType(v as "principal" | "vessel")} options={[{ value: "principal", label: "Principal" }, { value: "vessel", label: "Vessel" }]} required />
          <Input label="Field Key" value={form.fieldKey} onChange={v => setForm(p => ({ ...p, fieldKey: v }))} placeholder="tax_id" required />
          <Input label="Label" value={form.label} onChange={v => setForm(p => ({ ...p, label: v }))} placeholder="NPWP / Tax ID" required />
          <Select label="Field Type" value={form.type} onChange={v => setForm(p => ({ ...p, type: v }))} options={["text", "number", "date", "select", "textarea", "boolean"].map(type => ({ value: type, label: type }))} />
          <Select label="Required" value={form.required} onChange={v => setForm(p => ({ ...p, required: v }))} options={[{ value: "false", label: "No" }, { value: "true", label: "Yes" }]} />
          <Select label="Status" value={form.active} onChange={v => setForm(p => ({ ...p, active: v }))} options={[{ value: "true", label: "Active" }, { value: "false", label: "Inactive" }]} />
          <Textarea label="Options JSON" value={form.optionsJson} onChange={v => setForm(p => ({ ...p, optionsJson: v }))} className="col-span-2" placeholder='{"options":["A","B"]}' />
        </div>
      </Modal>
    </Card>
  );
}
function UsersPage() {
  const [search, setSearch] = useState("");
  const [page, setPageNum] = useState(1);
  const [showAdd, setShowAdd] = useState(false);
  const [showEdit, setShowEdit] = useState(false);
  const [editTarget, setEditTarget] = useState<ApiUser | null>(null);
  const [form, setForm] = useState({ name: "", email: "", password: "", role: "user" });
  const [loading2, setLoading2] = useState(false);
  const PAGE_SIZE = 10;

  const { data: users, loading, refresh } = useApiList<ApiUser, ApiUser>("/api/admin/users", u => u, []);

  const filtered = useMemo(() =>
    users.filter(u =>
      search === "" ||
      u.name.toLowerCase().includes(search.toLowerCase()) ||
      u.email.toLowerCase().includes(search.toLowerCase()) ||
      u.role.toLowerCase().includes(search.toLowerCase())
    ),
    [users, search]
  );

  function resetForm() { setForm({ name: "", email: "", password: "", role: "user" }); }

  async function createUser() {
    if (!form.name || !form.email || !form.password) { toast.error("Name, email and password are required"); return; }
    setLoading2(true);
    try {
      await apiPost("/api/admin/users", { name: form.name, email: form.email, password: form.password, role: form.role });
      toast.success("User created");
      setShowAdd(false); resetForm(); refresh();
    } catch (err) { toast.error(err instanceof Error ? err.message : "Failed to create user"); }
    finally { setLoading2(false); }
  }

  async function updateUser() {
    if (!editTarget) return;
    setLoading2(true);
    try {
      const body: Record<string, unknown> = { name: form.name, email: form.email, role: form.role };
      if (form.password) body.password = form.password;
      await apiPut(`/api/admin/users/${editTarget.id}`, body);
      toast.success("User updated");
      setShowEdit(false); setEditTarget(null); resetForm(); refresh();
    } catch (err) { toast.error(err instanceof Error ? err.message : "Failed to update user"); }
    finally { setLoading2(false); }
  }

  async function deleteUser(u: ApiUser) {
    if (!confirm(`Delete user "${u.name}"?`)) return;
    try {
      await apiDelete(`/api/admin/users/${u.id}`);
      toast.success("User deleted");
      refresh();
    } catch (err) { toast.error(err instanceof Error ? err.message : "Failed to delete user"); }
  }

  function openEdit(u: ApiUser) {
    setEditTarget(u);
    setForm({ name: u.name, email: u.email, password: "", role: u.role });
    setShowEdit(true);
  }

  const roleColors: Record<string, string> = {
    admin: "bg-violet-100 text-violet-700",
    superadmin: "bg-red-100 text-red-700",
    user: "bg-blue-100 text-blue-700",
  };

  return (
    <div>
      <PageHeader title="Users">
        <Btn variant="primary" size="sm" onClick={() => { resetForm(); setShowAdd(true); }}><Plus size={13} />Add User</Btn>
      </PageHeader>
      <Card>
        <div className="flex items-center gap-3 px-4 py-3 border-b border-slate-100">
          <SearchBar value={search} onChange={v => { setSearch(v); setPageNum(1); }} placeholder="Search users…" />
          <div className="ml-auto text-xs text-slate-400">{loading ? "Loading..." : `${filtered.length} users`}</div>
        </div>
        {filtered.length > 0 ? (
          <>
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-slate-100">
                  {["#", "Name", "Email", "Role", "Email Verified", "Joined", "Actions"].map(h => (
                    <th key={h} className="px-4 py-3 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-wide whitespace-nowrap">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {filtered.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE).map(u => (
                  <tr key={u.id} className="hover:bg-slate-50 group">
                    <td className="px-4 py-3 text-xs text-slate-400 font-mono">{u.id}</td>
                    <td className="px-4 py-3">
                      <div className="flex items-center gap-2">
                        <div className="w-7 h-7 rounded-full bg-gradient-to-br from-blue-500 to-violet-500 flex items-center justify-center shrink-0">
                          <span className="text-white text-[10px] font-bold">{u.name.charAt(0).toUpperCase()}</span>
                        </div>
                        <span className="font-medium text-slate-800">{u.name}</span>
                      </div>
                    </td>
                    <td className="px-4 py-3 text-slate-600">{u.email}</td>
                    <td className="px-4 py-3">
                      <span className={cn("px-2 py-0.5 rounded-full text-[11px] font-semibold capitalize", roleColors[u.role] ?? "bg-slate-100 text-slate-600")}>{u.role}</span>
                    </td>
                    <td className="px-4 py-3">
                      {u.is_email_verified
                        ? <span className="flex items-center gap-1 text-emerald-600 text-xs"><CheckCircle size={12} />Verified</span>
                        : <span className="flex items-center gap-1 text-amber-500 text-xs"><AlertCircle size={12} />Pending</span>}
                    </td>
                    <td className="px-4 py-3 text-xs text-slate-400">{u.created_at ? new Date(u.created_at).toLocaleDateString() : "—"}</td>
                    <td className="px-4 py-3">
                      <div className="flex items-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                        <button onClick={() => openEdit(u)} className="p-1.5 rounded-lg hover:bg-blue-50 text-blue-600 transition-colors"><Pencil size={13} /></button>
                        <button onClick={() => deleteUser(u)} className="p-1.5 rounded-lg hover:bg-red-50 text-red-500 transition-colors"><Trash2 size={13} /></button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <Pagination page={page} totalPages={Math.ceil(filtered.length / PAGE_SIZE)} onPageChange={setPageNum} />
          </>
        ) : (
          <EmptyState title="No users found" description={search ? `No users match "${search}"` : "No users registered yet."} />
        )}
      </Card>

      {/* Add User Modal */}
      <Modal open={showAdd} onClose={() => { setShowAdd(false); resetForm(); }} title="Add User" size="md"
        footer={<><Btn variant="secondary" size="sm" onClick={() => { setShowAdd(false); resetForm(); }}>Cancel</Btn><Btn variant="primary" size="sm" onClick={createUser} disabled={loading2}>{loading2 ? "Saving..." : "Create User"}</Btn></>}>
        <div className="grid grid-cols-1 gap-4">
          <Input label="Name" value={form.name} onChange={v => setForm(p => ({ ...p, name: v }))} required />
          <Input label="Email" value={form.email} onChange={v => setForm(p => ({ ...p, email: v }))} type="email" required />
          <Input label="Password" value={form.password} onChange={v => setForm(p => ({ ...p, password: v }))} type="password" required />
          <Select label="Role" value={form.role} onChange={v => setForm(p => ({ ...p, role: v }))} options={[{ value: "user", label: "User" }, { value: "admin", label: "Admin" }, { value: "superadmin", label: "Super Admin" }]} />
        </div>
      </Modal>

      {/* Edit User Modal */}
      <Modal open={showEdit} onClose={() => { setShowEdit(false); setEditTarget(null); resetForm(); }} title="Edit User" size="md"
        footer={<><Btn variant="secondary" size="sm" onClick={() => { setShowEdit(false); setEditTarget(null); resetForm(); }}>Cancel</Btn><Btn variant="primary" size="sm" onClick={updateUser} disabled={loading2}>{loading2 ? "Saving..." : "Save Changes"}</Btn></>}>
        <div className="grid grid-cols-1 gap-4">
          <Input label="Name" value={form.name} onChange={v => setForm(p => ({ ...p, name: v }))} required />
          <Input label="Email" value={form.email} onChange={v => setForm(p => ({ ...p, email: v }))} type="email" required />
          <Input label="New Password" value={form.password} onChange={v => setForm(p => ({ ...p, password: v }))} type="password" placeholder="Leave blank to keep current" />
          <Select label="Role" value={form.role} onChange={v => setForm(p => ({ ...p, role: v }))} options={[{ value: "user", label: "User" }, { value: "admin", label: "Admin" }, { value: "superadmin", label: "Super Admin" }]} />
        </div>
      </Modal>
    </div>
  );
}
function SettingsPage() {
  const [tab, setTab] = useState("account");
  const [profile, setProfile] = useState({ name: "", email: "" });
  const [password, setPassword] = useState({ current: "", next: "", confirm: "" });
  const [appSettings, setAppSettings] = useState({ company_name: "SKYagen", expiry_warning_days: "30" });
  const [storage, setStorage] = useState<ApiStorageSettings | null>(null);
  const [storagePath, setStoragePath] = useState("");
  const [storageLoading, setStorageLoading] = useState(false);
  const [uploadingDoc, setUploadingDoc] = useState(false);
  const [uploadedDoc, setUploadedDoc] = useState<ApiUploadedDocument | null>(null);

  async function loadStorage() {
    try {
      const res = await apiGet<ApiStorageSettings>("/api/admin/settings/storage");
      setStorage(res);
      setStoragePath(res.storage_path || "");
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Failed to load storage settings");
    }
  }

  useEffect(() => {
    apiGet<{ data: ApiProfile }>("/api/settings/profile").then(res => setProfile({ name: res.data.name, email: res.data.email })).catch(err => toast.error(err.message));
    apiGet<{ data: Record<string, { key: string; value: string }[]> }>("/api/admin/settings/app").then(res => {
      const flat = Object.values(res.data ?? {}).flat();
      setAppSettings(p => ({ ...p, ...Object.fromEntries(flat.map(x => [x.key, x.value])) }));
    }).catch(() => {});
    loadStorage();
  }, []);

  async function saveProfile() {
    try { await apiJson("/api/settings/profile", "PUT", profile); toast.success("Profile updated"); } catch (err) { toast.error(err instanceof Error ? err.message : "Save failed"); }
  }
  async function updatePassword() {
    if (password.next !== password.confirm) { toast.error("Password confirmation does not match"); return; }
    try { await apiJson("/api/auth/password", "PUT", { current_password: password.current, new_password: password.next }); toast.success("Password updated"); setPassword({ current: "", next: "", confirm: "" }); } catch (err) { toast.error(err instanceof Error ? err.message : "Password update failed"); }
  }
  async function saveAppSettings() {
    try { await apiJson("/api/admin/settings/app", "PUT", { settings: Object.entries(appSettings).map(([key, value]) => ({ key, value })) }); toast.success("App settings saved"); } catch (err) { toast.error(err instanceof Error ? err.message : "Save failed"); }
  }
  async function saveStorage() {
    if (!storagePath.trim()) { toast.error("Storage path is required"); return; }
    setStorageLoading(true);
    try {
      const next = await apiJson<ApiStorageSettings>("/api/admin/settings/storage", "PUT", { storage_path: storagePath.trim() });
      setStorage(next);
      setStoragePath(next.storage_path || storagePath.trim());
      toast.success("Storage path saved");
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Storage save failed");
    } finally {
      setStorageLoading(false);
    }
  }
  async function handleDocumentUpload(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    setUploadingDoc(true);
    try {
      const uploaded = await uploadDocument(file);
      setUploadedDoc(uploaded);
      toast.success("Document uploaded");
      loadStorage();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Upload failed");
    } finally {
      setUploadingDoc(false);
    }
  }

  const uploadedUrl = uploadedDoc?.file_url ? `${getApiBaseUrl()}${uploadedDoc.file_url}` : "";
  const tabs = [
    { id: "account", label: "Account", icon: User },
    { id: "security", label: "Security", icon: Shield },
    { id: "app", label: "App", icon: Settings },
    { id: "storage", label: "Storage", icon: Database },
    { id: "masters", label: "Masters", icon: Layers },
    { id: "payslips", label: "Payslips", icon: FileText },
  ];
  return <div><PageHeader title="Settings" /><div className="flex gap-4"><div className="w-44 shrink-0"><Card className="p-1.5">{tabs.map(({ id, label, icon: Icon }) => <button key={id} onClick={() => setTab(id)} className={cn("w-full flex items-center gap-2.5 px-3 py-2 rounded-lg text-xs font-medium transition-colors", tab === id ? "bg-blue-600 text-white" : "text-slate-600 hover:bg-slate-50")}><Icon size={13} />{label}</button>)}</Card></div><div className="flex-1">
    {tab === "account" && <Card className="p-5"><h3 className="font-semibold text-slate-800 text-sm mb-4">Account Settings</h3><div className="grid grid-cols-2 gap-4"><Input label="Full Name" value={profile.name} onChange={v => setProfile(p => ({ ...p, name: v }))} /><Input label="Email" type="email" value={profile.email} onChange={v => setProfile(p => ({ ...p, email: v }))} /></div><div className="flex justify-end mt-4 pt-4 border-t border-slate-100"><Btn variant="primary" onClick={saveProfile}>Save Changes</Btn></div></Card>}
    {tab === "security" && <Card className="p-5"><h3 className="font-semibold text-slate-800 text-sm mb-4">Security</h3><div className="space-y-4 max-w-sm"><Input label="Current Password" type="password" value={password.current} onChange={v => setPassword(p => ({ ...p, current: v }))} /><Input label="New Password" type="password" value={password.next} onChange={v => setPassword(p => ({ ...p, next: v }))} /><Input label="Confirm New Password" type="password" value={password.confirm} onChange={v => setPassword(p => ({ ...p, confirm: v }))} /><Btn variant="primary" onClick={updatePassword}>Update Password</Btn></div></Card>}
    {tab === "app" && <Card className="p-5"><h3 className="font-semibold text-slate-800 text-sm mb-4">App Settings</h3><div className="grid grid-cols-2 gap-4"><Input label="Company Name" value={appSettings.company_name} onChange={v => setAppSettings(p => ({ ...p, company_name: v }))} /><Input label="Expiry Warning Days" type="number" value={appSettings.expiry_warning_days} onChange={v => setAppSettings(p => ({ ...p, expiry_warning_days: v }))} /></div><div className="flex justify-end mt-4 pt-4 border-t border-slate-100"><Btn variant="primary" onClick={saveAppSettings}>Save App Settings</Btn></div></Card>}
    {tab === "storage" && <div className="space-y-4"><Card className="p-5"><div className="flex items-start justify-between mb-4"><div><h3 className="font-semibold text-slate-800 text-sm">Document Storage</h3><p className="text-xs text-slate-400 mt-1">Folder aktif untuk file dokumen yang di-upload</p></div><span className={cn("inline-flex items-center gap-1 rounded-full px-2.5 py-1 text-[11px] font-semibold", storage?.exists ? "bg-emerald-50 text-emerald-700" : "bg-amber-50 text-amber-700")}><span className={cn("h-1.5 w-1.5 rounded-full", storage?.exists ? "bg-emerald-500" : "bg-amber-500")} />{storage?.exists ? "Available" : "Not found"}</span></div><div className="grid grid-cols-[1fr_auto] gap-3"><Input label="Storage Path" value={storagePath} onChange={setStoragePath} placeholder="D:/Dokumen_SKYagen" /><div className="flex items-end"><Btn variant="primary" onClick={saveStorage} disabled={storageLoading}>{storageLoading ? <><RefreshCw size={13} className="animate-spin" />Saving...</> : "Save Path"}</Btn></div></div>{storage && <div className="mt-4 grid grid-cols-2 gap-3"><div className="rounded-lg border border-slate-100 bg-slate-50 p-3"><p className="text-[11px] font-semibold text-slate-400 uppercase">Absolute Path</p><p className="mt-1 break-all font-mono text-xs text-slate-700">{storage.absolute_path}</p></div><div className="rounded-lg border border-slate-100 bg-slate-50 p-3"><p className="text-[11px] font-semibold text-slate-400 uppercase">Rules</p><p className="mt-1 text-xs text-slate-700">Max {storage.max_file_size_mb} MB</p><p className="mt-1 text-[11px] text-slate-500">{storage.allowed_extensions.join(", ")}</p></div></div>}</Card><Card className="p-5"><div className="flex items-center justify-between mb-4"><div><h3 className="font-semibold text-slate-800 text-sm">Upload Document</h3><p className="text-xs text-slate-400 mt-1">Upload file ke storage aktif dan ambil URL preview</p></div><label className={cn("inline-flex items-center gap-2 rounded-lg px-3 py-2 text-xs font-semibold transition-colors", uploadingDoc ? "bg-slate-100 text-slate-400" : "cursor-pointer bg-blue-600 text-white hover:bg-blue-700")}><Upload size={13} />{uploadingDoc ? "Uploading..." : "Choose File"}<input type="file" className="hidden" disabled={uploadingDoc} accept={(storage?.allowed_extensions ?? []).join(",")} onChange={handleDocumentUpload} /></label></div>{uploadedDoc ? <div className="rounded-lg border border-slate-100 bg-slate-50 p-4"><div className="grid grid-cols-2 gap-3 text-xs"><div><p className="text-slate-400">Original Name</p><p className="mt-1 font-semibold text-slate-700">{uploadedDoc.original_name}</p></div><div><p className="text-slate-400">Stored Name</p><p className="mt-1 font-mono text-slate-700">{uploadedDoc.file_name}</p></div><div><p className="text-slate-400">Size</p><p className="mt-1 font-semibold text-slate-700">{Math.ceil(uploadedDoc.file_size / 1024)} KB</p></div><div><p className="text-slate-400">File URL</p><a className="mt-1 block break-all text-blue-600 hover:underline" href={uploadedUrl} target="_blank" rel="noreferrer">{uploadedDoc.file_url}</a></div></div><p className="mt-3 break-all rounded-md bg-white px-3 py-2 font-mono text-[11px] text-slate-500">{uploadedDoc.file_path}</p></div> : <div className="rounded-lg border border-dashed border-slate-200 bg-slate-50 p-8 text-center"><Upload size={24} className="mx-auto mb-2 text-slate-300" /><p className="text-sm font-medium text-slate-600">No uploaded document yet</p><p className="mt-1 text-xs text-slate-400">Pilih file untuk mengetes storage aktif</p></div>}</Card></div>}
    {tab === "payslips" && <PayslipTemplatesManager />}
    {tab === "masters" && <div className="space-y-4"><Card className="p-5"><h3 className="font-semibold text-slate-800 text-sm mb-4">Master Configuration</h3><div className="grid grid-cols-3 gap-3"><div className="p-3 bg-slate-50 rounded-xl border border-slate-100"><p className="text-xs font-semibold text-slate-700">Document Types</p><p className="text-[10px] text-slate-400 mt-0.5">Use Document Types page</p></div><div className="p-3 bg-slate-50 rounded-xl border border-slate-100"><p className="text-xs font-semibold text-slate-700">Document Names</p><p className="text-[10px] text-slate-400 mt-0.5">Use Document Names page</p></div><div className="p-3 bg-slate-50 rounded-xl border border-slate-100"><p className="text-xs font-semibold text-slate-700">Payslip Templates</p><p className="text-[10px] text-slate-400 mt-0.5">Kelola template Word (.docx)</p></div></div></Card><PayslipTemplatesManager /><CustomFieldsManager /><PrincipalRequirementsManager /></div>}
  </div></div></div>;
}

// ─── Import Modal ─────────────────────────────────────────────────────────────

function ImportModal({ open, onClose, entity, endpoint }: { open: boolean; onClose: () => void; entity: string; endpoint: string }) {
  const [step, setStep] = useState<"upload" | "preview" | "result">("upload");
  const [loading, setLoading] = useState(false);
  const [batchSize, setBatchSize] = useState("50");
  const [sheet, setSheet] = useState("");
  const [file, setFile] = useState<File | null>(null);
  const [result, setResult] = useState<{ inserted: number; skipped: number; errors: string[] } | null>(null);

  const errorCount = result?.errors?.length ?? 0;
  const totalRows = result ? result.inserted + result.skipped + errorCount : 0;

  function resetAndClose() {
    onClose();
    setStep("upload");
    setLoading(false);
    setBatchSize("50");
    setSheet("");
    setFile(null);
    setResult(null);
  }

  function continueToPreview() {
    if (!file) {
      toast.error("Pilih file Excel terlebih dahulu");
      return;
    }
    setStep("preview");
  }

  async function startImport() {
    if (!file) {
      toast.error("Pilih file Excel terlebih dahulu");
      setStep("upload");
      return;
    }

    setLoading(true);
    try {
      const response = await importExcel(endpoint, file, sheet, batchSize);
      setResult(response.result);
      toast.success(response.message || "Import selesai");
      setStep("result");
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Import gagal");
    } finally {
      setLoading(false);
    }
  }

  return (
    <Modal open={open} onClose={resetAndClose} title={`Import ${entity}`} size="lg"
      footer={
        step === "upload"
          ? <><Btn variant="secondary" size="sm" onClick={resetAndClose}>Cancel</Btn><Btn variant="primary" size="sm" onClick={continueToPreview}>Continue</Btn></>
          : step === "preview"
          ? <><Btn variant="secondary" size="sm" onClick={() => setStep("upload")}>Back</Btn><Btn variant="primary" size="sm" onClick={startImport} disabled={loading}>{loading ? <><RefreshCw size={12} className="animate-spin" />Importing...</> : "Start Import"}</Btn></>
          : <Btn variant="primary" size="sm" onClick={resetAndClose}>Done</Btn>
      }>
      {step === "upload" && (
        <div className="space-y-4">
          <div className="p-2.5 bg-slate-50 rounded-lg border border-slate-100 text-xs text-slate-500 font-mono">
            {endpoint}
          </div>
          <label className="block border-2 border-dashed border-slate-200 rounded-xl p-10 text-center cursor-pointer hover:border-blue-300 transition-colors">
            <Upload size={24} className="text-slate-300 mx-auto mb-2" />
            <p className="text-sm font-medium text-slate-600">{file ? file.name : "Drop your Excel file here"}</p>
            <p className="text-xs text-slate-400 mt-1">or click to browse, supports .xlsx and .xls</p>
            <input
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={event => setFile(event.target.files?.[0] ?? null)}
            />
          </label>
          <div className="grid grid-cols-2 gap-4">
            <Input label="Sheet Name" value={sheet} onChange={setSheet} placeholder="e.g. Sheet1" />
            <Input label="Batch Size" value={batchSize} onChange={setBatchSize} type="number" placeholder="50" />
          </div>
        </div>
      )}
      {step === "preview" && (
        <div className="space-y-3">
          <div className="flex items-center gap-3 p-3 bg-blue-50 rounded-xl border border-blue-100">
            <Info size={14} className="text-blue-500 shrink-0" />
            <p className="text-xs text-blue-700">Review file dan konfigurasi sebelum import ke database.</p>
          </div>
          <div className="grid grid-cols-2 gap-3">
            {[["File", file?.name || "-"], ["Sheet", sheet || "Default sheet"], ["Endpoint", endpoint], ["Batch Size", batchSize || "50"]].map(([k, v]) => (
              <div key={k} className="flex items-center justify-between gap-3 p-2.5 bg-slate-50 rounded-lg">
                <span className="text-xs text-slate-500">{k}</span>
                <span className="text-xs font-semibold text-slate-800 truncate">{v}</span>
              </div>
            ))}
          </div>
        </div>
      )}
      {step === "result" && (
        <div className="space-y-3">
          <div className="flex items-center gap-2 p-3 bg-emerald-50 rounded-xl border border-emerald-100">
            <CheckCircle size={14} className="text-emerald-500" />
            <p className="text-xs text-emerald-700 font-medium">Import completed</p>
          </div>
          <div className="grid grid-cols-2 gap-3">
            {[["Total Rows", String(totalRows)], ["Inserted", String(result?.inserted ?? 0)], ["Skipped", String(result?.skipped ?? 0)], ["Errors", String(errorCount)]].map(([k, v]) => (
              <div key={k} className={cn("p-3 rounded-xl text-center", k === "Errors" && v !== "0" ? "bg-red-50 border border-red-100" : "bg-slate-50 border border-slate-100")}>
                <p className={cn("text-xl font-bold", k === "Errors" && v !== "0" ? "text-red-600" : "text-slate-800")}>{v}</p>
                <p className="text-xs text-slate-500 mt-0.5">{k}</p>
              </div>
            ))}
          </div>
          {errorCount > 0 && (
            <div className="p-3 bg-amber-50 rounded-xl border border-amber-100 space-y-1">
              <p className="text-xs font-semibold text-amber-700">{errorCount} Error</p>
              {result?.errors.slice(0, 5).map((message, index) => (
                <p key={`${message}-${index}`} className="text-xs text-amber-600 font-mono">{message}</p>
              ))}
            </div>
          )}
        </div>
      )}
    </Modal>
  );
}

// ─── App ──────────────────────────────────────────────────────────────────────

export default function App() {
  const [isAuth, setIsAuth] = useState(() => !!getToken());
  const [authPage, setAuthPage] = useState<AuthPage>("login");
  const [page, setPage] = useState<Page>("dashboard");
  const [selectedCrew, setSelectedCrew] = useState<Crew | null>(null);
  const [selectedPrincipal, setSelectedPrincipal] = useState<Principal | null>(null);

  useEffect(() => {
    checkApiConnection();
  }, []);

  useEffect(() => {
    const handler = () => localStorage.setItem("skyagen_api_disconnected_page", page);
    window.addEventListener("skyagen:api-disconnected", handler);
    if (localStorage.getItem("skyagen_api_disconnected_at")) handler();
    return () => window.removeEventListener("skyagen:api-disconnected", handler);
  }, [page]);

  function renderAuthPage() {
    switch (authPage) {
      case "login": return <LoginPage setAuth={() => setIsAuth(true)} setAuthPage={setAuthPage} />;
      case "register": return <RegisterPage setAuthPage={setAuthPage} />;
      case "verify-email": return <VerifyEmailPage setAuthPage={setAuthPage} />;
      case "verify-success": return <VerifySuccessPage setAuthPage={setAuthPage} />;
      case "verify-failed": return <VerifyFailedPage setAuthPage={setAuthPage} />;
      case "update-password": return <UpdatePasswordPage setAuthPage={setAuthPage} />;
      case "error-401": return <ErrorPage code={401} setAuthPage={setAuthPage} />;
      case "error-403": return <ErrorPage code={403} setAuthPage={setAuthPage} />;
      case "error-404": return <ErrorPage code={404} setAuthPage={setAuthPage} />;
      default: return <LoginPage setAuth={() => setIsAuth(true)} setAuthPage={setAuthPage} />;
    }
  }

  function renderPage() {
    switch (page) {
      case "dashboard": return <DashboardPage setPage={setPage} />;
      case "crew-database": return <CrewDatabasePage setPage={setPage} setSelectedCrew={setSelectedCrew} />;
      case "crew-detail": return selectedCrew ? <CrewDetailPage crew={selectedCrew} setPage={setPage} /> : <CrewDatabasePage setPage={setPage} setSelectedCrew={setSelectedCrew} />;
      case "crew-form": return <CrewFormPage crew={selectedCrew} setPage={setPage} />;
      case "search-crew": return <SearchCrewPage setPage={setPage} setSelectedCrew={setSelectedCrew} />;
      case "available-crew": return <CrewStatusPage status="available" title="Available Crew" setPage={setPage} setSelectedCrew={setSelectedCrew} />;
      case "onboard-crew": return <CrewStatusPage status="onboard" title="Onboard Crew" setPage={setPage} setSelectedCrew={setSelectedCrew} />;
      case "waiting-crew": return <CrewStatusPage status="waiting" title="Waiting Crew" setPage={setPage} setSelectedCrew={setSelectedCrew} />;
      case "ex-crew": return <CrewStatusPage status="ex-crew" title="Ex Crew" setPage={setPage} setSelectedCrew={setSelectedCrew} />;
      case "blacklist": return <BlacklistPage />;
      case "joining-principal": return <JoiningPrincipalPage />;
      case "sign-on": return <SignOnPage />;
      case "sign-off": return <SignOffPage />;
      case "documents": return <DocumentsPage setPage={setPage} />;
      case "expiring-documents": return <ExpiringDocumentsPage />;
      case "document-types": return <DocumentTypesPage />;
      case "document-names": return <DocumentNamesPage />;
      case "vessels": return <VesselsPage />;
      case "principals": return <PrincipalsPage setPage={setPage} setSelectedPrincipal={setSelectedPrincipal} />;
      case "principal-detail": return selectedPrincipal ? <PrincipalDetailPage principal={selectedPrincipal} setPage={setPage} /> : <PrincipalsPage setPage={setPage} setSelectedPrincipal={setSelectedPrincipal} />;
      case "crew-reports": return <CrewReportsPage />;
      case "document-reports": return <DocumentReportsPage />;
      case "users": return <UsersPage />;
      case "settings": return <SettingsPage />;
      default: return <DashboardPage setPage={setPage} />;
    }
  }

  if (!isAuth) {
    return (
      <>
        <Toaster position="top-right" richColors />
        <ApiConnectionModal />
        {renderAuthPage()}
      </>
    );
  }

  return (
    <>
      <Toaster position="top-right" richColors />
      <ApiConnectionModal />
      <AdminLayout currentPage={page} setPage={setPage} onLogout={() => { clearToken(); setIsAuth(false); setAuthPage("login"); }}>
        <AppErrorBoundary>{renderPage()}</AppErrorBoundary>
      </AdminLayout>
    </>
  );
}
