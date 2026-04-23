const express = require("express");
const cors = require("cors");
const path = require("path");
const session = require("express-session");
const XLSX = require("xlsx");
const { query, hashPassword, initDb, withTransaction } = require("./db");
const {
  MEETING_STATUS,
  MEETING_TYPES,
  MEETING_REASONS,
  USER_ROLES,
  USER_ROLE_OPTIONS,
  OPPORTUNITY_STATUS,
  OPPORTUNITY_STATUS_OPTIONS,
  OPPORTUNITY_OPEN_STATUSES,
  OPPORTUNITY_SERVICE_LINES,
  OPPORTUNITY_TYPES,
  FOLLOW_UP_STATUS,
  FOLLOW_UP_STATUS_OPTIONS,
  FOLLOW_UP_TYPES
} = require("./constants");

const app = express();
const PORT = process.env.PORT || 3000;
const REMEMBER_ME_COOKIE_MAX_AGE_MS = 1000 * 60 * 60 * 24 * 30;
const VALID_RISKS = new Set(["Bajo", "Medio", "Alto"]);
const VALID_SEGMENTS = new Set(["A", "B", "C"]);
const VALID_COMPANY_TYPES = new Set(["Local", "Global"]);
const VALID_COUNTRIES = new Set(["Argentina", "Bolivia", "Mexico"]);
const VALID_ACCOUNT_STAGES = new Set(["Activa", "Prospecto"]);
const VALID_MEETING_STATUS = new Set(Object.values(MEETING_STATUS));
const MEETING_MODALITIES = ["Virtual", "Presencial"];
const VISIT_RULE_OBJECTIVES = ["Desarrollo de cuentas", "Nuevo negocio"];
const VALID_MEETING_MODALITIES = new Set(MEETING_MODALITIES);
const VALID_VISIT_RULE_OBJECTIVES = new Set(VISIT_RULE_OBJECTIVES);
const VALID_OPPORTUNITY_STATUS = new Set(OPPORTUNITY_STATUS_OPTIONS);
const VALID_SERVICE_LINES = new Set(OPPORTUNITY_SERVICE_LINES);
const VALID_OPPORTUNITY_TYPES = new Set(OPPORTUNITY_TYPES);
const VALID_FOLLOW_UP_STATUS = new Set(FOLLOW_UP_STATUS_OPTIONS);
const VALID_FOLLOW_UP_TYPES = new Set(FOLLOW_UP_TYPES);
const OPEN_OPPORTUNITY_STATUS_SET = new Set(OPPORTUNITY_OPEN_STATUSES);
const OPPORTUNITY_STATUS_SORT_ORDER = OPPORTUNITY_STATUS_OPTIONS.reduce((accumulator, status, index) => {
  accumulator[status] = index;
  return accumulator;
}, {});
const EXECUTIVE_ROLE = USER_ROLES.EXECUTIVE;
const VISIT_RULES_FEATURE_SETTING_KEY = "visit_rules_enabled";
const RULE_SEMAPHORE = {
  WHITE: "white",
  GREEN: "green",
  YELLOW: "yellow",
  RED: "red"
};
const BITRIX_COMPLAINT_GROUP_ID = 46;
const BITRIX_COMPLAINT_RESPONSIBLES = [
  { id: "44", name: "Gonzalo Garcia" },
  { id: "10542", name: "Guillermo Saralegui" },
  { id: "1720", name: "Guido Avalo Ceci" }
];
const VALID_BITRIX_COMPLAINT_RESPONSIBLE_IDS = new Set(BITRIX_COMPLAINT_RESPONSIBLES.map((item) => item.id));
const SUPERVISOR_ROLE_BY_SERVICE = {
  fixedFire: USER_ROLES.SUPERVISOR_IFCI,
  extinguishers: USER_ROLES.SUPERVISOR_EXT,
  works: USER_ROLES.SUPERVISOR_WORKS
};
let meetingTypeOptionsCache = [...MEETING_TYPES];
let meetingReasonOptionsCache = [...MEETING_REASONS];
let contactRoleOptionsCache = [];

app.use(cors());
app.use(express.json());
app.use(
  session({
    secret: process.env.SESSION_SECRET || "top-clientes-dev-secret",
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      sameSite: "lax"
    }
  })
);
app.use((_req, res, next) => {
  res.setHeader("Cache-Control", "no-store");
  next();
});
app.use(express.static(path.join(__dirname, "..", "public")));

function asyncHandler(handler) {
  return (req, res, next) => {
    Promise.resolve(handler(req, res, next)).catch(next);
  };
}

function sanitizeUser(row) {
  if (!row) return null;
  return {
    id: row.id,
    name: row.name,
    email: row.email,
    role: row.role,
    bitrixUserId: row.bitrix_user_id || ""
  };
}

async function refreshMeetingTypesCache() {
  const { rows } = await query(
    `
    SELECT id, value, label, color, sort_order
    FROM meeting_type_options
    ORDER BY sort_order ASC, id ASC
    `
  );

  meetingTypeOptionsCache = rows.length
    ? rows.map((row) => ({
        id: row.id,
        value: row.value,
        label: row.label,
        color: row.color
      }))
    : [...MEETING_TYPES];
}

function getMeetingTypes() {
  return meetingTypeOptionsCache;
}

async function refreshMeetingReasonsCache() {
  const { rows } = await query(
    `
    SELECT id, name, sort_order
    FROM meeting_reason_options
    ORDER BY sort_order ASC, id ASC
    `
  );

  meetingReasonOptionsCache = rows.length
    ? rows.map((row) => ({ id: row.id, name: row.name }))
    : [...MEETING_REASONS];
}

function getMeetingReasons() {
  return meetingReasonOptionsCache;
}

async function refreshContactRoleOptionsCache() {
  const { rows } = await query(
    `
    SELECT id, name, sort_order
    FROM contact_role_options
    ORDER BY sort_order ASC, id ASC
    `
  );

  contactRoleOptionsCache = rows.map((row) => ({ id: row.id, name: row.name }));
}

function getContactRoleOptions() {
  return contactRoleOptionsCache;
}

function normalizeBitrixWebhookRoot(input) {
  const raw = String(input || "").trim();
  if (!raw) {
    throw new Error("Debes indicar la URL del webhook de Bitrix");
  }

  let parsed;
  try {
    parsed = new URL(raw);
  } catch (_error) {
    throw new Error("La URL del webhook de Bitrix es inválida");
  }

  if (!/^https?:$/.test(parsed.protocol)) {
    throw new Error("La URL del webhook debe usar http o https");
  }

  if (!parsed.hostname.toLowerCase().includes("bitrix")) {
    throw new Error("La URL no parece pertenecer a Bitrix");
  }

  const pathSegments = parsed.pathname.split("/").filter(Boolean);
  const restIndex = pathSegments.findIndex((segment) => segment === "rest");

  if (restIndex === -1 || pathSegments.length < restIndex + 3) {
    throw new Error("La URL del webhook no tiene el formato esperado de Bitrix");
  }

  const rootSegments = pathSegments.slice(0, restIndex + 3);
  parsed.pathname = `/${rootSegments.join("/")}/`;
  parsed.search = "";
  parsed.hash = "";

  return parsed.toString();
}

function buildBitrixMethodUrl(bitrixRoot, method) {
  return `${String(bitrixRoot || "").replace(/\/+$/, "")}/${method}.json`;
}

function appendBitrixParams(searchParams, key, value) {
  if (value === undefined || value === null || value === "") return;

  if (Array.isArray(value)) {
    value.forEach((item) => {
      appendBitrixParams(searchParams, `${key}[]`, item);
    });
    return;
  }

  if (typeof value === "object") {
    Object.entries(value).forEach(([childKey, childValue]) => {
      appendBitrixParams(searchParams, `${key}[${childKey}]`, childValue);
    });
    return;
  }

  searchParams.append(key, String(value));
}

async function callBitrixMethod(bitrixRoot, method, payload = {}) {
  const url = buildBitrixMethodUrl(bitrixRoot, method);
  const formData = new URLSearchParams();

  Object.entries(payload || {}).forEach(([key, value]) => {
    appendBitrixParams(formData, key, value);
  });

  const response = await fetch(url, {
    method: "POST",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8"
    },
    body: formData.toString()
  });

  const rawBody = await response.text();
  let data = null;
  try {
    data = rawBody ? JSON.parse(rawBody) : null;
  } catch (_error) {
    data = null;
  }

  if (!response.ok || data?.error) {
    const error = new Error(
      data?.error_description || data?.error || `Bitrix respondió con error ${response.status}`
    );
    error.status = response.status || 502;
    error.code = data?.error || "";
    throw error;
  }

  return data;
}

async function fetchBitrixCollection(bitrixRoot, method, payload = {}, maxPages = 20) {
  const items = [];
  let next = 0;
  let guard = 0;

  while (guard < maxPages) {
    const data = await callBitrixMethod(bitrixRoot, method, {
      ...payload,
      start: next
    });

    const batch = Array.isArray(data?.result) ? data.result : [];
    items.push(...batch);

    if (data?.next === undefined || data?.next === null) {
      break;
    }

    next = Number(data.next);
    if (!Number.isFinite(next) || next <= 0) {
      break;
    }

    guard += 1;
  }

  return items;
}

function mapBitrixUser(user) {
  const activeRaw = user.ACTIVE;
  const isActive =
    activeRaw === true ||
    activeRaw === 1 ||
    String(activeRaw || "")
      .trim()
      .toUpperCase() === "Y" ||
    String(activeRaw || "")
      .trim()
      .toLowerCase() === "true";

  return {
    id: user.ID || user.id || "",
    name: user.NAME || "",
    lastName: user.LAST_NAME || "",
    email: user.EMAIL || "",
    active: isActive,
    admin: String(user.IS_ADMIN || "").toUpperCase() === "Y",
    workPosition: user.WORK_POSITION || "",
    fullName: [user.NAME, user.LAST_NAME].filter(Boolean).join(" ").trim() || user.EMAIL || `Usuario ${user.ID || ""}`
  };
}

async function fetchBitrixUsers(bitrixRoot) {
  const users = await fetchBitrixCollection(bitrixRoot, "user.search", {
    FILTER: { ACTIVE: true }
  });

  return users.map(mapBitrixUser);
}

function mapBitrixCompany(company) {
  return {
    id: company.ID || "",
    title: company.TITLE || company.COMPANY_TITLE || `Compañía ${company.ID || ""}`,
    assignedById: company.ASSIGNED_BY_ID || "",
    companyType: company.COMPANY_TYPE || ""
  };
}

async function fetchBitrixCompanies(bitrixRoot) {
  const companies = await fetchBitrixCollection(bitrixRoot, "crm.company.list", {
    order: { ID: "ASC" },
    select: ["ID", "TITLE", "ASSIGNED_BY_ID", "COMPANY_TYPE"]
  });

  return companies.map(mapBitrixCompany);
}

async function searchBitrixCompanies(bitrixRoot, searchTerm) {
  const normalizedSearchTerm = String(searchTerm || "").trim();
  if (!normalizedSearchTerm) return [];

  const companies = await fetchBitrixCollection(
    bitrixRoot,
    "crm.company.list",
    {
      order: { TITLE: "ASC" },
      select: ["ID", "TITLE", "ASSIGNED_BY_ID", "COMPANY_TYPE"],
      filter: {
        "%TITLE": normalizedSearchTerm
      }
    },
    1
  );

  return companies.map(mapBitrixCompany);
}

function formatBitrixDirectoryError(error, resourceLabel) {
  if (!error) {
    return `No se pudieron cargar ${resourceLabel} de Bitrix.`;
  }

  if (String(error.code || "").trim() === "insufficient_scope") {
    return `El webhook actual no tiene permisos para leer ${resourceLabel} de Bitrix.`;
  }

  return error.message || `No se pudieron cargar ${resourceLabel} de Bitrix.`;
}

async function getBitrixDirectoryOptions() {
  const webhookUrl = await getAppSetting("bitrix_webhook_url", "");
  if (!webhookUrl) {
    return {
      users: [],
      companies: [],
      permissions: {
        users: false,
        companies: false
      },
      errors: {
        users: "No hay webhook de Bitrix configurado.",
        companies: "No hay webhook de Bitrix configurado."
      }
    };
  }

  const bitrixRoot = normalizeBitrixWebhookRoot(webhookUrl);
  const [usersResult, companiesResult] = await Promise.all([
    fetchBitrixUsers(bitrixRoot).then((value) => ({ ok: true, value })).catch((error) => ({ ok: false, error })),
    fetchBitrixCompanies(bitrixRoot).then((value) => ({ ok: true, value })).catch((error) => ({ ok: false, error }))
  ]);

  return {
    users: (usersResult.ok ? usersResult.value : [])
      .filter((user) => Boolean(user.active))
      .map((user) => ({
        id: String(user.id || ""),
        label: `${user.fullName}${user.email ? ` · ${user.email}` : ""}${user.id ? ` · ${user.id}` : ""}`.trim(),
        name: user.fullName,
        email: user.email || ""
      })),
    companies: (companiesResult.ok ? companiesResult.value : []).map((company) => ({
      id: String(company.id || ""),
      label: `${company.title}${company.id ? ` · ${company.id}` : ""}`.trim(),
      title: company.title
    })),
    permissions: {
      users: usersResult.ok,
      companies: companiesResult.ok
    },
    errors: {
      users: usersResult.ok ? "" : formatBitrixDirectoryError(usersResult.error, "usuarios"),
      companies: companiesResult.ok ? "" : formatBitrixDirectoryError(companiesResult.error, "compañías")
    }
  };
}

function mapBitrixLead(lead) {
  return {
    id: lead.ID || "",
    title: lead.TITLE || `Lead ${lead.ID || ""}`,
    assignedById: lead.ASSIGNED_BY_ID || "",
    statusId: lead.STATUS_ID || ""
  };
}

async function fetchBitrixLeads(bitrixRoot) {
  const leads = await fetchBitrixCollection(bitrixRoot, "crm.lead.list", {
    order: { ID: "ASC" },
    select: ["ID", "TITLE", "ASSIGNED_BY_ID", "STATUS_ID"]
  });

  return leads.map(mapBitrixLead);
}

async function buildBitrixMappings(bitrixRoot) {
  const [bitrixUsersResult, bitrixCompaniesResult, bitrixLeadsResult, usersResult, clientsResult] = await Promise.all([
    fetchBitrixUsers(bitrixRoot).then((value) => ({ ok: true, value })).catch((error) => ({ ok: false, error })),
    fetchBitrixCompanies(bitrixRoot).then((value) => ({ ok: true, value })).catch((error) => ({ ok: false, error })),
    fetchBitrixLeads(bitrixRoot).then((value) => ({ ok: true, value })).catch((error) => ({ ok: false, error })),
    query("SELECT id, name, email, role, bitrix_user_id FROM users ORDER BY name ASC, id ASC"),
    query(
      `
      SELECT id, name, bitrix_company_id, bitrix_lead_id
      FROM clients
      WHERE is_hidden = FALSE
      ORDER BY name ASC, id ASC
      `
    )
  ]);

  if (!bitrixUsersResult.ok) {
    throw bitrixUsersResult.error;
  }

  const bitrixUsers = bitrixUsersResult.value;
  const bitrixCompanies = bitrixCompaniesResult.ok ? bitrixCompaniesResult.value : [];
  const bitrixLeads = bitrixLeadsResult.ok ? bitrixLeadsResult.value : [];
  const bitrixUsersById = new Map(bitrixUsers.map((user) => [String(user.id), user]));
  const bitrixCompaniesById = new Map(bitrixCompanies.map((company) => [String(company.id), company]));
  const bitrixLeadsById = new Map(bitrixLeads.map((lead) => [String(lead.id), lead]));

  return {
    bitrixUsers,
    bitrixCompanies,
    userMappings: usersResult.rows.map((row) => {
      const bitrixUserId = String(row.bitrix_user_id || "").trim();
      const bitrixUser = bitrixUserId ? bitrixUsersById.get(bitrixUserId) || null : null;

      return {
        appUser: sanitizeUser(row),
        localBitrixUserId: bitrixUserId,
        bitrixUser,
        status: bitrixUserId ? (bitrixUser ? "mapped" : "missing") : "unmapped"
      };
    }),
    clientMappings: clientsResult.rows.map((row) => {
      const bitrixCompanyId = String(row.bitrix_company_id || "").trim();
      const bitrixLeadId = String(row.bitrix_lead_id || "").trim();
      const company = bitrixCompanyId ? bitrixCompaniesById.get(bitrixCompanyId) || null : null;
      const lead = bitrixLeadId ? bitrixLeadsById.get(bitrixLeadId) || null : null;

      return {
        clientId: row.id,
        clientName: row.name,
        bitrixCompanyId,
        bitrixLeadId,
        matchedEntity: company || lead,
        matchedEntityType: company ? "company" : lead ? "lead" : "",
        status:
          bitrixCompaniesResult.ok || bitrixLeadsResult.ok
            ? bitrixCompanyId || bitrixLeadId
              ? company || lead
                ? "mapped"
                : "missing"
              : "unmapped"
            : "scope_missing"
      };
    }),
    counts: {
      localUsers: usersResult.rows.length,
      localClients: clientsResult.rows.length,
      bitrixUsers: bitrixUsers.length,
      bitrixCompanies: bitrixCompanies.length,
      bitrixLeads: bitrixLeads.length
    },
    permissions: {
      users: bitrixUsersResult.ok,
      companies: bitrixCompaniesResult.ok,
      leads: bitrixLeadsResult.ok
    },
    errors: {
      companies: bitrixCompaniesResult.ok ? "" : bitrixCompaniesResult.error?.message || "",
      leads: bitrixLeadsResult.ok ? "" : bitrixLeadsResult.error?.message || ""
    }
  };
}

function formatBitrixDeadline(daysFromNow = 7) {
  const deadline = new Date();
  deadline.setDate(deadline.getDate() + Number(daysFromNow || 0));
  deadline.setHours(23, 59, 0, 0);

  const year = deadline.getFullYear();
  const month = String(deadline.getMonth() + 1).padStart(2, "0");
  const day = String(deadline.getDate()).padStart(2, "0");
  const hours = String(deadline.getHours()).padStart(2, "0");
  const minutes = String(deadline.getMinutes()).padStart(2, "0");
  const offsetMinutes = -deadline.getTimezoneOffset();
  const sign = offsetMinutes >= 0 ? "+" : "-";
  const absOffsetMinutes = Math.abs(offsetMinutes);
  const offsetHours = String(Math.floor(absOffsetMinutes / 60)).padStart(2, "0");
  const offsetRemainderMinutes = String(absOffsetMinutes % 60).padStart(2, "0");

  return `${year}-${month}-${day}T${hours}:${minutes}:00${sign}${offsetHours}:${offsetRemainderMinutes}`;
}

function buildComplaintBitrixTaskTitle(clientName, scheduledFor) {
  return `Reclamo - ${String(clientName || "").trim()} - ${String(scheduledFor || "").trim()}`;
}

async function createBitrixTask(bitrixRoot, { title, description, responsibleId, groupId, deadline, createdById = "" }) {
  const fields = {
    TITLE: title,
    DESCRIPTION: description,
    RESPONSIBLE_ID: responsibleId,
    GROUP_ID: groupId,
    DEADLINE: deadline
  };

  if (String(createdById || "").trim()) {
    fields.CREATED_BY = String(createdById || "").trim();
  }

  const data = await callBitrixMethod(bitrixRoot, "tasks.task.add", { fields });

  return {
    id: data?.result?.task?.id || data?.result?.id || null,
    raw: data?.result || null
  };
}

async function updateBitrixTask(bitrixRoot, { taskId, title, description, responsibleId, groupId, deadline }) {
  const data = await callBitrixMethod(bitrixRoot, "tasks.task.update", {
    taskId,
    fields: {
      TITLE: title,
      DESCRIPTION: description,
      RESPONSIBLE_ID: responsibleId,
      GROUP_ID: groupId,
      DEADLINE: deadline
    }
  });

  return {
    id: String(taskId || ""),
    raw: data?.result || null
  };
}

async function getAppSetting(key, fallback = "") {
  const { rows } = await query("SELECT value FROM app_settings WHERE key = $1 LIMIT 1", [String(key || "").trim()]);
  if (!rows[0]) return fallback;
  return String(rows[0].value || "");
}

function parseAppSettingBoolean(value, fallback = true) {
  if (value === undefined || value === null || value === "") return fallback;
  const normalized = String(value).trim().toLowerCase();
  if (["1", "true", "si", "sí", "yes", "y", "on"].includes(normalized)) return true;
  if (["0", "false", "no", "n", "off"].includes(normalized)) return false;
  return fallback;
}

async function isVisitRulesFeatureEnabled() {
  const rawValue = await getAppSetting(VISIT_RULES_FEATURE_SETTING_KEY, "true");
  return parseAppSettingBoolean(rawValue, true);
}

async function setAppSetting(key, value) {
  await query(
    `
    INSERT INTO app_settings (key, value, updated_at)
    VALUES ($1, $2, NOW())
    ON CONFLICT (key)
    DO UPDATE SET value = EXCLUDED.value, updated_at = NOW()
    `,
    [String(key || "").trim(), String(value || "")]
  );
}

function isValidMeetingReason(subject) {
  return getMeetingReasons().some((reason) => (typeof reason === "string" ? reason : reason.name) === subject);
}

function isValidMeetingType(kind) {
  return getMeetingTypes().some((type) => type.value === kind);
}

function validateUserPayload(payload, { requirePassword = true } = {}) {
  const errors = [];
  const name = String(payload?.name || "").trim();
  const email = String(payload?.email || "").trim().toLowerCase();
  const password = String(payload?.password || "");
  const role = String(payload?.role || "").trim();
  const bitrixUserId = String(payload?.bitrixUserId || "").trim();

  if (!name) errors.push("El nombre es obligatorio");
  if (!email || !email.includes("@")) errors.push("El email es inválido");
  if (requirePassword && password.length < 6) errors.push("La contraseña debe tener al menos 6 caracteres");
  if (!requirePassword && password && password.length < 6) {
    errors.push("La contraseña debe tener al menos 6 caracteres");
  }
  if (!USER_ROLE_OPTIONS.includes(role)) errors.push("Rol inválido");

  return {
    errors,
    values: { name, email, password, role, bitrixUserId }
  };
}

function requireAuth(req, res, next) {
  if (!req.session.userId) {
    res.status(401).json({ error: "Sesión requerida" });
    return;
  }
  next();
}

function isSettingsAdminUser(user) {
  return (
    String(user?.email || "").trim().toLowerCase() === "nicolas@maxiseguridad.com" &&
    String(user?.name || "").trim() === "Nicolas Beguelman"
  );
}

async function requireSettingsAdmin(req, res, next) {
  const { rows } = await query("SELECT id, name, email, role FROM users WHERE id = $1", [req.session.userId]);
  const user = rows[0];

  if (!user) {
    req.session.destroy(() => {});
    res.status(401).json({ error: "Sesión inválida" });
    return;
  }

  if (!isSettingsAdminUser(user)) {
    res.status(403).json({ error: "No tenés permisos para acceder a esta sección" });
    return;
  }

  req.authUser = user;
  next();
}

function isExecutiveScopedUser(user) {
  return String(user?.role || "").trim() === EXECUTIVE_ROLE;
}

function normalizeOptionalIdFilter(value) {
  const normalized = String(value || "").trim();
  if (!normalized || normalized === "todos") return null;
  const numericValue = Number(normalized);
  return Number.isInteger(numericValue) && numericValue > 0 ? numericValue : null;
}

async function getSessionUser(req) {
  if (!req.session.userId) return null;
  return getUserById(req.session.userId);
}

async function requireClientAccess(req, res, next) {
  const user = await getSessionUser(req);

  if (!user) {
    req.session.destroy(() => {});
    res.status(401).json({ error: "Sesión inválida" });
    return;
  }

  const clientId = Number(req.params.id);
  if (!Number.isInteger(clientId) || clientId <= 0) {
    res.status(400).json({ error: "Cliente inválido" });
    return;
  }

  const clientResult = await query("SELECT id, executive_user_id, is_hidden FROM clients WHERE id = $1", [clientId]);
  const client = clientResult.rows[0];

  if (!client || client.is_hidden) {
    res.status(404).json({ error: "Cliente no encontrado" });
    return;
  }

  if (isExecutiveScopedUser(user) && Number(client.executive_user_id) !== Number(user.id)) {
    res.status(403).json({ error: "No tenés permisos para acceder a esta compañía" });
    return;
  }

  req.authUser = sanitizeUser(user);
  req.accessClient = {
    id: client.id,
    executiveUserId: client.executive_user_id
  };
  next();
}

async function requireOpportunityAccess(req, res, next) {
  const user = await getSessionUser(req);

  if (!user) {
    req.session.destroy(() => {});
    res.status(401).json({ error: "Sesión inválida" });
    return;
  }

  const opportunityId = Number(req.params.opportunityId);
  if (!Number.isInteger(opportunityId) || opportunityId <= 0) {
    res.status(400).json({ error: "Oportunidad inválida" });
    return;
  }

  const result = await query(
    `
    SELECT opportunities.id, opportunities.client_id, clients.executive_user_id, clients.is_hidden
    FROM opportunities
    INNER JOIN clients ON clients.id = opportunities.client_id
    WHERE opportunities.id = $1
    `,
    [opportunityId]
  );
  const opportunity = result.rows[0];

  if (!opportunity || opportunity.is_hidden) {
    res.status(404).json({ error: "Oportunidad no encontrada" });
    return;
  }

  if (isExecutiveScopedUser(user) && Number(opportunity.executive_user_id) !== Number(user.id)) {
    res.status(403).json({ error: "No tenés permisos para acceder a esta oportunidad" });
    return;
  }

  req.authUser = sanitizeUser(user);
  req.accessOpportunity = {
    id: opportunity.id,
    clientId: opportunity.client_id,
    executiveUserId: opportunity.executive_user_id
  };
  next();
}

async function logAuditEvent({
  req = null,
  user = null,
  action,
  entityType = "",
  entityId = null,
  entityName = "",
  details = {}
}) {
  const auditUser = user || (req ? await getSessionUser(req) : null);

  await query(
    `
    INSERT INTO audit_logs (
      user_id, user_name, user_email, action, entity_type, entity_id, entity_name, details, created_at
    ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8::jsonb, NOW())
    `,
    [
      auditUser?.id || null,
      String(auditUser?.name || "").trim(),
      String(auditUser?.email || "").trim().toLowerCase(),
      String(action || "").trim(),
      String(entityType || "").trim(),
      entityId === null || entityId === undefined ? null : Number(entityId),
      String(entityName || "").trim(),
      JSON.stringify(details || {})
    ]
  );
}

function toClientObject(row) {
  return {
    id: row.id,
    rank: row.position,
    name: row.name,
    sector: row.sector,
    companyType: row.company_type || "Local",
    country: row.country || "Argentina",
    accountStage: row.account_stage || "Activa",
    bitrixLeadId: row.bitrix_lead_id || "",
    bitrixCompanyId: row.bitrix_company_id || "",
    manager: row.manager,
    executiveUserId: row.executive_user_id,
    executiveName: row.executive_name || row.manager,
    billing: Number(row.billing_2025),
    potential: Math.round(Number(row.billing_2025) * 1.12),
    segment: row.segment,
    risk: row.risk,
    services: {
      fixedFire: !!row.service_fixed_fire,
      extinguishers: !!row.service_extinguishers,
      works: !!row.service_works
    },
    supervisors: {
      fixedFire: row.supervisor_ifci_user_id
        ? {
            userId: row.supervisor_ifci_user_id,
            name: row.supervisor_ifci_name || "",
            role: USER_ROLES.SUPERVISOR_IFCI
          }
        : null,
      extinguishers: row.supervisor_ext_user_id
        ? {
            userId: row.supervisor_ext_user_id,
            name: row.supervisor_ext_name || "",
            role: USER_ROLES.SUPERVISOR_EXT
          }
        : null,
      works: row.supervisor_works_user_id
        ? {
            userId: row.supervisor_works_user_id,
            name: row.supervisor_works_name || "",
            role: USER_ROLES.SUPERVISOR_WORKS
          }
        : null
    },
    walletShare: row.wallet_share,
    nps: row.nps,
    openOpportunities: row.open_opportunities,
    notes: row.notes
  };
}

function toBranchObject(row) {
  return {
    id: row.id,
    clientId: row.client_id,
    name: row.name,
    bitrixCompanyId: row.bitrix_company_id || "",
    sector: row.sector,
    manager: row.manager,
    executiveUserId: row.parent_executive_user_id,
    executiveName: row.parent_executive_name || row.parent_manager || row.manager,
    segment: row.segment,
    risk: row.risk,
    services: {
      fixedFire: !!row.service_fixed_fire,
      extinguishers: !!row.service_extinguishers,
      works: !!row.service_works
    },
    supervisors: {
      fixedFire: row.supervisor_ifci_user_id
        ? {
            userId: row.supervisor_ifci_user_id,
            name: row.supervisor_ifci_name || "",
            role: USER_ROLES.SUPERVISOR_IFCI
          }
        : null,
      extinguishers: row.supervisor_ext_user_id
        ? {
            userId: row.supervisor_ext_user_id,
            name: row.supervisor_ext_name || "",
            role: USER_ROLES.SUPERVISOR_EXT
          }
        : null,
      works: row.supervisor_works_user_id
        ? {
            userId: row.supervisor_works_user_id,
            name: row.supervisor_works_name || "",
            role: USER_ROLES.SUPERVISOR_WORKS
          }
        : null
    },
    notes: row.notes,
    createdAt: row.created_at,
    updatedAt: row.updated_at
  };
}

function buildClientSelectQuery(whereClause = "", orderClause = "ORDER BY clients.billing_2025 DESC, clients.id ASC") {
  return `
    SELECT
      clients.*,
      executive_user.name AS executive_name,
      supervisor_ifci.name AS supervisor_ifci_name,
      supervisor_ext.name AS supervisor_ext_name,
      supervisor_works.name AS supervisor_works_name
    FROM clients
    LEFT JOIN users AS executive_user ON executive_user.id = clients.executive_user_id
    LEFT JOIN users AS supervisor_ifci ON supervisor_ifci.id = clients.supervisor_ifci_user_id
    LEFT JOIN users AS supervisor_ext ON supervisor_ext.id = clients.supervisor_ext_user_id
    LEFT JOIN users AS supervisor_works ON supervisor_works.id = clients.supervisor_works_user_id
    ${whereClause}
    ${orderClause}
  `;
}

function buildBranchSelectQuery(whereClause = "", orderClause = "ORDER BY client_branches.id ASC") {
  return `
    SELECT
      client_branches.*,
      parent_client.manager AS parent_manager,
      parent_client.executive_user_id AS parent_executive_user_id,
      parent_executive_user.name AS parent_executive_name,
      supervisor_ifci.name AS supervisor_ifci_name,
      supervisor_ext.name AS supervisor_ext_name,
      supervisor_works.name AS supervisor_works_name
    FROM client_branches
    INNER JOIN clients AS parent_client ON parent_client.id = client_branches.client_id
    LEFT JOIN users AS parent_executive_user ON parent_executive_user.id = parent_client.executive_user_id
    LEFT JOIN users AS supervisor_ifci ON supervisor_ifci.id = client_branches.supervisor_ifci_user_id
    LEFT JOIN users AS supervisor_ext ON supervisor_ext.id = client_branches.supervisor_ext_user_id
    LEFT JOIN users AS supervisor_works ON supervisor_works.id = client_branches.supervisor_works_user_id
    ${whereClause}
    ${orderClause}
  `;
}

function parseBooleanCell(value) {
  const normalized = String(value ?? "").trim().toLowerCase();
  return ["si", "sí", "true", "1", "x"].includes(normalized);
}

async function findUserByEmailOrName(value, expectedRole = null) {
  const normalized = String(value || "").trim();
  if (!normalized) return null;
  const { rows } = await query(
    `
    SELECT id, name, email, role
    FROM users
    WHERE lower(email) = lower($1) OR lower(name) = lower($1)
    ORDER BY id ASC
    LIMIT 1
    `,
    [normalized]
  );
  const user = rows[0] || null;
  if (!user) return null;
  if (expectedRole && user.role !== expectedRole) return null;
  return user;
}

async function createClientRecord(payload) {
  const nextPositionResult = await query("SELECT COALESCE(MAX(position), 0) + 1 AS next FROM clients");
  const nextPosition = nextPositionResult.rows[0].next;
  const {
    name,
    billing,
    sector,
    companyType,
    country,
    accountStage,
    bitrixLeadId,
    bitrixCompanyId,
    executiveUserId,
    risk,
    segment,
    services,
    supervisors,
    notes
  } = payload;

  const normalizedBilling = typeof billing === "number" && !Number.isNaN(billing) ? billing : 0;
  const executiveResult = await query("SELECT name FROM users WHERE id = $1", [Number(executiveUserId)]);
  const executiveName = executiveResult.rows[0]?.name || "";
  const fixedFireSupervisorId = services.fixedFire ? normalizeNullableUserId(supervisors.fixedFire.userId) : null;
  const extinguishersSupervisorId = services.extinguishers ? normalizeNullableUserId(supervisors.extinguishers.userId) : null;
  const worksSupervisorId = services.works ? normalizeNullableUserId(supervisors.works.userId) : null;

  const createdResult = await query(
    `
    INSERT INTO clients (
      position, name, billing_2025, sector, company_type, country, account_stage, manager, executive_user_id, risk, segment,
      bitrix_lead_id, bitrix_company_id, service_fixed_fire, service_extinguishers, service_works,
      wallet_share, nps, open_opportunities, notes,
      supervisor_ifci_user_id, supervisor_ext_user_id, supervisor_works_user_id
    ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, $18, $19, $20, $21, $22, $23)
    RETURNING *
    `,
    [
      nextPosition,
      name.trim(),
      normalizedBilling,
      sector.trim(),
      companyType,
      country,
      accountStage,
      executiveName,
      Number(executiveUserId),
      risk,
      segment,
      String(bitrixLeadId || "").trim(),
      String(bitrixCompanyId || "").trim(),
      services.fixedFire,
      services.extinguishers,
      services.works,
      "",
      0,
      0,
      notes.trim(),
      fixedFireSupervisorId,
      extinguishersSupervisorId,
      worksSupervisorId
    ]
  );

  const selectedCreated = await query(buildClientSelectQuery("WHERE clients.id = $1", ""), [createdResult.rows[0].id]);
  return buildClientResponse(selectedCreated.rows[0]);
}

async function createUserRecord(payload) {
  const { values } = validateUserPayload(payload, { requirePassword: true });
  const result = await query(
    `
    INSERT INTO users (name, email, password_hash, role, bitrix_user_id)
    VALUES ($1, $2, $3, $4, $5)
    RETURNING id, name, email, role, bitrix_user_id, created_at
    `,
    [values.name, values.email, hashPassword(values.password), values.role, values.bitrixUserId]
  );

  return {
    ...sanitizeUser(result.rows[0]),
    createdAt: result.rows[0].created_at
  };
}

async function getUsersByRoles(roles) {
  const { rows } = await query(
    `
    SELECT id, name, email, role
    FROM users
    WHERE role = ANY($1::text[])
    ORDER BY name ASC, id ASC
    `,
    [roles]
  );
  return rows.map(sanitizeUser);
}

async function getSectorOptions() {
  const { rows } = await query(
    `
    SELECT id, name, created_at
    FROM sector_options
    ORDER BY lower(name) ASC, id ASC
    `
  );

  return rows.map((row) => ({
    id: row.id,
    name: row.name,
    createdAt: row.created_at
  }));
}

async function getContactRoleOptionByName(name) {
  const normalizedName = String(name || "").trim();
  if (!normalizedName) return null;

  const { rows } = await query("SELECT id, name FROM contact_role_options WHERE lower(name) = lower($1) LIMIT 1", [normalizedName]);
  return rows[0] || null;
}

async function getUserById(userId) {
  const { rows } = await query(
    `
    SELECT id, name, email, role, bitrix_user_id, created_at
    FROM users
    WHERE id = $1
    `,
    [userId]
  );
  return rows[0] || null;
}

async function getUserAssignments(userId) {
  const { rows } = await query(
    `
    SELECT
      id,
      name,
      'company' AS entity_type,
      executive_user_id,
      supervisor_ifci_user_id,
      supervisor_ext_user_id,
      supervisor_works_user_id
    FROM clients
    WHERE executive_user_id = $1
       OR supervisor_ifci_user_id = $1
       OR supervisor_ext_user_id = $1
       OR supervisor_works_user_id = $1
    UNION ALL
    SELECT
      id,
      name,
      'branch' AS entity_type,
      executive_user_id,
      supervisor_ifci_user_id,
      supervisor_ext_user_id,
      supervisor_works_user_id
    FROM client_branches
    WHERE executive_user_id = $1
       OR supervisor_ifci_user_id = $1
       OR supervisor_ext_user_id = $1
       OR supervisor_works_user_id = $1
    ORDER BY name ASC, id ASC
    `,
    [userId]
  );

  return rows.map((row) => ({
    clientId: row.id,
    clientName: `${row.name}${row.entity_type === "branch" ? " (Sucursal)" : ""}`,
    asExecutive: Number(row.executive_user_id) === Number(userId),
    asSupervisorIfci: Number(row.supervisor_ifci_user_id) === Number(userId),
    asSupervisorExt: Number(row.supervisor_ext_user_id) === Number(userId),
    asSupervisorWorks: Number(row.supervisor_works_user_id) === Number(userId)
  }));
}

async function fetchVisitRules(entityType, entityId, db = { query }) {
  const { rows } = await db.query(
    `
    SELECT id, entity_type, entity_id, periodicity_days, contact_role, objective, created_at, updated_at
    FROM visit_rules
    WHERE entity_type = $1 AND entity_id = $2
    ORDER BY id ASC
    `,
    [entityType, entityId]
  );

  return rows.map((row) => ({
    id: row.id,
    entityType: row.entity_type,
    entityId: row.entity_id,
    periodicityDays: Number(row.periodicity_days || 0),
    contactRole: row.contact_role || "",
    objective: row.objective || "",
    createdAt: row.created_at,
    updatedAt: row.updated_at
  }));
}

async function replaceVisitRules(db, entityType, entityId, rules) {
  await db.query("DELETE FROM visit_rules WHERE entity_type = $1 AND entity_id = $2", [entityType, entityId]);

  for (const rule of rules) {
    await db.query(
      `
      INSERT INTO visit_rules (entity_type, entity_id, periodicity_days, contact_role, objective, created_at, updated_at)
      VALUES ($1, $2, $3, $4, $5, NOW(), NOW())
      `,
      [entityType, entityId, rule.periodicityDays, rule.contactRole, rule.objective]
    );
  }
}

function addDays(dateString, days) {
  const base = dateString ? new Date(`${dateString}T00:00:00`) : new Date();
  if (Number.isNaN(base.getTime())) {
    const fallback = new Date();
    fallback.setDate(fallback.getDate() + days);
    return fallback.toISOString().slice(0, 10);
  }
  base.setDate(base.getDate() + days);
  return base.toISOString().slice(0, 10);
}

async function getAutomaticMeetingEntityContext(db, entityType, entityId) {
  if (entityType === "client") {
    const { rows } = await db.query(
      `
      SELECT id, executive_user_id
      FROM clients
      WHERE id = $1 AND is_hidden = FALSE
      `,
      [entityId]
    );
    const client = rows[0];
    if (!client) return null;
    return {
      clientId: client.id,
      branchId: null,
      executiveUserId: client.executive_user_id || null
    };
  }

  if (entityType === "branch") {
    const { rows } = await db.query(
      `
      SELECT client_branches.id, client_branches.client_id, clients.executive_user_id
      FROM client_branches
      INNER JOIN clients ON clients.id = client_branches.client_id
      WHERE client_branches.id = $1 AND clients.is_hidden = FALSE
      `,
      [entityId]
    );
    const branch = rows[0];
    if (!branch) return null;
    return {
      clientId: branch.client_id,
      branchId: branch.id,
      executiveUserId: branch.executive_user_id || null
    };
  }

  return null;
}

async function createAutomaticMeetingForRule(db, rule, baseDate) {
  const context = await getAutomaticMeetingEntityContext(db, rule.entityType, rule.entityId);
  if (!context?.clientId || !context.executiveUserId) return null;

  const participantNames = await resolveParticipantNames([context.executiveUserId]);
  const scheduledFor = addDays(baseDate, Number(rule.periodicityDays || 0));

  const result = await db.query(
    `
    INSERT INTO meetings (
      client_id, branch_id, opportunity_id, kind, subject, objective, scheduled_for, participants, participant_user_ids,
      contact_name, contact_role, modality, next_meeting_date, minutes, findings,
      active_negotiations_status, opportunities, substitute_recovery, global_contacts, service_health_status, service_status,
      created_by, status, follow_up_from_meeting_id, is_automatic, automatic_rule_id, automatic_rule_entity_type, automatic_rule_entity_id,
      created_at, updated_at
    ) VALUES (
      $1, $2, NULL, $3, $4, $5, $6, $7, $8,
      $9, $10, $11, '', '', '',
      '', '', '', '', '', '',
      $12, $13, NULL, TRUE, $14, $15, $16,
      NOW(), NOW()
    )
    RETURNING id
    `,
    [
      context.clientId,
      context.branchId,
      "Comercial",
      "Acercamiento al cliente",
      rule.objective,
      scheduledFor,
      participantNames.join(", "),
      [context.executiveUserId],
      "",
      rule.contactRole,
      "Presencial",
      "Regla automática",
      MEETING_STATUS.SCHEDULED,
      rule.id,
      rule.entityType,
      rule.entityId
    ]
  );

  return result.rows[0]?.id || null;
}

async function syncAutomaticMeetingsForEntity(db, entityType, entityId) {
  if (!(await isVisitRulesFeatureEnabled())) return;
  const rules = await fetchVisitRules(entityType, entityId, db);
  if (!rules.length) return;

  for (const rule of rules) {
    const existingPending = await db.query(
      `
      SELECT id
      FROM meetings
      WHERE automatic_rule_id = $1
        AND COALESCE(is_deleted, FALSE) = FALSE
        AND status = ANY($2::text[])
      ORDER BY id ASC
      LIMIT 1
      `,
      [rule.id, [MEETING_STATUS.SCHEDULED, MEETING_STATUS.CONFIRMED]]
    );

    if (!existingPending.rows[0]) {
      await createAutomaticMeetingForRule(db, rule, new Date().toISOString().slice(0, 10));
    }
  }
}

async function regenerateAutomaticMeetingIfNeeded(db, meeting) {
  if (!(await isVisitRulesFeatureEnabled())) return;
  if (!meeting?.isAutomatic || meeting.status !== MEETING_STATUS.COMPLETED || !meeting.automaticRuleId) return;

  const ruleResult = await db.query(
    `
    SELECT id, entity_type, entity_id, periodicity_days, contact_role, objective
    FROM visit_rules
    WHERE id = $1
    `,
    [meeting.automaticRuleId]
  );
  const row = ruleResult.rows[0];
  if (!row) return;

  const existingPending = await db.query(
    `
    SELECT id
    FROM meetings
    WHERE automatic_rule_id = $1
      AND COALESCE(is_deleted, FALSE) = FALSE
      AND status = ANY($2::text[])
    ORDER BY id ASC
    LIMIT 1
    `,
    [row.id, [MEETING_STATUS.SCHEDULED, MEETING_STATUS.CONFIRMED]]
  );
  if (existingPending.rows[0]) return;

  await createAutomaticMeetingForRule(
    db,
    {
      id: row.id,
      entityType: row.entity_type,
      entityId: row.entity_id,
      periodicityDays: row.periodicity_days,
      contactRole: row.contact_role,
      objective: row.objective
    },
    new Date().toISOString().slice(0, 10)
  );
}

function validateRoleChangeAgainstAssignments(assignments, nextRole) {
  if (!assignments.length) return null;

  if (assignments.some((assignment) => assignment.asExecutive) && nextRole !== USER_ROLES.EXECUTIVE) {
    const company = assignments.find((assignment) => assignment.asExecutive);
    return `No podés cambiar el rol: el usuario está asignado como Ejecutivo en ${company.clientName}`;
  }
  if (assignments.some((assignment) => assignment.asSupervisorIfci) && nextRole !== USER_ROLES.SUPERVISOR_IFCI) {
    const company = assignments.find((assignment) => assignment.asSupervisorIfci);
    return `No podés cambiar el rol: el usuario está asignado como Supervisor IFCI en ${company.clientName}`;
  }
  if (assignments.some((assignment) => assignment.asSupervisorExt) && nextRole !== USER_ROLES.SUPERVISOR_EXT) {
    const company = assignments.find((assignment) => assignment.asSupervisorExt);
    return `No podés cambiar el rol: el usuario está asignado como Supervisor EXT en ${company.clientName}`;
  }
  if (assignments.some((assignment) => assignment.asSupervisorWorks) && nextRole !== USER_ROLES.SUPERVISOR_WORKS) {
    const company = assignments.find((assignment) => assignment.asSupervisorWorks);
    return `No podés cambiar el rol: el usuario está asignado como Supervisor Obra en ${company.clientName}`;
  }

  return null;
}

async function validateUserAssignment(userId, expectedRole, label) {
  if (userId === null || userId === undefined) return null;
  const numericUserId = Number(userId);
  if (!Number.isInteger(numericUserId) || numericUserId <= 0) {
    return `${label} inválido`;
  }

  const { rows } = await query("SELECT id, role FROM users WHERE id = $1", [numericUserId]);
  const user = rows[0];
  if (!user) return `${label} no encontrado`;
  if (user.role !== expectedRole) return `${label} debe tener el rol ${expectedRole}`;
  return null;
}

function normalizeNullableUserId(value) {
  if (value === null || value === undefined || value === "") return null;
  const numericValue = Number(value);
  if (!Number.isInteger(numericValue) || numericValue <= 0) return NaN;
  return numericValue;
}

async function validateSectorExists(sectorName) {
  const normalizedSector = String(sectorName || "").trim();
  if (!normalizedSector) return "El campo sector es obligatorio";

  const { rows } = await query("SELECT id FROM sector_options WHERE lower(name) = lower($1)", [normalizedSector]);
  if (!rows[0]) {
    return "El sector seleccionado no existe en la configuración";
  }
  return null;
}

async function fetchClientBranches(clientId) {
  const { rows } = await query(buildBranchSelectQuery("WHERE client_branches.client_id = $1"), [clientId]);
  const branches = rows.map(toBranchObject);
  return Promise.all(
    branches.map(async (branch) => ({
      ...branch,
      visitRules: await fetchVisitRules("branch", branch.id)
    }))
  );
}

async function withBillingRank(client) {
  const { rows } = await query(
    `
    SELECT COUNT(*)::int + 1 AS rank
    FROM clients
    WHERE billing_2025 > $1
       OR (billing_2025 = $2 AND id < $3)
    `,
    [client.billing, client.billing, client.id]
  );

  return {
    ...client,
    rank: rows[0]?.rank || 1
  };
}

function toMeetingObject(row) {
  return {
    id: row.id,
    clientId: row.client_id,
    clientName: row.client_name || "",
    branchId: row.branch_id,
    opportunityId: row.opportunity_id || null,
    kind: row.kind,
    subject: row.subject,
    objective: row.objective,
    scheduledFor: row.scheduled_for,
    participants: row.participants,
    participantUserIds: Array.isArray(row.participant_user_ids) ? row.participant_user_ids.map((id) => Number(id)) : [],
    contactName: row.contact_name,
    contactRole: row.contact_role,
    modality: row.modality || "Presencial",
    nextMeetingDate: row.next_meeting_date || "",
    followUpFromMeetingId: row.follow_up_from_meeting_id || null,
    minutes: row.minutes,
    findings: row.findings,
    activeNegotiationsStatus: row.active_negotiations_status || row.findings || "",
    opportunities: row.opportunities || "",
    substituteRecovery: row.substitute_recovery || "",
    globalContacts: row.global_contacts || "",
    serviceHealthStatus: row.service_health_status || "",
    serviceStatus: row.service_status || "",
    complaintBitrixResponsibleId: row.complaint_bitrix_responsible_id || "",
    complaintBitrixTaskId: row.complaint_bitrix_task_id || "",
    isAutomatic: !!row.is_automatic,
    automaticRuleId: row.automatic_rule_id || null,
    automaticRuleEntityType: row.automatic_rule_entity_type || "",
    automaticRuleEntityId: row.automatic_rule_entity_id || null,
    createdBy: row.created_by,
    status: row.status,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
    opportunityTitle: row.opportunity_title || ""
  };
}

function toOpportunityObject(row) {
  return {
    id: row.id,
    clientId: row.client_id,
    clientName: row.client_name || "",
    branchId: row.branch_id,
    branchName: row.branch_name || "",
    title: row.title,
    opportunityType: row.opportunity_type,
    serviceLine: row.service_line,
    status: row.status,
    amount: Number(row.amount || 0),
    probability: Number(row.probability || 0),
    expectedCloseDate: row.expected_close_date || "",
    ownerUserId: row.owner_user_id || null,
    ownerName: row.owner_name || "",
    source: row.source || "",
    description: row.description || "",
    lossReason: row.loss_reason || "",
    createdBy: row.created_by || "",
    createdAt: row.created_at,
    updatedAt: row.updated_at
  };
}

function toFollowUpObject(row) {
  return {
    id: row.id,
    opportunityId: row.opportunity_id,
    type: row.type,
    title: row.title,
    dueDate: row.due_date || "",
    status: row.status,
    assignedUserId: row.assigned_user_id || null,
    assignedUserName: row.assigned_user_name || "",
    notes: row.notes || "",
    completedAt: row.completed_at,
    createdBy: row.created_by || "",
    createdAt: row.created_at,
    updatedAt: row.updated_at
  };
}

function todayKey() {
  return new Date().toISOString().slice(0, 10);
}

function isOpenOpportunityStatus(status) {
  return OPEN_OPPORTUNITY_STATUS_SET.has(status);
}

function classifyFollowUpStatus(followUp) {
  if (!followUp) return FOLLOW_UP_STATUS.PENDING;
  if (followUp.status === FOLLOW_UP_STATUS.DONE) return FOLLOW_UP_STATUS.DONE;
  if (followUp.dueDate && followUp.dueDate < todayKey()) return "Vencido";
  return FOLLOW_UP_STATUS.PENDING;
}

function enrichOpportunity(opportunity, followUps = []) {
  const sortedFollowUps = [...followUps].sort((left, right) => {
    const leftPriority = classifyFollowUpStatus(left) === "Vencido" ? 0 : left.status === FOLLOW_UP_STATUS.PENDING ? 1 : 2;
    const rightPriority = classifyFollowUpStatus(right) === "Vencido" ? 0 : right.status === FOLLOW_UP_STATUS.PENDING ? 1 : 2;
    if (leftPriority !== rightPriority) return leftPriority - rightPriority;
    return String(left.dueDate || "").localeCompare(String(right.dueDate || ""));
  });

  const pendingFollowUps = sortedFollowUps.filter((followUp) => followUp.status === FOLLOW_UP_STATUS.PENDING);
  const overdueFollowUps = pendingFollowUps.filter((followUp) => followUp.dueDate && followUp.dueDate < todayKey());
  const nextFollowUp = sortedFollowUps.find((followUp) => followUp.status === FOLLOW_UP_STATUS.PENDING) || null;

  return {
    ...opportunity,
    followUps: sortedFollowUps.map((followUp) => ({
      ...followUp,
      visualStatus: classifyFollowUpStatus(followUp)
    })),
    pendingFollowUps: pendingFollowUps.length,
    overdueFollowUps: overdueFollowUps.length,
    nextFollowUp: nextFollowUp
      ? {
          ...nextFollowUp,
          visualStatus: classifyFollowUpStatus(nextFollowUp)
        }
      : null,
    weightedAmount: Math.round(opportunity.amount * (opportunity.probability / 100))
  };
}

function buildCrmSummary(opportunities) {
  const openOpportunities = opportunities.filter((opportunity) => isOpenOpportunityStatus(opportunity.status));
  const overdueFollowUps = openOpportunities.reduce((total, opportunity) => total + Number(opportunity.overdueFollowUps || 0), 0);
  const dueThisWeek = openOpportunities.reduce((total, opportunity) => {
    const nextFollowUp = opportunity.nextFollowUp;
    if (!nextFollowUp || nextFollowUp.status !== FOLLOW_UP_STATUS.PENDING || !nextFollowUp.dueDate) return total;
    const delta = Math.floor((new Date(`${nextFollowUp.dueDate}T00:00:00`) - new Date(`${todayKey()}T00:00:00`)) / 86400000);
    return delta >= 0 && delta <= 7 ? total + 1 : total;
  }, 0);

  return {
    openCount: openOpportunities.length,
    totalAmount: openOpportunities.reduce((total, opportunity) => total + opportunity.amount, 0),
    weightedAmount: openOpportunities.reduce((total, opportunity) => total + opportunity.weightedAmount, 0),
    overdueFollowUps,
    dueThisWeek
  };
}

async function attachFollowUpsToOpportunities(opportunities) {
  if (!opportunities.length) return [];

  const opportunityIds = opportunities.map((opportunity) => opportunity.id);
  const { rows } = await query(
    `
    SELECT
      opportunity_follow_ups.*,
      assigned_user.name AS assigned_user_name
    FROM opportunity_follow_ups
    LEFT JOIN users AS assigned_user ON assigned_user.id = opportunity_follow_ups.assigned_user_id
    WHERE opportunity_follow_ups.opportunity_id = ANY($1::INTEGER[])
    ORDER BY opportunity_follow_ups.due_date ASC, opportunity_follow_ups.id ASC
    `,
    [opportunityIds]
  );

  const followUpsByOpportunity = rows.reduce((accumulator, row) => {
    const followUp = toFollowUpObject(row);
    if (!accumulator.has(followUp.opportunityId)) {
      accumulator.set(followUp.opportunityId, []);
    }
    accumulator.get(followUp.opportunityId).push(followUp);
    return accumulator;
  }, new Map());

  return opportunities.map((opportunity) => enrichOpportunity(opportunity, followUpsByOpportunity.get(opportunity.id) || []));
}

async function fetchOpportunities({
  clientId = null,
  search = "",
  ownerUserId = null,
  status = null,
  executiveScopedUserId = null
} = {}) {
  const where = ["clients.is_hidden = FALSE"];
  const params = [];

  if (clientId !== null) {
    params.push(Number(clientId));
    where.push(`opportunities.client_id = $${params.length}`);
  }

  if (search.trim()) {
    params.push(`%${search.trim()}%`);
    where.push(
      `(opportunities.title ILIKE $${params.length} OR COALESCE(opportunities.source, '') ILIKE $${params.length} OR clients.name ILIKE $${params.length})`
    );
  }

  if (ownerUserId !== null) {
    params.push(Number(ownerUserId));
    where.push(`opportunities.owner_user_id = $${params.length}`);
  }

  if (status !== null) {
    params.push(status);
    where.push(`opportunities.status = $${params.length}`);
  }

  if (executiveScopedUserId !== null) {
    params.push(Number(executiveScopedUserId));
    where.push(`clients.executive_user_id = $${params.length}`);
  }

  const { rows } = await query(
    `
    SELECT
      opportunities.*,
      clients.name AS client_name,
      client_branches.name AS branch_name,
      owner_user.name AS owner_name
    FROM opportunities
    INNER JOIN clients ON clients.id = opportunities.client_id
    LEFT JOIN client_branches ON client_branches.id = opportunities.branch_id
    LEFT JOIN users AS owner_user ON owner_user.id = opportunities.owner_user_id
    WHERE ${where.join(" AND ")}
    ORDER BY
      CASE opportunities.status
        ${OPPORTUNITY_STATUS_OPTIONS.map((currentStatus, index) => `WHEN '${currentStatus}' THEN ${index + 1}`).join(" ")}
        ELSE 99
      END ASC,
      opportunities.expected_close_date ASC,
      opportunities.updated_at DESC,
      opportunities.id DESC
    `,
    params
  );

  return attachFollowUpsToOpportunities(rows.map(toOpportunityObject));
}

function normalizeParticipantUserIds(values) {
  return [...new Set((values || []).map((value) => Number(value)).filter((value) => Number.isInteger(value) && value > 0))];
}

async function resolveParticipantNames(participantUserIds) {
  const normalizedIds = normalizeParticipantUserIds(participantUserIds);
  if (!normalizedIds.length) return [];
  const { rows } = await query(
    `
    SELECT id, name
    FROM users
    WHERE id = ANY($1::INTEGER[])
    ORDER BY name ASC, id ASC
    `,
    [normalizedIds]
  );
  return rows.map((row) => row.name);
}

function getMeetingColor(kind) {
  return getMeetingTypes().find((type) => type.value === kind)?.color || "yellow";
}

function getMeetingLabel(kind) {
  return getMeetingTypes().find((type) => type.value === kind)?.label || kind;
}

function buildVisitsWhereClause(filters = {}) {
  const {
    search = "",
    status = "todos",
    kind = "todos",
    modality = "todos",
    executiveUserId = "todos",
    supervisorUserId = "todos",
    participantUserId = "todos",
    dateFrom = "",
    dateTo = "",
    scopedExecutiveUserId = null
  } = filters;

  const where = ["COALESCE(meetings.is_deleted, FALSE) = FALSE"];
  const params = [];

  if (search) {
    params.push(`%${String(search).trim()}%`);
    where.push(
      `(clients.name ILIKE $${params.length} OR COALESCE(client_branches.name, '') ILIKE $${params.length} OR COALESCE(meetings.contact_name, '') ILIKE $${params.length} OR COALESCE(meetings.participants, '') ILIKE $${params.length})`
    );
  }

  if (status !== "todos") {
    params.push(status);
    where.push(`meetings.status = $${params.length}`);
  }

  if (kind !== "todos") {
    params.push(kind);
    where.push(`meetings.kind = $${params.length}`);
  }

  if (modality !== "todos") {
    params.push(modality);
    where.push(`meetings.modality = $${params.length}`);
  }

  if (executiveUserId !== "todos") {
    params.push(Number(executiveUserId));
    where.push(`clients.executive_user_id = $${params.length}`);
  }

  if (scopedExecutiveUserId !== null) {
    params.push(Number(scopedExecutiveUserId));
    where.push(`clients.executive_user_id = $${params.length}`);
  }

  if (supervisorUserId !== "todos") {
    params.push(Number(supervisorUserId));
    where.push(`(
      COALESCE(client_branches.supervisor_ifci_user_id, clients.supervisor_ifci_user_id) = $${params.length}
      OR COALESCE(client_branches.supervisor_ext_user_id, clients.supervisor_ext_user_id) = $${params.length}
      OR COALESCE(client_branches.supervisor_works_user_id, clients.supervisor_works_user_id) = $${params.length}
    )`);
  }

  if (participantUserId !== "todos") {
    params.push(Number(participantUserId));
    where.push(`meetings.participant_user_ids @> ARRAY[$${params.length}]::INTEGER[]`);
  }

  if (dateFrom) {
    params.push(dateFrom);
    where.push(`meetings.scheduled_for >= $${params.length}`);
  }

  if (dateTo) {
    params.push(dateTo);
    where.push(`meetings.scheduled_for <= $${params.length}`);
  }

  return { where, params };
}

async function fetchVisits(filters = {}) {
  const { where, params } = buildVisitsWhereClause(filters);
  const { rows } = await query(
    `
    SELECT
      meetings.id,
      meetings.client_id,
      meetings.branch_id,
      meetings.opportunity_id,
      meetings.kind,
      meetings.subject,
      meetings.objective,
      meetings.scheduled_for,
      meetings.participants,
      meetings.participant_user_ids,
      meetings.contact_name,
      meetings.contact_role,
      meetings.modality,
      meetings.next_meeting_date,
      meetings.follow_up_from_meeting_id,
      meetings.minutes,
      meetings.findings,
      meetings.created_by,
      meetings.status,
      meetings.created_at,
      meetings.updated_at,
      clients.name AS client_name,
      clients.executive_user_id AS executive_user_id,
      executive_user.name AS executive_name,
      client_branches.name AS branch_name,
      COALESCE(branch_supervisor_ifci.name, client_supervisor_ifci.name) AS supervisor_ifci_name,
      COALESCE(branch_supervisor_ext.name, client_supervisor_ext.name) AS supervisor_ext_name,
      COALESCE(branch_supervisor_works.name, client_supervisor_works.name) AS supervisor_works_name
    FROM meetings
    INNER JOIN clients ON clients.id = meetings.client_id
    LEFT JOIN users AS executive_user ON executive_user.id = clients.executive_user_id
    LEFT JOIN client_branches ON client_branches.id = meetings.branch_id
    LEFT JOIN users AS client_supervisor_ifci ON client_supervisor_ifci.id = clients.supervisor_ifci_user_id
    LEFT JOIN users AS client_supervisor_ext ON client_supervisor_ext.id = clients.supervisor_ext_user_id
    LEFT JOIN users AS client_supervisor_works ON client_supervisor_works.id = clients.supervisor_works_user_id
    LEFT JOIN users AS branch_supervisor_ifci ON branch_supervisor_ifci.id = client_branches.supervisor_ifci_user_id
    LEFT JOIN users AS branch_supervisor_ext ON branch_supervisor_ext.id = client_branches.supervisor_ext_user_id
    LEFT JOIN users AS branch_supervisor_works ON branch_supervisor_works.id = client_branches.supervisor_works_user_id
    ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
    ORDER BY meetings.scheduled_for DESC, meetings.id DESC
    `,
    params
  );

  return rows.map((row) => ({
    ...toMeetingObject(row),
    clientName: row.client_name,
    branchName: row.branch_name,
    scopeLabel: row.branch_name ? `${row.client_name} · ${row.branch_name}` : `${row.client_name} · Casa matriz`,
    kindLabel: getMeetingLabel(row.kind),
    color: getMeetingColor(row.kind),
    executiveName: row.executive_name || "",
    supervisors: [row.supervisor_ifci_name, row.supervisor_ext_name, row.supervisor_works_name].filter(Boolean)
  }));
}

function isUpcomingMeetingStatus(status) {
  return status === MEETING_STATUS.SCHEDULED || status === MEETING_STATUS.CONFIRMED;
}

function buildMeetingDots(meetings) {
  return meetings
    .filter((meeting) => meeting.status === MEETING_STATUS.COMPLETED)
    .slice(0, 8)
    .map((meeting) => ({
      kind: meeting.kind,
      color: getMeetingColor(meeting.kind)
    }));
}

function buildMeetingSummary(meetings) {
  const upcoming = meetings
    .filter((meeting) => meeting.scheduledFor && isUpcomingMeetingStatus(meeting.status))
    .sort((a, b) => a.scheduledFor.localeCompare(b.scheduledFor));

  const latestCompleted = meetings
    .filter((meeting) => meeting.scheduledFor && meeting.status === MEETING_STATUS.COMPLETED)
    .sort((a, b) => b.scheduledFor.localeCompare(a.scheduledFor));

  if (upcoming.length) {
    return `${upcoming[0].subject} · ${upcoming[0].scheduledFor}`;
  }

  if (latestCompleted.length) {
    return `Última: ${latestCompleted[0].scheduledFor}`;
  }

  return "Sin reunión agendada";
}

async function attachMeetingSummary(client) {
  const { rows } = await query(
    `
    SELECT meetings.id, meetings.client_id, meetings.branch_id, meetings.opportunity_id, meetings.kind, meetings.subject, meetings.objective, meetings.scheduled_for, meetings.participants, meetings.participant_user_ids, meetings.contact_name, meetings.contact_role,
           meetings.modality, meetings.next_meeting_date, meetings.follow_up_from_meeting_id, meetings.minutes, meetings.findings,
           meetings.active_negotiations_status, meetings.opportunities, meetings.substitute_recovery, meetings.global_contacts, meetings.service_health_status, meetings.service_status,
           meetings.complaint_bitrix_responsible_id, meetings.complaint_bitrix_task_id,
           meetings.created_by, meetings.status, meetings.created_at, meetings.updated_at,
           opportunities_table.title AS opportunity_title
    FROM meetings
    LEFT JOIN opportunities AS opportunities_table ON opportunities_table.id = meetings.opportunity_id
    WHERE meetings.client_id = $1
      AND COALESCE(meetings.is_deleted, FALSE) = FALSE
    ORDER BY meetings.scheduled_for DESC, meetings.id DESC
    `,
    [client.id]
  );

  const meetings = rows.map(toMeetingObject);

  return {
    ...client,
    meetings,
    nextMeeting: buildMeetingSummary(meetings),
    meetingDots: buildMeetingDots(meetings),
    meetingsCount: meetings.length,
    meetingsCompleted: meetings.filter((meeting) => meeting.status === MEETING_STATUS.COMPLETED).length,
    upcomingMeetings: meetings.filter((meeting) => isUpcomingMeetingStatus(meeting.status)).length
  };
}

async function buildClientResponse(row, { includeCrm = true } = {}) {
  const ranked = await withBillingRank(toClientObject(row));
  const clientWithMeetings = await attachMeetingSummary(ranked);
  const clientWithRules = {
    ...clientWithMeetings,
    visitRules: await fetchVisitRules("client", clientWithMeetings.id)
  };
  const branches = await fetchClientBranches(clientWithMeetings.id);

  if (!includeCrm) {
    return {
      ...clientWithRules,
      branches
    };
  }

  const opportunities = await fetchOpportunities({ clientId: clientWithMeetings.id });
  const crmSummary = buildCrmSummary(opportunities);
  return {
    ...clientWithRules,
    branches,
    opportunities,
    crmSummary,
    openOpportunities: crmSummary.openCount
  };
}

function mergeWhereClause(baseCondition, whereClause = "") {
  if (!whereClause) return `WHERE ${baseCondition}`;
  if (/^where\s+/i.test(whereClause.trim())) {
    return `${whereClause} AND ${baseCondition}`;
  }
  return `WHERE ${baseCondition} AND (${whereClause})`;
}

async function validateAccountPayload(payload, { requireCommercialMetrics = false, requireBilling = true, requireExecutive = true } = {}) {
  const errors = [];
  const {
    name,
    billing,
    sector,
    companyType,
    country,
    accountStage,
    executiveUserId,
    risk,
    segment,
    services,
    supervisors,
    notes
  } = payload || {};

  if (typeof name !== "string" || !name.trim()) errors.push("El campo name es obligatorio");
  if (requireBilling && (typeof billing !== "number" || Number.isNaN(billing) || billing < 0)) {
    errors.push("El campo billing debe ser un número mayor o igual a 0");
  }
  const sectorError = await validateSectorExists(sector);
  if (sectorError) errors.push(sectorError);
  if (requireExecutive && !VALID_COMPANY_TYPES.has(companyType)) errors.push("El tipo de compañía es inválido");
  if (requireExecutive && !VALID_COUNTRIES.has(country)) errors.push("El país de la compañía es inválido");
  if (requireExecutive && !VALID_ACCOUNT_STAGES.has(accountStage)) errors.push("El estado de la compañía es inválido");
  if (!VALID_RISKS.has(risk)) errors.push("El campo risk es inválido");
  if (!VALID_SEGMENTS.has(segment)) errors.push("El campo segment es inválido");

  if (!services || typeof services !== "object") {
    errors.push("El campo services es obligatorio");
  } else {
    const { fixedFire, extinguishers, works } = services;
    if (typeof fixedFire !== "boolean" || typeof extinguishers !== "boolean" || typeof works !== "boolean") {
      errors.push("Los servicios deben ser booleanos");
    }
  }

  if (requireCommercialMetrics) {
    const walletShare = String(payload?.walletShare || "").trim();
    const nps = Number(payload?.nps);
    const openOpportunities = Number(payload?.openOpportunities);
    if (!walletShare) errors.push("El campo walletShare es obligatorio");
    if (!Number.isInteger(nps) || nps < 0 || nps > 100) errors.push("El campo nps debe ser un entero entre 0 y 100");
    if (!Number.isInteger(openOpportunities) || openOpportunities < 0) {
      errors.push("El campo openOpportunities debe ser un entero mayor o igual a 0");
    }
  }
  if (typeof notes !== "string") errors.push("El campo notes debe ser texto");
  if (payload?.bitrixLeadId !== undefined && typeof payload.bitrixLeadId !== "string") {
    errors.push("El campo bitrixLeadId debe ser texto");
  }
  if (payload?.bitrixCompanyId !== undefined && typeof payload.bitrixCompanyId !== "string") {
    errors.push("El campo bitrixCompanyId debe ser texto");
  }

  if (requireExecutive) {
    const executiveError = await validateUserAssignment(executiveUserId, EXECUTIVE_ROLE, "El ejecutivo");
    if (executiveError) errors.push(executiveError);
  }

  if (!supervisors || typeof supervisors !== "object") {
    errors.push("El bloque de supervisores es obligatorio");
  } else {
    for (const [serviceKey, expectedRole] of Object.entries(SUPERVISOR_ROLE_BY_SERVICE)) {
      if (services?.[serviceKey]) {
        const normalizedSupervisorId = normalizeNullableUserId(supervisors[serviceKey]?.userId);
        if (normalizedSupervisorId === null) {
          errors.push(
            `Debes asignar un supervisor de ${serviceKey === "fixedFire" ? "IFCI" : serviceKey === "extinguishers" ? "EXT" : "Obra"}`
          );
          continue;
        }
        if (Number.isNaN(normalizedSupervisorId)) {
          errors.push(
            `El supervisor de ${serviceKey === "fixedFire" ? "IFCI" : serviceKey === "extinguishers" ? "EXT" : "Obra"} es inválido`
          );
          continue;
        }
        const error = await validateUserAssignment(
          normalizedSupervisorId,
          expectedRole,
          `El supervisor de ${serviceKey === "fixedFire" ? "IFCI" : serviceKey === "extinguishers" ? "EXT" : "Obra"}`
        );
        if (error) errors.push(error);
      }
    }
  }

  return errors;
}

async function validateClientPayload(payload) {
  return validateAccountPayload(payload, { requireCommercialMetrics: false, requireBilling: false, requireExecutive: true });
}

async function validateBranchPayload(payload) {
  return validateAccountPayload(payload, { requireCommercialMetrics: false, requireBilling: false, requireExecutive: false });
}

async function validateVisitRulesPayload(rules) {
  if (rules === undefined) return [];
  if (!Array.isArray(rules)) return ["Las reglas de visita son inválidas"];

  const errors = [];
  for (const rule of rules) {
    const periodicityDays = Number(rule?.periodicityDays);
    const contactRole = String(rule?.contactRole || "").trim();
    const objective = String(rule?.objective || "").trim();

    if (!Number.isInteger(periodicityDays) || periodicityDays <= 0) {
      errors.push("La periodicidad de visita debe ser un número entero mayor a 0");
      break;
    }
    if (!contactRole) {
      errors.push("Debes indicar a quién ir a ver en cada regla");
      break;
    }
    if (!getContactRoleOptions().some((role) => role.name === contactRole)) {
      errors.push("La función del contacto definida en una regla es inválida");
      break;
    }
    if (!objective) {
      errors.push("Debes indicar el objetivo de cada regla");
      break;
    }
    if (!VALID_VISIT_RULE_OBJECTIVES.has(objective)) {
      errors.push("El objetivo de la regla es inválido");
      break;
    }
  }

  return errors;
}

async function validateParticipantUsers(participantUserIds) {
  const normalizedIds = normalizeParticipantUserIds(participantUserIds);
  if (!normalizedIds.length) {
    return "Debes seleccionar al menos un participante";
  }

  const { rows } = await query(
    `
    SELECT id
    FROM users
    WHERE id = ANY($1::INTEGER[])
    `,
    [normalizedIds]
  );

  if (rows.length !== normalizedIds.length) {
    return "Uno o más participantes no existen en el sistema";
  }

  return null;
}

async function validateAnyUser(userId, label) {
  if (userId === null || userId === undefined || userId === "") return null;
  const numericUserId = Number(userId);
  if (!Number.isInteger(numericUserId) || numericUserId <= 0) {
    return `${label} inválido`;
  }

  const { rows } = await query("SELECT id FROM users WHERE id = $1", [numericUserId]);
  if (!rows[0]) return `${label} no encontrado`;
  return null;
}

async function validateOpportunityPayload(payload, clientId) {
  const errors = [];

  if (typeof payload?.title !== "string" || !payload.title.trim()) {
    errors.push("El nombre de la oportunidad es obligatorio");
  }

  if (!VALID_OPPORTUNITY_TYPES.has(payload?.opportunityType)) {
    errors.push("El tipo de oportunidad es inválido");
  }

  if (!VALID_SERVICE_LINES.has(payload?.serviceLine)) {
    errors.push("La línea de servicio es inválida");
  }

  if (!VALID_OPPORTUNITY_STATUS.has(payload?.status)) {
    errors.push("El estado de la oportunidad es inválido");
  }

  const amount = Number(payload?.amount);
  if (!Number.isFinite(amount) || amount < 0) {
    errors.push("El monto debe ser un número mayor o igual a 0");
  }

  const probability = Number(payload?.probability);
  if (!Number.isInteger(probability) || probability < 0 || probability > 100) {
    errors.push("La probabilidad debe ser un entero entre 0 y 100");
  }

  if (payload?.expectedCloseDate !== undefined && typeof payload.expectedCloseDate !== "string") {
    errors.push("La fecha estimada de cierre es inválida");
  }

  if (payload?.source !== undefined && typeof payload.source !== "string") {
    errors.push("El origen debe ser texto");
  }

  if (payload?.description !== undefined && typeof payload.description !== "string") {
    errors.push("La descripción debe ser texto");
  }

  if (payload?.lossReason !== undefined && typeof payload.lossReason !== "string") {
    errors.push("El motivo de pérdida debe ser texto");
  }

  if (payload?.status === OPPORTUNITY_STATUS.LOST && !String(payload?.lossReason || "").trim()) {
    errors.push("Debes indicar el motivo de pérdida");
  }

  const ownerError = await validateAnyUser(payload?.ownerUserId, "El responsable");
  if (ownerError) errors.push(ownerError);

  const branchId = payload?.branchId ? Number(payload.branchId) : null;
  if (branchId) {
    const branchResult = await query("SELECT id FROM client_branches WHERE id = $1 AND client_id = $2", [branchId, clientId]);
    if (!branchResult.rows[0]) {
      errors.push("La sucursal seleccionada no pertenece a esta compañía");
    }
  }

  return errors;
}

async function validateFollowUpPayload(payload) {
  const errors = [];

  if (typeof payload?.title !== "string" || !payload.title.trim()) {
    errors.push("El título del seguimiento es obligatorio");
  }

  if (typeof payload?.dueDate !== "string" || !payload.dueDate.trim()) {
    errors.push("La fecha del seguimiento es obligatoria");
  }

  if (!VALID_FOLLOW_UP_TYPES.has(payload?.type)) {
    errors.push("El tipo de seguimiento es inválido");
  }

  if (!VALID_FOLLOW_UP_STATUS.has(payload?.status)) {
    errors.push("El estado del seguimiento es inválido");
  }

  if (payload?.notes !== undefined && typeof payload.notes !== "string") {
    errors.push("Los comentarios del seguimiento deben ser texto");
  }

  const assignedError = await validateAnyUser(payload?.assignedUserId, "El usuario asignado");
  if (assignedError) errors.push(assignedError);

  return errors;
}

function normalizeMeetingStatus(payload) {
  const hasContent = [payload.minutes, payload.findings].some((value) => String(value || "").trim());
  if (payload.status && VALID_MEETING_STATUS.has(payload.status)) return payload.status;
  const today = new Date().toISOString().slice(0, 10);
  if (payload.scheduledFor && payload.scheduledFor > today) return MEETING_STATUS.SCHEDULED;
  return hasContent ? MEETING_STATUS.COMPLETED : MEETING_STATUS.SCHEDULED;
}

async function validateMeetingPayload(payload) {
  const errors = [];

  if (typeof payload?.subject !== "string" || !payload.subject.trim()) {
    errors.push("El motivo de la reunión es obligatorio");
  }
  if (typeof payload?.subject === "string" && payload.subject.trim() && !isValidMeetingReason(payload.subject.trim())) {
    errors.push("El motivo de la reunión es inválido");
  }
  if (typeof payload?.objective !== "string" || !payload.objective.trim()) {
    errors.push("El objetivo de la reunión es obligatorio");
  }
  if (typeof payload?.scheduledFor !== "string" || !payload.scheduledFor.trim()) {
    errors.push("La fecha de la reunión es obligatoria");
  }
  const participantError = await validateParticipantUsers(payload?.participantUserIds);
  if (participantError) errors.push(participantError);
  if (typeof payload?.contactName !== "string" || !payload.contactName.trim()) {
    errors.push("Debes indicar con quién nos vamos a reunir");
  }
  if (typeof payload?.contactRole !== "string" || !payload.contactRole.trim()) {
    errors.push("Debes indicar la función del contacto");
  }
  if (!VALID_MEETING_MODALITIES.has(payload?.modality)) {
    errors.push("La modalidad de la reunión es inválida");
  }
  if (!isValidMeetingType(payload?.kind)) {
    errors.push("El tipo de reunión es inválido");
  }
  if (payload?.createdBy !== undefined && typeof payload.createdBy !== "string") {
    errors.push("El campo cargado por debe ser texto");
  }
  if (payload?.minutes !== undefined && typeof payload.minutes !== "string") {
    errors.push("La minuta debe ser texto");
  }
  if (payload?.findings !== undefined && typeof payload.findings !== "string") {
    errors.push("Los hallazgos deben ser texto");
  }
  for (const field of ["activeNegotiationsStatus", "opportunities", "substituteRecovery", "globalContacts", "serviceHealthStatus", "serviceStatus"]) {
    if (payload?.[field] !== undefined && typeof payload[field] !== "string") {
      errors.push(`El campo ${field} debe ser texto`);
    }
  }
  if (payload?.complaintBitrixResponsibleId !== undefined && typeof payload.complaintBitrixResponsibleId !== "string") {
    errors.push("El responsable Bitrix del reclamo es inválido");
  }
  if (
    typeof payload?.complaintBitrixResponsibleId === "string" &&
    payload.complaintBitrixResponsibleId.trim() &&
    !VALID_BITRIX_COMPLAINT_RESPONSIBLE_IDS.has(payload.complaintBitrixResponsibleId.trim())
  ) {
    errors.push("El responsable Bitrix del reclamo es inválido");
  }
  if (String(payload?.serviceStatus || "").trim() && !String(payload?.complaintBitrixResponsibleId || "").trim()) {
    errors.push("Debes seleccionar un responsable Bitrix para el reclamo");
  }
  if (payload?.status && !VALID_MEETING_STATUS.has(payload.status)) {
    errors.push("El estado de la reunión es inválido");
  }
  if (payload?.nextMeetingDate !== undefined && typeof payload.nextMeetingDate !== "string") {
    errors.push("La fecha de próxima reunión es inválida");
  }
  if (typeof payload?.scheduledFor === "string" && typeof payload?.nextMeetingDate === "string") {
    const nextMeetingDate = payload.nextMeetingDate.trim();
    if (nextMeetingDate && nextMeetingDate < payload.scheduledFor.trim()) {
      errors.push("La próxima reunión no puede ser anterior a la reunión actual");
    }
  }

  return errors;
}

async function upsertMeetingFollowUp(client, meeting, createdBy) {
  const nextMeetingDate = String(meeting.nextMeetingDate || "").trim();
  const followUpResult = await client.query(
    "SELECT id FROM meetings WHERE follow_up_from_meeting_id = $1 AND COALESCE(is_deleted, FALSE) = FALSE ORDER BY id ASC LIMIT 1",
    [meeting.id]
  );
  const existingFollowUp = followUpResult.rows[0] || null;

  if (meeting.status !== MEETING_STATUS.COMPLETED || !nextMeetingDate) {
    if (existingFollowUp) {
      await client.query("DELETE FROM meetings WHERE id = $1", [existingFollowUp.id]);
    }
    return;
  }

  const followUpValues = [
    meeting.clientId,
    meeting.branchId,
    meeting.opportunityId,
    meeting.kind,
    meeting.subject,
    meeting.objective,
    nextMeetingDate,
    meeting.participants,
    meeting.participantUserIds,
    meeting.contactName,
    meeting.contactRole,
    meeting.modality,
    "",
    "",
    meeting.activeNegotiationsStatus || "",
    meeting.opportunities || "",
    meeting.substituteRecovery || "",
    meeting.globalContacts || "",
    meeting.serviceHealthStatus || "",
    meeting.serviceStatus || "",
    meeting.complaintBitrixResponsibleId || "",
    meeting.complaintBitrixTaskId || "",
    String(createdBy || "").trim(),
    MEETING_STATUS.SCHEDULED,
    meeting.id
  ];

  if (existingFollowUp) {
    await client.query(
      `
      UPDATE meetings
      SET
        client_id = $1,
        branch_id = $2,
        opportunity_id = $3,
        kind = $4,
        subject = $5,
        objective = $6,
        scheduled_for = $7,
        participants = $8,
        participant_user_ids = $9,
        contact_name = $10,
        contact_role = $11,
        modality = $12,
        minutes = $13,
        findings = $14,
        active_negotiations_status = $15,
        opportunities = $16,
        substitute_recovery = $17,
        global_contacts = $18,
        service_health_status = $19,
        service_status = $20,
        complaint_bitrix_responsible_id = $21,
        complaint_bitrix_task_id = $22,
        created_by = $23,
        status = $24,
        updated_at = NOW()
      WHERE id = $25
      `,
      [...followUpValues.slice(0, 24), existingFollowUp.id]
    );
    return;
  }

  await client.query(
    `
    INSERT INTO meetings (
      client_id, branch_id, opportunity_id, kind, subject, objective, scheduled_for, participants, participant_user_ids,
      contact_name, contact_role, modality, minutes, findings,
      active_negotiations_status, opportunities, substitute_recovery, global_contacts, service_health_status, service_status,
      complaint_bitrix_responsible_id, complaint_bitrix_task_id, created_by, status, follow_up_from_meeting_id, created_at, updated_at
    ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, $18, $19, $20, $21, $22, $23, $24, $25, NOW(), NOW())
    `,
    followUpValues
  );
}

async function syncComplaintBitrixTask({ clientId, scheduledFor, serviceStatus, complaintBitrixResponsibleId, complaintBitrixTaskId }) {
  const complaintText = String(serviceStatus || "").trim();
  const responsibleId = String(complaintBitrixResponsibleId || "").trim();
  const existingTaskId = String(complaintBitrixTaskId || "").trim();

  if (!complaintText) {
    return { taskId: existingTaskId };
  }

  if (!responsibleId) {
    throw new Error("Debes seleccionar un responsable Bitrix para el reclamo");
  }

  const webhookUrl = await getAppSetting("bitrix_webhook_url", "");
  if (!webhookUrl) {
    throw new Error("No hay un webhook de Bitrix configurado para generar reclamos");
  }

  const bitrixRoot = normalizeBitrixWebhookRoot(webhookUrl);
  const clientResult = await query(
    `
    SELECT
      clients.name,
      clients.executive_user_id,
      executive_user.bitrix_user_id AS executive_bitrix_user_id
    FROM clients
    LEFT JOIN users AS executive_user ON executive_user.id = clients.executive_user_id
    WHERE clients.id = $1
    LIMIT 1
    `,
    [clientId]
  );
  const client = clientResult.rows[0];
  if (!client) {
    throw new Error("No se pudo identificar la compañía para generar el reclamo en Bitrix");
  }

  const executiveBitrixUserId = String(client.executive_bitrix_user_id || "").trim();

  const taskPayload = {
    title: buildComplaintBitrixTaskTitle(client.name, scheduledFor),
    description: complaintText,
    responsibleId,
    groupId: BITRIX_COMPLAINT_GROUP_ID,
    deadline: formatBitrixDeadline(7),
    createdById: executiveBitrixUserId
  };

  if (existingTaskId) {
    await updateBitrixTask(bitrixRoot, {
      taskId: existingTaskId,
      ...taskPayload
    });
    return { taskId: existingTaskId };
  }

  const createdTask = await createBitrixTask(bitrixRoot, taskPayload);
  if (!createdTask.id) {
    throw new Error("Bitrix no devolvió el identificador de la tarea de reclamo");
  }

  return { taskId: String(createdTask.id) };
}

app.get("/api/health", (_req, res) => {
  res.json({ ok: true });
});

app.get(
  "/api/sector-options",
  asyncHandler(async (_req, res) => {
    res.json({ sectors: await getSectorOptions() });
  })
);

app.get(
  "/api/auth/me",
  asyncHandler(async (req, res) => {
    if (!req.session.userId) {
      res.status(401).json({ error: "Sesión requerida" });
      return;
    }

    const { rows } = await query("SELECT id, name, email, role FROM users WHERE id = $1", [req.session.userId]);
    const user = rows[0];

    if (!user) {
      req.session.destroy(() => {});
      res.status(401).json({ error: "Sesión inválida" });
      return;
    }

    const visitRulesEnabled = await isVisitRulesFeatureEnabled();

    res.json({
      user: {
        ...sanitizeUser(user),
        canManageSettings: isSettingsAdminUser(user)
      },
      featureFlags: {
        visitRulesEnabled
      }
    });
  })
);

app.post(
  "/api/auth/login",
  asyncHandler(async (req, res) => {
    const email = String(req.body?.email || "").trim().toLowerCase();
    const password = String(req.body?.password || "");
    const rememberMe = Boolean(req.body?.rememberMe);

    if (!email || !password) {
      res.status(400).json({ error: "Email y contraseña son obligatorios" });
      return;
    }

    const { rows } = await query("SELECT * FROM users WHERE lower(email) = $1", [email]);
    const user = rows[0];

    if (!user || user.password_hash !== hashPassword(password)) {
      res.status(401).json({ error: "Credenciales inválidas" });
      return;
    }

    req.session.userId = user.id;
    if (rememberMe) {
      req.session.cookie.maxAge = REMEMBER_ME_COOKIE_MAX_AGE_MS;
    } else {
      req.session.cookie.expires = false;
      req.session.cookie.maxAge = null;
    }
    await logAuditEvent({
      user,
      action: "auth.login",
      entityType: "session",
      entityId: user.id,
      entityName: user.email
    });
    const visitRulesEnabled = await isVisitRulesFeatureEnabled();

    res.json({
      user: {
        ...sanitizeUser(user),
        canManageSettings: isSettingsAdminUser(user)
      },
      featureFlags: {
        visitRulesEnabled
      }
    });
  })
);

app.post("/api/auth/logout", (req, res) => {
  getSessionUser(req)
    .then((user) =>
      logAuditEvent({
        user,
        action: "auth.logout",
        entityType: "session",
        entityId: user?.id || null,
        entityName: user?.email || ""
      })
    )
    .catch(() => {})
    .finally(() => {
      req.session.destroy(() => {
        res.json({ ok: true });
      });
    });
});

app.get("/api/meeting-types", (_req, res) => {
  res.json({
    meetingTypes: getMeetingTypes(),
    meetingReasons: getMeetingReasons(),
    contactRoles: getContactRoleOptions(),
    statuses: Object.values(MEETING_STATUS),
    modalities: MEETING_MODALITIES
  });
});

app.use("/api", requireAuth);
app.use("/api/settings", requireSettingsAdmin);

app.get(
  "/api/bitrix/options",
  asyncHandler(async (_req, res) => {
    const directory = await getBitrixDirectoryOptions();
    res.json(directory);
  })
);

app.get(
  "/api/bitrix/companies-search",
  asyncHandler(async (req, res) => {
    const queryText = String(req.query?.q || "").trim();
    if (!queryText) {
      res.json({ companies: [] });
      return;
    }

    const webhookUrl = await getAppSetting("bitrix_webhook_url", "");
    if (!webhookUrl) {
      res.json({
        companies: [],
        permissions: { companies: false },
        errors: { companies: "No hay webhook de Bitrix configurado." }
      });
      return;
    }

    const bitrixRoot = normalizeBitrixWebhookRoot(webhookUrl);
    try {
      const companies = await searchBitrixCompanies(bitrixRoot, queryText);
      res.json({
        companies: companies.map((company) => ({
          id: String(company.id || ""),
          label: `${company.title}${company.id ? ` · ${company.id}` : ""}`.trim(),
          title: company.title
        })),
        permissions: { companies: true },
        errors: { companies: "" }
      });
    } catch (error) {
      res.json({
        companies: [],
        permissions: { companies: false },
        errors: { companies: formatBitrixDirectoryError(error, "compañías") }
      });
    }
  })
);

app.get(
  "/api/crm/catalogs",
  asyncHandler(async (_req, res) => {
    res.json({
      opportunityStatuses: OPPORTUNITY_STATUS_OPTIONS,
      openOpportunityStatuses: OPPORTUNITY_OPEN_STATUSES,
      opportunityTypes: OPPORTUNITY_TYPES,
      serviceLines: OPPORTUNITY_SERVICE_LINES,
      followUpStatuses: FOLLOW_UP_STATUS_OPTIONS,
      followUpTypes: FOLLOW_UP_TYPES
    });
  })
);

app.get(
  "/api/settings/sectors",
  asyncHandler(async (_req, res) => {
    res.json({ sectors: await getSectorOptions() });
  })
);

app.get(
  "/api/settings/meeting-types",
  asyncHandler(async (_req, res) => {
    res.json({ meetingTypes: getMeetingTypes() });
  })
);

app.get(
  "/api/settings/meeting-reasons",
  asyncHandler(async (_req, res) => {
    res.json({ meetingReasons: getMeetingReasons() });
  })
);

app.get(
  "/api/settings/clients-import-template",
  asyncHandler(async (_req, res) => {
    const rows = [
      {
        name: "Cliente Ejemplo",
        sector: "Energía",
        companyType: "Local",
        country: "Argentina",
        accountStage: "Activa",
        executive: "comercial@maxi.local",
        risk: "Medio",
        segment: "B",
        fixedFire: "No",
        extinguishers: "Si",
        works: "No",
        supervisorIfci: "",
        supervisorExt: "ext@maxi.local",
        supervisorWorks: "",
        notes: "Importado desde Excel"
      }
    ];

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Clientes");
    const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="plantilla-clientes.xlsx"');
    res.send(buffer);
  })
);

app.get(
  "/api/settings/users-import-template",
  asyncHandler(async (_req, res) => {
    const rows = [
      {
        name: "Nuevo Comercial",
        email: "nuevo.comercial@empresa.com",
        role: USER_ROLES.EXECUTIVE,
        password: "Temporal1234"
      }
    ];

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Usuarios");
    const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="plantilla-usuarios.xlsx"');
    res.send(buffer);
  })
);

app.get(
  "/api/settings/branches-import-template",
  asyncHandler(async (_req, res) => {
    const rows = [
      {
        clientName: "Cliente Ejemplo",
        branchName: "Planta Córdoba",
        sector: "Energía",
        risk: "Medio",
        segment: "B",
        fixedFire: "No",
        extinguishers: "Si",
        works: "No",
        supervisorIfci: "",
        supervisorExt: "ext@maxi.local",
        supervisorWorks: "",
        notes: "Sucursal importada desde Excel"
      }
    ];

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sucursales");
    const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="plantilla-sucursales.xlsx"');
    res.send(buffer);
  })
);

app.post(
  "/api/settings/clients-import",
  asyncHandler(async (req, res) => {
    const fileData = String(req.body?.fileData || "").trim();
    if (!fileData) {
      res.status(400).json({ error: "Debes adjuntar un archivo Excel" });
      return;
    }

    const workbook = XLSX.read(Buffer.from(fileData, "base64"), { type: "buffer" });
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
    if (!rows.length) {
      res.status(400).json({ error: "El archivo no tiene filas para importar" });
      return;
    }

    let createdCount = 0;
    const errors = [];

    for (const [index, row] of rows.entries()) {
      const line = index + 2;
      const executive = await findUserByEmailOrName(row.executive, EXECUTIVE_ROLE);
      if (!executive) {
        errors.push(`Fila ${line}: no se encontró un Ejecutivo válido en "executive"`);
        continue;
      }

      const services = {
        fixedFire: parseBooleanCell(row.fixedFire),
        extinguishers: parseBooleanCell(row.extinguishers),
        works: parseBooleanCell(row.works)
      };

      const supervisors = {
        fixedFire: { userId: null },
        extinguishers: { userId: null },
        works: { userId: null }
      };

      if (services.fixedFire) {
        const supervisor = await findUserByEmailOrName(row.supervisorIfci, USER_ROLES.SUPERVISOR_IFCI);
        supervisors.fixedFire.userId = supervisor?.id || null;
      }
      if (services.extinguishers) {
        const supervisor = await findUserByEmailOrName(row.supervisorExt, USER_ROLES.SUPERVISOR_EXT);
        supervisors.extinguishers.userId = supervisor?.id || null;
      }
      if (services.works) {
        const supervisor = await findUserByEmailOrName(row.supervisorWorks, USER_ROLES.SUPERVISOR_WORKS);
        supervisors.works.userId = supervisor?.id || null;
      }

      const payload = {
        name: String(row.name || "").trim(),
        billing: 0,
        sector: String(row.sector || "").trim(),
        companyType: String(row.companyType || "Local").trim() || "Local",
        country: String(row.country || "Argentina").trim() || "Argentina",
        accountStage: String(row.accountStage || "Activa").trim() || "Activa",
        executiveUserId: executive.id,
        risk: String(row.risk || "Bajo").trim() || "Bajo",
        segment: String(row.segment || "C").trim() || "C",
        services,
        supervisors,
        notes: String(row.notes || "").trim()
      };

      const duplicateResult = await query("SELECT id FROM clients WHERE lower(name) = lower($1) LIMIT 1", [payload.name]);
      if (duplicateResult.rows[0]) {
        errors.push(`Fila ${line}: ya existe un cliente con el nombre "${payload.name}"`);
        continue;
      }

      const validationErrors = await validateClientPayload(payload);
      if (validationErrors.length) {
        errors.push(`Fila ${line}: ${validationErrors[0]}`);
        continue;
      }

      await createClientRecord(payload);
      createdCount += 1;
    }

    res.json({
      createdCount,
      errorCount: errors.length,
      errors
    });
  })
);

app.post(
  "/api/settings/branches-import",
  asyncHandler(async (req, res) => {
    const fileData = String(req.body?.fileData || "").trim();
    if (!fileData) {
      res.status(400).json({ error: "Debes adjuntar un archivo Excel" });
      return;
    }

    const workbook = XLSX.read(Buffer.from(fileData, "base64"), { type: "buffer" });
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
    if (!rows.length) {
      res.status(400).json({ error: "El archivo no tiene filas para importar" });
      return;
    }

    let createdCount = 0;
    const errors = [];

    for (const [index, row] of rows.entries()) {
      const line = index + 2;
      const clientName = String(row.clientName || "").trim();
      const branchName = String(row.branchName || "").trim();

      if (!clientName) {
        errors.push(`Fila ${line}: debés indicar la compañía en "clientName"`);
        continue;
      }

      const clientResult = await query(
        `
        SELECT id, manager, billing_2025, executive_user_id
        FROM clients
        WHERE lower(name) = lower($1)
        LIMIT 1
        `,
        [clientName]
      );
      const client = clientResult.rows[0];
      if (!client) {
        errors.push(`Fila ${line}: no existe una compañía llamada "${clientName}"`);
        continue;
      }

      const services = {
        fixedFire: parseBooleanCell(row.fixedFire),
        extinguishers: parseBooleanCell(row.extinguishers),
        works: parseBooleanCell(row.works)
      };

      const supervisors = {
        fixedFire: { userId: null },
        extinguishers: { userId: null },
        works: { userId: null }
      };

      if (services.fixedFire) {
        const supervisor = await findUserByEmailOrName(row.supervisorIfci, USER_ROLES.SUPERVISOR_IFCI);
        supervisors.fixedFire.userId = supervisor?.id || null;
      }
      if (services.extinguishers) {
        const supervisor = await findUserByEmailOrName(row.supervisorExt, USER_ROLES.SUPERVISOR_EXT);
        supervisors.extinguishers.userId = supervisor?.id || null;
      }
      if (services.works) {
        const supervisor = await findUserByEmailOrName(row.supervisorWorks, USER_ROLES.SUPERVISOR_WORKS);
        supervisors.works.userId = supervisor?.id || null;
      }

      const payload = {
        name: branchName,
        sector: String(row.sector || "").trim(),
        risk: String(row.risk || "Bajo").trim() || "Bajo",
        segment: String(row.segment || "C").trim() || "C",
        services,
        supervisors,
        notes: String(row.notes || "").trim()
      };

      const duplicateResult = await query(
        `
        SELECT id
        FROM client_branches
        WHERE client_id = $1 AND lower(name) = lower($2)
        LIMIT 1
        `,
        [client.id, payload.name]
      );
      if (duplicateResult.rows[0]) {
        errors.push(`Fila ${line}: ya existe la sucursal "${payload.name}" para "${clientName}"`);
        continue;
      }

      const validationErrors = await validateBranchPayload(payload);
      if (validationErrors.length) {
        errors.push(`Fila ${line}: ${validationErrors[0]}`);
        continue;
      }

      const fixedFireSupervisorId = services.fixedFire ? normalizeNullableUserId(supervisors.fixedFire.userId) : null;
      const extinguishersSupervisorId = services.extinguishers ? normalizeNullableUserId(supervisors.extinguishers.userId) : null;
      const worksSupervisorId = services.works ? normalizeNullableUserId(supervisors.works.userId) : null;

      const result = await query(
        `
        INSERT INTO client_branches (
          client_id, name, billing_2025, sector, manager, executive_user_id, risk, segment,
          service_fixed_fire, service_extinguishers, service_works, notes,
          supervisor_ifci_user_id, supervisor_ext_user_id, supervisor_works_user_id,
          created_at, updated_at
        ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, NOW(), NOW())
        RETURNING id
        `,
        [
          client.id,
          payload.name.trim(),
          Number(client.billing_2025 || 0),
          payload.sector.trim(),
          String(client.manager || "").trim(),
          client.executive_user_id || null,
          payload.risk,
          payload.segment,
          services.fixedFire,
          services.extinguishers,
          services.works,
          payload.notes.trim(),
          fixedFireSupervisorId,
          extinguishersSupervisorId,
          worksSupervisorId
        ]
      );

      await logAuditEvent({
        req,
        action: "branch.create",
        entityType: "branch",
        entityId: result.rows[0].id,
        entityName: payload.name,
        details: { clientId: client.id, imported: true }
      });
      createdCount += 1;
    }

    res.json({
      createdCount,
      errorCount: errors.length,
      errors
    });
  })
);

app.post(
  "/api/settings/users-import",
  asyncHandler(async (req, res) => {
    const fileData = String(req.body?.fileData || "").trim();
    if (!fileData) {
      res.status(400).json({ error: "Debes adjuntar un archivo Excel" });
      return;
    }

    const workbook = XLSX.read(Buffer.from(fileData, "base64"), { type: "buffer" });
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
    if (!rows.length) {
      res.status(400).json({ error: "El archivo no tiene filas para importar" });
      return;
    }

    let createdCount = 0;
    const errors = [];

    for (const [index, row] of rows.entries()) {
      const line = index + 2;
      const payload = {
        name: String(row.name || "").trim(),
        email: String(row.email || "").trim(),
        role: String(row.role || "").trim(),
        password: String(row.password || "").trim()
      };

      const validation = validateUserPayload(payload, { requirePassword: true });
      if (validation.errors.length) {
        errors.push(`Fila ${line}: ${validation.errors[0]}`);
        continue;
      }

      const existing = await query("SELECT id FROM users WHERE lower(email) = lower($1) LIMIT 1", [payload.email]);
      if (existing.rows[0]) {
        errors.push(`Fila ${line}: ya existe un usuario con el email "${payload.email}"`);
        continue;
      }

      try {
        await createUserRecord(payload);
        createdCount += 1;
      } catch (error) {
        if (error.code === "23505") {
          errors.push(`Fila ${line}: ya existe un usuario con el email "${payload.email}"`);
          continue;
        }
        throw error;
      }
    }

    res.json({
      createdCount,
      errorCount: errors.length,
      errors
    });
  })
);

app.post(
  "/api/settings/meeting-types",
  asyncHandler(async (req, res) => {
    const value = String(req.body?.value || "").trim();
    const label = String(req.body?.label || "").trim();
    const color = String(req.body?.color || "yellow").trim();

    if (!value || !label) {
      res.status(400).json({ error: "El nombre interno y la etiqueta son obligatorios" });
      return;
    }

    await query(
      `
      INSERT INTO meeting_type_options (value, label, color, sort_order)
      VALUES ($1, $2, $3, COALESCE((SELECT MAX(sort_order) FROM meeting_type_options), 0) + 1)
      `,
      [value, label, color]
    );

    await logAuditEvent({
      req,
      action: "settings.meeting_type.create",
      entityType: "meeting_type",
      entityName: label,
      details: { value, color }
    });
    await refreshMeetingTypesCache();
    res.status(201).json({ meetingTypes: getMeetingTypes() });
  })
);

app.patch(
  "/api/settings/meeting-types/:id",
  asyncHandler(async (req, res) => {
    const id = Number(req.params.id);
    const value = String(req.body?.value || "").trim();
    const label = String(req.body?.label || "").trim();
    const color = String(req.body?.color || "yellow").trim();

    if (!Number.isInteger(id) || id <= 0) {
      res.status(400).json({ error: "Tipo de reunión inválido" });
      return;
    }

    if (!value || !label) {
      res.status(400).json({ error: "El nombre interno y la etiqueta son obligatorios" });
      return;
    }

    const result = await query(
      `
      UPDATE meeting_type_options
      SET value = $1, label = $2, color = $3
      WHERE id = $4
      RETURNING id
      `,
      [value, label, color, id]
    );

    if (!result.rows[0]) {
      res.status(404).json({ error: "Tipo de reunión no encontrado" });
      return;
    }

    await logAuditEvent({
      req,
      action: "settings.meeting_type.update",
      entityType: "meeting_type",
      entityId: id,
      entityName: label,
      details: { value, color }
    });
    await refreshMeetingTypesCache();
    res.json({ meetingTypes: getMeetingTypes() });
  })
);

app.delete(
  "/api/settings/meeting-types/:id",
  asyncHandler(async (req, res) => {
    const id = Number(req.params.id);
    if (!Number.isInteger(id) || id <= 0) {
      res.status(400).json({ error: "Tipo de reunión inválido" });
      return;
    }

    const existing = getMeetingTypes().find((type) => Number(type.id) === id);
    if (!existing) {
      res.status(404).json({ error: "Tipo de reunión no encontrado" });
      return;
    }

    const usage = await query("SELECT id FROM meetings WHERE kind = $1 LIMIT 1", [existing.value]);
    if (usage.rows[0]) {
      res.status(409).json({ error: "No se puede eliminar un tipo de reunión que ya está en uso" });
      return;
    }

    await query("DELETE FROM meeting_type_options WHERE id = $1", [id]);
    await logAuditEvent({
      req,
      action: "settings.meeting_type.delete",
      entityType: "meeting_type",
      entityId: id,
      entityName: existing.label
    });
    await refreshMeetingTypesCache();
    res.json({ meetingTypes: getMeetingTypes() });
  })
);

app.post(
  "/api/settings/meeting-reasons",
  asyncHandler(async (req, res) => {
    const name = String(req.body?.name || "").trim();
    if (!name) {
      res.status(400).json({ error: "El motivo de reunión es obligatorio" });
      return;
    }

    await query(
      `
      INSERT INTO meeting_reason_options (name, sort_order)
      VALUES ($1, COALESCE((SELECT MAX(sort_order) FROM meeting_reason_options), 0) + 1)
      `,
      [name]
    );

    await logAuditEvent({
      req,
      action: "settings.meeting_reason.create",
      entityType: "meeting_reason",
      entityName: name
    });
    await refreshMeetingReasonsCache();
    res.status(201).json({ meetingReasons: getMeetingReasons() });
  })
);

app.patch(
  "/api/settings/meeting-reasons/:id",
  asyncHandler(async (req, res) => {
    const id = Number(req.params.id);
    const name = String(req.body?.name || "").trim();
    if (!Number.isInteger(id) || id <= 0) {
      res.status(400).json({ error: "Motivo de reunión inválido" });
      return;
    }
    if (!name) {
      res.status(400).json({ error: "El motivo de reunión es obligatorio" });
      return;
    }

    const result = await query(
      `
      UPDATE meeting_reason_options
      SET name = $1
      WHERE id = $2
      RETURNING id
      `,
      [name, id]
    );

    if (!result.rows[0]) {
      res.status(404).json({ error: "Motivo de reunión no encontrado" });
      return;
    }

    await logAuditEvent({
      req,
      action: "settings.meeting_reason.update",
      entityType: "meeting_reason",
      entityId: id,
      entityName: name
    });
    await refreshMeetingReasonsCache();
    res.json({ meetingReasons: getMeetingReasons() });
  })
);

app.delete(
  "/api/settings/meeting-reasons/:id",
  asyncHandler(async (req, res) => {
    const id = Number(req.params.id);
    if (!Number.isInteger(id) || id <= 0) {
      res.status(400).json({ error: "Motivo de reunión inválido" });
      return;
    }

    const existing = getMeetingReasons().find((reason) => Number(reason.id) === id);
    if (!existing) {
      res.status(404).json({ error: "Motivo de reunión no encontrado" });
      return;
    }

    const usage = await query("SELECT id FROM meetings WHERE subject = $1 LIMIT 1", [existing.name]);
    if (usage.rows[0]) {
      res.status(409).json({ error: "No se puede eliminar un motivo de reunión que ya está en uso" });
      return;
    }

    await query("DELETE FROM meeting_reason_options WHERE id = $1", [id]);
    await logAuditEvent({
      req,
      action: "settings.meeting_reason.delete",
      entityType: "meeting_reason",
      entityId: id,
      entityName: existing.name
    });
    await refreshMeetingReasonsCache();
    res.json({ meetingReasons: getMeetingReasons() });
  })
);

app.get(
  "/api/settings/contact-roles",
  asyncHandler(async (_req, res) => {
    res.json({ contactRoles: getContactRoleOptions() });
  })
);

app.post(
  "/api/settings/contact-roles",
  asyncHandler(async (req, res) => {
    const name = String(req.body?.name || "").trim();
    if (!name) {
      res.status(400).json({ error: "La función del contacto es obligatoria" });
      return;
    }

    await query(
      `
      INSERT INTO contact_role_options (name, sort_order)
      VALUES ($1, COALESCE((SELECT MAX(sort_order) FROM contact_role_options), 0) + 1)
      `,
      [name]
    );

    await logAuditEvent({
      req,
      action: "settings.contact_role.create",
      entityType: "contact_role",
      entityName: name
    });
    await refreshContactRoleOptionsCache();
    res.status(201).json({ contactRoles: getContactRoleOptions() });
  })
);

app.patch(
  "/api/settings/contact-roles/:id",
  asyncHandler(async (req, res) => {
    const id = Number(req.params.id);
    const name = String(req.body?.name || "").trim();
    if (!Number.isInteger(id) || id <= 0) {
      res.status(400).json({ error: "Función de contacto inválida" });
      return;
    }
    if (!name) {
      res.status(400).json({ error: "La función del contacto es obligatoria" });
      return;
    }

    const existingResult = await query("SELECT id, name FROM contact_role_options WHERE id = $1", [id]);
    const existing = existingResult.rows[0];
    if (!existing) {
      res.status(404).json({ error: "Función de contacto no encontrada" });
      return;
    }

    await withTransaction(async (client) => {
      await client.query(
        `
        UPDATE contact_role_options
        SET name = $1
        WHERE id = $2
        `,
        [name, id]
      );

      if (String(existing.name).trim().toLowerCase() !== name.toLowerCase()) {
        await client.query(
          `
          UPDATE meetings
          SET contact_role = $1
          WHERE lower(contact_role) = lower($2)
          `,
          [name, existing.name]
        );
        await client.query(
          `
          UPDATE visit_rules
          SET contact_role = $1, updated_at = NOW()
          WHERE lower(contact_role) = lower($2)
          `,
          [name, existing.name]
        );
      }
    });

    await logAuditEvent({
      req,
      action: "settings.contact_role.update",
      entityType: "contact_role",
      entityId: id,
      entityName: name
    });
    await refreshContactRoleOptionsCache();
    res.json({ contactRoles: getContactRoleOptions() });
  })
);

app.delete(
  "/api/settings/contact-roles/:id",
  asyncHandler(async (req, res) => {
    const id = Number(req.params.id);
    if (!Number.isInteger(id) || id <= 0) {
      res.status(400).json({ error: "Función de contacto inválida" });
      return;
    }

    const existing = getContactRoleOptions().find((role) => Number(role.id) === id);
    if (!existing) {
      res.status(404).json({ error: "Función de contacto no encontrada" });
      return;
    }

    const usage = await query("SELECT id FROM meetings WHERE lower(contact_role) = lower($1) LIMIT 1", [existing.name]);
    if (usage.rows[0]) {
      res.status(409).json({ error: "No se puede eliminar una función del contacto que ya está en uso" });
      return;
    }

    const rulesUsage = await query("SELECT id FROM visit_rules WHERE lower(contact_role) = lower($1) LIMIT 1", [existing.name]);
    if (rulesUsage.rows[0]) {
      res.status(409).json({ error: "No se puede eliminar una función del contacto que ya está asignada a reglas" });
      return;
    }

    await query("DELETE FROM contact_role_options WHERE id = $1", [id]);
    await logAuditEvent({
      req,
      action: "settings.contact_role.delete",
      entityType: "contact_role",
      entityId: id,
      entityName: existing.name
    });
    await refreshContactRoleOptionsCache();
    res.json({ contactRoles: getContactRoleOptions() });
  })
);

app.post(
  "/api/settings/sectors",
  asyncHandler(async (req, res) => {
    const name = String(req.body?.name || "").trim();
    if (!name) {
      res.status(400).json({ error: "El nombre del sector es obligatorio" });
      return;
    }

    try {
      const result = await query(
        `
        INSERT INTO sector_options (name)
        VALUES ($1)
        RETURNING id, name, created_at
        `,
        [name]
      );

      res.status(201).json({
        sector: {
          id: result.rows[0].id,
          name: result.rows[0].name,
          createdAt: result.rows[0].created_at
        }
      });
      await logAuditEvent({
        req,
        action: "settings.sector.create",
        entityType: "sector",
        entityId: result.rows[0].id,
        entityName: result.rows[0].name
      });
    } catch (error) {
      if (error.code === "23505") {
        res.status(409).json({ error: "Ese sector ya existe" });
        return;
      }
      throw error;
    }
  })
);

app.delete(
  "/api/settings/sectors/:id",
  asyncHandler(async (req, res) => {
    const sectorId = Number(req.params.id);
    if (!Number.isInteger(sectorId) || sectorId <= 0) {
      res.status(400).json({ error: "Sector inválido" });
      return;
    }

    const sectorResult = await query("SELECT id, name FROM sector_options WHERE id = $1", [sectorId]);
    const sector = sectorResult.rows[0];
    if (!sector) {
      res.status(404).json({ error: "Sector no encontrado" });
      return;
    }

    const usageResult = await query("SELECT id FROM clients WHERE lower(sector) = lower($1) LIMIT 1", [sector.name]);
    if (usageResult.rows[0]) {
      res.status(409).json({ error: "No se puede eliminar un sector que ya está asignado a compañías" });
      return;
    }

    const branchUsageResult = await query("SELECT id FROM client_branches WHERE lower(sector) = lower($1) LIMIT 1", [sector.name]);
    if (branchUsageResult.rows[0]) {
      res.status(409).json({ error: "No se puede eliminar un sector que ya está asignado a sucursales" });
      return;
    }

    await query("DELETE FROM sector_options WHERE id = $1", [sectorId]);
    await logAuditEvent({
      req,
      action: "settings.sector.delete",
      entityType: "sector",
      entityId: sectorId,
      entityName: sector.name
    });
    res.json({ ok: true });
  })
);

app.get(
  "/api/settings/trash/meetings",
  asyncHandler(async (_req, res) => {
    const { rows } = await query(
      `
      SELECT
        meetings.id,
        meetings.client_id,
        meetings.branch_id,
        meetings.kind,
        meetings.subject,
        meetings.scheduled_for,
        meetings.status,
        meetings.deleted_at,
        meetings.deleted_by,
        clients.name AS client_name,
        client_branches.name AS branch_name
      FROM meetings
      INNER JOIN clients ON clients.id = meetings.client_id
      LEFT JOIN client_branches ON client_branches.id = meetings.branch_id
      WHERE COALESCE(meetings.is_deleted, FALSE) = TRUE
      ORDER BY meetings.deleted_at DESC NULLS LAST, meetings.id DESC
      `
    );

    res.json({
      meetings: rows.map((row) => ({
        id: row.id,
        clientId: row.client_id,
        branchId: row.branch_id,
        clientName: row.client_name,
        branchName: row.branch_name || "",
        kind: row.kind,
        kindLabel: getMeetingLabel(row.kind),
        subject: row.subject,
        scheduledFor: row.scheduled_for,
        status: row.status,
        deletedAt: row.deleted_at,
        deletedBy: row.deleted_by || ""
      }))
    });
  })
);

app.post(
  "/api/settings/trash/meetings/:id/restore",
  asyncHandler(async (req, res) => {
    const meetingId = Number(req.params.id);
    if (!Number.isInteger(meetingId) || meetingId <= 0) {
      res.status(400).json({ error: "Reunión inválida" });
      return;
    }

    const result = await query(
      `
      UPDATE meetings
      SET is_deleted = FALSE, deleted_at = NULL, deleted_by = '', updated_at = NOW()
      WHERE id = $1 AND COALESCE(is_deleted, FALSE) = TRUE
      RETURNING id, subject
      `,
      [meetingId]
    );

    const meeting = result.rows[0];
    if (!meeting) {
      res.status(404).json({ error: "Reunión no encontrada en la papelera" });
      return;
    }

    await logAuditEvent({
      req,
      action: "meeting.restore",
      entityType: "meeting",
      entityId: meeting.id,
      entityName: meeting.subject,
      details: { source: "trash" }
    });

    res.json({ ok: true });
  })
);

app.get(
  "/api/settings/audit-logs",
  asyncHandler(async (req, res) => {
    const limit = Math.min(Math.max(Number(req.query.limit || 100), 1), 300);
    const { rows } = await query(
      `
      SELECT id, user_id, user_name, user_email, action, entity_type, entity_id, entity_name, details, created_at
      FROM audit_logs
      ORDER BY created_at DESC, id DESC
      LIMIT $1
      `,
      [limit]
    );

    res.json({
      logs: rows.map((row) => ({
        id: row.id,
        userId: row.user_id,
        userName: row.user_name,
        userEmail: row.user_email,
        action: row.action,
        entityType: row.entity_type,
        entityId: row.entity_id,
        entityName: row.entity_name,
        details: row.details || {},
        createdAt: row.created_at
      }))
    });
  })
);

app.get(
  "/api/settings/bitrix/config",
  asyncHandler(async (_req, res) => {
    const webhookUrl = await getAppSetting("bitrix_webhook_url", "");
    res.json({
      webhookUrl
    });
  })
);

app.get(
  "/api/settings/features",
  asyncHandler(async (_req, res) => {
    res.json({
      featureFlags: {
        visitRulesEnabled: await isVisitRulesFeatureEnabled()
      }
    });
  })
);

app.patch(
  "/api/settings/features",
  asyncHandler(async (req, res) => {
    const visitRulesEnabled = Boolean(req.body?.visitRulesEnabled);
    await setAppSetting(VISIT_RULES_FEATURE_SETTING_KEY, visitRulesEnabled ? "true" : "false");

    await logAuditEvent({
      req,
      action: "settings.features.update",
      entityType: "settings",
      entityName: "feature_flags",
      details: { visitRulesEnabled }
    });

    res.json({
      featureFlags: {
        visitRulesEnabled
      }
    });
  })
);

app.put(
  "/api/settings/bitrix/config",
  asyncHandler(async (req, res) => {
    const normalizedWebhook = normalizeBitrixWebhookRoot(req.body?.webhookUrl);
    await setAppSetting("bitrix_webhook_url", normalizedWebhook);

    await logAuditEvent({
      req,
      action: "settings.bitrix.config.update",
      entityType: "bitrix",
      entityName: "webhook"
    });

    res.json({
      webhookUrl: normalizedWebhook
    });
  })
);

app.post(
  "/api/settings/bitrix/users-preview",
  asyncHandler(async (req, res) => {
    const bitrixRoot = normalizeBitrixWebhookRoot(req.body?.webhookUrl);
    const users = await fetchBitrixUsers(bitrixRoot);

    await logAuditEvent({
      req,
      action: "settings.bitrix.users_preview",
      entityType: "bitrix",
      entityName: "user.get",
      details: { usersCount: users.length }
    });

    res.json({
      users,
      count: users.length
    });
  })
);

app.post(
  "/api/settings/bitrix/mappings-preview",
  asyncHandler(async (req, res) => {
    const bitrixRoot = normalizeBitrixWebhookRoot(req.body?.webhookUrl);
    const mappings = await buildBitrixMappings(bitrixRoot);

    await logAuditEvent({
      req,
      action: "settings.bitrix.mappings_preview",
      entityType: "bitrix",
      entityName: "mappings",
      details: mappings.counts
    });

    res.json(mappings);
  })
);

app.post(
  "/api/settings/bitrix/tasks-test",
  asyncHandler(async (req, res) => {
    const bitrixRoot = normalizeBitrixWebhookRoot(req.body?.webhookUrl);
    const title = String(req.body?.title || "").trim();
    const description = String(req.body?.description || "").trim();
    const responsibleId = String(req.body?.responsibleId || "").trim();

    if (!title) {
      res.status(400).json({ error: "El título de la tarea es obligatorio" });
      return;
    }
    if (!responsibleId) {
      res.status(400).json({ error: "Debes indicar un responsable de Bitrix" });
      return;
    }

    const createdTask = await createBitrixTask(bitrixRoot, {
      title,
      description,
      responsibleId
    });

    await logAuditEvent({
      req,
      action: "settings.bitrix.task_test_create",
      entityType: "bitrix_task",
      entityId: createdTask.id,
      entityName: title,
      details: { responsibleId }
    });

    res.status(201).json({
      task: createdTask
    });
  })
);

app.get(
  "/api/users",
  requireSettingsAdmin,
  asyncHandler(async (_req, res) => {
    const { rows } = await query(
      `
      SELECT id, name, email, role, bitrix_user_id, created_at
      FROM users
      ORDER BY name ASC, id ASC
      `
    );

    res.json({
      users: rows.map((row) => ({
        ...sanitizeUser(row),
        createdAt: row.created_at
      })),
      roles: USER_ROLE_OPTIONS
    });
  })
);

app.post(
  "/api/users",
  requireSettingsAdmin,
  asyncHandler(async (req, res) => {
    const { errors, values } = validateUserPayload(req.body, { requirePassword: true });
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    try {
      const user = await createUserRecord(values);
      await logAuditEvent({
        req,
        action: "user.create",
        entityType: "user",
        entityId: user.id,
        entityName: user.email,
        details: { role: user.role }
      });
      res.status(201).json({
        user
      });
    } catch (error) {
      if (error.code === "23505") {
        res.status(409).json({ error: "Ya existe un usuario con ese email" });
        return;
      }
      throw error;
    }
  })
);

app.patch(
  "/api/users/:id",
  requireSettingsAdmin,
  asyncHandler(async (req, res) => {
    const userId = Number(req.params.id);
    if (!Number.isInteger(userId) || userId <= 0) {
      res.status(400).json({ error: "Usuario inválido" });
      return;
    }

    const existingUser = await getUserById(userId);
    if (!existingUser) {
      res.status(404).json({ error: "Usuario no encontrado" });
      return;
    }

    const { errors, values } = validateUserPayload(req.body, { requirePassword: false });
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    const assignments = await getUserAssignments(userId);
    const assignmentError = validateRoleChangeAgainstAssignments(assignments, values.role);
    if (assignmentError) {
      res.status(409).json({ error: assignmentError });
      return;
    }

    const updateFields = ["name = $1", "email = $2", "role = $3", "bitrix_user_id = $4"];
    const params = [values.name, values.email, values.role, values.bitrixUserId];

    if (values.password) {
      updateFields.push(`password_hash = $${params.length + 1}`);
      params.push(hashPassword(values.password));
    }

    params.push(userId);

    let result;
    try {
      result = await query(
        `
        UPDATE users
        SET ${updateFields.join(", ")}
        WHERE id = $${params.length}
        RETURNING id, name, email, role, bitrix_user_id, created_at
        `,
        params
      );
    } catch (error) {
      if (error.code === "23505") {
        res.status(409).json({ error: "Ya existe un usuario con ese email" });
        return;
      }
      throw error;
    }

    const user = {
      ...sanitizeUser(result.rows[0]),
      createdAt: result.rows[0].created_at
    };
    await logAuditEvent({
      req,
      action: "user.update",
      entityType: "user",
      entityId: user.id,
      entityName: user.email,
      details: { role: user.role }
    });
    res.json({ user });
  })
);

app.delete(
  "/api/users/:id",
  requireSettingsAdmin,
  asyncHandler(async (req, res) => {
    const userId = Number(req.params.id);
    if (!Number.isInteger(userId) || userId <= 0) {
      res.status(400).json({ error: "Usuario inválido" });
      return;
    }

    if (Number(req.session.userId) === userId) {
      res.status(400).json({ error: "No podés eliminar tu propio usuario mientras tenés la sesión abierta" });
      return;
    }

    const existingUser = await getUserById(userId);
    if (!existingUser) {
      res.status(404).json({ error: "Usuario no encontrado" });
      return;
    }

    const assignments = await getUserAssignments(userId);
    if (assignments.length) {
      res.status(409).json({
        error: `No podés eliminar este usuario porque está asignado a ${assignments[0].clientName}`
      });
      return;
    }

    await query("DELETE FROM users WHERE id = $1", [userId]);
    await logAuditEvent({
      req,
      action: "user.delete",
      entityType: "user",
      entityId: userId,
      entityName: existingUser.email,
      details: { role: existingUser.role }
    });
    res.json({ ok: true });
  })
);

app.get(
  "/api/company-assignment-options",
  asyncHandler(async (_req, res) => {
    res.json({
      allUsers: await getUsersByRoles(USER_ROLE_OPTIONS),
      executives: await getUsersByRoles([EXECUTIVE_ROLE]),
      supervisors: {
        fixedFire: await getUsersByRoles([USER_ROLES.SUPERVISOR_IFCI]),
        extinguishers: await getUsersByRoles([USER_ROLES.SUPERVISOR_EXT]),
        works: await getUsersByRoles([USER_ROLES.SUPERVISOR_WORKS])
      }
    });
  })
);

app.get(
  "/api/pipeline",
  asyncHandler(async (req, res) => {
    const search = String(req.query.search || "").trim();
    const ownerUserId = String(req.query.ownerUserId || "todos").trim();
    const status = String(req.query.status || "todos").trim();

    const opportunities = await fetchOpportunities({
      search,
      ownerUserId: ownerUserId !== "todos" ? Number(ownerUserId) : null,
      status: status !== "todos" ? status : null
    });

    res.json({
      opportunities,
      summary: buildCrmSummary(opportunities)
    });
  })
);

app.get(
  "/api/clients",
  asyncHandler(async (req, res) => {
    const currentUser = await getSessionUser(req);
    if (!currentUser) {
      req.session.destroy(() => {});
      res.status(401).json({ error: "Sesión inválida" });
      return;
    }

    const {
      search = "",
      risk = "todos",
      segment = "todos",
      fixedFire = "todos",
      extinguishers = "todos",
      works = "todos"
    } = req.query;

    const where = [];
    const params = [];

    if (search.trim()) {
      const like = `%${search.trim()}%`;
      params.push(like, like, like);
      where.push(
        `(clients.name ILIKE $${params.length - 2} OR clients.sector ILIKE $${params.length - 1} OR COALESCE(executive_user.name, clients.manager) ILIKE $${params.length})`
      );
    }
    if (risk !== "todos") {
      params.push(risk);
      where.push(`clients.risk = $${params.length}`);
    }
    if (segment !== "todos") {
      params.push(segment);
      where.push(`clients.segment = $${params.length}`);
    }
    if (fixedFire !== "todos") {
      params.push(fixedFire === "si");
      where.push(`clients.service_fixed_fire = $${params.length}`);
    }
    if (extinguishers !== "todos") {
      params.push(extinguishers === "si");
      where.push(`clients.service_extinguishers = $${params.length}`);
    }
    if (works !== "todos") {
      params.push(works === "si");
      where.push(`clients.service_works = $${params.length}`);
    }
    if (isExecutiveScopedUser(currentUser)) {
      params.push(Number(currentUser.id));
      where.push(`clients.executive_user_id = $${params.length}`);
    }

    const { rows } = await query(
      buildClientSelectQuery(mergeWhereClause("clients.is_hidden = FALSE", where.length ? `WHERE ${where.join(" AND ")}` : "")),
      params
    );

    const clients = await Promise.all(rows.map((row) => buildClientResponse(row, { includeCrm: false })));

    const totalsWhere = ["is_hidden = FALSE"];
    const totalsParams = [];
    if (isExecutiveScopedUser(currentUser)) {
      totalsParams.push(Number(currentUser.id));
      totalsWhere.push(`executive_user_id = $${totalsParams.length}`);
    }
    const totalsResult = await query(
      `SELECT COUNT(*)::int AS total_clients, COALESCE(SUM(billing_2025), 0) AS total_billing FROM clients WHERE ${totalsWhere.join(" AND ")}`,
      totalsParams
    );
    const totals = totalsResult.rows[0];

    const meetingSummaryParams = [];
    const meetingSummaryWhere = [];
    if (isExecutiveScopedUser(currentUser)) {
      meetingSummaryParams.push(Number(currentUser.id));
      meetingSummaryWhere.push(`clients.executive_user_id = $${meetingSummaryParams.length}`);
    }
    const meetingSummaryResult = await query(`
      SELECT
        COUNT(*)::int AS total_meetings,
        COALESCE(SUM(CASE WHEN status = 'Agendada' THEN 1 ELSE 0 END), 0)::int AS scheduled_meetings,
        COALESCE(SUM(CASE WHEN status = 'Confirmada' THEN 1 ELSE 0 END), 0)::int AS confirmed_meetings,
        COALESCE(SUM(CASE WHEN status = 'Realizada' THEN 1 ELSE 0 END), 0)::int AS completed_meetings
        ,COUNT(DISTINCT CASE WHEN status = 'Realizada' THEN client_id END)::int AS visited_clients
      FROM meetings
      INNER JOIN clients ON clients.id = meetings.client_id
      WHERE COALESCE(meetings.is_deleted, FALSE) = FALSE
      ${meetingSummaryWhere.length ? `AND ${meetingSummaryWhere.join(" AND ")}` : ""}
    `, meetingSummaryParams);
    const meetingSummary = meetingSummaryResult.rows[0];
    const crmSummary = buildCrmSummary(
      await fetchOpportunities({ executiveScopedUserId: isExecutiveScopedUser(currentUser) ? Number(currentUser.id) : null })
    );

    res.json({
      clients,
      globalKpis: {
        totalClients: totals.total_clients || 0,
        totalBilling: Number(totals.total_billing) || 0,
        totalPotential: Math.round((Number(totals.total_billing) || 0) * 1.12),
        meetingsScheduled: meetingSummary.scheduled_meetings || 0,
        meetingsConfirmed: meetingSummary.confirmed_meetings || 0,
        meetingsCompleted: meetingSummary.completed_meetings || 0,
        meetingsTotal: meetingSummary.total_meetings || 0,
        clientsVisited: meetingSummary.visited_clients || 0,
        openOpportunities: crmSummary.openCount || 0,
        pipelineAmount: crmSummary.totalAmount || 0,
        weightedPipelineAmount: crmSummary.weightedAmount || 0,
        overdueFollowUps: crmSummary.overdueFollowUps || 0,
        visitsPerVisitedClient: meetingSummary.visited_clients
          ? Number(meetingSummary.completed_meetings || 0) / Number(meetingSummary.visited_clients)
          : 0
      }
    });
  })
);

app.get(
  "/api/calendar",
  asyncHandler(async (req, res) => {
    const currentUser = await getSessionUser(req);
    if (!currentUser) {
      req.session.destroy(() => {});
      res.status(401).json({ error: "Sesión inválida" });
      return;
    }

    const month = String(req.query.month || "").trim();
    const participantUserId = String(req.query.participantUserId || "todos").trim();
    const where = ["meetings.status = ANY($1::text[])", "COALESCE(meetings.is_deleted, FALSE) = FALSE"];
    const params = [[MEETING_STATUS.SCHEDULED, MEETING_STATUS.CONFIRMED, MEETING_STATUS.COMPLETED]];

    if (/^\d{4}-\d{2}$/.test(month)) {
      params.push(month);
      where.push(`LEFT(meetings.scheduled_for, 7) = $${params.length}`);
    }

    if (participantUserId !== "todos") {
      params.push(Number(participantUserId));
      where.push(`meetings.participant_user_ids @> ARRAY[$${params.length}]::INTEGER[]`);
    }
    if (isExecutiveScopedUser(currentUser)) {
      params.push(Number(currentUser.id));
      where.push(`clients.executive_user_id = $${params.length}`);
    }

    const { rows } = await query(
      `
      SELECT
        meetings.id,
        meetings.client_id,
        meetings.branch_id,
        meetings.opportunity_id,
        meetings.kind,
        meetings.subject,
        meetings.objective,
        meetings.scheduled_for,
        meetings.participants,
        meetings.participant_user_ids,
        meetings.contact_name,
        meetings.contact_role,
        meetings.modality,
        meetings.next_meeting_date,
        meetings.follow_up_from_meeting_id,
        meetings.minutes,
        meetings.findings,
        meetings.created_by,
        meetings.status,
        meetings.created_at,
        meetings.updated_at,
        clients.name AS client_name,
        client_branches.name AS branch_name
      FROM meetings
      INNER JOIN clients ON clients.id = meetings.client_id
      LEFT JOIN client_branches ON client_branches.id = meetings.branch_id
      WHERE ${where.join(" AND ")}
      ORDER BY meetings.scheduled_for ASC, meetings.id ASC
      `,
      params
    );

    const meetings = rows.map((row) => ({
      ...toMeetingObject(row),
      clientId: row.client_id,
      branchId: row.branch_id,
      clientName: row.client_name,
      branchName: row.branch_name,
      scopeLabel: row.branch_name ? `${row.client_name} · ${row.branch_name}` : row.client_name,
      color: getMeetingColor(row.kind)
    }));

    res.json({ meetings });
  })
);

app.get(
  "/api/visits",
  asyncHandler(async (req, res) => {
    const currentUser = await getSessionUser(req);
    if (!currentUser) {
      req.session.destroy(() => {});
      res.status(401).json({ error: "Sesión inválida" });
      return;
    }

    res.json({
      visits: await fetchVisits({
        ...req.query,
        scopedExecutiveUserId: isExecutiveScopedUser(currentUser) ? Number(currentUser.id) : null
      })
    });
  })
);

app.get(
  "/api/visits/export",
  asyncHandler(async (req, res) => {
    const currentUser = await getSessionUser(req);
    if (!currentUser) {
      req.session.destroy(() => {});
      res.status(401).json({ error: "Sesión inválida" });
      return;
    }

    const visits = await fetchVisits({
      ...req.query,
      scopedExecutiveUserId: isExecutiveScopedUser(currentUser) ? Number(currentUser.id) : null
    });
    const rows = visits.map((visit) => ({
      Fecha: visit.scheduledFor,
      Cliente: visit.clientName,
      Alcance: visit.branchName || "Casa matriz",
      Tipo: visit.kindLabel || visit.kind,
      Modalidad: visit.modality || "",
      Estado: visit.status,
      Contacto: visit.contactName || "",
      Funcion: visit.contactRole || "",
      Participantes: visit.participants || "",
      Ejecutivo: visit.executiveName || "",
      Supervisores: visit.supervisors.join(" / ")
    }));

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Visitas");
    const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="visitas.xlsx"');
    res.send(buffer);
  })
);

app.get(
  "/api/visits/stats",
  asyncHandler(async (req, res) => {
    const currentUser = await getSessionUser(req);
    if (!currentUser) {
      req.session.destroy(() => {});
      res.status(401).json({ error: "Sesión inválida" });
      return;
    }

    const visitRulesEnabled = await isVisitRulesFeatureEnabled();
    const dateFrom = String(req.query.dateFrom || "").trim();
    const dateTo = String(req.query.dateTo || "").trim();
    const requestedExecutiveUserId = normalizeOptionalIdFilter(req.query.executiveUserId);
    const effectiveExecutiveUserId = isExecutiveScopedUser(currentUser) ? Number(currentUser.id) : requestedExecutiveUserId;
    const params = [];
    const dateConditions = ["COALESCE(meetings.is_deleted, FALSE) = FALSE"];

    if (dateFrom) {
      params.push(dateFrom);
      dateConditions.push(`meetings.scheduled_for >= $${params.length}`);
    }

    if (dateTo) {
      params.push(dateTo);
      dateConditions.push(`meetings.scheduled_for <= $${params.length}`);
    }

    if (effectiveExecutiveUserId) {
      params.push(effectiveExecutiveUserId);
      dateConditions.push(`
        EXISTS (
          SELECT 1
          FROM clients AS scope_clients
          WHERE scope_clients.id = meetings.client_id
            AND scope_clients.executive_user_id = $${params.length}
            AND scope_clients.is_hidden = FALSE
        )
      `);
    }

    const meetingsJoinClause = dateConditions.join(" AND ");

    const usersResult = await query(
      `
      SELECT
        users.id,
        users.name,
        users.role,
        COUNT(DISTINCT meetings.id) FILTER (WHERE meetings.status = 'Agendada')::int AS scheduled_count,
        COUNT(DISTINCT meetings.id) FILTER (WHERE meetings.status = 'Confirmada')::int AS confirmed_count,
        COUNT(DISTINCT meetings.id) FILTER (WHERE meetings.status = 'Realizada')::int AS completed_count,
        COUNT(DISTINCT meetings.id)::int AS total_count
      FROM users
      LEFT JOIN meetings ON users.id = ANY(meetings.participant_user_ids) AND ${meetingsJoinClause}
      GROUP BY users.id, users.name, users.role
      HAVING COUNT(DISTINCT meetings.id) > 0
      ORDER BY total_count DESC, users.name ASC
      `,
      params
    );

    const typesResult = await query(
      `
      SELECT
        meeting_type_options.id,
        meeting_type_options.value,
        meeting_type_options.label,
        meeting_type_options.color,
        COUNT(meetings.id) FILTER (WHERE meetings.status = 'Agendada')::int AS scheduled_count,
        COUNT(meetings.id) FILTER (WHERE meetings.status = 'Confirmada')::int AS confirmed_count,
        COUNT(meetings.id) FILTER (WHERE meetings.status = 'Realizada')::int AS completed_count,
        COUNT(meetings.id)::int AS total_count
      FROM meeting_type_options
      LEFT JOIN meetings ON meetings.kind = meeting_type_options.value AND ${meetingsJoinClause}
      GROUP BY meeting_type_options.id, meeting_type_options.value, meeting_type_options.label, meeting_type_options.color, meeting_type_options.sort_order
      HAVING COUNT(meetings.id) > 0
      ORDER BY meeting_type_options.sort_order ASC, meeting_type_options.id ASC
      `,
      params
    );

    let automaticTotal = 0;
    let automaticCompleted = 0;
    let totalRuleSemaphores = 0;
    let whiteCount = 0;
    let greenCount = 0;
    let yellowCount = 0;
    let redCount = 0;

    if (visitRulesEnabled) {
      const automaticRulesResult = await query(
        `
        SELECT
          COUNT(meetings.id)::int AS total_count,
          COUNT(meetings.id) FILTER (WHERE meetings.status = 'Realizada')::int AS completed_count
        FROM meetings
        WHERE ${meetingsJoinClause}
          AND meetings.is_automatic = TRUE
        `,
        params
      );
      const automaticRuleSummary = automaticRulesResult.rows[0] || { total_count: 0, completed_count: 0 };
      automaticTotal = Number(automaticRuleSummary.total_count || 0);
      automaticCompleted = Number(automaticRuleSummary.completed_count || 0);

      const ruleStatusParams = [];
      const ruleStatusConditions = ["rule_context.client_is_hidden = FALSE"];
      if (effectiveExecutiveUserId) {
        ruleStatusParams.push(effectiveExecutiveUserId);
        ruleStatusConditions.push(`rule_context.executive_user_id = $${ruleStatusParams.length}`);
      }

      const automaticRuleSemaphoresResult = await query(
        `
        WITH rule_context AS (
          SELECT
            visit_rules.id,
            CASE
              WHEN visit_rules.entity_type = 'client' THEN client_context.executive_user_id
              WHEN visit_rules.entity_type = 'branch' THEN branch_client_context.executive_user_id
              ELSE NULL
            END AS executive_user_id,
            CASE
              WHEN visit_rules.entity_type = 'client' THEN COALESCE(client_context.is_hidden, TRUE)
              WHEN visit_rules.entity_type = 'branch' THEN COALESCE(branch_client_context.is_hidden, TRUE)
              ELSE TRUE
            END AS client_is_hidden
          FROM visit_rules
          LEFT JOIN clients AS client_context
            ON visit_rules.entity_type = 'client'
           AND client_context.id = visit_rules.entity_id
          LEFT JOIN client_branches AS branch_context
            ON visit_rules.entity_type = 'branch'
           AND branch_context.id = visit_rules.entity_id
          LEFT JOIN clients AS branch_client_context
            ON branch_client_context.id = branch_context.client_id
        ),
        rule_statuses AS (
          SELECT
            rule_context.id,
            CASE
              WHEN pending_meeting.scheduled_for IS NOT NULL
                AND pending_meeting.scheduled_for < CURRENT_DATE - INTERVAL '10 days' THEN '${RULE_SEMAPHORE.RED}'
              WHEN pending_meeting.scheduled_for IS NOT NULL
                AND pending_meeting.scheduled_for < CURRENT_DATE THEN '${RULE_SEMAPHORE.YELLOW}'
              WHEN completed_meeting.completed_meeting_id IS NOT NULL THEN '${RULE_SEMAPHORE.GREEN}'
              ELSE '${RULE_SEMAPHORE.WHITE}'
            END AS semaphore_status
          FROM rule_context
          LEFT JOIN LATERAL (
            SELECT
              meetings.id,
              NULLIF(meetings.scheduled_for, '')::date AS scheduled_for
            FROM meetings
            WHERE meetings.automatic_rule_id = rule_context.id
              AND COALESCE(meetings.is_deleted, FALSE) = FALSE
              AND meetings.status = ANY($${ruleStatusParams.length + 1}::text[])
            ORDER BY NULLIF(meetings.scheduled_for, '')::date ASC NULLS LAST, meetings.id ASC
            LIMIT 1
          ) AS pending_meeting ON TRUE
          LEFT JOIN LATERAL (
            SELECT meetings.id AS completed_meeting_id
            FROM meetings
            WHERE meetings.automatic_rule_id = rule_context.id
              AND COALESCE(meetings.is_deleted, FALSE) = FALSE
              AND meetings.status = '${MEETING_STATUS.COMPLETED}'
            ORDER BY NULLIF(meetings.scheduled_for, '')::date DESC NULLS LAST, meetings.id DESC
            LIMIT 1
          ) AS completed_meeting ON TRUE
          WHERE ${ruleStatusConditions.join(" AND ")}
        )
        SELECT
          COUNT(*)::int AS total_rules,
          COUNT(*) FILTER (WHERE semaphore_status = '${RULE_SEMAPHORE.WHITE}')::int AS white_count,
          COUNT(*) FILTER (WHERE semaphore_status = '${RULE_SEMAPHORE.GREEN}')::int AS green_count,
          COUNT(*) FILTER (WHERE semaphore_status = '${RULE_SEMAPHORE.YELLOW}')::int AS yellow_count,
          COUNT(*) FILTER (WHERE semaphore_status = '${RULE_SEMAPHORE.RED}')::int AS red_count
        FROM rule_statuses
        `,
        [...ruleStatusParams, [MEETING_STATUS.SCHEDULED, MEETING_STATUS.CONFIRMED]]
      );
      const ruleSemaphoreSummary = automaticRuleSemaphoresResult.rows[0] || {
        total_rules: 0,
        white_count: 0,
        green_count: 0,
        yellow_count: 0,
        red_count: 0
      };
      totalRuleSemaphores = Number(ruleSemaphoreSummary.total_rules || 0);
      whiteCount = Number(ruleSemaphoreSummary.white_count || 0);
      greenCount = Number(ruleSemaphoreSummary.green_count || 0);
      yellowCount = Number(ruleSemaphoreSummary.yellow_count || 0);
      redCount = Number(ruleSemaphoreSummary.red_count || 0);
    }

    res.json({
      period: { dateFrom, dateTo },
      byRule: {
        totalCount: automaticTotal,
        completedCount: automaticCompleted,
        completionRate: automaticTotal > 0 ? (automaticCompleted / automaticTotal) * 100 : 0,
        semaphores: {
          totalRules: totalRuleSemaphores,
          whiteCount,
          greenCount,
          yellowCount,
          redCount,
          whiteRate: totalRuleSemaphores > 0 ? (whiteCount / totalRuleSemaphores) * 100 : 0,
          greenRate: totalRuleSemaphores > 0 ? (greenCount / totalRuleSemaphores) * 100 : 0,
          yellowRate: totalRuleSemaphores > 0 ? (yellowCount / totalRuleSemaphores) * 100 : 0,
          redRate: totalRuleSemaphores > 0 ? (redCount / totalRuleSemaphores) * 100 : 0
        }
      },
      byUser: usersResult.rows.map((row) => ({
        id: row.id,
        name: row.name,
        role: row.role,
        scheduledCount: row.scheduled_count,
        confirmedCount: row.confirmed_count,
        completedCount: row.completed_count,
        totalCount: row.total_count
      })),
      byType: typesResult.rows.map((row) => ({
        id: row.id,
        value: row.value,
        label: row.label,
        color: row.color,
        scheduledCount: row.scheduled_count,
        confirmedCount: row.confirmed_count,
        completedCount: row.completed_count,
        totalCount: row.total_count
      })),
      featureFlags: {
        visitRulesEnabled
      }
    });
  })
);

app.patch(
  "/api/meetings/:meetingId/status",
  asyncHandler(async (req, res, next) => {
    const user = await getSessionUser(req);

    if (!user) {
      req.session.destroy(() => {});
      res.status(401).json({ error: "Sesión inválida" });
      return;
    }

    if (isExecutiveScopedUser(user)) {
      const meetingId = Number(req.params.meetingId);
      const accessResult = await query(
        `
        SELECT meetings.id
        FROM meetings
        INNER JOIN clients ON clients.id = meetings.client_id
        WHERE meetings.id = $1
          AND clients.executive_user_id = $2
          AND clients.is_hidden = FALSE
          AND COALESCE(meetings.is_deleted, FALSE) = FALSE
        `,
        [meetingId, user.id]
      );

      if (!accessResult.rows[0]) {
        res.status(403).json({ error: "No tenés permisos para actualizar esta reunión" });
        return;
      }
    }

    req.authUser = sanitizeUser(user);
    next();
  }),
  asyncHandler(async (req, res) => {
    const meetingId = Number(req.params.meetingId);
    const status = String(req.body?.status || "").trim();
    const nextMeetingDate = String(req.body?.nextMeetingDate || "").trim();

    if (!Number.isInteger(meetingId) || meetingId <= 0) {
      res.status(400).json({ error: "Reunión inválida" });
      return;
    }

    if (!VALID_MEETING_STATUS.has(status)) {
      res.status(400).json({ error: "Estado de visita inválido" });
      return;
    }

    const existingResult = await query(
      `
      SELECT id, client_id, branch_id, opportunity_id, kind, subject, objective, scheduled_for, participants, participant_user_ids, contact_name, contact_role,
             modality, next_meeting_date, follow_up_from_meeting_id, minutes, findings,
             active_negotiations_status, opportunities, substitute_recovery, global_contacts, service_health_status, service_status,
             complaint_bitrix_responsible_id, complaint_bitrix_task_id,
             created_by, status, is_automatic, automatic_rule_id, automatic_rule_entity_type, automatic_rule_entity_id, created_at, updated_at
      FROM meetings
      WHERE id = $1
        AND COALESCE(is_deleted, FALSE) = FALSE
      `,
      [meetingId]
    );
    const existingMeeting = existingResult.rows[0];

    if (!existingMeeting) {
      res.status(404).json({ error: "Reunión no encontrada" });
      return;
    }

    const updatedMeeting = await withTransaction(async (client) => {
      const result = await client.query(
        `
        UPDATE meetings
        SET
          status = $1,
          next_meeting_date = $2,
          updated_at = NOW()
        WHERE id = $3
        RETURNING id, client_id, branch_id, opportunity_id, kind, subject, objective, scheduled_for, participants, participant_user_ids, contact_name, contact_role,
                  modality, next_meeting_date, follow_up_from_meeting_id, minutes, findings,
                  active_negotiations_status, opportunities, substitute_recovery, global_contacts, service_health_status, service_status,
                  complaint_bitrix_responsible_id, complaint_bitrix_task_id,
                  created_by, status, is_automatic, automatic_rule_id, automatic_rule_entity_type, automatic_rule_entity_id, created_at, updated_at
        `,
        [status, nextMeetingDate, meetingId]
      );

      const meeting = toMeetingObject(result.rows[0]);
      await upsertMeetingFollowUp(client, meeting, meeting.createdBy || "");
      await regenerateAutomaticMeetingIfNeeded(client, meeting);
      return meeting;
    });

    res.json({ meeting: updatedMeeting });
  })
);

app.get(
  "/api/clients/:id",
  requireClientAccess,
  asyncHandler(async (req, res) => {
    const { rows } = await query(buildClientSelectQuery("WHERE clients.id = $1 AND clients.is_hidden = FALSE", ""), [Number(req.params.id)]);
    const row = rows[0];

    if (!row) {
      res.status(404).json({ error: "Cliente no encontrado" });
      return;
    }

    res.json({ client: await buildClientResponse(row) });
  })
);

app.post(
  "/api/clients",
  asyncHandler(async (req, res) => {
    const currentUser = await getSessionUser(req);
    if (!currentUser) {
      req.session.destroy(() => {});
      res.status(401).json({ error: "Sesión inválida" });
      return;
    }

    const payload = req.body || {};
    const visitRulesEnabled = await isVisitRulesFeatureEnabled();
    const errors = await validateClientPayload(payload);
    const ruleErrors = visitRulesEnabled ? await validateVisitRulesPayload(payload.visitRules) : [];
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }
    if (ruleErrors.length) {
      res.status(400).json({ error: ruleErrors[0] });
      return;
    }

    if (isExecutiveScopedUser(currentUser) && Number(payload.executiveUserId) !== Number(currentUser.id)) {
      res.status(403).json({ error: "Solo podés crear compañías asignadas a tu usuario" });
      return;
    }

    const created = await createClientRecord(payload);
    if (visitRulesEnabled) {
      await replaceVisitRules({ query }, "client", created.id, payload.visitRules || []);
      await syncAutomaticMeetingsForEntity({ query }, "client", created.id);
    }
    const selectedCreated = await query(buildClientSelectQuery("WHERE clients.id = $1", ""), [created.id]);
    const client = await buildClientResponse(selectedCreated.rows[0]);
    await logAuditEvent({
      req,
      action: "client.create",
      entityType: "client",
      entityId: client.id,
      entityName: client.name
    });
    res.status(201).json({ client });
  })
);

app.patch(
  "/api/clients/:id",
  requireClientAccess,
  asyncHandler(async (req, res) => {
    const clientId = Number(req.params.id);
    const existingClientResult = await query("SELECT id, billing_2025 FROM clients WHERE id = $1", [clientId]);

    if (!existingClientResult.rows[0]) {
      res.status(404).json({ error: "Cliente no encontrado" });
      return;
    }

    const payload = req.body || {};
    const visitRulesEnabled = await isVisitRulesFeatureEnabled();
    const errors = await validateClientPayload(payload);
    const ruleErrors = visitRulesEnabled ? await validateVisitRulesPayload(payload.visitRules) : [];
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }
    if (ruleErrors.length) {
      res.status(400).json({ error: ruleErrors[0] });
      return;
    }

    if (isExecutiveScopedUser(req.authUser) && Number(payload.executiveUserId) !== Number(req.authUser.id)) {
      res.status(403).json({ error: "No podés reasignar esta compañía a otro ejecutivo" });
      return;
    }

    const {
      name,
      billing,
      sector,
      companyType,
      country,
      accountStage,
      bitrixLeadId,
      bitrixCompanyId,
      executiveUserId,
      risk,
      segment,
      services,
      supervisors,
      notes
    } = payload;

    const normalizedBilling =
      typeof billing === "number" && !Number.isNaN(billing) ? billing : Number(existingClientResult.rows[0].billing_2025 || 0);
    const executiveResult = await query("SELECT name FROM users WHERE id = $1", [Number(executiveUserId)]);
    const executiveName = executiveResult.rows[0]?.name || "";
    const fixedFireSupervisorId = services.fixedFire ? normalizeNullableUserId(supervisors.fixedFire.userId) : null;
    const extinguishersSupervisorId = services.extinguishers ? normalizeNullableUserId(supervisors.extinguishers.userId) : null;
    const worksSupervisorId = services.works ? normalizeNullableUserId(supervisors.works.userId) : null;

    const updatedResult = await query(
      `
      UPDATE clients
      SET
        name = $1,
        billing_2025 = $2,
        sector = $3,
        company_type = $4,
        country = $5,
        account_stage = $6,
        manager = $7,
        executive_user_id = $8,
        risk = $9,
        segment = $10,
        bitrix_lead_id = $11,
        bitrix_company_id = $12,
        service_fixed_fire = $13,
        service_extinguishers = $14,
        service_works = $15,
        wallet_share = $16,
        nps = $17,
        open_opportunities = $18,
        notes = $19,
        supervisor_ifci_user_id = $20,
        supervisor_ext_user_id = $21,
        supervisor_works_user_id = $22
      WHERE id = $23
      RETURNING *
      `,
      [
        name.trim(),
        normalizedBilling,
        sector.trim(),
        companyType,
        country,
        accountStage,
        executiveName,
        Number(executiveUserId),
        risk,
        segment,
        String(bitrixLeadId || "").trim(),
        String(bitrixCompanyId || "").trim(),
        services.fixedFire,
        services.extinguishers,
        services.works,
        "",
        0,
        0,
        notes.trim(),
        fixedFireSupervisorId,
        extinguishersSupervisorId,
        worksSupervisorId,
        clientId
      ]
    );

    if (visitRulesEnabled) {
      await query(
        `
        DELETE FROM meetings
        WHERE is_automatic = TRUE
          AND automatic_rule_entity_type = 'client'
          AND automatic_rule_entity_id = $1
          AND COALESCE(is_deleted, FALSE) = FALSE
          AND status = ANY($2::text[])
        `,
        [clientId, [MEETING_STATUS.SCHEDULED, MEETING_STATUS.CONFIRMED]]
      );
      await replaceVisitRules({ query }, "client", clientId, payload.visitRules || []);
      await syncAutomaticMeetingsForEntity({ query }, "client", clientId);
    }
    const selectedUpdated = await query(buildClientSelectQuery("WHERE clients.id = $1", ""), [updatedResult.rows[0].id]);
    const client = await buildClientResponse(selectedUpdated.rows[0]);
    await logAuditEvent({
      req,
      action: "client.update",
      entityType: "client",
      entityId: client.id,
      entityName: client.name
    });
    res.json({ client });
  })
);

app.patch(
  "/api/clients/:id/hide",
  requireClientAccess,
  asyncHandler(async (req, res) => {
    const clientId = Number(req.params.id);
    const existingClientResult = await query("SELECT id, name FROM clients WHERE id = $1 AND is_hidden = FALSE", [clientId]);

    if (!existingClientResult.rows[0]) {
      res.status(404).json({ error: "Cliente no encontrado" });
      return;
    }

    await query(
      `
      UPDATE clients
      SET is_hidden = TRUE
      WHERE id = $1
      `,
      [clientId]
    );

    await logAuditEvent({
      req,
      action: "client.hide",
      entityType: "client",
      entityId: clientId,
      entityName: existingClientResult.rows[0].name
    });

    res.json({ ok: true });
  })
);

app.post(
  "/api/clients/:id/branches",
  requireClientAccess,
  asyncHandler(async (req, res) => {
    const clientId = Number(req.params.id);
    const existingClientResult = await query("SELECT id FROM clients WHERE id = $1", [clientId]);

    if (!existingClientResult.rows[0]) {
      res.status(404).json({ error: "Cliente no encontrado" });
      return;
    }

    const payload = req.body || {};
    const visitRulesEnabled = await isVisitRulesFeatureEnabled();
    const errors = await validateBranchPayload(payload);
    const ruleErrors = visitRulesEnabled ? await validateVisitRulesPayload(payload.visitRules) : [];
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }
    if (ruleErrors.length) {
      res.status(400).json({ error: ruleErrors[0] });
      return;
    }

    const { name, sector, risk, segment, bitrixCompanyId, services, supervisors, notes } = payload;
    const parentClientResult = await query("SELECT manager, billing_2025, executive_user_id FROM clients WHERE id = $1", [clientId]);
    const executiveName = parentClientResult.rows[0]?.manager || "";
    const parentBilling = Number(parentClientResult.rows[0]?.billing_2025 || 0);
    const parentExecutiveUserId = parentClientResult.rows[0]?.executive_user_id || null;
    const fixedFireSupervisorId = services.fixedFire ? normalizeNullableUserId(supervisors.fixedFire.userId) : null;
    const extinguishersSupervisorId = services.extinguishers ? normalizeNullableUserId(supervisors.extinguishers.userId) : null;
    const worksSupervisorId = services.works ? normalizeNullableUserId(supervisors.works.userId) : null;

    const result = await query(
      `
      INSERT INTO client_branches (
        client_id, name, billing_2025, sector, manager, executive_user_id, risk, segment,
        bitrix_company_id, service_fixed_fire, service_extinguishers, service_works, notes,
        supervisor_ifci_user_id, supervisor_ext_user_id, supervisor_works_user_id,
        created_at, updated_at
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, NOW(), NOW())
      RETURNING id
      `,
      [
        clientId,
        name.trim(),
        parentBilling,
        sector.trim(),
        executiveName,
        parentExecutiveUserId,
        risk,
        segment,
        String(bitrixCompanyId || "").trim(),
        services.fixedFire,
        services.extinguishers,
        services.works,
        notes.trim(),
        fixedFireSupervisorId,
        extinguishersSupervisorId,
        worksSupervisorId
      ]
    );

    if (visitRulesEnabled) {
      await replaceVisitRules({ query }, "branch", result.rows[0].id, payload.visitRules || []);
      await syncAutomaticMeetingsForEntity({ query }, "branch", result.rows[0].id);
    }
    const selectedBranch = await query(buildBranchSelectQuery("WHERE client_branches.id = $1", ""), [result.rows[0].id]);
    const branch = {
      ...toBranchObject(selectedBranch.rows[0]),
      visitRules: await fetchVisitRules("branch", result.rows[0].id)
    };
    await logAuditEvent({
      req,
      action: "branch.create",
      entityType: "branch",
      entityId: branch.id,
      entityName: branch.name,
      details: { clientId }
    });
    res.status(201).json({ branch });
  })
);

app.patch(
  "/api/clients/:id/branches/:branchId",
  requireClientAccess,
  asyncHandler(async (req, res) => {
    const clientId = Number(req.params.id);
    const branchId = Number(req.params.branchId);
    const existingBranchResult = await query("SELECT id FROM client_branches WHERE id = $1 AND client_id = $2", [branchId, clientId]);

    if (!existingBranchResult.rows[0]) {
      res.status(404).json({ error: "Sucursal no encontrada" });
      return;
    }

    const payload = req.body || {};
    const visitRulesEnabled = await isVisitRulesFeatureEnabled();
    const errors = await validateBranchPayload(payload);
    const ruleErrors = visitRulesEnabled ? await validateVisitRulesPayload(payload.visitRules) : [];
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }
    if (ruleErrors.length) {
      res.status(400).json({ error: ruleErrors[0] });
      return;
    }

    const { name, sector, risk, segment, bitrixCompanyId, services, supervisors, notes } = payload;
    const parentClientResult = await query("SELECT manager, billing_2025, executive_user_id FROM clients WHERE id = $1", [clientId]);
    const executiveName = parentClientResult.rows[0]?.manager || "";
    const parentBilling = Number(parentClientResult.rows[0]?.billing_2025 || 0);
    const parentExecutiveUserId = parentClientResult.rows[0]?.executive_user_id || null;
    const fixedFireSupervisorId = services.fixedFire ? normalizeNullableUserId(supervisors.fixedFire.userId) : null;
    const extinguishersSupervisorId = services.extinguishers ? normalizeNullableUserId(supervisors.extinguishers.userId) : null;
    const worksSupervisorId = services.works ? normalizeNullableUserId(supervisors.works.userId) : null;

    const result = await query(
      `
      UPDATE client_branches
      SET
        name = $1,
        billing_2025 = $2,
        sector = $3,
        manager = $4,
        executive_user_id = $5,
        risk = $6,
        segment = $7,
        bitrix_company_id = $8,
        service_fixed_fire = $9,
        service_extinguishers = $10,
        service_works = $11,
        notes = $12,
        supervisor_ifci_user_id = $13,
        supervisor_ext_user_id = $14,
        supervisor_works_user_id = $15,
        updated_at = NOW()
      WHERE id = $16 AND client_id = $17
      RETURNING id
      `,
      [
        name.trim(),
        parentBilling,
        sector.trim(),
        executiveName,
        parentExecutiveUserId,
        risk,
        segment,
        String(bitrixCompanyId || "").trim(),
        services.fixedFire,
        services.extinguishers,
        services.works,
        notes.trim(),
        fixedFireSupervisorId,
        extinguishersSupervisorId,
        worksSupervisorId,
        branchId,
        clientId
      ]
    );

    if (visitRulesEnabled) {
      await query(
        `
        DELETE FROM meetings
        WHERE is_automatic = TRUE
          AND automatic_rule_entity_type = 'branch'
          AND automatic_rule_entity_id = $1
          AND COALESCE(is_deleted, FALSE) = FALSE
          AND status = ANY($2::text[])
        `,
        [branchId, [MEETING_STATUS.SCHEDULED, MEETING_STATUS.CONFIRMED]]
      );
      await replaceVisitRules({ query }, "branch", branchId, payload.visitRules || []);
      await syncAutomaticMeetingsForEntity({ query }, "branch", branchId);
    }
    const selectedBranch = await query(buildBranchSelectQuery("WHERE client_branches.id = $1", ""), [result.rows[0].id]);
    const branch = {
      ...toBranchObject(selectedBranch.rows[0]),
      visitRules: await fetchVisitRules("branch", branchId)
    };
    await logAuditEvent({
      req,
      action: "branch.update",
      entityType: "branch",
      entityId: branch.id,
      entityName: branch.name,
      details: { clientId }
    });
    res.json({ branch });
  })
);

app.post(
  "/api/clients/:id/opportunities",
  requireClientAccess,
  asyncHandler(async (req, res) => {
    const clientId = Number(req.params.id);
    const existingClientResult = await query("SELECT id FROM clients WHERE id = $1 AND is_hidden = FALSE", [clientId]);

    if (!existingClientResult.rows[0]) {
      res.status(404).json({ error: "Cliente no encontrado" });
      return;
    }

    const payload = req.body || {};
    const errors = await validateOpportunityPayload(payload, clientId);
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    const currentUserResult = await query("SELECT name FROM users WHERE id = $1", [req.session.userId]);
    const currentUser = currentUserResult.rows[0];
    const result = await query(
      `
      INSERT INTO opportunities (
        client_id, branch_id, title, opportunity_type, service_line, status, amount, probability, expected_close_date,
        owner_user_id, source, description, loss_reason, created_by, created_at, updated_at
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, NOW(), NOW())
      RETURNING id
      `,
      [
        clientId,
        payload.branchId ? Number(payload.branchId) : null,
        payload.title.trim(),
        payload.opportunityType,
        payload.serviceLine,
        payload.status,
        Number(payload.amount),
        Number(payload.probability),
        String(payload.expectedCloseDate || "").trim(),
        payload.ownerUserId ? Number(payload.ownerUserId) : null,
        String(payload.source || "").trim(),
        String(payload.description || "").trim(),
        String(payload.lossReason || "").trim(),
        String(currentUser?.name || "").trim()
      ]
    );

    const opportunities = await fetchOpportunities({ clientId });
    const opportunity = opportunities.find((item) => Number(item.id) === Number(result.rows[0].id));
    res.status(201).json({ opportunity });
  })
);

app.patch(
  "/api/clients/:id/opportunities/:opportunityId",
  requireClientAccess,
  asyncHandler(async (req, res) => {
    const clientId = Number(req.params.id);
    const opportunityId = Number(req.params.opportunityId);
    const existingOpportunityResult = await query(
      "SELECT id FROM opportunities WHERE id = $1 AND client_id = $2",
      [opportunityId, clientId]
    );

    if (!existingOpportunityResult.rows[0]) {
      res.status(404).json({ error: "Oportunidad no encontrada" });
      return;
    }

    const payload = req.body || {};
    const errors = await validateOpportunityPayload(payload, clientId);
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    await query(
      `
      UPDATE opportunities
      SET
        branch_id = $1,
        title = $2,
        opportunity_type = $3,
        service_line = $4,
        status = $5,
        amount = $6,
        probability = $7,
        expected_close_date = $8,
        owner_user_id = $9,
        source = $10,
        description = $11,
        loss_reason = $12,
        updated_at = NOW()
      WHERE id = $13 AND client_id = $14
      `,
      [
        payload.branchId ? Number(payload.branchId) : null,
        payload.title.trim(),
        payload.opportunityType,
        payload.serviceLine,
        payload.status,
        Number(payload.amount),
        Number(payload.probability),
        String(payload.expectedCloseDate || "").trim(),
        payload.ownerUserId ? Number(payload.ownerUserId) : null,
        String(payload.source || "").trim(),
        String(payload.description || "").trim(),
        String(payload.lossReason || "").trim(),
        opportunityId,
        clientId
      ]
    );

    const opportunities = await fetchOpportunities({ clientId });
    const opportunity = opportunities.find((item) => Number(item.id) === opportunityId);
    res.json({ opportunity });
  })
);

app.post(
  "/api/opportunities/:opportunityId/follow-ups",
  requireOpportunityAccess,
  asyncHandler(async (req, res) => {
    const opportunityId = Number(req.params.opportunityId);
    const existingOpportunityResult = await query("SELECT id, client_id FROM opportunities WHERE id = $1", [opportunityId]);
    const existingOpportunity = existingOpportunityResult.rows[0];

    if (!existingOpportunity) {
      res.status(404).json({ error: "Oportunidad no encontrada" });
      return;
    }

    const payload = req.body || {};
    const errors = await validateFollowUpPayload(payload);
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    const currentUserResult = await query("SELECT name FROM users WHERE id = $1", [req.session.userId]);
    const currentUser = currentUserResult.rows[0];
    const result = await query(
      `
      INSERT INTO opportunity_follow_ups (
        opportunity_id, type, title, due_date, status, assigned_user_id, notes, completed_at, created_by, created_at, updated_at
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, NOW(), NOW())
      RETURNING id
      `,
      [
        opportunityId,
        payload.type,
        payload.title.trim(),
        payload.dueDate.trim(),
        payload.status,
        payload.assignedUserId ? Number(payload.assignedUserId) : null,
        String(payload.notes || "").trim(),
        payload.status === FOLLOW_UP_STATUS.DONE ? new Date().toISOString() : null,
        String(currentUser?.name || "").trim()
      ]
    );

    const opportunities = await fetchOpportunities({ clientId: existingOpportunity.client_id });
    const opportunity = opportunities.find((item) => Number(item.id) === opportunityId);
    const followUp = opportunity?.followUps?.find((item) => Number(item.id) === Number(result.rows[0].id)) || null;
    res.status(201).json({ followUp, opportunity });
  })
);

app.patch(
  "/api/opportunities/:opportunityId/follow-ups/:followUpId",
  requireOpportunityAccess,
  asyncHandler(async (req, res) => {
    const opportunityId = Number(req.params.opportunityId);
    const followUpId = Number(req.params.followUpId);
    const existingResult = await query(
      `
      SELECT opportunity_follow_ups.id, opportunities.client_id
      FROM opportunity_follow_ups
      INNER JOIN opportunities ON opportunities.id = opportunity_follow_ups.opportunity_id
      WHERE opportunity_follow_ups.id = $1 AND opportunity_follow_ups.opportunity_id = $2
      `,
      [followUpId, opportunityId]
    );
    const existingFollowUp = existingResult.rows[0];

    if (!existingFollowUp) {
      res.status(404).json({ error: "Seguimiento no encontrado" });
      return;
    }

    const payload = req.body || {};
    const errors = await validateFollowUpPayload(payload);
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    await query(
      `
      UPDATE opportunity_follow_ups
      SET
        type = $1,
        title = $2,
        due_date = $3,
        status = $4,
        assigned_user_id = $5,
        notes = $6,
        completed_at = CASE
          WHEN $4 = '${FOLLOW_UP_STATUS.DONE}' THEN COALESCE(completed_at, NOW())
          ELSE NULL
        END,
        updated_at = NOW()
      WHERE id = $7 AND opportunity_id = $8
      `,
      [
        payload.type,
        payload.title.trim(),
        payload.dueDate.trim(),
        payload.status,
        payload.assignedUserId ? Number(payload.assignedUserId) : null,
        String(payload.notes || "").trim(),
        followUpId,
        opportunityId
      ]
    );

    const opportunities = await fetchOpportunities({ clientId: existingFollowUp.client_id });
    const opportunity = opportunities.find((item) => Number(item.id) === opportunityId);
    const followUp = opportunity?.followUps?.find((item) => Number(item.id) === followUpId) || null;
    res.json({ followUp, opportunity });
  })
);

app.post(
  "/api/clients/:id/meetings",
  requireClientAccess,
  asyncHandler(async (req, res) => {
    const clientId = Number(req.params.id);
    const existingClientResult = await query("SELECT id FROM clients WHERE id = $1", [clientId]);
    if (!existingClientResult.rows[0]) {
      res.status(404).json({ error: "Cliente no encontrado" });
      return;
    }

    const payload = req.body || {};
    const errors = await validateMeetingPayload(payload);
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    const branchId = payload.branchId ? Number(payload.branchId) : null;
    if (branchId) {
      const branchResult = await query("SELECT id FROM client_branches WHERE id = $1 AND client_id = $2", [branchId, clientId]);
      if (!branchResult.rows[0]) {
        res.status(400).json({ error: "La sucursal seleccionada no pertenece a esta compañía" });
        return;
      }
    }
    const opportunityId = payload.opportunityId ? Number(payload.opportunityId) : null;
    if (opportunityId) {
      const opportunityResult = await query("SELECT id FROM opportunities WHERE id = $1 AND client_id = $2", [opportunityId, clientId]);
      if (!opportunityResult.rows[0]) {
        res.status(400).json({ error: "La oportunidad seleccionada no pertenece a esta compañía" });
        return;
      }
    }

    const currentUserResult = await query("SELECT name FROM users WHERE id = $1", [req.session.userId]);
    const currentUser = currentUserResult.rows[0];
    const status = normalizeMeetingStatus(payload);
    const nextMeetingDate = String(payload.nextMeetingDate || "").trim();
    const participantUserIds = normalizeParticipantUserIds(payload.participantUserIds);
    const participantNames = await resolveParticipantNames(participantUserIds);
    const complaintTask = await syncComplaintBitrixTask({
      clientId,
      scheduledFor: payload.scheduledFor.trim(),
      serviceStatus: payload.serviceStatus,
      complaintBitrixResponsibleId: payload.complaintBitrixResponsibleId,
      complaintBitrixTaskId: ""
    });

    const createdMeeting = await withTransaction(async (client) => {
      const result = await client.query(
        `
        INSERT INTO meetings (
          client_id, branch_id, opportunity_id, kind, subject, objective, scheduled_for, participants, participant_user_ids,
          contact_name, contact_role, modality, next_meeting_date, minutes, findings,
          active_negotiations_status, opportunities, substitute_recovery, global_contacts, service_health_status, service_status,
          complaint_bitrix_responsible_id, complaint_bitrix_task_id, created_by, status, created_at, updated_at
        ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, $18, $19, $20, $21, $22, $23, $24, $25, NOW(), NOW())
        RETURNING id, client_id, branch_id, opportunity_id, kind, subject, objective, scheduled_for, participants, participant_user_ids, contact_name, contact_role,
                  modality, next_meeting_date, follow_up_from_meeting_id, minutes, findings,
                  active_negotiations_status, opportunities, substitute_recovery, global_contacts, service_health_status, service_status,
                  complaint_bitrix_responsible_id, complaint_bitrix_task_id,
                  created_by, status, is_automatic, automatic_rule_id, automatic_rule_entity_type, automatic_rule_entity_id, created_at, updated_at
        `,
        [
          clientId,
          branchId,
          opportunityId,
          payload.kind,
          payload.subject.trim(),
          payload.objective.trim(),
          payload.scheduledFor.trim(),
          participantNames.join(", "),
          participantUserIds,
          payload.contactName.trim(),
          payload.contactRole.trim(),
          payload.modality,
          nextMeetingDate,
          String(payload.minutes || "").trim(),
          String(payload.findings || "").trim(),
          String(payload.activeNegotiationsStatus || "").trim(),
          String(payload.opportunities || "").trim(),
          String(payload.substituteRecovery || "").trim(),
          String(payload.globalContacts || "").trim(),
          String(payload.serviceHealthStatus || "").trim(),
          String(payload.serviceStatus || "").trim(),
          String(payload.complaintBitrixResponsibleId || "").trim(),
          String(complaintTask.taskId || "").trim(),
          String(currentUser?.name || "").trim(),
          status
        ]
      );

      const meeting = toMeetingObject(result.rows[0]);
      await upsertMeetingFollowUp(client, meeting, currentUser?.name || "");
      await regenerateAutomaticMeetingIfNeeded(client, meeting);
      return meeting;
    });

    await logAuditEvent({
      req,
      action: "meeting.create",
      entityType: "meeting",
      entityId: createdMeeting.id,
      entityName: createdMeeting.subject,
      details: { clientId }
    });
    res.status(201).json({ meeting: createdMeeting });
  })
);

app.patch(
  "/api/clients/:id/meetings/:meetingId",
  requireClientAccess,
  asyncHandler(async (req, res) => {
    const clientId = Number(req.params.id);
    const meetingId = Number(req.params.meetingId);
    const existingResult = await query(
      "SELECT id FROM meetings WHERE id = $1 AND client_id = $2 AND COALESCE(is_deleted, FALSE) = FALSE",
      [meetingId, clientId]
    );

    if (!existingResult.rows[0]) {
      res.status(404).json({ error: "Reunión no encontrada" });
      return;
    }

    const payload = req.body || {};
    const errors = await validateMeetingPayload(payload);
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    const branchId = payload.branchId ? Number(payload.branchId) : null;
    if (branchId) {
      const branchResult = await query("SELECT id FROM client_branches WHERE id = $1 AND client_id = $2", [branchId, clientId]);
      if (!branchResult.rows[0]) {
        res.status(400).json({ error: "La sucursal seleccionada no pertenece a esta compañía" });
        return;
      }
    }
    const opportunityId = payload.opportunityId ? Number(payload.opportunityId) : null;
    if (opportunityId) {
      const opportunityResult = await query("SELECT id FROM opportunities WHERE id = $1 AND client_id = $2", [opportunityId, clientId]);
      if (!opportunityResult.rows[0]) {
        res.status(400).json({ error: "La oportunidad seleccionada no pertenece a esta compañía" });
        return;
      }
    }

    const currentUserResult = await query("SELECT name FROM users WHERE id = $1", [req.session.userId]);
    const currentUser = currentUserResult.rows[0];
    const status = normalizeMeetingStatus(payload);
    const nextMeetingDate = String(payload.nextMeetingDate || "").trim();
    const participantUserIds = normalizeParticipantUserIds(payload.participantUserIds);
    const participantNames = await resolveParticipantNames(participantUserIds);
    const meetingForBitrixResult = await query(
      `
      SELECT complaint_bitrix_task_id
      FROM meetings
      WHERE id = $1 AND client_id = $2
      LIMIT 1
      `,
      [meetingId, clientId]
    );
    const complaintTask = await syncComplaintBitrixTask({
      clientId,
      scheduledFor: payload.scheduledFor.trim(),
      serviceStatus: payload.serviceStatus,
      complaintBitrixResponsibleId: payload.complaintBitrixResponsibleId,
      complaintBitrixTaskId: meetingForBitrixResult.rows[0]?.complaint_bitrix_task_id || ""
    });

    const updatedMeeting = await withTransaction(async (client) => {
      const result = await client.query(
        `
        UPDATE meetings
        SET
          branch_id = $1,
          opportunity_id = $2,
          kind = $3,
          subject = $4,
          objective = $5,
          scheduled_for = $6,
          participants = $7,
          participant_user_ids = $8,
          contact_name = $9,
          contact_role = $10,
          modality = $11,
          next_meeting_date = $12,
          minutes = $13,
          findings = $14,
          active_negotiations_status = $15,
          opportunities = $16,
          substitute_recovery = $17,
          global_contacts = $18,
          service_health_status = $19,
          service_status = $20,
          complaint_bitrix_responsible_id = $21,
          complaint_bitrix_task_id = $22,
          created_by = $23,
          status = $24,
          updated_at = NOW()
        WHERE id = $25 AND client_id = $26
        RETURNING id, client_id, branch_id, opportunity_id, kind, subject, objective, scheduled_for, participants, participant_user_ids, contact_name, contact_role,
                  modality, next_meeting_date, follow_up_from_meeting_id, minutes, findings,
                  active_negotiations_status, opportunities, substitute_recovery, global_contacts, service_health_status, service_status,
                  complaint_bitrix_responsible_id, complaint_bitrix_task_id,
                  created_by, status, created_at, updated_at
        `,
        [
          branchId,
          opportunityId,
          payload.kind,
          payload.subject.trim(),
          payload.objective.trim(),
          payload.scheduledFor.trim(),
          participantNames.join(", "),
          participantUserIds,
          payload.contactName.trim(),
          payload.contactRole.trim(),
          payload.modality,
          nextMeetingDate,
          String(payload.minutes || "").trim(),
          String(payload.findings || "").trim(),
          String(payload.activeNegotiationsStatus || "").trim(),
          String(payload.opportunities || "").trim(),
          String(payload.substituteRecovery || "").trim(),
          String(payload.globalContacts || "").trim(),
          String(payload.serviceHealthStatus || "").trim(),
          String(payload.serviceStatus || "").trim(),
          String(payload.complaintBitrixResponsibleId || "").trim(),
          String(complaintTask.taskId || "").trim(),
          String(currentUser?.name || "").trim(),
          status,
          meetingId,
          clientId
        ]
      );

      const meeting = toMeetingObject(result.rows[0]);
      await upsertMeetingFollowUp(client, meeting, currentUser?.name || "");
      return meeting;
    });

    await logAuditEvent({
      req,
      action: "meeting.update",
      entityType: "meeting",
      entityId: updatedMeeting.id,
      entityName: updatedMeeting.subject,
      details: { clientId }
    });
    res.json({ meeting: updatedMeeting });
  })
);

app.delete(
  "/api/clients/:id/meetings/:meetingId",
  requireClientAccess,
  asyncHandler(async (req, res) => {
    const clientId = Number(req.params.id);
    const meetingId = Number(req.params.meetingId);
    const existingResult = await query(
      "SELECT id, subject FROM meetings WHERE id = $1 AND client_id = $2 AND COALESCE(is_deleted, FALSE) = FALSE",
      [meetingId, clientId]
    );

    const meeting = existingResult.rows[0];
    if (!meeting) {
      res.status(404).json({ error: "Reunión no encontrada" });
      return;
    }

    await query(
      `
      UPDATE meetings
      SET is_deleted = TRUE, deleted_at = NOW(), deleted_by = $3, updated_at = NOW()
      WHERE id = $1 AND client_id = $2
      `,
      [meetingId, clientId, String(req.authUser?.name || "").trim()]
    );
    await logAuditEvent({
      req,
      action: "meeting.delete",
      entityType: "meeting",
      entityId: meeting.id,
      entityName: meeting.subject,
      details: { clientId }
    });
    res.json({ ok: true });
  })
);

app.use("/api", (_req, res) => {
  res.status(404).json({ error: "Ruta API no encontrada" });
});

app.use((_req, res) => {
  res.sendFile(path.join(__dirname, "..", "public", "index.html"));
});

app.use((error, _req, res, _next) => {
  console.error(error);
  res.status(Number(error?.status) || 500).json({
    error: String(error?.message || "Error interno del servidor")
  });
});

async function start() {
  await initDb();
  await refreshMeetingTypesCache();
  await refreshMeetingReasonsCache();
  await refreshContactRoleOptionsCache();
  app.listen(PORT, () => {
    console.log(`Servidor iniciado en http://localhost:${PORT}`);
  });
}

start().catch((error) => {
  console.error("No se pudo iniciar el servidor:", error);
  process.exit(1);
});
